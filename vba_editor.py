"""
vba_editor.py
=============
Programa para alterar os modulos de codigo VBA de um workbook do Excel via Python.

Usa pywin32 (win32com) para controlar o Excel e acessar o VBProject de um arquivo
.xlsm / .xlsb / .xls. Permite:

  - exportar  : fazer backup de todos os modulos para arquivos (.bas/.cls/.frm)
  - importar  : inserir um modulo a partir de um arquivo
  - substituir: trocar todo o codigo de um modulo por um novo conteudo
  - editar    : buscar e substituir texto dentro do codigo de um modulo
  - adicionar : criar um modulo novo (standard, classe ou UserForm vazio)
  - criar-form: criar um UserForm com controles a partir de um spec JSON
  - adicionar-controle : acrescentar um controle a um UserForm que ja existe
    (criar-form so monta form novo; recriar um form perderia layout e .frx)
  - config-controle : alterar propriedades de um controle ja existente
    (legenda, posicao, tamanho) sem remover e recriar
  - remover   : apagar um modulo
  - listar    : mostrar os modulos existentes
  - ler       : imprimir o codigo de um modulo (ou de um procedimento)
  - procurar  : buscar texto em todos os modulos (modulo:linha: trecho)
  - verificar : compilar o projeto VBA e informar se ha erro
  - corrigir-nomes : remover nomes '_FilterDatabase' duplicados (causa do
    dialogo 'Conflito de nome' que trava a abertura via automacao)

Os dialogos modais exibidos durante o Open sao respondidos automaticamente
por uma thread vigia (ver _vigiar_dialogos_conflito):
  - 'Conflito de nome' (nomes definidos duplicados, ex. '_FilterDatabase');
  - o erro "O nome nao pode ser igual a um nome interno, de objeto ou de
    intervalo", que o Excel dispara quando recusa o nome oferecido no
    dialogo acima.
Qualquer outro dialogo que trave a abertura e reportado com titulo, texto e
botoes, em vez de deixar a automacao esperando calada.

REQUISITOS
----------
1. Excel instalado.
2. pywin32 instalado:  pip install pywin32
3. No Excel:  Arquivo > Opcoes > Central de Confiabilidade > Configuracoes da
   Central de Confiabilidade > Configuracoes de Macro >
   [x] "Confiar no acesso ao modelo de objeto do projeto do VBA".
   (Voce ja habilitou isso.)

EXEMPLOS DE USO (linha de comando)
----------------------------------
    python vba_editor.py "086- SGQ Control Factory.xlsm" listar
    python vba_editor.py "086- SGQ Control Factory.xlsm" exportar --pasta backup_vba
    python vba_editor.py arquivo.xlsm importar --arquivo Modulo1.bas
    python vba_editor.py arquivo.xlsm substituir --modulo Module1 --codigo novo.bas
    python vba_editor.py arquivo.xlsm editar --modulo Module1 --de "ValorAntigo" --para "ValorNovo"
    python vba_editor.py arquivo.xlsm adicionar --modulo MeuModulo --tipo std
    python vba_editor.py arquivo.xlsm adicionar --modulo frmVazio --tipo form
    python vba_editor.py arquivo.xlsm criar-form --spec form.json
    python vba_editor.py arquivo.xlsm remover --modulo MeuModulo
    python vba_editor.py arquivo.xlsm ler --modulo Module1 --proc MinhaSub
    python vba_editor.py arquivo.xlsm procurar --texto "tblProgramacao"
    python vba_editor.py arquivo.xlsm verificar
"""

import os
import re
import sys
import json
import time
import ctypes
import shutil
import argparse
import threading
import unicodedata
from datetime import datetime

import win32com.client
import win32con
import win32gui


def _tirar_cabecalho_export(codigo):
    """Remove o cabecalho que o VBE escreve ao EXPORTAR um modulo.

    Um .bas exportado comeca com 'Attribute VB_Name = "..."'; um .cls/.frm
    comeca com o bloco 'VERSION 1.0 CLASS / BEGIN / ... / END' seguido de
    varios 'Attribute VB_*'. Essas linhas sao metadados do ARQUIVO, nao codigo:
    o painel de codigo do VBE nao as exibe. Reinseri-las por AddFromString e
    pedir erro de compilacao ("Attribute VB_Name ... invalido"), ainda mais no
    caso .cls, onde 'VERSION'/'BEGIN'/'END' viram linhas de codigo invalidas.

    Retorna (codigo_limpo, linhas_removidas).
    """
    if not codigo:
        return codigo, []
    linhas = codigo.splitlines()
    removidas = []
    i = 0
    # bloco BEGIN...END das classes/forms (vem logo apos o VERSION)
    while i < len(linhas):
        s = linhas[i].strip()
        if not s:
            i += 1
            continue
        if s.upper().startswith("VERSION "):
            removidas.append(s)
            i += 1
            if i < len(linhas) and linhas[i].strip().upper() == "BEGIN":
                while i < len(linhas):
                    removidas.append(linhas[i].strip())
                    fim = linhas[i].strip().upper() == "END"
                    i += 1
                    if fim:
                        break
            continue
        if s.startswith("Attribute "):
            removidas.append(s)
            i += 1
            continue
        break  # primeira linha que e codigo de verdade: para aqui
    if not removidas:
        return codigo, []
    # nao come linhas em branco significativas depois do cabecalho
    while i < len(linhas) and not linhas[i].strip():
        i += 1
    return "\n".join(linhas[i:]), removidas


class CompilacaoIndisponivel(RuntimeError):
    """O ambiente nao permite acionar o 'Compilar' do VBE.

    Distinta de "o projeto nao compilou": aqui o teste sequer pode ser feito,
    entao tratar como falha de compilacao seria mentira.
    """


def _nome_seguro(nome):
    """Remove caracteres invalidos para nome de arquivo no Windows."""
    return re.sub(r'[<>:"/\\|?*]', "_", nome)

# Tipos de componente VBA (enum vbext_ComponentType)
VBEXT_CT_STD_MODULE = 1    # Modulo padrao (.bas)
VBEXT_CT_CLASS_MODULE = 2  # Modulo de classe (.cls)
VBEXT_CT_MSFORM = 3        # UserForm (.frm)
VBEXT_CT_DOCUMENT = 100    # Modulo de planilha / ThisWorkbook (nao pode ser removido)

TIPO_NOME = {
    VBEXT_CT_STD_MODULE: "Modulo padrao",
    VBEXT_CT_CLASS_MODULE: "Modulo de classe",
    VBEXT_CT_MSFORM: "UserForm",
    VBEXT_CT_DOCUMENT: "Documento (planilha/workbook)",
}

# ProgId dos controles de UserForm (biblioteca MSForms). Aceita nomes
# amigaveis (pt/en) -> o identificador que o Designer.Controls.Add espera.
PROGID_CONTROLE = {
    "label": "Forms.Label.1",
    "rotulo": "Forms.Label.1",
    "textbox": "Forms.TextBox.1",
    "texto": "Forms.TextBox.1",
    "commandbutton": "Forms.CommandButton.1",
    "botao": "Forms.CommandButton.1",
    "combobox": "Forms.ComboBox.1",
    "listbox": "Forms.ListBox.1",
    "lista": "Forms.ListBox.1",
    "checkbox": "Forms.CheckBox.1",
    "caixaselecao": "Forms.CheckBox.1",
    "optionbutton": "Forms.OptionButton.1",
    "opcao": "Forms.OptionButton.1",
    "togglebutton": "Forms.ToggleButton.1",
    "frame": "Forms.Frame.1",
    "quadro": "Forms.Frame.1",
    "image": "Forms.Image.1",
    "imagem": "Forms.Image.1",
    "spinbutton": "Forms.SpinButton.1",
    "scrollbar": "Forms.ScrollBar.1",
    "multipage": "Forms.MultiPage.1",
    "tabstrip": "Forms.TabStrip.1",
}

# Aliases pt -> nome real da propriedade do form/controle. Chaves nao
# listadas aqui sao repassadas como estao (ex: "Font", "BackColor").
ALIAS_PROP = {
    "nome": "Name",
    "caption": "Caption",
    "titulo": "Caption",
    "texto": "Text",
    "valor": "Value",
    "left": "Left",
    "esquerda": "Left",
    "top": "Top",
    "topo": "Top",
    "width": "Width",
    "largura": "Width",
    "height": "Height",
    "altura": "Height",
}

# Palavras reservadas do VBA que NAO podem ser nome de variavel, constante ou
# parametro. Usar uma delas so estoura como "Erro de sintaxe" quando o VBA
# compila o procedimento -- e o 'verificar' nao compila em toda maquina. Caso
# real que motivou a checagem: "Dim imp As ..." ('Imp' e o operador de
# implicacao logica, como 'Eqv').
PALAVRAS_RESERVADAS_VBA = {
    "and", "as", "boolean", "byref", "byte", "byval", "call", "case", "const",
    "currency", "date", "declare", "dim", "do", "double", "each", "else",
    "elseif", "end", "enum", "eqv", "erase", "event", "exit", "false", "for",
    "friend", "function", "get", "global", "gosub", "goto", "if", "imp",
    "implements", "in", "integer", "is", "let", "lib", "like", "long", "loop",
    "lset", "me", "mod", "new", "next", "not", "nothing", "null", "on",
    "option", "optional", "or", "paramarray", "preserve", "private",
    "property", "public", "raiseevent", "redim", "rem", "resume", "return",
    "rset", "select", "set", "single", "static", "step", "stop", "string",
    "sub", "then", "to", "true", "type", "typeof", "until", "variant", "wend",
    "while", "with", "withevents", "xor",
}

_RE_DECL_VAR = re.compile(
    r"^\s*(?:Dim|Static|Private|Public|Global)\s+"
    r"(?!(?:Sub|Function|Property|Declare|Type|Enum|Const|Event)\b)"
    r"(?:WithEvents\s+)?(.+)$", re.IGNORECASE)
_RE_DECL_CONST = re.compile(r"^\s*(?:(?:Private|Public|Global)\s+)?Const\s+(\w+)", re.IGNORECASE)
_RE_DECL_PROC = re.compile(
    r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+\s*\((.*)\)", re.IGNORECASE)


def _separar_virgulas(texto):
    """Divide por virgula FORA de parenteses (Dim a(1, 2), b)."""
    partes, nivel, atual = [], 0, ""
    for ch in texto:
        if ch == "(":
            nivel += 1
        elif ch == ")":
            nivel -= 1
        if ch == "," and nivel == 0:
            partes.append(atual)
            atual = ""
        else:
            atual += ch
    partes.append(atual)
    return partes


def nomes_reservados_usados(codigo):
    """[(numero da linha, nome)] de palavras reservadas usadas como identificador.

    Olha declaracoes (Dim/Static/Private/Public/Global/Const) e parametros de
    Sub/Function/Property. Junta as linhas continuadas com ' _' antes, e ignora
    comentarios e strings.
    """
    achados = []
    linhas = codigo.replace("\r\n", "\n").split("\n")
    i = 0
    while i < len(linhas):
        inicio = i
        linha = linhas[i]
        while linha.rstrip().endswith(" _") and i + 1 < len(linhas):
            i += 1
            linha = linha.rstrip()[:-2] + " " + linhas[i].strip()
        i += 1
        linha = re.sub(r'"[^"]*"', '""', linha)          # strings fora
        linha = re.split(r"(?:^|\s)'|\bRem\b", linha)[0]  # comentario fora
        nomes = []
        m = _RE_DECL_PROC.match(linha)
        if m:
            for p in _separar_virgulas(m.group(1)):
                p = re.sub(r"^\s*(?:Optional\s+)?(?:ByVal\s+|ByRef\s+)?(?:ParamArray\s+)?", "",
                           p, flags=re.IGNORECASE)
                mm = re.match(r"(\w+)", p)
                if mm:
                    nomes.append(mm.group(1))
        else:
            m = _RE_DECL_CONST.match(linha)
            if m:
                nomes.append(m.group(1))
            else:
                m = _RE_DECL_VAR.match(linha)
                if m:
                    for p in _separar_virgulas(m.group(1)):
                        mm = re.match(r"\s*(\w+)", p)
                        if mm:
                            nomes.append(mm.group(1))
        for n in nomes:
            if n.lower() in PALAVRAS_RESERVADAS_VBA:
                achados.append((inicio + 1, n))
    return achados


# Id do botao 'Depurar > Compilar VBAProject' nos menus do VBE
# (independe do idioma do Office).
ID_COMANDO_COMPILAR = 578

# Modos de calculo do Excel (enum XlCalculation)
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105


# ----------------------------------------------------------------------
# Vigia do dialogo 'Conflito de nome' (_FilterDatabase etc.)
# ----------------------------------------------------------------------
# Workbooks com nomes definidos duplicados (classico: '_FilterDatabase')
# fazem o Excel exibir um dialogo modal 'Conflito de nome' / 'Name Conflict'
# durante o Workbooks.Open. DisplayAlerts=False NAO suprime esse dialogo e,
# com o Excel invisivel, a automacao trava para sempre esperando resposta.
# O vigia roda numa thread paralela durante o Open e responde sozinho.

_user32 = ctypes.windll.user32

VK_RETURN = 0x0D

# Prefixo dos nomes gerados pelo vigia ao renomear um nome em conflito.
# NAO pode comecar com '_': o Excel trata nomes iniciados por underscore como
# nomes internos e rejeita a renomeacao com o erro "O nome nao pode ser igual
# a um nome interno, de objeto ou de intervalo", abrindo um SEGUNDO dialogo
# (de erro) que travava a automacao para sempre.
PREFIXO_RENOMEACAO = "NomeDuplicado_"
# Prefixo antigo, mantido so para o corrigir-nomes limpar sobras de versoes
# anteriores da ferramenta.
PREFIXO_RENOMEACAO_LEGADO = "_NomeRenomeado_"

# Trechos (sem acento, minusculos) que identificam o dialogo de ERRO exibido
# quando o nome digitado no dialogo de conflito e recusado pelo Excel.
_ERRO_NOME_RECUSADO = (
    "nome interno",          # "nao pode ser igual a um nome interno..."
    "built-in name",
    "nome nao e valido",
    "not valid",
    "ja existe",
    "already exists",
)


def _sem_acento(texto):
    """Normaliza para comparacao: minusculo e sem acentos."""
    normalizado = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in normalizado
                   if not unicodedata.combining(c)).lower()


def _texto_controle(hwnd, tam=1024):
    """Le o texto de um controle de OUTRO processo.

    win32gui.GetWindowText nao funciona para controles (Edit/Static) de outro
    processo - retorna vazio. WM_GETTEXT e marshalado pelo Windows e funciona.
    """
    buf = ctypes.create_unicode_buffer(tam)
    _user32.SendMessageW(hwnd, win32con.WM_GETTEXT, tam, ctypes.byref(buf))
    return buf.value


def _filhos(hwnd):
    """Lista os hwnds filhos de uma janela."""
    acc = []
    try:
        win32gui.EnumChildWindows(hwnd, lambda h, a: a.append(h), acc)
    except Exception:
        pass  # janela pode ter fechado no meio da enumeracao
    return acc


def _classificar_dialogo(hwnd):
    """Diz que tipo de dialogo e este hwnd: 'conflito', 'erro_nome' ou None."""
    if not win32gui.IsWindowVisible(hwnd):
        return None
    try:
        classe = win32gui.GetClassName(hwnd)
    except Exception:
        return None
    # O dialogo de conflito de nomes do Excel NAO e um dialogo Win32 comum:
    # ele vem na classe propria 'bosa_sdm_XL9' (nao '#32770') e seus controles
    # tambem sao proprios ('EDTBX' no lugar de 'Edit', botoes desenhados sem
    # janela). Procurar so por '#32770'/'Button' nunca o encontra.
    if not (classe.startswith("bosa_sdm") or classe == "#32770"):
        return None
    titulo = _sem_acento(win32gui.GetWindowText(hwnd))
    # cobre 'Nomes em conflito' (pt-BR) e 'Conflito de nome'/'Name Conflict'
    if "conflito" in titulo or "conflict" in titulo:
        return "conflito"
    # O dialogo de erro costuma ter titulo generico ("Microsoft Excel"), entao
    # a identificacao vai pelo texto dos Static.
    for h in _filhos(hwnd):
        try:
            if win32gui.GetClassName(h) != "Static":
                continue
        except Exception:
            continue
        corpo = _sem_acento(_texto_controle(h))
        if any(t in corpo for t in _ERRO_NOME_RECUSADO):
            return "erro_nome"
    return None


def _achar_dialogos():
    """Retorna [(hwnd, tipo)] dos dialogos de conflito/erro de nome visiveis.

    O dialogo de ERRO vem primeiro: ele e modal sobre o de conflito, entao
    precisa ser fechado antes de qualquer tentativa no de conflito.
    """
    achados = []

    def cb(hwnd, acc):
        tipo = _classificar_dialogo(hwnd)
        if tipo:
            acc.append((hwnd, tipo))

    win32gui.EnumWindows(cb, achados)
    achados.sort(key=lambda par: 0 if par[1] == "erro_nome" else 1)
    return achados


def _achar_dialogos_desconhecidos():
    """Dialogos modais visiveis que o vigia NAO sabe responder.

    Servem para diagnostico: se o Open nao volta, e quase sempre porque um
    destes esta na tela esperando resposta com o Excel invisivel.
    """
    achados = []

    def cb(hwnd, acc):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            classe = win32gui.GetClassName(hwnd)
        except Exception:
            return
        if not (classe.startswith("bosa_sdm") or classe == "#32770"):
            return
        if _classificar_dialogo(hwnd):
            return  # e um dos que sabemos tratar
        acc.append(hwnd)

    win32gui.EnumWindows(cb, achados)
    return achados


def _descrever_dialogo(hwnd):
    """Titulo + textos + botoes de um dialogo, para mensagem de diagnostico."""
    titulo = win32gui.GetWindowText(hwnd)
    textos, botoes = [], []
    for h in _filhos(hwnd):
        try:
            cls = win32gui.GetClassName(h)
        except Exception:
            continue
        if cls == "Static":
            t = _texto_controle(h).strip()
            if t:
                textos.append(t)
        elif cls == "Button":
            r = win32gui.GetWindowText(h).replace("&", "").strip()
            if r:
                botoes.append(r)
    return (f"titulo='{titulo}' | texto={' / '.join(textos)[:300]} "
            f"| botoes={botoes}")


def _mapear_controles(hwnd):
    """Devolve (edit, {rotulo_botao: hwnd}) de um dialogo.

    Reconhece tanto controles Win32 padrao ('Edit'/'Button') quanto os
    proprios do Excel ('EDTBX'), usados nos dialogos classe 'bosa_sdm_XL9'.
    """
    edit = None
    botoes = {}
    for h in _filhos(hwnd):
        try:
            cls = win32gui.GetClassName(h)
        except Exception:
            continue
        if cls in ("Edit", "EDTBX") and edit is None:
            edit = h
        elif cls == "Button":
            rotulo = _sem_acento(
                win32gui.GetWindowText(h).replace("&", "").strip())
            botoes[rotulo] = h
    return edit, botoes


def _digitar(hwnd, texto):
    """Digita texto num controle por WM_CHAR (nao exige foco no desktop).

    SendInput/keybd_event nao servem aqui: dependem de a janela ser a
    foreground do desktop interativo, o que nao vale para um Excel invisivel
    rodando por automacao. WM_CHAR postado direto no controle funciona.
    """
    for ch in texto:
        win32gui.PostMessage(hwnd, win32con.WM_CHAR, ord(ch), 0)


def _tecla(hwnd, vk):
    """Envia uma tecla (down+up) para um controle, sem exigir foco."""
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk, 0)


def _clicar(botao):
    """Clica um botao do dialogo.

    PostMessage (assincrono) de proposito: com SendMessage a thread vigia
    ficaria bloqueada ate o Excel terminar de tratar o clique - e se esse
    clique abrir OUTRO modal (o erro de nome recusado), quem teria de fechar
    esse modal e justamente o vigia. Daria deadlock.
    """
    win32gui.PostMessage(botao, win32con.BM_CLICK, 0, 0)


def _fechar_dialogo_erro(hwnd):
    """Fecha o dialogo de erro ('nome nao pode ser igual a um nome interno')."""
    _, botoes = _mapear_controles(hwnd)
    for rotulo in ("ok", "encerrar", "cancelar", "cancel"):
        if rotulo in botoes:
            _clicar(botoes[rotulo])
            return True
    # dialogo proprio do Excel, sem botao-janela: confirma pelo teclado
    _tecla(hwnd, VK_RETURN)
    return True


def _responder_dialogo_conflito(hwnd, contador):
    """Responde um dialogo de conflito de nomes. Retorna True se respondeu.

    Variantes tratadas:
      - caixa de texto (Edit ou EDTBX): digita um nome unico e confirma. No
        dialogo 'bosa_sdm_XL9' do Excel nao ha botao-janela para clicar, e a
        confirmacao sai por ENTER na propria caixa;
      - so botoes Sim/Nao ('usar a versao existente do nome?'): clica Sim.

    Cancelar NAO serve: o Cancel aborta o Workbooks.Open inteiro, que estoura
    0x800A03EC ('O metodo Open da classe Workbooks falhou').
    """
    edit, botoes = _mapear_controles(hwnd)

    if edit is not None:
        # NAO pode comecar com '_': o Excel recusa como nome interno.
        novo_nome = f"{PREFIXO_RENOMEACAO}{contador}"
        _digitar(edit, novo_nome)
        time.sleep(0.25)  # deixa o Excel consumir os WM_CHAR antes do ENTER
        if "ok" in botoes:
            _clicar(botoes["ok"])
        else:
            _tecla(edit, VK_RETURN)
        return True

    for rotulo in ("sim", "yes", "ok"):
        if rotulo in botoes:
            _clicar(botoes[rotulo])
            return True
    return False


def _vigiar_dialogos_conflito(parar, resolvidos, limite=200):
    """Loop da thread vigia: resolve dialogos ate 'parar' ser setado.

    Trata os DOIS dialogos que aparecem na abertura:
      - 'Conflito de nome': digita um nome novo e confirma;
      - erro de nome recusado: fecha e deixa o vigia tentar outro nome.

    resolvidos: lista compartilhada; cada resposta dada vira um item nela.
    limite: teto de respostas, para nao girar para sempre se o Excel recusar
    todo nome oferecido.
    """
    tentativas = 0
    desconhecidos_avisados = set()
    ocioso_desde = time.monotonic()
    while not parar.is_set():
        dialogos = _achar_dialogos()
        if not dialogos:
            # Nada que saibamos responder. Se algo modal esta na tela ha mais
            # de 10s, o Open esta travado nele: avisa uma unica vez, com os
            # dados do dialogo, em vez de esperar calado para sempre.
            if time.monotonic() - ocioso_desde > 10:
                for hwnd in _achar_dialogos_desconhecidos():
                    if hwnd in desconhecidos_avisados:
                        continue
                    desconhecidos_avisados.add(hwnd)
                    print("  [AVISO] dialogo do Excel travando a abertura e "
                          "o vigia nao sabe responde-lo:\n"
                          f"          {_descrever_dialogo(hwnd)}\n"
                          "          Rode com --visivel para responder na mao.")
            parar.wait(0.5)
            continue
        ocioso_desde = time.monotonic()

        hwnd, tipo = dialogos[0]
        tentativas += 1
        if tentativas > limite:
            print(f"  [conflito] limite de {limite} dialogos atingido; "
                  "parei de responder. Rode 'corrigir-nomes'.")
            return

        if tipo == "erro_nome":
            if _fechar_dialogo_erro(hwnd):
                print("  [conflito] dialogo de erro de nome fechado; "
                      "tentando outro nome.")
        else:
            n = len(resolvidos) + 1
            if _responder_dialogo_conflito(hwnd, n):
                resolvidos.append(hwnd)
                print(f"  [conflito] dialogo 'Conflito de nome' respondido "
                      f"automaticamente ({n})")
        # da tempo do dialogo fechar antes de procurar o proximo
        parar.wait(0.3)
        ocioso_desde = time.monotonic()


class VBAEditor:
    """Abre um workbook e expoe operacoes sobre seus modulos VBA.

    Use como context manager para garantir que o Excel seja fechado:

        with VBAEditor("arquivo.xlsm") as ed:
            ed.listar()
    """

    def __init__(self, caminho_workbook, visivel=False,
                 auto_backup=True, pasta_backup="backups",
                 somente_leitura=False):
        self.caminho = os.path.abspath(caminho_workbook)
        if not os.path.exists(self.caminho):
            raise FileNotFoundError(f"Workbook nao encontrado: {self.caminho}")
        self.visivel = visivel
        self.auto_backup = auto_backup
        self.pasta_backup = pasta_backup
        self.somente_leitura = somente_leitura
        self._fez_backup = False
        self.excel = None
        self.wb = None
        self._wb_descartavel = None

    # ---- ciclo de vida -------------------------------------------------
    def __enter__(self):
        self.abrir()
        return self

    def __exit__(self, exc_type, exc, tb):
        # Salva so se nao houve excecao
        self.fechar(salvar=exc_type is None)

    def abrir(self):
        # Em somente-leitura o lock nao importa: nada sera salvo, entao o
        # workbook pode estar aberto no Excel do usuario sem problema.
        if not self.somente_leitura:
            self._verificar_lock()
        # DispatchEx: SEMPRE cria uma instancia nova e isolada do Excel.
        # Dispatch reaproveitaria uma instancia ja aberta pelo usuario, e o
        # Quit() do fechar() derrubaria a sessao dele (podendo descartar
        # trabalho nao salvo).
        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.Visible = self.visivel
        self.excel.DisplayAlerts = False
        self.excel.AskToUpdateLinks = False
        # Macros NAO devem rodar durante a edicao de modulos: um
        # Workbook_Open que mostra forms/muda Visible/da erro derruba a
        # automacao (alem de poder alterar dados). ForceDisable = 3.
        try:
            self.excel.AutomationSecurity = 3
        except Exception:
            pass  # versoes antigas sem a propriedade: segue com o padrao
        self.excel.EnableEvents = False

        # Recalculo desligado ANTES do Open: com o modo automatico, abrir um
        # workbook grande recalcula a pasta inteira durante o proprio Open e
        # ele leva minutos (ou parece travado). O modo original e restaurado
        # antes de salvar (ver fechar()).
        self._calc_original = self._desligar_recalculo()

        # O Open pode travar num dialogo modal 'Conflito de nome' (nomes
        # duplicados tipo '_FilterDatabase'). O vigia responde por nos.
        parar = threading.Event()
        resolvidos = []
        vigia = threading.Thread(
            target=_vigiar_dialogos_conflito, args=(parar, resolvidos),
            daemon=True)
        vigia.start()
        try:
            self.wb = self._abrir_workbook()
        finally:
            parar.set()
            vigia.join(timeout=2)
            # workbook alvo aberto (ou falhou): o vazio ja cumpriu o papel
            self._fechar_descartavel()
        if resolvidos:
            print(f"  [conflito] {len(resolvidos)} dialogo(s) de conflito de "
                  "nome respondido(s) na abertura. Rode o comando "
                  "'corrigir-nomes' para eliminar a causa em definitivo.")
        self._verificar_acesso_vbproject()
        # Reforca o modo manual: se o Calculation nao pode ser ajustado antes
        # (Excel sem workbook aberto em versoes antigas), agora da.
        if self._calc_original is None:
            self._calc_original = self._desligar_recalculo()
        return self

    def _desligar_recalculo(self):
        """Poe o Excel em calculo manual. Retorna o modo anterior, ou None.

        Application.Calculation so aceita leitura/escrita com algum workbook
        aberto - por isso, quando ainda nao ha nenhum, abrimos um workbook
        vazio descartavel so para poder mudar o modo, e o fechamos em seguida.
        Recalculo automatico num arquivo grande faz o Excel gastar minutos no
        Open e ainda rejeitar chamadas COM por estar 'ocupado'
        (RPC_E_CALL_REJECTED).
        """
        try:
            if self._com_retry(lambda: self.excel.Workbooks.Count) == 0:
                # NAO fechar aqui: sem nenhum workbook aberto o Excel volta o
                # Calculation para automatico. O descartavel fica de pe ate o
                # workbook alvo abrir (ver _fechar_descartavel()).
                self._wb_descartavel = self._com_retry(
                    lambda: self.excel.Workbooks.Add())
            anterior = self._com_retry(lambda: self.excel.Calculation)
            self._com_retry(lambda: setattr(
                self.excel, "Calculation", XL_CALCULATION_MANUAL))
            return anterior
        except Exception as e:
            print(f"  [aviso] nao consegui desligar o recalculo: {e}")
            return None

    def _fechar_descartavel(self):
        """Fecha o workbook vazio usado para ajustar o Calculation."""
        wb = getattr(self, "_wb_descartavel", None)
        if wb is None:
            return
        self._wb_descartavel = None
        try:
            self._com_retry(lambda: wb.Close(SaveChanges=False))
        except Exception:
            pass  # workbook vazio sobrando nao atrapalha a sessao

    def _abrir_workbook(self):
        """Abre o workbook, em somente-leitura quando a sessao nao vai salvar.

        Alguns workbooks recusam a abertura para escrita via automacao (o Open
        estoura 0x800A03EC) mas abrem sem problema em somente-leitura. Como os
        comandos de leitura (listar/ler/exportar/procurar/verificar) nunca
        salvam, abri-los como ReadOnly evita esse erro por completo - e ainda
        dispensa fechar o arquivo antes de rodar.
        """
        if self.somente_leitura:
            return self.excel.Workbooks.Open(
                self.caminho, UpdateLinks=0, ReadOnly=True,
                IgnoreReadOnlyRecommended=True)
        try:
            return self.excel.Workbooks.Open(self.caminho, UpdateLinks=0)
        except Exception as e:
            raise RuntimeError(
                f"Nao consegui abrir o workbook para escrita: {e}\n"
                "  - se o arquivo estiver aberto em outro Excel, feche-o;\n"
                "  - se for so leitura que voce precisa, use um comando de "
                "leitura (listar/ler/exportar/procurar), que abre o arquivo "
                "em modo somente-leitura e nao esbarra nesse erro."
            ) from e

    def _verificar_lock(self):
        """Aborta se o workbook ja estiver aberto no Excel.

        O Excel cria um arquivo de lock '~$<nome>' ao lado do workbook
        enquanto ele esta aberto. Abrir por cima (mesmo em outra instancia)
        seria somente-leitura e o salvamento falharia no final.
        """
        lock = os.path.join(os.path.dirname(self.caminho),
                            "~$" + os.path.basename(self.caminho))
        if os.path.exists(lock):
            raise RuntimeError(
                "O workbook parece estar aberto no Excel (existe o lock "
                f"'{os.path.basename(lock)}'). Feche o Excel e rode novamente. "
                "Se o Excel nao estiver aberto, o lock e sobra de um "
                "encerramento anormal: apague o arquivo e tente de novo."
            )

    # RPC_E_CALL_REJECTED: o Excel rejeita chamadas COM enquanto esta ocupado
    # (recalculando, rodando Workbook_Open, etc.). E transitorio: basta
    # esperar e tentar de novo.
    _RPC_E_CALL_REJECTED = -2147418111

    def _com_retry(self, fn, tentativas=30, espera=2.0):
        """Executa fn() repetindo enquanto o Excel rejeitar a chamada.

        A rejeicao aparece como com_error RPC_E_CALL_REJECTED ou, quando
        acontece na resolucao do nome da propriedade, como AttributeError.
        """
        for i in range(tentativas):
            try:
                return fn()
            except Exception as e:
                transitorio = (
                    getattr(e, "hresult", None) == self._RPC_E_CALL_REJECTED
                    or isinstance(e, AttributeError))
                if not transitorio or i == tentativas - 1:
                    raise
                if i == 0:
                    print("  [aguardando] Excel ocupado; tentando de novo...")
                time.sleep(espera)

    def _verificar_acesso_vbproject(self, tentativas=30, espera=2.0):
        ultimo_erro = None
        for i in range(tentativas):
            try:
                _ = self.wb.VBProject.VBComponents.Count
                return
            except Exception as e:
                ultimo_erro = e
                # Excel ocupado aparece como com_error RPC_E_CALL_REJECTED ou,
                # quando a rejeicao acontece na resolucao do nome da
                # propriedade, como AttributeError. Ambos sao transitorios.
                transitorio = (
                    getattr(e, "hresult", None) == self._RPC_E_CALL_REJECTED
                    or isinstance(e, AttributeError))
                if transitorio:
                    if i == 0:
                        print("  [aguardando] Excel ocupado (recalculo/macros); "
                              "tentando de novo...")
                    time.sleep(espera)
                    continue
                break  # outro erro: e falta de permissao mesmo, nao adianta insistir
        raise PermissionError(
            "Nao foi possivel acessar o VBProject. Habilite no Excel: "
            "Central de Confiabilidade > Configuracoes de Macro > "
            "'Confiar no acesso ao modelo de objeto do projeto do VBA'.\n"
            f"Detalhe: {ultimo_erro}"
        )

    def fechar(self, salvar=True):
        if self.somente_leitura:
            salvar = False  # sessao de leitura nunca grava
        self._fechar_descartavel()  # caso o Open tenha falhado antes da hora
        try:
            if self.wb is not None:
                if salvar:
                    # restaura o modo de calculo original para nao gravar o
                    # workbook em modo manual (ver abrir())
                    calc = getattr(self, "_calc_original", None)
                    if calc is not None and calc != XL_CALCULATION_MANUAL:
                        try:
                            self._com_retry(lambda: setattr(
                                self.excel, "Calculation", calc))
                        except Exception as e:
                            print(f"  [aviso] nao restaurei o recalculo: {e}")
                    self._com_retry(lambda: self.wb.Save())
                self._com_retry(lambda: self.wb.Close(SaveChanges=False))
        finally:
            if self.excel is not None:
                try:
                    self._com_retry(lambda: self.excel.Quit())
                except Exception as e:
                    print(f"  [aviso] Excel nao encerrou de forma limpa: {e}. "
                          "Se ficar um processo EXCEL orfao, finalize-o.")
            self.wb = None
            self.excel = None

    # ---- backup --------------------------------------------------------
    def _garantir_backup(self):
        """Cria uma copia datada do workbook antes da PRIMEIRA alteracao.

        Copia o arquivo do disco (estado anterior as alteracoes desta sessao).
        Roda apenas uma vez por sessao e so para operacoes que modificam.
        """
        if not self.auto_backup or self._fez_backup:
            return
        pasta = self.pasta_backup
        if not os.path.isabs(pasta):
            pasta = os.path.join(os.path.dirname(self.caminho), pasta)
        os.makedirs(pasta, exist_ok=True)
        nome = os.path.basename(self.caminho)
        base, ext = os.path.splitext(nome)
        carimbo = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        destino = os.path.join(pasta, f"{base}_backup_{carimbo}{ext}")
        shutil.copy2(self.caminho, destino)
        self._fez_backup = True
        print(f"  [backup] copia criada: {destino}")

    # ---- helpers internos ----------------------------------------------
    @property
    def _componentes(self):
        return self.wb.VBProject.VBComponents

    def _achar(self, nome_modulo):
        """Retorna o VBComponent pelo nome, ou levanta erro."""
        for comp in self._componentes:
            if comp.Name.lower() == nome_modulo.lower():
                return comp
        raise KeyError(f"Modulo '{nome_modulo}' nao encontrado.")

    # ---- operacoes -----------------------------------------------------
    def _aba_do_modulo(self, nome_modulo):
        """Nome da aba correspondente a um modulo de documento (ou '').

        O modulo de planilha aparece no VBE pelo CodeName ('Plan13'), que nao
        diz nada sobre qual aba ele controla. Sem esse mapeamento e preciso
        adivinhar pela leitura do codigo - erro facil de cometer e caro de
        descobrir depois.
        """
        try:
            for ws in self._com_retry(lambda: self.wb.Worksheets):
                if self._com_retry(lambda: ws.CodeName) == nome_modulo:
                    return self._com_retry(lambda: ws.Name)
        except Exception:
            pass  # workbook sem planilhas acessiveis: segue sem o mapeamento
        return ""

    def listar(self):
        """Lista os modulos do projeto VBA."""
        comps = self._componentes
        print(f"\nProjeto VBA de: {os.path.basename(self.caminho)}")
        print(f"Total de modulos: {comps.Count}\n")
        print(f"{'NOME':<24} {'ABA':<18} {'TIPO':<28} {'LINHAS':>7}")
        print("-" * 81)
        infos = []
        for comp in comps:
            tipo_num = self._com_retry(lambda: comp.Type)
            tipo = TIPO_NOME.get(tipo_num, f"Tipo {tipo_num}")
            linhas = self._com_retry(lambda: comp.CodeModule.CountOfLines)
            nome = self._com_retry(lambda: comp.Name)
            aba = self._aba_do_modulo(nome) if tipo_num == VBEXT_CT_DOCUMENT else ""
            print(f"{nome:<24} {aba:<18} {tipo:<28} {linhas:>7}")
            infos.append((nome, aba, tipo, linhas))
        print()
        return infos

    def exportar(self, pasta_destino):
        """Exporta todos os modulos com codigo para arquivos (backup).

        Modulos que falharem (ex: nome com caractere problematico) sao pulados
        e reportados ao final, sem interromper o backup dos demais.
        """
        # Caminho ABSOLUTO: o Export do COM resolve caminho relativo contra o
        # diretorio de trabalho do Excel (ex: Documentos), nao o do Python --
        # com caminho relativo todos os modulos falhariam.
        pasta_destino = os.path.abspath(pasta_destino)
        os.makedirs(pasta_destino, exist_ok=True)
        exportados = []
        falhas = []
        for comp in self._componentes:
            if comp.CodeModule.CountOfLines == 0 and comp.Type == VBEXT_CT_DOCUMENT:
                continue  # documento vazio: nada para salvar
            ext = {
                VBEXT_CT_STD_MODULE: ".bas",
                VBEXT_CT_CLASS_MODULE: ".cls",
                VBEXT_CT_MSFORM: ".frm",
                VBEXT_CT_DOCUMENT: ".cls",
            }.get(comp.Type, ".txt")
            destino = os.path.join(pasta_destino, _nome_seguro(comp.Name) + ext)
            try:
                comp.Export(destino)
                exportados.append(destino)
            except Exception as e:
                falhas.append((comp.Name, str(e)))
                print(f"  [FALHOU] {comp.Name}: {e}")
        print(f"\n{len(exportados)} modulo(s) exportado(s) para '{pasta_destino}'.")
        if falhas:
            print(f"{len(falhas)} modulo(s) NAO exportado(s): "
                  + ", ".join(n for n, _ in falhas))
        return exportados

    def importar(self, caminho_arquivo):
        """Importa um modulo a partir de um arquivo .bas/.cls/.frm."""
        caminho_arquivo = os.path.abspath(caminho_arquivo)
        if not os.path.exists(caminho_arquivo):
            raise FileNotFoundError(caminho_arquivo)
        self._garantir_backup()
        comp = self._componentes.Import(caminho_arquivo)
        print(f"Modulo '{comp.Name}' importado de '{caminho_arquivo}'.")
        return comp.Name

    def adicionar(self, nome_modulo, tipo="std"):
        """Cria um modulo novo. tipo: 'std', 'classe' ou 'form' (UserForm vazio).

        Para criar um UserForm ja com controles, use criar_form().
        """
        mapa = {
            "std": VBEXT_CT_STD_MODULE,
            "classe": VBEXT_CT_CLASS_MODULE,
            "form": VBEXT_CT_MSFORM,
        }
        if tipo not in mapa:
            raise ValueError("tipo deve ser 'std', 'classe' ou 'form'.")
        self._garantir_backup()
        comp = self._componentes.Add(mapa[tipo])
        comp.Name = nome_modulo
        print(f"Modulo '{nome_modulo}' ({tipo}) criado.")
        return comp.Name

    @staticmethod
    def _aplicar_props(obj, props):
        """Aplica um dict de propriedades a um controle do form.

        Traduz aliases em portugues (ver ALIAS_PROP) e ignora, com aviso,
        qualquer propriedade que o objeto nao aceite -- assim um nome de
        propriedade errado no spec nao aborta a criacao do form inteiro.
        """
        for chave, valor in props.items():
            nome = ALIAS_PROP.get(str(chave).lower(), chave)
            try:
                # "Font.Size": propriedade de um sub-objeto. Sem isto, mexer na
                # fonte de um controle exigia um driver COM a parte.
                alvo = obj
                partes = str(nome).split(".")
                for p in partes[:-1]:
                    alvo = getattr(alvo, p)
                setattr(alvo, partes[-1], valor)
            except Exception as e:
                print(f"  [aviso] propriedade '{nome}'={valor!r} ignorada: {e}")

    # Aparencia que adicionar_controle(modelo=...) copia de um controle
    # existente. Tamanho e posicao NAO entram: quem posiciona e o chamador.
    PROPS_APARENCIA = ("ForeColor", "BackColor", "BackStyle", "BorderStyle",
                       "BorderColor", "SpecialEffect", "TextAlign", "WordWrap")
    PROPS_FONTE = ("Name", "Size", "Bold", "Italic", "Underline")

    @classmethod
    def _copiar_aparencia(cls, modelo, ctrl):
        copiadas = []
        for p in cls.PROPS_APARENCIA:
            try:
                setattr(ctrl, p, getattr(modelo, p))
                copiadas.append(p)
            except Exception:
                pass  # nem todo tipo de controle tem toda propriedade
        for p in cls.PROPS_FONTE:
            try:
                setattr(ctrl.Font, p, getattr(modelo.Font, p))
                copiadas.append("Font." + p)
            except Exception:
                pass
        return copiadas

    @staticmethod
    def _aplicar_props_form(comp, designer, props):
        """Aplica propriedades ao proprio UserForm.

        As propriedades de projeto do form (Caption, Width, Height, ...) vivem
        na colecao VBComponent.Properties -- setar via Designer nao persiste
        depois de salvar. Tenta Properties primeiro e cai para o Designer no
        que nao existir la.
        """
        for chave, valor in props.items():
            nome = ALIAS_PROP.get(str(chave).lower(), chave)
            try:
                comp.Properties(nome).Value = valor
                continue
            except Exception:
                pass
            try:
                setattr(designer, nome, valor)
            except Exception as e:
                print(f"  [aviso] propriedade '{nome}'={valor!r} ignorada: {e}")

    def criar_form(self, spec):
        """Cria um UserForm completo a partir de uma especificacao (dict).

        Formato do spec:
            {
              "nome":     "frmExemplo",      # nome do UserForm (opcional)
              "caption":  "Titulo",          # legenda da janela (opcional)
              "largura":  300, "altura": 200,# tamanho do form (opcional)
              "propriedades": { ... },       # props extras do form (opcional)
              "controles": [                 # lista de controles (opcional)
                 {"tipo": "label",  "nome": "lbl1", "caption": "Nome:",
                  "left": 12, "top": 12, "width": 60, "height": 18},
                 {"tipo": "textbox","nome": "txtNome",
                  "left": 80, "top": 10, "width": 180, "height": 20},
                 {"tipo": "botao",  "nome": "btnOK", "caption": "OK",
                  "left": 80, "top": 50, "width": 80, "height": 24}
              ],
              "codigo": "Private Sub btnOK_Click()\\n...\\nEnd Sub"  # opcional
            }

        Retorna o nome final do UserForm criado.
        """
        self._garantir_backup()
        comp = self._componentes.Add(VBEXT_CT_MSFORM)
        if spec.get("nome"):
            comp.Name = spec["nome"]
        designer = comp.Designer

        # Propriedades do proprio form (caption/tamanho + extras).
        props_form = {}
        for chave in ("caption", "titulo", "largura", "altura", "width", "height"):
            if chave in spec:
                props_form[chave] = spec[chave]
        props_form.update(spec.get("propriedades", {}))
        self._aplicar_props_form(comp, designer, props_form)

        # Controles.
        controles = spec.get("controles", [])
        for c in controles:
            tipo = str(c.get("tipo", "")).lower()
            progid = PROGID_CONTROLE.get(tipo)
            if progid is None:
                raise ValueError(
                    f"Tipo de controle desconhecido: '{c.get('tipo')}'. "
                    f"Validos: {', '.join(sorted(set(PROGID_CONTROLE)))}."
                )
            ctrl = designer.Controls.Add(progid)
            resto = {k: v for k, v in c.items() if k != "tipo"}
            self._aplicar_props(ctrl, resto)

        # Codigo VBA do modulo do form (ex: handlers _Click).
        if spec.get("codigo"):
            comp.CodeModule.AddFromString(spec["codigo"])

        print(f"UserForm '{comp.Name}' criado com {len(controles)} controle(s).")
        return comp.Name

    def _controle_do_form(self, designer, nome, nome_form):
        for c in designer.Controls:
            if c.Name.lower() == str(nome).lower():
                return c
        existentes = ", ".join(c.Name for c in designer.Controls)
        raise KeyError(f"O form '{nome_form}' nao tem controle '{nome}'. Tem: {existentes}.")

    def adicionar_controle(self, nome_form, tipo, nome_controle, props=None,
                           dentro_de=None, modelo=None):
        """Acrescenta um controle a um UserForm que JA EXISTE.

        criar_form() so monta form novo. Para mexer num form ja pronto ela
        nao serve: recriar o form perderia o layout e o .frx atuais. Aqui o
        alvo e localizado por _achar() e o controle entra no Designer dele,
        preservando tudo o que ja estava la.

        dentro_de: nome de um Frame (ou outro container) do form. O controle
            entra NELE, e Left/Top passam a ser relativos a ele -- sem isto, um
            campo novo numa tela montada dentro de um Frame caia no form, por
            baixo do Frame, invisivel.
        modelo: nome de um controle existente cuja APARENCIA (fonte, cores,
            borda, efeito) e copiada antes das props -- para o campo novo
            nascer igual aos vizinhos. Tamanho e posicao nao sao copiados.

        Retorna o nome do controle criado.
        """
        comp = self._achar(nome_form)
        if comp.Type != VBEXT_CT_MSFORM:
            raise ValueError(
                f"'{nome_form}' nao e um UserForm (Type={comp.Type}). "
                "So UserForms tem Designer para receber controles."
            )

        chave = str(tipo).lower()
        progid = PROGID_CONTROLE.get(chave)
        if progid is None:
            raise ValueError(
                f"Tipo de controle desconhecido: '{tipo}'. "
                f"Validos: {', '.join(sorted(set(PROGID_CONTROLE)))}."
            )

        designer = comp.Designer
        for existente in designer.Controls:
            if existente.Name.lower() == str(nome_controle).lower():
                raise ValueError(
                    f"O form '{comp.Name}' ja tem um controle chamado "
                    f"'{existente.Name}'. Escolha outro nome."
                )

        container = designer
        if dentro_de:
            container = self._controle_do_form(designer, dentro_de, comp.Name)
        ref_modelo = self._controle_do_form(designer, modelo, comp.Name) if modelo else None

        self._garantir_backup()
        ctrl = container.Controls.Add(progid)
        ctrl.Name = nome_controle
        if ref_modelo is not None:
            copiadas = self._copiar_aparencia(ref_modelo, ctrl)
            print(f"  aparencia copiada de '{ref_modelo.Name}': {len(copiadas)} propriedade(s)")
        self._aplicar_props(ctrl, props or {})
        onde = f"'{comp.Name}.{container.Name}'" if dentro_de else f"'{comp.Name}'"
        print(f"Controle '{ctrl.Name}' ({chave}) adicionado a {onde}.")
        return ctrl.Name

    def configurar_controle(self, nome_form, nome_controle, props):
        """Altera propriedades de um controle que JA EXISTE num UserForm.

        Complementa adicionar_controle(): sem isto, corrigir uma legenda ou
        reposicionar um controle exigiria remover e recriar, perdendo
        qualquer outra propriedade ja ajustada a mao no Designer.
        """
        comp = self._achar(nome_form)
        if comp.Type != VBEXT_CT_MSFORM:
            raise ValueError(
                f"'{nome_form}' nao e um UserForm (Type={comp.Type})."
            )
        designer = comp.Designer
        alvo = None
        for c in designer.Controls:
            if c.Name.lower() == str(nome_controle).lower():
                alvo = c
                break
        if alvo is None:
            existentes = ", ".join(c.Name for c in designer.Controls)
            raise KeyError(
                f"O form '{comp.Name}' nao tem controle '{nome_controle}'. "
                f"Tem: {existentes}."
            )
        if not props:
            raise ValueError("Nada a alterar: --propriedades veio vazio.")

        self._garantir_backup()
        self._aplicar_props(alvo, props)
        print(f"Controle '{alvo.Name}' do form '{comp.Name}' atualizado "
              f"({len(props)} propriedade(s)).")
        return alvo.Name

    def configurar_form(self, nome_form, props):
        """Altera propriedades do PROPRIO UserForm (Height, Width, Caption...).

        configurar_controle() so alcanca controles; crescer a janela para caber
        um campo novo exigia um driver COM a parte. Usa o mesmo caminho de
        criar_form: VBComponent.Properties primeiro (o que persiste ao salvar),
        o Designer no que nao existir la.
        """
        comp = self._achar(nome_form)
        if comp.Type != VBEXT_CT_MSFORM:
            raise ValueError(f"'{nome_form}' nao e um UserForm (Type={comp.Type}).")
        if not props:
            raise ValueError("Nada a alterar: props veio vazio.")
        self._garantir_backup()
        self._aplicar_props_form(comp, comp.Designer, props)
        print(f"UserForm '{comp.Name}' atualizado ({len(props)} propriedade(s)).")
        return comp.Name

    def remover(self, nome_modulo):
        """Remove um modulo. Documentos (planilhas/ThisWorkbook) nao podem ser removidos."""
        comp = self._achar(nome_modulo)
        if comp.Type == VBEXT_CT_DOCUMENT:
            raise ValueError(
                f"'{nome_modulo}' e um modulo de documento e nao pode ser removido. "
                "Use 'substituir' para limpar o codigo dele."
            )
        self._garantir_backup()
        self._componentes.Remove(comp)
        print(f"Modulo '{nome_modulo}' removido.")

    def ler_codigo(self, nome_modulo):
        """Retorna o codigo completo de um modulo como string."""
        cm = self._achar(nome_modulo).CodeModule
        if cm.CountOfLines == 0:
            return ""
        return cm.Lines(1, cm.CountOfLines)

    def substituir_codigo(self, nome_modulo, novo_codigo):
        """Troca TODO o codigo de um modulo pelo conteudo informado (string).

        O conteudo pode vir direto de um arquivo exportado (.bas/.cls/.frm):
        o cabecalho de export e removido antes de inserir (ver
        _tirar_cabecalho_export).
        """
        novo_codigo, removidas = _tirar_cabecalho_export(novo_codigo)
        self._recusar_reservadas(novo_codigo, nome_modulo)
        self._garantir_backup()
        cm = self._achar(nome_modulo).CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        if novo_codigo:
            cm.AddFromString(novo_codigo)
        if removidas:
            print(f"  [cabecalho] {len(removidas)} linha(s) de cabecalho de "
                  f"export descartada(s): {', '.join(removidas[:3])}"
                  + (" ..." if len(removidas) > 3 else ""))
        print(f"Codigo do modulo '{nome_modulo}' substituido "
              f"({cm.CountOfLines} linha(s)).")

    def adicionar_codigo(self, nome_modulo, codigo):
        """Anexa codigo (string) ao final de um modulo existente."""
        self._recusar_reservadas(codigo, nome_modulo)
        self._garantir_backup()
        cm = self._achar(nome_modulo).CodeModule
        cm.InsertLines(cm.CountOfLines + 1, codigo)
        print(f"Codigo anexado ao modulo '{nome_modulo}'.")

    _FIM_PROC = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)

    def _faixa_da_declaracao(self, cm, nome_proc, kind):
        """(declaracao, End) 1-based. ProcBodyLine e a linha do Sub/Function."""
        decl = cm.ProcBodyLine(nome_proc, kind)
        for i in range(decl, cm.CountOfLines + 1):
            if self._FIM_PROC.match(cm.Lines(i, 1)):
                return decl, i
        raise RuntimeError(f"End de '{nome_proc}' nao encontrado.")

    @staticmethod
    def _crlf(texto):
        return texto.replace("\r\n", "\n").replace("\n", "\r\n")

    # Liga/desliga a recusa de codigo com palavra reservada usada como nome.
    checar_reservadas = True

    def _recusar_reservadas(self, codigo, onde):
        """Recusa ANTES de gravar: o erro so apareceria ao compilar o procedimento."""
        if not self.checar_reservadas:
            return
        achados = nomes_reservados_usados(codigo)
        if achados:
            lista = ", ".join(f"linha {n}: '{nome}'" for n, nome in achados)
            raise ValueError(
                f"Codigo para '{onde}' usa palavra reservada do VBA como nome ({lista}). "
                "Isso da 'Erro de sintaxe' so na compilacao do procedimento. Renomeie. "
                "(Para forcar, ed.checar_reservadas = False.)")

    def substituir_procedimento(self, nome_modulo, nome_proc, novo_codigo, kind=0,
                                manter_cabecalho=False):
        """Substitui um procedimento inteiro (Sub/Function/Property) pelo nome.

        Usa ProcStartLine/ProcCountLines do VBE -> robusto, nao depende de casar
        texto (evita problemas com acentos/comentarios). kind: 0=Proc/Sub/Function,
        1=Let, 2=Set, 3=Get (vbext_ProcKind).

        manter_cabecalho=True: troca so da DECLARACAO ao End. O VBE conta como
        parte do procedimento os comentarios e linhas vazias ACIMA dele
        (ProcStartLine), e o modo padrao os apaga -- o que destruia os blocos
        de comentario de especificacao que os projetos poem acima de cada
        procedimento. Nesse modo, novo_codigo deve comecar na declaracao.
        """
        self._recusar_reservadas(novo_codigo, f"{nome_modulo}.{nome_proc}")
        self._garantir_backup()
        cm = self._achar(nome_modulo).CodeModule
        if manter_cabecalho:
            inicio, fim = self._faixa_da_declaracao(cm, nome_proc, kind)
            qtd = fim - inicio + 1
        else:
            inicio = cm.ProcStartLine(nome_proc, kind)
            qtd = cm.ProcCountLines(nome_proc, kind)
        cm.DeleteLines(inicio, qtd)
        cm.InsertLines(inicio, self._crlf(novo_codigo.rstrip("\r\n")))
        print(f"Procedimento '{nome_proc}' de '{nome_modulo}' substituido "
              f"(linhas {inicio}..{inicio + qtd - 1}"
              f"{', cabecalho preservado' if manter_cabecalho else ''}).")

    def trocar_linha(self, nome_modulo, antiga, nova, proc=None, kind=0):
        """Troca UMA linha de codigo, casada inteira, sem olhar caixa nem espacos de borda.

        Feita para ancoras: o VBE re-capitaliza identificadores no projeto
        todo, entao casar com caixa falha a toa. Exige exatamente UMA
        ocorrencia (no modulo, ou so dentro de `proc`) -- ancora ambigua ja
        trocou a linha errada. O recuo da linha original e aplicado a cada
        linha de `nova` (que pode ter varias linhas).
        """
        self._recusar_reservadas(nova, nome_modulo)
        self._garantir_backup()
        cm = self._achar(nome_modulo).CodeModule
        if proc:
            a, b = self._faixa_da_declaracao(cm, proc, kind)
        else:
            a, b = 1, cm.CountOfLines
        alvo = antiga.strip().lower()
        achadas = [i for i in range(a, b + 1) if cm.Lines(i, 1).strip().lower() == alvo]
        if len(achadas) != 1:
            onde = f"'{proc}'" if proc else f"'{nome_modulo}'"
            raise ValueError(f"Linha {antiga!r} achada {len(achadas)} vez(es) em {onde}; "
                             "esperado exatamente 1.")
        i = achadas[0]
        recuo = re.match(r"^\s*", cm.Lines(i, 1)).group(0)
        linhas = [(recuo + l) if l.strip() else l for l in nova.replace("\r\n", "\n").split("\n")]
        cm.DeleteLines(i, 1)
        cm.InsertLines(i, "\r\n".join(linhas))
        print(f"Linha {i} de '{nome_modulo}' trocada.")
        return i

    def ler(self, nome_modulo, proc=None, numerar=False):
        """Imprime (e retorna) o codigo de um modulo ou de um procedimento.

        proc: nome de um Sub/Function do modulo. Atencao: o VBE considera os
        comentarios imediatamente acima do procedimento como parte dele, entao
        eles aparecem junto. numerar=True prefixa o numero real de cada linha.
        """
        cm = self._achar(nome_modulo).CodeModule
        if proc:
            try:
                inicio = cm.ProcStartLine(proc, 0)  # 0 = vbext_pk_Proc (Sub/Function)
                qtd = cm.ProcCountLines(proc, 0)
            except Exception:
                raise KeyError(
                    f"Procedimento '{proc}' nao encontrado em '{nome_modulo}'.")
            codigo = cm.Lines(inicio, qtd)
            base = inicio
        else:
            if cm.CountOfLines == 0:
                print(f"(modulo '{nome_modulo}' esta vazio)")
                return ""
            codigo = cm.Lines(1, cm.CountOfLines)
            base = 1
        if numerar:
            codigo = "\n".join(f"{base + i:5d}  {linha}"
                               for i, linha in enumerate(codigo.splitlines()))
        print(codigo)
        return codigo

    def procurar(self, texto, modulo=None, sensivel=False):
        """Busca texto no codigo de todos os modulos (ou de apenas um).

        Imprime cada ocorrencia como 'modulo:linha: conteudo' e retorna a
        lista [(modulo, linha, conteudo), ...]. Por padrao ignora
        maiusculas/minusculas (como o proprio VBA).
        """
        alvo = texto if sensivel else texto.lower()
        comps = [self._achar(modulo)] if modulo else list(self._componentes)
        ocorrencias = []
        modulos_com_match = set()
        for comp in comps:
            cm = comp.CodeModule
            if cm.CountOfLines == 0:
                continue
            linhas = cm.Lines(1, cm.CountOfLines).splitlines()
            for n, linha in enumerate(linhas, start=1):
                pesquisa = linha if sensivel else linha.lower()
                if alvo in pesquisa:
                    ocorrencias.append((comp.Name, n, linha))
                    modulos_com_match.add(comp.Name)
                    print(f"{comp.Name}:{n}: {linha.strip()}")
        if ocorrencias:
            print(f"\n{len(ocorrencias)} ocorrencia(s) em "
                  f"{len(modulos_com_match)} modulo(s).")
        else:
            print(f"Nenhuma ocorrencia de '{texto}'.")
        return ocorrencias

    def _controle_compilar(self, vbe):
        """Localiza o botao 'Compilar' (Id 578) nos menus do VBE.

        Busca pelo Id numerico, que e o mesmo em qualquer idioma do Office
        (nao depende do menu se chamar 'Debug' ou 'Depurar').
        """
        def busca(controles, nivel):
            for c in controles:
                try:
                    if c.Id == ID_COMANDO_COMPILAR:
                        return c
                except Exception:
                    continue
                if nivel < 2:  # entra em submenus (popups), sem descer demais
                    try:
                        achado = busca(c.Controls, nivel + 1)
                    except Exception:
                        achado = None
                    if achado is not None:
                        return achado
            return None

        # Em algumas instalacoes do Office o VBE.CommandBars nao resolve
        # ('Erro ao carregar a biblioteca/DLL de tipo' / TYPE_E_CANTLOADLIBRARY).
        # Tenta os caminhos conhecidos antes de desistir, e desiste com uma
        # excecao propria - o chamador precisa distinguir "nao compilou" de
        # "nao consegui testar a compilacao".
        candidatos = (
            lambda: [vbe.CommandBars("Menu Bar")],  # nome interno, nao localizado
            lambda: list(vbe.CommandBars),
            lambda: [self.excel.VBE.CommandBars("Menu Bar")],
            lambda: list(self.excel.VBE.CommandBars),
        )
        ultimo_erro = None
        for obter in candidatos:
            try:
                barras = obter()
            except Exception as e:
                ultimo_erro = e
                continue
            for barra in barras:
                try:
                    achado = busca(barra.Controls, 0)
                except Exception as e:
                    ultimo_erro = e
                    continue
                if achado is not None:
                    return achado
        if ultimo_erro is not None:
            raise CompilacaoIndisponivel(
                "nao consegui acessar os menus do VBE para acionar o "
                f"'Compilar' ({ultimo_erro})")
        return None

    @staticmethod
    def _dialogo_erro_compilacao():
        """Procura a caixa de erro do VBE. Retorna (hwnd, mensagem) ou None."""
        achados = []

        def enum_cb(hwnd, acc):
            if (win32gui.IsWindowVisible(hwnd)
                    and win32gui.GetClassName(hwnd) == "#32770"
                    and "Visual Basic" in win32gui.GetWindowText(hwnd)):
                acc.append(hwnd)

        win32gui.EnumWindows(enum_cb, achados)
        if not achados:
            return None
        hwnd = achados[0]
        textos = []
        win32gui.EnumChildWindows(
            hwnd,
            lambda h, acc: acc.append(
                (win32gui.GetClassName(h), win32gui.GetWindowText(h))),
            textos)
        mensagem = " ".join(t.replace("\r", " ").replace("\n", " ").strip()
                            for cls, t in textos
                            if cls == "Static" and t.strip())
        return hwnd, mensagem

    def verificar(self):
        """Compila o projeto VBA (Depurar > Compilar) e informa se ha erro.

        Em caso de erro de compilacao, captura a mensagem da caixa de
        dialogo do VBE, fecha-a sozinho e informa modulo/linha do erro
        (o VBE deixa o cursor exatamente nela). Retorna True se compilou.
        Nao altera nem salva o workbook.
        """
        vbe = self.excel.VBE
        try:
            vbe.ActiveVBProject = self.wb.VBProject
        except Exception:
            pass  # com um unico workbook aberto o projeto ativo ja e o dele
        try:
            ctrl = self._controle_compilar(vbe)
        except CompilacaoIndisponivel as e:
            print(f"INDETERMINADO: {e}.\n"
                  "  O codigo foi gravado, mas a compilacao NAO foi testada.\n"
                  "  Verifique no Excel: Alt+F11 > Depurar > Compilar VBAProject.")
            return None  # nem True nem False: teste nao realizado
        if ctrl is None:
            print("INDETERMINADO: botao 'Compilar' nao encontrado nos menus do "
                  "VBE.\n  O codigo foi gravado, mas a compilacao NAO foi "
                  "testada.")
            return None
        if not ctrl.Enabled:
            # O VBE desabilita 'Compilar' quando o projeto ja esta compilado
            print("OK: projeto ja estava compilado.")
            return True

        vbe.MainWindow.Visible = True  # a caixa de erro so aparece com o VBE visivel
        try:
            ctrl.Execute()
            dialogo = None
            for _ in range(40):  # ate 10 s
                time.sleep(0.25)
                dialogo = self._dialogo_erro_compilacao()
                if dialogo is not None or not ctrl.Enabled:
                    break

            if dialogo is None and not ctrl.Enabled:
                print("OK: projeto compilado sem erros.")
                return True

            mensagem = ""
            if dialogo is not None:
                hwnd, mensagem = dialogo
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                time.sleep(0.5)

            print("FALHOU: o projeto NAO compilou.")
            if mensagem:
                print(f"  Mensagem do VBE: {mensagem}")
            try:
                painel = vbe.ActiveCodePane
                cm = painel.CodeModule
                linha = painel.GetSelection()[0]
                print(f"  Local: {cm.Parent.Name}, linha {linha}:")
                print(f"    {cm.Lines(linha, 1).strip()}")
            except Exception:
                pass  # sem localizacao, a mensagem ja ajuda
            return False
        finally:
            vbe.MainWindow.Visible = False

    def corrigir_nomes(self, padrao="_FilterDatabase"):
        """Remove nomes definidos problematicos do workbook (causa do
        dialogo 'Conflito de nome' na abertura).

        Apaga todo nome definido cujo nome contem 'padrao' (padrao:
        '_FilterDatabase', nome oculto que o Excel cria para AutoFiltro e que,
        duplicado/corrompido, gera o conflito). Remover e seguro: o Excel
        recria o nome quando um filtro e aplicado de novo. Tambem apaga os
        nomes 'NomeDuplicado_N' criados pelo vigia ao responder o dialogo
        (e os '_NomeRenomeado_N' de versoes antigas da ferramenta).
        Retorna a quantidade removida.
        """
        self._garantir_backup()
        removidos = []
        falhas = []
        # Materializa a lista antes: deletar enquanto itera a colecao COM
        # pula itens.
        nomes = [self.wb.Names(i + 1) for i in range(self.wb.Names.Count)]
        for nm in nomes:
            try:
                rotulo = nm.Name  # sheet-level vem como 'Planilha!_FilterDatabase'
            except Exception:
                continue
            if (padrao.lower() in rotulo.lower()
                    or PREFIXO_RENOMEACAO in rotulo
                    or PREFIXO_RENOMEACAO_LEGADO in rotulo):
                try:
                    nm.Delete()
                    removidos.append(rotulo)
                except Exception as e:
                    falhas.append((rotulo, str(e)))
        for rotulo in removidos:
            print(f"  removido: {rotulo}")
        for rotulo, erro in falhas:
            print(f"  [FALHOU] {rotulo}: {erro}")
        print(f"\n{len(removidos)} nome(s) removido(s), "
              f"{len(falhas)} falha(s).")
        return len(removidos)

    def editar(self, nome_modulo, texto_de, texto_para, todas=True):
        """Busca e substitui texto dentro do codigo de um modulo.

        Retorna a quantidade de ocorrencias substituidas.
        """
        codigo = self.ler_codigo(nome_modulo)
        ocorrencias = codigo.count(texto_de)
        if ocorrencias == 0:
            print(f"Nenhuma ocorrencia de '{texto_de}' em '{nome_modulo}'.")
            return 0
        novo = (codigo.replace(texto_de, texto_para)
                if todas else codigo.replace(texto_de, texto_para, 1))
        self.substituir_codigo(nome_modulo, novo)
        feitas = ocorrencias if todas else 1
        print(f"{feitas} ocorrencia(s) de '{texto_de}' substituida(s) por "
              f"'{texto_para}' em '{nome_modulo}'.")
        return feitas


# ----------------------------------------------------------------------
# Interface de linha de comando
# ----------------------------------------------------------------------
def main(argv=None):
    p = argparse.ArgumentParser(
        description="Edita os modulos de codigo VBA de um workbook do Excel.")
    p.add_argument("workbook", help="Caminho do arquivo .xlsm/.xlsb/.xls")
    p.add_argument("--visivel", action="store_true",
                   help="Mostra o Excel durante a execucao (util para depurar).")

    sub = p.add_subparsers(dest="comando", required=True)

    sub.add_parser("listar", help="Lista os modulos do projeto.")

    sub.add_parser("backup", help="Cria uma copia datada do workbook em backups/.")

    sp = sub.add_parser("exportar", help="Exporta todos os modulos (backup).")
    sp.add_argument("--pasta", default="backup_vba", help="Pasta de destino.")

    sp = sub.add_parser("importar", help="Importa um modulo de um arquivo.")
    sp.add_argument("--arquivo", required=True, help="Arquivo .bas/.cls/.frm.")

    sp = sub.add_parser("substituir", help="Substitui todo o codigo de um modulo.")
    sp.add_argument("--modulo", required=True)
    sp.add_argument("--codigo", required=True,
                    help="Caminho de um arquivo de texto com o novo codigo.")

    sp = sub.add_parser("editar", help="Busca e substitui texto no codigo.")
    sp.add_argument("--modulo", required=True)
    sp.add_argument("--de", required=True, help="Texto a procurar.")
    sp.add_argument("--para", required=True, help="Texto de substituicao.")
    sp.add_argument("--primeira", action="store_true",
                    help="Substitui apenas a primeira ocorrencia.")

    sp = sub.add_parser("adicionar", help="Cria um modulo novo (std/classe/form vazio).")
    sp.add_argument("--modulo", required=True)
    sp.add_argument("--tipo", choices=["std", "classe", "form"], default="std")

    sp = sub.add_parser("criar-form",
                        help="Cria um UserForm com controles a partir de um spec JSON.")
    sp.add_argument("--spec", required=True,
                    help="Arquivo JSON com a especificacao do form e seus controles.")

    sp = sub.add_parser("adicionar-controle",
                        help="Acrescenta um controle a um UserForm existente.")
    sp.add_argument("--modulo", required=True, help="UserForm que recebe o controle.")
    sp.add_argument("--tipo", required=True,
                    help="Tipo do controle (ex: checkbox, textbox, botao).")
    sp.add_argument("--nome", required=True, help="Nome do controle criado.")
    sp.add_argument("--propriedades",
                    help="JSON com propriedades (caption, left, top, width, "
                         "height, value, \"Font.Size\"...).")
    sp.add_argument("--dentro-de",
                    help="Frame (ou outro container) que recebe o controle; "
                         "left/top ficam relativos a ele.")
    sp.add_argument("--modelo",
                    help="Controle existente cuja aparencia (fonte, cores, "
                         "borda) o novo copia.")

    sp = sub.add_parser("trocar-linha",
                        help="Troca UMA linha casada inteira, sem olhar caixa.")
    sp.add_argument("--modulo", required=True)
    sp.add_argument("--de", required=True, help="Linha atual (sem o recuo).")
    sp.add_argument("--para", required=True, help="Linha(s) nova(s); o recuo e mantido.")
    sp.add_argument("--proc", help="Procura so dentro deste procedimento.")

    sp = sub.add_parser("substituir-proc",
                        help="Substitui um procedimento inteiro pelo nome.")
    sp.add_argument("--modulo", required=True)
    sp.add_argument("--proc", required=True)
    sp.add_argument("--codigo", required=True,
                    help="Arquivo de texto com o procedimento novo.")
    sp.add_argument("--manter-cabecalho", action="store_true",
                    help="Preserva os comentarios acima da declaracao.")

    sp = sub.add_parser("config-form",
                        help="Altera propriedades do proprio UserForm (Height, Width, Caption).")
    sp.add_argument("--modulo", required=True, help="UserForm a alterar.")
    sp.add_argument("--propriedades", required=True, help="JSON com as propriedades.")

    sp = sub.add_parser("config-controle",
                        help="Altera propriedades de um controle existente.")
    sp.add_argument("--modulo", required=True, help="UserForm que contem o controle.")
    sp.add_argument("--nome", required=True, help="Nome do controle a alterar.")
    sp.add_argument("--propriedades", required=True,
                    help="JSON com as propriedades a alterar.")

    sp = sub.add_parser("remover", help="Remove um modulo.")
    sp.add_argument("--modulo", required=True)

    sp = sub.add_parser("ler", help="Imprime o codigo de um modulo ou procedimento.")
    sp.add_argument("--modulo", required=True)
    sp.add_argument("--proc", help="Nome de um Sub/Function especifico.")
    sp.add_argument("--numerar", action="store_true",
                    help="Prefixa cada linha com o numero real dela no modulo.")

    sp = sub.add_parser("procurar", help="Busca texto no codigo dos modulos.")
    sp.add_argument("--texto", required=True, help="Texto a procurar.")
    sp.add_argument("--modulo", help="Limita a busca a um unico modulo.")
    sp.add_argument("--sensivel", action="store_true",
                    help="Diferencia maiusculas de minusculas.")

    sp = sub.add_parser("corrigir-nomes",
                        help="Remove nomes '_FilterDatabase' duplicados que "
                             "causam o dialogo 'Conflito de nome' na abertura.")
    sp.add_argument("--padrao", default="_FilterDatabase",
                    help="Trecho do nome definido a remover "
                         "(padrao: _FilterDatabase).")

    sub.add_parser("verificar", help="Compila o projeto VBA e informa se ha erro.")

    args = p.parse_args(argv)

    # Acentos fora do charset do console nao devem derrubar o programa
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")

    # Somente comandos que alteram o projeto salvam o workbook ao fechar.
    # Salvar em comando de leitura reescreveria o arquivo inteiro a toa
    # (mtime e bytes mudam, e o git acusaria alteracao sem haver edicao).
    comandos_que_salvam = {"importar", "substituir", "editar", "adicionar",
                           "remover", "criar-form", "adicionar-controle",
                           "config-controle", "config-form", "corrigir-nomes", "trocar-linha",
                           "substituir-proc"}

    # Comandos que nao salvam abrem o arquivo em somente-leitura: dispensa
    # fechar o workbook antes de rodar e contorna workbooks que recusam a
    # abertura para escrita via automacao.
    ed = VBAEditor(args.workbook, visivel=args.visivel,
                   somente_leitura=args.comando not in comandos_que_salvam)
    try:
        ed.abrir()
    except BaseException:
        ed.fechar(salvar=False)  # nao deixa um Excel orfao segurando o arquivo
        raise
    verificacao_ok = True
    try:
        if args.comando == "listar":
            ed.listar()
        elif args.comando == "backup":
            ed._garantir_backup()
        elif args.comando == "exportar":
            ed.exportar(args.pasta)
        elif args.comando == "importar":
            ed.importar(args.arquivo)
        elif args.comando == "substituir":
            with open(args.codigo, "r", encoding="utf-8") as f:
                ed.substituir_codigo(args.modulo, f.read())
        elif args.comando == "editar":
            ed.editar(args.modulo, args.de, args.para, todas=not args.primeira)
        elif args.comando == "adicionar":
            ed.adicionar(args.modulo, args.tipo)
        elif args.comando == "criar-form":
            with open(args.spec, "r", encoding="utf-8") as f:
                ed.criar_form(json.load(f))
        elif args.comando == "adicionar-controle":
            props = json.loads(args.propriedades) if args.propriedades else {}
            ed.adicionar_controle(args.modulo, args.tipo, args.nome, props,
                                  dentro_de=args.dentro_de, modelo=args.modelo)
        elif args.comando == "trocar-linha":
            ed.trocar_linha(args.modulo, args.de, args.para, proc=args.proc)
        elif args.comando == "substituir-proc":
            with open(args.codigo, "r", encoding="utf-8") as f:
                ed.substituir_procedimento(args.modulo, args.proc, f.read(),
                                           manter_cabecalho=args.manter_cabecalho)
        elif args.comando == "config-form":
            ed.configurar_form(args.modulo, json.loads(args.propriedades))
        elif args.comando == "config-controle":
            ed.configurar_controle(args.modulo, args.nome,
                                   json.loads(args.propriedades))
        elif args.comando == "remover":
            ed.remover(args.modulo)
        elif args.comando == "ler":
            ed.ler(args.modulo, proc=args.proc, numerar=args.numerar)
        elif args.comando == "procurar":
            ed.procurar(args.texto, modulo=args.modulo, sensivel=args.sensivel)
        elif args.comando == "corrigir-nomes":
            ed.corrigir_nomes(args.padrao)
        elif args.comando == "verificar":
            verificacao_ok = ed.verificar()
    except BaseException:
        ed.fechar(salvar=False)  # erro: nunca salva
        raise
    ed.fechar(salvar=args.comando in comandos_que_salvam)
    # exit codes do 'verificar': 0 = compilou, 1 = NAO compilou,
    # 2 = indeterminado (nao foi possivel acionar o Compilar neste ambiente).
    # Sao distintos de proposito: um script nao pode tratar "nao testei" como
    # "esta tudo certo" nem como "quebrou".
    if verificacao_ok is None:
        sys.exit(2)
    if not verificacao_ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
