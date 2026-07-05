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
  - remover   : apagar um modulo
  - listar    : mostrar os modulos existentes
  - ler       : imprimir o codigo de um modulo (ou de um procedimento)
  - procurar  : buscar texto em todos os modulos (modulo:linha: trecho)
  - verificar : compilar o projeto VBA e informar se ha erro

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
import shutil
import argparse
from datetime import datetime

import win32com.client
import win32con
import win32gui


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

# Id do botao 'Depurar > Compilar VBAProject' nos menus do VBE
# (independe do idioma do Office).
ID_COMANDO_COMPILAR = 578


class VBAEditor:
    """Abre um workbook e expoe operacoes sobre seus modulos VBA.

    Use como context manager para garantir que o Excel seja fechado:

        with VBAEditor("arquivo.xlsm") as ed:
            ed.listar()
    """

    def __init__(self, caminho_workbook, visivel=False,
                 auto_backup=True, pasta_backup="backups"):
        self.caminho = os.path.abspath(caminho_workbook)
        if not os.path.exists(self.caminho):
            raise FileNotFoundError(f"Workbook nao encontrado: {self.caminho}")
        self.visivel = visivel
        self.auto_backup = auto_backup
        self.pasta_backup = pasta_backup
        self._fez_backup = False
        self.excel = None
        self.wb = None

    # ---- ciclo de vida -------------------------------------------------
    def __enter__(self):
        self.abrir()
        return self

    def __exit__(self, exc_type, exc, tb):
        # Salva so se nao houve excecao
        self.fechar(salvar=exc_type is None)

    def abrir(self):
        self._verificar_lock()
        # DispatchEx: SEMPRE cria uma instancia nova e isolada do Excel.
        # Dispatch reaproveitaria uma instancia ja aberta pelo usuario, e o
        # Quit() do fechar() derrubaria a sessao dele (podendo descartar
        # trabalho nao salvo).
        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.Visible = self.visivel
        self.excel.DisplayAlerts = False
        self.wb = self.excel.Workbooks.Open(self.caminho)
        self._verificar_acesso_vbproject()
        return self

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

    def _verificar_acesso_vbproject(self):
        try:
            _ = self.wb.VBProject.VBComponents.Count
        except Exception as e:
            raise PermissionError(
                "Nao foi possivel acessar o VBProject. Habilite no Excel: "
                "Central de Confiabilidade > Configuracoes de Macro > "
                "'Confiar no acesso ao modelo de objeto do projeto do VBA'.\n"
                f"Detalhe: {e}"
            )

    def fechar(self, salvar=True):
        try:
            if self.wb is not None:
                if salvar:
                    self.wb.Save()
                self.wb.Close(SaveChanges=False)
        finally:
            if self.excel is not None:
                self.excel.Quit()
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
    def listar(self):
        """Lista os modulos do projeto VBA."""
        comps = self._componentes
        print(f"\nProjeto VBA de: {os.path.basename(self.caminho)}")
        print(f"Total de modulos: {comps.Count}\n")
        print(f"{'NOME':<30} {'TIPO':<28} {'LINHAS':>7}")
        print("-" * 67)
        infos = []
        for comp in comps:
            tipo = TIPO_NOME.get(comp.Type, f"Tipo {comp.Type}")
            linhas = comp.CodeModule.CountOfLines
            print(f"{comp.Name:<30} {tipo:<28} {linhas:>7}")
            infos.append((comp.Name, tipo, linhas))
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
                setattr(obj, nome, valor)
            except Exception as e:
                print(f"  [aviso] propriedade '{nome}'={valor!r} ignorada: {e}")

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
        """Troca TODO o codigo de um modulo pelo conteudo informado (string)."""
        self._garantir_backup()
        cm = self._achar(nome_modulo).CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        if novo_codigo:
            cm.AddFromString(novo_codigo)
        print(f"Codigo do modulo '{nome_modulo}' substituido "
              f"({cm.CountOfLines} linha(s)).")

    def adicionar_codigo(self, nome_modulo, codigo):
        """Anexa codigo (string) ao final de um modulo existente."""
        self._garantir_backup()
        cm = self._achar(nome_modulo).CodeModule
        cm.InsertLines(cm.CountOfLines + 1, codigo)
        print(f"Codigo anexado ao modulo '{nome_modulo}'.")

    def substituir_procedimento(self, nome_modulo, nome_proc, novo_codigo, kind=0):
        """Substitui um procedimento inteiro (Sub/Function/Property) pelo nome.

        Usa ProcStartLine/ProcCountLines do VBE -> robusto, nao depende de casar
        texto (evita problemas com acentos/comentarios). kind: 0=Proc/Sub/Function,
        1=Set, 2=Get, 3=Let.
        """
        self._garantir_backup()
        cm = self._achar(nome_modulo).CodeModule
        inicio = cm.ProcStartLine(nome_proc, kind)
        qtd = cm.ProcCountLines(nome_proc, kind)
        cm.DeleteLines(inicio, qtd)
        cm.InsertLines(inicio, novo_codigo)
        print(f"Procedimento '{nome_proc}' de '{nome_modulo}' substituido "
              f"(linhas {inicio}..{inicio + qtd - 1}).")

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

        try:
            barras = [vbe.CommandBars("Menu Bar")]  # nome interno, nao localizado
        except Exception:
            barras = list(vbe.CommandBars)
        for barra in barras:
            achado = busca(barra.Controls, 0)
            if achado is not None:
                return achado
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
        ctrl = self._controle_compilar(vbe)
        if ctrl is None:
            raise RuntimeError("Botao 'Compilar' nao encontrado nos menus do VBE.")
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

    sub.add_parser("verificar", help="Compila o projeto VBA e informa se ha erro.")

    args = p.parse_args(argv)

    # Acentos fora do charset do console nao devem derrubar o programa
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")

    # Somente comandos que alteram o projeto salvam o workbook ao fechar.
    # Salvar em comando de leitura reescreveria o arquivo inteiro a toa
    # (mtime e bytes mudam, e o git acusaria alteracao sem haver edicao).
    comandos_que_salvam = {"importar", "substituir", "editar", "adicionar",
                           "remover", "criar-form"}

    ed = VBAEditor(args.workbook, visivel=args.visivel)
    ed.abrir()
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
        elif args.comando == "remover":
            ed.remover(args.modulo)
        elif args.comando == "ler":
            ed.ler(args.modulo, proc=args.proc, numerar=args.numerar)
        elif args.comando == "procurar":
            ed.procurar(args.texto, modulo=args.modulo, sensivel=args.sensivel)
        elif args.comando == "verificar":
            verificacao_ok = ed.verificar()
    except BaseException:
        ed.fechar(salvar=False)  # erro: nunca salva
        raise
    ed.fechar(salvar=args.comando in comandos_que_salvam)
    if not verificacao_ok:
        sys.exit(1)  # permite usar 'verificar' em scripts: exit code 1 = nao compilou


if __name__ == "__main__":
    main()
