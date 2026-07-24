# vba-editor

Ferramenta em Python para ler e alterar os **módulos de código VBA** de um workbook
do Excel (`.xlsm`/`.xlsb`/`.xls`) via `pywin32` (COM), sem abrir o editor manualmente.

Permite: listar, exportar (backup), importar, substituir o código de um módulo,
editar (buscar/substituir texto), substituir um procedimento inteiro pelo nome,
adicionar e remover módulos, e **criar UserForms com controles** a partir de uma
especificação JSON. Faz **backup datado automático** antes da primeira alteração.

## Requisitos

- Windows com Excel instalado.
- `pip install pywin32`
- No Excel: *Arquivo > Opções > Central de Confiabilidade > Configurações de Macro >
  "Confiar no acesso ao modelo de objeto do projeto do VBA"*.

## Uso (linha de comando)

```bash
python vba_editor.py "C:\caminho\workbook.xlsm" listar
python vba_editor.py "workbook.xlsm" backup
python vba_editor.py "workbook.xlsm" exportar --pasta backup_vba
python vba_editor.py "workbook.xlsm" editar --modulo Conexao --de "x" --para "y"
python vba_editor.py "workbook.xlsm" substituir --modulo Module1 --codigo novo.txt
python vba_editor.py "workbook.xlsm" adicionar --modulo MeuModulo --tipo std
python vba_editor.py "workbook.xlsm" adicionar --modulo frmVazio  --tipo form
python vba_editor.py "workbook.xlsm" criar-form --spec form.json
python vba_editor.py "workbook.xlsm" remover --modulo MeuModulo
```

## Criar um UserForm com controles

O comando `criar-form` monta um UserForm inteiro (janela + controles + código de
eventos) a partir de um arquivo JSON. Exemplo de `form.json`:

```json
{
  "nome": "frmCadastro",
  "caption": "Cadastro de Cliente",
  "largura": 300,
  "altura": 180,
  "controles": [
    {"tipo": "label",   "nome": "lblNome", "caption": "Nome:",  "left": 12, "top": 16, "width": 60,  "height": 18},
    {"tipo": "textbox", "nome": "txtNome",                      "left": 80, "top": 14, "width": 190, "height": 20},
    {"tipo": "botao",   "nome": "btnOK",   "caption": "Gravar", "left": 110, "top": 110, "width": 75, "height": 26}
  ],
  "codigo": "Private Sub btnOK_Click()\n    MsgBox txtNome.Text\n    Unload Me\nEnd Sub\n"
}
```

- **`tipo`** aceita nomes em pt/en: `label`/`rotulo`, `textbox`/`texto`, `botao`/`commandbutton`,
  `combobox`, `listbox`/`lista`, `checkbox`, `optionbutton`/`opcao`, `togglebutton`,
  `frame`/`quadro`, `image`/`imagem`, `spinbutton`, `scrollbar`, `multipage`, `tabstrip`.
- **Propriedades**: `nome`, `caption`, `left`/`top`/`width`/`height` (ou `esquerda`/`topo`/`largura`/`altura`),
  `texto`, `valor`. Qualquer outra chave é repassada como propriedade nativa do MSForms
  (ex.: `"BackColor"`, `"Enabled"`); se o controle não a aceitar, ela é ignorada com aviso.
- **`codigo`** (opcional) é o código VBA do módulo do form — use `\n` para quebras de linha.

> A ferramenta cria uma instância própria e isolada do Excel (`DispatchEx`) —
> ela nunca toca numa sessão do Excel que você já tenha aberta. Se o **próprio
> workbook alvo** estiver aberto, ela aborta com aviso (detecta o lock `~$`);
> feche-o antes de rodar. Isso vale só para os comandos que gravam: os de
> leitura abrem em somente-leitura e não se importam com o lock.

## Abertura de workbooks problemáticos

Alguns workbooks travavam a automação na abertura. O que a ferramenta faz hoje:

- **Diálogo "Nomes em conflito"** (nomes definidos duplicados, típico
  `_FilterDatabase`): uma thread vigia responde sozinha, digitando um nome
  novo e confirmando. Detalhes que importam para quem for mexer nisso:
  - o diálogo **não** é um `#32770`: ele vem na classe própria do Excel
    `bosa_sdm_XL9`, e sua caixa de texto é `EDTBX`, não `Edit`;
  - ele **não tem botões-janela** — a confirmação sai por `ENTER` na caixa;
  - a digitação usa `WM_CHAR` via `PostMessage`. `SendInput`/`keybd_event` não
    funcionam: exigem que a janela seja a foreground do desktop interativo, o
    que nunca acontece com um Excel invisível rodando por automação;
  - o nome digitado **não pode começar com `_`** — o Excel recusa como nome
    interno e abre um segundo diálogo de erro (que o vigia também trata);
  - **cancelar não resolve**: o Cancel aborta o `Workbooks.Open` inteiro, que
    estoura `0x800A03EC` ("O método Open da classe Workbooks falhou").
  - use `corrigir-nomes` para eliminar a causa e não depender mais do vigia.
- **Qualquer outro diálogo** que trave a abertura é reportado com título,
  texto e botões, em vez de deixar a automação esperando calada.
- **Recálculo desligado antes do `Open`**: em modo automático, abrir um arquivo
  grande recalcula a pasta inteira durante o próprio `Open`.
- **Comandos de leitura** (`listar`, `ler`, `exportar`, `procurar`,
  `verificar`, `backup`) abrem o arquivo em somente-leitura. Além de dispensar
  fechar o workbook antes, isso contorna workbooks que recusam a abertura para
  escrita via automação.

## Cabeçalho de export é descartado automaticamente

`substituir` aceita um arquivo `.bas`/`.cls`/`.frm` **exportado** direto, sem
edição manual: o cabeçalho de export é removido antes de o código entrar no
módulo.

Isso importa porque esse cabeçalho é metadado do **arquivo**, não código — o
painel do VBE nem o exibe. Reinseri-lo por `AddFromString` deixa uma linha
inválida dentro do módulo e o projeto para de compilar:

- `.bas` → `Attribute VB_Name = "..."`
- `.cls` / `.frm` → bloco `VERSION 1.0 CLASS` / `BEGIN` … `END` mais os
  vários `Attribute VB_*` (caso pior: `VERSION`, `BEGIN` e `END` viram
  linhas de código sem sentido)

Quando algo é descartado, a ferramenta informa:

```
  [cabecalho] 1 linha(s) de cabecalho de export descartada(s): Attribute VB_Name = "Inversao_lotes"
```

Código que já vem sem cabeçalho passa intacto.

## Uso como biblioteca

```python
from vba_editor import VBAEditor

with VBAEditor("workbook.xlsm") as ed:          # salva ao sair, só se não houver erro
    ed.substituir_procedimento("ClsX", "MinhaFunc", novo_codigo)  # troca por nome
    print(ed.ler_codigo("ClsX"))
```

## Observações

- `*.xlsm` e `backups/` estão no `.gitignore` — este repositório versiona **apenas a ferramenta**.
- `substituir_procedimento` usa `ProcStartLine`/`ProcCountLines` do VBE, evitando casar
  texto (robusto a acentos/comentários).
