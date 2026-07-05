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
> feche-o antes de rodar.

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
