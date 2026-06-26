# vba-editor

Ferramenta em Python para ler e alterar os **módulos de código VBA** de um workbook
do Excel (`.xlsm`/`.xlsb`/`.xls`) via `pywin32` (COM), sem abrir o editor manualmente.

Permite: listar, exportar (backup), importar, substituir o código de um módulo,
editar (buscar/substituir texto), substituir um procedimento inteiro pelo nome,
adicionar e remover módulos. Faz **backup datado automático** antes da primeira alteração.

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
python vba_editor.py "workbook.xlsm" remover --modulo MeuModulo
```

> Feche o Excel antes de rodar (instâncias abertas travam o arquivo).

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
