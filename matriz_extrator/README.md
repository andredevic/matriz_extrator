# Matriz Extrator (Fontes de Energia / Equipamentos)

Projeto para consolidar Matrizes Energéticas (Excel) em uma planilha única, pronta para importação no sistema de cadastro de Fontes de Energia e Equipamentos.

O script lê todas as matrizes colocadas na pasta `planilhas/`, extrai apenas os dados relevantes da matriz e gera um arquivo consolidado.

---

## Estrutura do projeto

matriz_extrator/<br>
├─ planilhas/  # COLOQUE AQUI TODAS AS MATRIZES (.xls, .xlsx, .xlsm) <br>
├─ convertidos/ # gerado automaticamente (conversão de .xls → .xlsx)<br>
├─ saida/ # arquivo final será gerado aqui<br>
├─ extrair_matriz.py # script principal<br>
└─ README.md<br>
---

## Requisitos

- Windows
- Python 3.9 ou superior
- Microsoft Excel instalado (necessário para converter arquivos `.xls`)

### Bibliotecas Python

Instale uma única vez:

```bash
python -m pip install pandas openpyxl pywin32
```
### Como usar

- 1 Copie todas as matrizes energéticas para a pasta:

```bash
planilhas/
```
Pode colocar quantos arquivos quiser.
Extensões suportadas: .xls, .xlsx, .xlsm.

- 2 No terminal, dentro da pasta do projeto, execute:
```bash
python extrair_matriz.py
```

- 3 Ao final do processo, o arquivo será gerado em:
```bash
saida/matriz_consolidada.xlsx
```
