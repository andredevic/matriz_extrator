# Matriz Extrator (Fontes de Energia / Equipamentos)

Projeto para consolidar Matrizes Energéticas (Excel) em uma planilha única, pronta para importação no sistema.

## Estrutura

- `planilhas/` → coloque aqui todas as matrizes `.xlsx`
- `saida/` → o script gera aqui o arquivo final
- `extrair_matriz.py` → script principal

## Saída gerada

- `saida/matriz_consolidada.xlsx`

## Como usar

1. Copie suas matrizes para a pasta `planilhas/`
2. Rode:

```bash
python extrair_matriz.py
```

3. Pegue o arquivo:

- `saida/matriz_consolidada.xlsx`

## Regras importantes

- Linhas começam na **linha 11**.
- Campos de TAG/descrição/processo que resultarem vazios viram **null** (None).
- `Tag do Equipamento` (B) e `Descrição do Equipamento` (C+D) fazem **fill-down** (herdam de cima se vierem vazias).
- Linhas sem nenhuma informação de fonte/processo são ignoradas.

