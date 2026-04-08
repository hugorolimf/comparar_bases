# Comparador Inteligente de Excel

Script Python modular para comparar dois arquivos Excel de forma automática, com detecção de cabeçalho, inferência básica de tipos, validação prévia e geração de diff em Excel e JSON.

## Execução

```bash
python main.py
```

O script vai perguntar:

1. Número do Excel base, escolhido a partir da pasta `excel`.
2. Número do Excel de comparação, escolhido a partir da mesma pasta.
3. Número da aba de cada arquivo.
4. Chave principal para parear os registros entre os dois arquivos.
5. Colunas de diff da base e colunas de diff da comparação, selecionadas separadamente e pareadas por ordem.
6. Pasta e nome de saída.

Antes da escolha da chave, o script mostra os melhores matches encontrados entre as colunas de ambos os arquivos e destaca a melhor sugestão. Depois disso, ele lista as colunas de cada arquivo para a seleção dos identificadores do diff.

## Saídas

- Arquivo `.xlsx` com abas `Resumo`, `Mapeamento`, `Adição`, `Exclusão`, `Alteração` e `Igual`.
- Arquivo `.json` com a mesma estrutura para automação.

## Exportação

- `excel_diff/reporting/excel_report.py`: exportação exclusiva para Excel.
- `excel_diff/reporting/json_report.py`: exportação exclusiva para JSON.
- `excel_diff/reporting/report_writer.py`: orquestra os dois módulos sem misturar a lógica.

## Estrutura

- `main.py`: ponto de entrada.
- `excel_diff/cli.py`: fluxo interativo.
- `excel_diff/analysis/`: detecção de cabeçalhos, schema e compatibilidade.
- `excel_diff/comparison/`: motor de diff.
- `excel_diff/reporting/`: geração de Excel e JSON.
- `excel_diff/io/`: leitura dos workbooks.
- `excel_diff/utils/`: normalização e helpers.

## Dependências

- `openpyxl`
