let
    // Carrega a planilha Excel
    Fonte = Excel.Workbook(File.Contents("DadosFinanceiros.xlsx"), null, true),
    
    // Seleciona a aba específica 'DadosFinanceiros' da planilha
    DadosFinanceiros_Sheet = Fonte{[Item="DadosFinanceiros",Kind="Sheet"]}[Data],
    
    // Aplica a função 'Unpivot' para transformar colunas em linhas
    // Importante: A função Table.UnpivotOtherColumns espera a tabela como seu primeiro argumento
    #"Tabela Transposta" = Table.UnpivotOtherColumns(DadosFinanceiros_Sheet, {"Tipo", "Componente"}, "Data", "Valor")
in
    #"Tabela Transposta"

