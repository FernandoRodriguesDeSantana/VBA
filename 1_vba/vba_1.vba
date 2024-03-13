Attribute VB_Name = "Módulo1"
'Apertar F8 para debugar o código linha por linha'

Sub compilação()
'Sub "compilação" é o nome do código que irei desenvolver'

linha = 1

linha_fim = Range("A1").End(xlDown).Row
'Range seleciona uma célula. .End(xlDown) desce até a última linha preenchida. .Row "pega" o número da linha que contém a informação'
'Se ao invés de .Row eu usasse .Column iria "pegar" o número da coluna que a célula selecionada se encontra'

Range("C2:C" & linha_fim).Copy 'Seleciona a célula C2 até a célula C(linha_fim) e faz uma cópia'
Range("J1").PasteSpecial 'Seleciona a célula J1 e a partir dela cola o conteúdo que foi copiado'
Application.CutCopyMode = False 'funcionalidade para "apertar" ESC para descelecionar o conteúdo copiado'
ActiveSheet.Range("$J$1:$J$" & linha_fim).RemoveDuplicates Columns:=1, Header:=xlNo 'Remove os valores duplicados de uma seleção'
linha_fim = Range("J1").End(xlDown).Row


While linha <= linha_fim
    Sheets.Add After:=ActiveSheet 'Para cada iteração criará uma nova aba'
    ActiveSheet.Name = Sheets("Base de Dados").Cells(linha, 10) 'ActiveSheet.Name nomeia a nova aba com o conteúdo selecionado por .Cells(linha, coluna)'
    Sheets("Base de Dados").Range("A1:C1").Copy  'Entra na aba Base de Dados, seleciona as células de A1:C1 e as copia'
    ActiveSheet.Range("A1").PasteSpecial 'Seleciona a aba "Base de Dados" e cola a partir da coluna A1 o que foi copiado'
    
    linha = linha + 1 'Iteração'
Wend

Sheets("Base de Dados").Range("J:J").Clear
linha = 2

While Sheets("Base de Dados").Cells(linha, 1) <> "" '<> significa "diferente de" e "" significa "vazio"'
    Sheets("Base de Dados").Range("A" & linha & ":C" & linha).Copy
    bairro = Sheets("Base de Dados").Cells(linha, 3)
    Sheets(bairro).Select 'Seleciona a aba com o nome do conteúdo da variável bairro'
    Range("A10000").End(xlUp).Offset(1, 0).PasteSpecial 'Vai para a célula 10000, sobe até encontrar uma informação, desce uma linha'
    Application.CutCopyMode = False
    Sheets("Base de Dados").Select
    
    linha = linha + 1
Wend

For Each aba In ThisWorkbook.Sheets
    aba.Columns("A:C").AutoFit
Next

End Sub
