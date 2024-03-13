Attribute VB_Name = "M�dulo1"
'Apertar F8 para debugar o c�digo linha por linha'

Sub compila��o()
'Sub "compila��o" � o nome do c�digo que irei desenvolver'

linha = 1

linha_fim = Range("A1").End(xlDown).Row
'Range seleciona uma c�lula. .End(xlDown) desce at� a �ltima linha preenchida. .Row "pega" o n�mero da linha que cont�m a informa��o'
'Se ao inv�s de .Row eu usasse .Column iria "pegar" o n�mero da coluna que a c�lula selecionada se encontra'

Range("C2:C" & linha_fim).Copy 'Seleciona a c�lula C2 at� a c�lula C(linha_fim) e faz uma c�pia'
Range("J1").PasteSpecial 'Seleciona a c�lula J1 e a partir dela cola o conte�do que foi copiado'
Application.CutCopyMode = False 'funcionalidade para "apertar" ESC para descelecionar o conte�do copiado'
ActiveSheet.Range("$J$1:$J$" & linha_fim).RemoveDuplicates Columns:=1, Header:=xlNo 'Remove os valores duplicados de uma sele��o'
linha_fim = Range("J1").End(xlDown).Row


While linha <= linha_fim
    Sheets.Add After:=ActiveSheet 'Para cada itera��o criar� uma nova aba'
    ActiveSheet.Name = Sheets("Base de Dados").Cells(linha, 10) 'ActiveSheet.Name nomeia a nova aba com o conte�do selecionado por .Cells(linha, coluna)'
    Sheets("Base de Dados").Range("A1:C1").Copy  'Entra na aba Base de Dados, seleciona as c�lulas de A1:C1 e as copia'
    ActiveSheet.Range("A1").PasteSpecial 'Seleciona a aba "Base de Dados" e cola a partir da coluna A1 o que foi copiado'
    
    linha = linha + 1 'Itera��o'
Wend

Sheets("Base de Dados").Range("J:J").Clear
linha = 2

While Sheets("Base de Dados").Cells(linha, 1) <> "" '<> significa "diferente de" e "" significa "vazio"'
    Sheets("Base de Dados").Range("A" & linha & ":C" & linha).Copy
    bairro = Sheets("Base de Dados").Cells(linha, 3)
    Sheets(bairro).Select 'Seleciona a aba com o nome do conte�do da vari�vel bairro'
    Range("A10000").End(xlUp).Offset(1, 0).PasteSpecial 'Vai para a c�lula 10000, sobe at� encontrar uma informa��o, desce uma linha'
    Application.CutCopyMode = False
    Sheets("Base de Dados").Select
    
    linha = linha + 1
Wend

For Each aba In ThisWorkbook.Sheets
    aba.Columns("A:C").AutoFit
Next

End Sub
