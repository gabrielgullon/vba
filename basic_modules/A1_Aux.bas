'namespace=basic_modules/A1_Aux/
Attribute VB_Name = "A1_Aux"

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveCell.FormulaR1C1 = "Gabriel Gullon Gonzalez"
    Range("A1").Select
End Sub

Sub completaTabDin()

Dim ultima_linha As Long
ultima_linha = ActiveSheet.Range("L3").CurrentRegion.Rows.Count

Dim cell As Range
lastRowConsolidado = planCons.Cells(Rows.Count, 1).End(xlUp).Row

For Each cell In Range("i3:i" & ultima_linha)
    If cell = "" Then
        cell.Value = cell.Offset(-1, 0).Value
    End If
Next

End Sub


Option Explicit

' Na cel A2
'=SUBSTITUIR(RemoveAcentos(SUBSTITUIR(SUBSTITUIR(A2;" ";);"-";));ESQUERDA(A2;3);"")

Function RemoveAcentos(sString As String) As String
     
    Dim sAcentos As String
    Dim sSemAcentos As String
    Dim sTemp As String
    Dim i As Long
  
    'Liste nesta vari�vel todos os acentos poss�veis
    sAcentos = "�����������������������������������������������"
      
    'Letras sem acentua��o correspondentes para substitui��o
    sSemAcentos = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
      
    'Armazena em sTemp a string recebida
    sTemp = sString
      
    'Loop que percorrer� todas as letras da vari�vel 'sAcentos',
    'subtituindo pelo caractere correspondente em 'sSemAcentos'
    For i = 1 To Len(sAcentos)
        sTemp = Replace(sTemp, Mid(sAcentos, i, 1), Mid(sSemAcentos, i, 1))
    Next i
      
    'Retorna a nova string
    RemoveAcentos = sTemp
      
End Function


Sub auxiliar()
Dim plan As Worksheet
Set plan = ActiveWorkbook.Sheets("Planilha4")
Dim linout As Long

lin = plan.Range("A1").CurrentRegion.Rows.Count
dat = plan.Range("C1").Value

For i = 2 To lin
'    A = plan.Range("A" & i).Value
'    B = plan.Range("B" & i).Value
'    C = plan.Range("C" & i).Value
    
    a = "#" & plan.Range("A" & i).Value
    B = dat & "G_" & plan.Range("B" & i).Value
    c = "-"
    
    linout = plan.Range("G1").CurrentRegion.Rows.Count + 1
    plan.Range("G" & linout).Value = a
    linout = plan.Range("G1").CurrentRegion.Rows.Count + 1
    plan.Range("G" & linout).Value = B
    linout = plan.Range("G1").CurrentRegion.Rows.Count + 1
    plan.Range("G" & linout).Value = c
    'plan.Range("G" & linout).Value = dat & "Valid_" & B & " = " & dat & B
    
Next i


End Sub


Sub juntaStrings()
Dim cell As Range
Dim str As String
Dim Rng As Range

str = ""

Set Rng = Selection

For Each cell In Rng
    str = str & ";" & cell.Value
Next

Range("B1").Value = str

End Sub


Sub separaStringACada()

Dim cell As Range
Dim tamanhoCell As Integer

Dim Rng As Range
    
    Set Rng = Selection
    
    For Each cell In Rng
        If Not IsEmpty(cell) Then
            Cells(cell.Row, cell.Column + 2).Value _
            = Replace(Trim(cell), vbLf, ";")
        End If
    Next

End Sub

Function splitLineBreaks(ByVal str As String) As String()
    str = Replace(str, vbCrLf, vbCr)
    str = Replace(str, vbLf, vbCr)

    splitLineBreaks = Split(str, vbCr)
    
End Function