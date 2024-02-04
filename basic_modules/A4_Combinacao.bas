Attribute VB_Name = "A4_Combinacao"
Sub combinacao()
Dim wkb As Workbook
Dim wks As Worksheet

Set wkb = ActiveWorkbook
Set wks = wkb.Sheets("Consolidado")


Dim linhasData As Long
linhasData = wks.Range("A1").CurrentRegion.Rows.Count

Dim linhasAGV As Long
linhasAGV = wks.Range("C1").CurrentRegion.Rows.Count

Dim linhasFV As Long
linhasFV = wks.Range("E1").CurrentRegion.Rows.Count

Dim d As Long, a As Long, f As Long, c As Long

c = 2

For d = 2 To linhasData
   For a = 2 To linhasAGV
       For f = 2 To linhasFV
           
           wks.Range("G" & c).Value = wks.Range("A" & d).Value
           wks.Range("H" & c).Value = wks.Range("C" & a).Value
           wks.Range("I" & c).Value = wks.Range("E" & f).Value
           
           c = c + 1
       Next f
   Next a
Next d

End Sub