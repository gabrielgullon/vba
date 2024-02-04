Attribute VB_Name = "A3_Collections"

'Testa se uma determinada chave (key) existe em uma Collection (col)
'Se for uma Collection de objetos, ent�o vObject = True
Public Function Collection_Test_Key(col As Collection, key As Variant, Optional vObject As Boolean = False) As Boolean
    Dim obj As Variant
    Dim obj2 As Object
    On Error GoTo Err
    Collection_Test_Key = True
    If vObject Then
        Set obj2 = col(key)
        Set obj2 = Nothing
    Else
        obj = col(key)
    End If
    Exit Function
Err:
    Collection_Test_Key = False
End Function

Sub PROCV()
    Dim base As Variant
    Dim Output As Variant
    Dim colBusca As New Collection
    Dim colAux As Collection
    
    Dim orig As Worksheet, dest As Worksheet
    
    Set orig = ThisWorkbook.Sheets("LT")
    Set dest = ThisWorkbook.Sheets("BASE_NS")
    
    
    'Coloca a base do Excel para dentro de uma Matriz em VBA
    ultima_linha = orig.Range("A1").CurrentRegion.Rows.Count
    base = orig.Range("I1:J" & ultima_linha).Value
    
   
    'Percorre base, na dimens�o 1 (linhas)
    For i = 2 To UBound(base, 1)
        'Pega o campo da base, que servir� de chave => base(linha, coluna)
        
        chave = CStr(base(i, 1))
        
        'Set colAux = New Collection
        
        valor = base(i, 2)
                
        'testa se chave n�o existe na collection
        If Not Collection_Test_Key(colBusca, chave) Then
            colBusca.Add valor, chave
        End If
            
    Next i
    
    ultima_linha = dest.Range("A1").CurrentRegion.Rows.Count
    'Coloca a base do Excel para dentro de uma Matriz em VBA
    
    base = dest.Range("FT3:FT" & ultima_linha).Value
    
    'Cria uma Matriz para guardar o retorno da fun��o
    ReDim Output(1 To UBound(base, 1), 1 To 1)
    
    For i = 1 To UBound(base, 1)
        'Pega o campo da base, que servir� de chave => base(linha, coluna)
        chave = CStr(base(i, 1))
        
        'testa se chave n�o existe na collection
        If Collection_Test_Key(colBusca, chave) Then
            Output(i, 1) = colBusca(chave)
        End If
            
    Next i
    
    'Joga valor da Matriz em um Range Excel
    dest.Range("FV3:FV" & ultima_linha).Value = Output

End Sub

Sub SOMASE()
    Dim base As Variant
    Dim Output As Variant
    Dim colBusca As New Collection
    Dim colAux As Collection
    
    
    'Coloca a base do Excel para dentro de uma Matriz em VBA
    ultima_linha = Sheets("Meta").Range("A1").CurrentRegion.Rows.Count
    base = Sheets("Meta").Range("A1:K" & ultima_linha).Value
    
   
    'Percorre base, na dimens�o 1 (linhas)
    For i = 2 To UBound(base, 1)
        'Pega o campo da base, que servir� de chave => base(linha, coluna)
        
        chave = CStr(base(i, 1) & base(i, 5))
        
        'Set colAux = New Collection
        
        qtd = base(i, 7)
                
        'testa se chave n�o existe na collection
        If Not Collection_Test_Key(colBusca, chave) Then
            colBusca.Add qtd, chave
        Else
            qtd = qtd + colBusca(chave)
            colBusca.Remove (chave)
            colBusca.Add qtd, chave
        End If
            
    Next i
    
    ultima_linha = Sheets("Meta").Range("A1").CurrentRegion.Rows.Count
    'Coloca a base do Excel para dentro de uma Matriz em VBA
    base = Range("A2:e" & ultima_linha).Value
    'Cria uma Matriz para guardar o retorno da fun��o
    ReDim Output(1 To UBound(base, 1), 1 To 1)
    
    For i = 1 To UBound(base, 1)
        'Pega o campo da base, que servir� de chave => base(linha, coluna)
        chave = CStr(base(i, 1) & base(i, 5))
        
        'testa se chave n�o existe na collection
        If Collection_Test_Key(colBusca, chave) Then
            Output(i, 1) = colBusca(chave)
        End If
            
    Next i
    
    'Joga valor da Matriz em um Range Excel
    Sheets("Meta").Range("i2:i" & ultima_linha).Value = Output

End Sub

Sub mediaPonderada()
    Dim base As Variant
    Dim Output As Variant
    Dim colBusca As New Collection
    Dim colAux As Collection
    
    
    'Coloca a base do Excel para dentro de uma Matriz em VBA
    ultima_linha = Sheets("Plan1").Range("X1").CurrentRegion.Rows.Count
    base = Sheets("Plan1").Range("B1:X" & ultima_linha).Value
    
   
    'Percorre base, na dimens�o 1 (linhas)
    For i = 2 To UBound(base, 1)
        'Pega o campo da base, que servir� de chave => base(linha, coluna)
        chave = CStr(base(i, 1))
        
        Set colAux = New Collection
        
  '      qtd = base(i, 19)
'        vol = base(i, 20)
        fDesp = base(i, 21)
       fLiq = base(i, 22)
    '    fCta = base(i, 23)
        
        planejados = fDesp
                
        'testa se chave n�o existe na collection
        If Not Collection_Test_Key(colBusca, chave) Then
            colBusca.Add planejados, chave
        Else
            planejados = planejados + colBusca(chave)
            colBusca.Remove (chave)
            colBusca.Add planejados, chave
        End If
            
    Next i
    
    ultima_linha = Sheets("Tr&Cliente").Range("B1").CurrentRegion.Rows.Count
    'Coloca a base do Excel para dentro de uma Matriz em VBA
    base = Range("B2:B" & ultima_linha).Value
    'Cria uma Matriz para guardar o retorno da fun��o
    ReDim Output(1 To UBound(base, 1), 1 To 1)
    
    For i = 1 To UBound(base, 1)
        'Pega o campo da base, que servir� de chave => base(linha, coluna)
        chave = CStr(base(i, 1))
        
        'testa se chave n�o existe na collection
        If Collection_Test_Key(colBusca, chave) Then
            Output(i, 1) = colBusca(chave)
        End If
            
    Next i
    
    'Joga valor da Matriz em um Range Excel
    Sheets("Tr&Cliente").Range("v2:v" & ultima_linha).Value = Output
    
End Sub


Sub CONTSE()
    Dim base As Variant
    Dim Output As Variant
    Dim colBusca As New Collection
    Dim colAux As Collection
    
    
    'Coloca a base do Excel para dentro de uma Matriz em VBA
    ultima_linha = Sheets("Consolidado").Range("A1").CurrentRegion.Rows.Count
    base = Sheets("Consolidado").Range("A1:R" & ultima_linha).Value
    
   
    'Percorre base, na dimens�o 1 (linhas)
    For i = 2 To UBound(base, 1)
        'Pega o campo da base, que servir� de chave => base(linha, coluna)
        chave = CStr(base(i, 18))
                
        'testa se chave n�o existe na collection
        If Not Collection_Test_Key(colBusca, chave) Then
            colBusca.Add 1, chave
        Else
            planejados = 1 + colBusca(chave)
            colBusca.Remove (chave)
            colBusca.Add planejados, chave
        End If
            
    Next i
    
    ultima_linha = Sheets("Receb Cli Dia").Range("A1").CurrentRegion.Rows.Count
    'Coloca a base do Excel para dentro de uma Matriz em VBA
    base = Range("A2:A" & ultima_linha).Value
    'Cria uma Matriz para guardar o retorno da fun��o
    ReDim Output(1 To UBound(base, 1), 1 To 1)
    
    For i = 1 To UBound(base, 1)
        'Pega o campo da base, que servir� de chave => base(linha, coluna)
        chave = CStr(base(i, 1))
        
        'testa se chave n�o existe na collection
        If Collection_Test_Key(colBusca, chave) Then
            Output(i, 1) = colBusca(chave)
        End If
            
    Next i
    
    'Joga valor da Matriz em um Range Excel
    Sheets("Receb Cli Dia").Range("E2:E" & ultima_linha).Value = Output

End Sub