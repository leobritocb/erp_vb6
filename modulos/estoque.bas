Attribute VB_Name = "estoque"
Sub baixaEstoque(tela, ID, quant)

Dim baixar As Long
Dim estoque As Long

Dim rs As DAO.Recordset
 
conectaBD
Set rs = dbCon.OpenRecordset("SELECT ESTOQUE FROM TB_Produtos where ID = " & ID)


estoque = rs.Fields("Estoque")

While Not rs.EOF
rs.MoveNext 'Move-se para o próximo registro do recordset
Wend

baixar = estoque - quant
dbCon.Execute ("UPDATE TB_Produtos SET estoque= " & baixar & " WHERE ID=" & ID)
   rs.Close

End Sub

Sub sobeEstoque(tela, ID, quant)

Dim subir As Long
Dim estoque As Long

Dim rs As DAO.Recordset
 
conectaBD
Set rs = dbCon.OpenRecordset("SELECT ESTOQUE FROM TB_Produtos where ID = " & ID)


estoque = rs.Fields("Estoque")

While Not rs.EOF
rs.MoveNext 'Move-se para o próximo registro do recordset
Wend

subir = estoque + quant
dbCon.Execute ("UPDATE TB_Produtos SET estoque= " & subir & " WHERE ID=" & ID)
   rs.Close

End Sub

Sub kardex(cod, movimento, op, quant, estoque, tela, Optional obs = "")

Dim strKardex As String
    Dim dadosBD(1 To 10) As Variant
    Dim dadosForm(1 To 11) As Variant
    Dim i As Integer
    
    dadosForm(1) = cod
    dadosForm(2) = movimento
    If op = "+" Then
              dadosForm(3) = "Entrada"
            End If
            If op = "-" Then
              dadosForm(3) = "Saida"
            End If
    
    dadosForm(4) = CInt(quant)
    dadosForm(5) = CInt(estoque)
    
    If op = "+" Then
        dadosForm(6) = estoque + quant
    End If
    If op = "-" Then
        dadosForm(6) = estoque - quant
    End If
            
    dadosForm(7) = Date
    dadosForm(8) = Time
    dadosForm(9) = "1"
    dadosForm(10) = tela
    If obs <> "" Then
    dadosForm(11) = obs
    Else: dadosForm(11) = "null"
    End If
    
    strKardex = "INSERT INTO TB_kardex " & _
        "VALUES('" & _
          dadosForm(1) & "', '" & _
          dadosForm(2) & "', '" & _
          dadosForm(3) & "'," & _
          dadosForm(4) & "," & _
          dadosForm(5) & "," & _
          dadosForm(6) & ", '" & _
          dadosForm(7) & "', '" & _
          dadosForm(8) & "', " & _
          dadosForm(9) & ", '" & _
          dadosForm(10) & "', '" & _
          dadosForm(11) & "')"

    conectaBD
        dbCon.Execute (strKardex)
    encerraBD

End Sub
