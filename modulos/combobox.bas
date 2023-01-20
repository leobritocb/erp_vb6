Attribute VB_Name = "combobox"
Sub popularCombobox(Optional cbxOrigem As Object, Optional cbxFormaPag As Object, Optional cbxUf As Object, Optional cbxUn As Object, Optional cbxCat As Object, Optional cbxMarca As Object, Optional cbxSexo As Object, Optional cbxEstCivil As Object)

cbxOrigem.Clear
cbxFormaPag.Clear
cbxUf.Clear
cbxUn.Clear
cbxCat.Clear
cbxMarca.Clear
cbxSexo.Clear
cbxEstCivil.Clear

linha = 3
plan = "combobox"
coluna = 2

Do Until Sheets(plan).Cells(linha, coluna) = ""
cbxOrigem.AddItem Sheets(plan).Cells(linha, coluna)
linha = linha + 1
Loop

linha = 3
coluna = 4

Do Until Sheets(plan).Cells(linha, coluna) = ""
cbxFormaPag.AddItem Sheets(plan).Cells(linha, coluna)
linha = linha + 1
Loop

linha = 3
coluna = 6

Do Until Sheets(plan).Cells(linha, coluna) = ""
cbxUf.AddItem Sheets(plan).Cells(linha, coluna)
linha = linha + 1
Loop

linha = 3
coluna = 8

Do Until Sheets(plan).Cells(linha, coluna) = ""
cbxUn.AddItem Sheets(plan).Cells(linha, coluna)
linha = linha + 1
Loop

linha = 3
coluna = 10

Do Until Sheets(plan).Cells(linha, coluna) = ""
cbxCat.AddItem Sheets(plan).Cells(linha, coluna)
linha = linha + 1
Loop

linha = 3
coluna = 12

Do Until Sheets(plan).Cells(linha, coluna) = ""
cbxMarca.AddItem Sheets(plan).Cells(linha, coluna)
linha = linha + 1
Loop

linha = 3
coluna = 14

Do Until Sheets(plan).Cells(linha, coluna) = ""
cbxSexo.AddItem Sheets(plan).Cells(linha, coluna)
linha = linha + 1
Loop

linha = 3
coluna = 16

Do Until Sheets(plan).Cells(linha, coluna) = ""
cbxEstCivil.AddItem Sheets(plan).Cells(linha, coluna)
linha = linha + 1
Loop
End Sub

Sub popularCbxUF(cbxUf As Object)

cbxUf.Clear
linha = 3
plan = "combobox"
coluna = 6

Do Until Sheets(plan).Cells(linha, coluna) = ""
cbxUf.AddItem Sheets(plan).Cells(linha, coluna)
linha = linha + 1
Loop
End Sub

Sub popularCbxFormaPagamento(cbx As Object, formaPag As Integer)

cbx.Clear
linha = 3
plan = "combobox"
coluna = 4

Do Until Sheets(plan).Cells(linha, coluna) = ""
If Sheets(plan).Cells(linha, coluna - 1) = formaPag Then
cbx.AddItem Sheets(plan).Cells(linha, coluna)
End If
If formaPag = 0 Then
cbx.AddItem Sheets(plan).Cells(linha, coluna)
End If
linha = linha + 1
Loop
End Sub

'Sub origemcombo()
'pedido.origem.Clear
'linha = 3
'Do Until Sheets(plan).Cells(linha, coluna) = ""
'pedido.origem.AddItem Sheets(plan).Cells(linha, coluna)
'linha = linha + 1
'Loop
'Entrada.origem.Clear
'linha = 3
'coluna = 8
'Do Until Sheets(plan).Cells(linha, coluna) = ""
'Entrada.origem.AddItem Sheets(plan).Cells(linha, coluna)
'linha = linha + 1
'Loop
'End Sub
'
'Sub statuscombo()
'
'pedido.status.Clear
'linha = 3
'Do Until Sheets(plan).Cells(linha, coluna) = ""
'pedido.status.AddItem Sheets(plan).Cells(linha, coluna)
'linha = linha + 1
'Loop
'End Sub
'
'
'Sub especiecombo()
'Dim linha As Integer
'linha = 3
'plan = "combobox"
'coluna = 10
'Entrada.especie.Clear
'Do Until Sheets(plan).Cells(linha, coluna) = ""
'Entrada.especie.AddItem Sheets(plan).Cells(linha, coluna)
'linha = linha + 1
'Loop
'
'linha = 3
'pedido.especie.Clear
'Do Until Sheets(plan).Cells(linha, coluna) = ""
'pedido.especie.AddItem Sheets(plan).Cells(linha, coluna)
'linha = linha + 1
'Loop
'
'End Sub
'
'
Public Function ufCombobox()




End Function

