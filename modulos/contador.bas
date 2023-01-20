Attribute VB_Name = "contador"
Public Function clientesConta()
BDCliente.Select
linha = 3
coluna = 1
cont = 1
Do Until Sheets("cliente").Cells(linha, coluna) = ""
linha = linha + 1
cont = cont + 1
Loop
cont = cont + 1
clientesConta = cont
End Function

Public Function fornConta()
BDCliente.Select
linha = 3
coluna = 24
cont = 1
Do Until Sheets("cliente").Cells(linha, coluna) = ""
linha = linha + 1
cont = cont + 1
Loop
cont = cont + 1
fornConta = cont
End Function

Public Function pedidosconta()

BDpedidos.Select
linha = 3
coluna = 1
cont = Sheets("pedidos").Cells(linha, coluna).Value
pedidosconta = cont + 1
End Function

Public Function posicaoconta()

Dim linha As Integer
BDestoque.Select
coluna = 6
Do Until Sheets("estoque").Cells(1, coluna) = ""
cont = cont + 1
coluna = coluna + 1
Loop
localProduto.n = cont

End Function

Public Function ultimoProduto()
Dim res As String
res = ""
linha = 3
coluna = 1
Do Until Sheets("estoque").Cells(linha, coluna) = ""
linha = linha + 1
Loop
If linha > 3 Then linha = linha - 1
ordenarProdutosCOD
res = Sheets("estoque").Cells(linha, coluna).Value
If res = "" Then
res = "1"
res = CInt(res)
Else: res = CInt(res) + 1
End If
ultimoProduto = res
End Function

Public Function ultimoCliente()
Dim res As String
res = ""
linha = 3
coluna = 1
Do Until Sheets("cliente").Cells(linha, coluna) = ""
linha = linha + 1
Loop
ordenarclientesCOD
If linha > 3 Then linha = linha - 1
res = Sheets("cliente").Cells(linha, coluna).Value
If res = "" Then
res = "1"
res = CInt(res)
Else: res = CInt(res) + 1
End If
ultimoCliente = res
End Function

Public Function ultimoFornecedor()
Dim res As String
res = ""
linha = 3
coluna = 24
Do Until Sheets("cliente").Cells(linha, coluna) = ""
linha = linha + 1
Loop
ordenarFornecedorCOD
If linha > 3 Then linha = linha - 1
res = Sheets("cliente").Cells(linha, coluna).Value
If res = "" Then
res = "1"
res = CInt(res)
Else: res = CInt(res) + 1
End If
ultimoFornecedor = res
End Function

Public Function ultimoEntregador()
Dim res As String
res = ""
linha = 3
coluna = 42
Do Until Sheets("cliente").Cells(linha, coluna) = ""
linha = linha + 1
Loop
If linha > 3 Then linha = linha - 1
ordenarEntregadorCOD
res = Sheets("cliente").Cells(linha, coluna).Value
If res = "" Then
res = "1"
res = CInt(res)
Else: res = CInt(res) + 1
End If
ultimoEntregador = res
End Function

Public Function contaRegistros(plan As String, linha As Integer, coluna As Integer)
cont = linha
Do Until Sheets(plan).Cells(cont, coluna) = ""
cont = cont + 1
Loop
contaRegistros = cont - linha
End Function

Public Function ultimaVenda()
Dim res As String
res = ""
ordenarVendasCOD

res = Sheets("vendas").Cells(2, 1).Value
If res = "" Then
res = "1"
res = CInt(res)
Else: res = CInt(res) + 1
End If
ultimaVenda = res
End Function
