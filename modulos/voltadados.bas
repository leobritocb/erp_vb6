Attribute VB_Name = "voltadados"
Public Function linhaGerente(cargo)

   'Declaração de variáveis
Dim wksOrigem As Worksheet
Dim rowData As Range
Dim linha As Integer
Dim coluna As Integer

coluna = 3
linha = 2
'Definição da planilha de origem
Set wksOrigem = Worksheets("usuario")
 
'Definição do range de origem
Set rowData = wksOrigem.Range("A1").CurrentRegion
 
'Alimentar variável linCont com número de linhas do intervalo fonte
 
'Alimentar variável colCont com número de linhas do intervalo fonte
'Popular os textboxes
Sheets("usuario").Select
     
    With wksOrigem
           While .Cells(linha, coluna).Value <> Empty
            valorCelula = .Cells(linha, coluna).Value
            If UCase(valorCelula) = UCase(cargo) Then
                
                linhaGerente = linha
                Exit Function
            End If
            linha = linha + 1
        Wend
    End With

End Function

Public Function linhaproduto(cod)

   'Declaração de variáveis
Dim wksOrigem As Worksheet
Dim rowData As Range
Dim linha As Integer
Dim coluna As Integer

coluna = 1
linha = 2
'Definição da planilha de origem
Set wksOrigem = Worksheets("estoque")
 
'Definição do range de origem
Set rowData = wksOrigem.Range("A1").CurrentRegion
 
'Alimentar variável linCont com número de linhas do intervalo fonte
 
'Alimentar variável colCont com número de linhas do intervalo fonte
'Popular os textboxes
Sheets("estoque").Select
     
    With wksOrigem
           While .Cells(linha, coluna).Value <> Empty
            valorCelula = .Cells(linha, coluna).Value
            If UCase(valorCelula) = UCase(cod) Then
                
                linhaproduto = linha
                Exit Function
            End If
            linha = linha + 1
        Wend
    End With

End Function

Public Function linhaCliente(cod)

   'Declaração de variáveis
Dim wksOrigem As Worksheet
Dim rowData As Range
Dim linha As Integer
Dim coluna As Integer

coluna = 1
linha = 2
'Definição da planilha de origem
Set wksOrigem = Worksheets("cliente")
 
'Definição do range de origem
Set rowData = wksOrigem.Range("A1").CurrentRegion
 
'Alimentar variável linCont com número de linhas do intervalo fonte
 
'Alimentar variável colCont com número de linhas do intervalo fonte
'Popular os textboxes
Sheets("cliente").Select
     
    With wksOrigem
           While .Cells(linha, coluna).Value <> Empty
            valorCelula = .Cells(linha, coluna).Value
            If UCase(valorCelula) = UCase(cod) Then
                
                linhaCliente = linha
            End If
            linha = linha + 1
        Wend
    End With

End Function

Sub produtoVenda()

Dim valor_pesq As String
valor_pesq = pdv.codProduto
   'Declaração de variáveis
Dim wksOrigem As Worksheet
Dim rData As Range
Dim rCell As Range
Dim LstItem As ListItem
Dim linha As Integer
Dim coluna As Integer
Dim lincont As Long
Dim colCont As Long
Dim viewCont As Long
Dim imagem As String
Dim i As Long
Dim j As Long
Dim plan As String

plan = "estoque"
coluna = 1
linha = linhaproduto(valor_pesq)

    If linha > 0 And pdv.codProduto <> "" Then
    imagem = Sheets(plan).Cells(linha, 13)
    pdv.nomeProduto.ForeColor = &H0&
    pdv.nomeProduto = Sheets(plan).Cells(linha, 3)
    pdv.vlrunProduto = Format(CDbl(Sheets(plan).Cells(linha, 11)), "#0.00")
    pdv.estoque = Sheets(plan).Cells(linha, 12)
    
    pdv.lblUn.Caption = Sheets(plan).Cells(linha, 4)
    If pdv.quantProduto = "" Then
    pdv.quantProduto = 1
    End If
    
    pdv.vlrTotalProduto = Format(pdv.quantProduto * pdv.vlrunProduto, "#0.00")
    If imagem <> "" Then
    pdv.imgProduto.Picture = LoadPicture()
    End If
    Exit Sub
    End If
    If pdv.codProduto <> "" Then
    pdv.nomeProduto.ForeColor = &HFF&
    pdv.nomeProduto = "Produto não encontrado"
    pdv.vlrunProduto = "0.00"
    pdv.estoque = 0
    pdv.quantProduto = "1"
    pdv.vlrTotalProduto = "0.00"
    'pdv.imgProduto.Picture = LoadPicture("C:\GettingTec\logo.bmp")
    End If
End Sub

Sub dadoscliente()
Dim linha As Integer
Dim cliente As String
Dim cod As String
linha = 2
cliente = novoPedido.cliente.Text
cod = novoPedido.codCliente.Text
plan = "cliente"
BDCliente.Select

Do Until Sheets(plan).Cells(linha, 1) = ""
If Sheets(plan).Cells(linha, 2) = cliente Then
novoPedido.codCliente = Sheets(plan).Cells(linha, 1)
Exit Sub

End If
linha = linha + 1
Loop

linha = 2
Do Until Sheets(plan).Cells(linha, 2) = ""
If Sheets(plan).Cells(linha, 1) = cod Then
novoPedido.cliente = Sheets(plan).Cells(linha, 2)
Exit Sub

End If
linha = linha + 1
Loop

End Sub

Sub dadoscodCliente()
Dim linha As Integer
Dim cod As String
linha = 2
cod = novoPedido.codCliente.Text
plan = "cliente"
BDCliente.Select

Do Until Sheets(plan).Cells(linha, 2) = ""
If Sheets(plan).Cells(linha, 1) = cod Then
novoPedido.cliente = Sheets(plan).Cells(linha, 2)
Exit Sub
Else: novoPedido.cliente = ""
End If
linha = linha + 1
Loop

cod = pedido.codCliente.Text
linha = 2
Do Until Sheets(plan).Cells(linha, 2) = ""
If Sheets(plan).Cells(linha, 1) = cod Then
pedido.cliente = Sheets(plan).Cells(linha, 2)
Exit Sub

End If
linha = linha + 1
Loop


End Sub

Sub dadosOrigem()
Dim linha As Integer
Dim origem As String
Dim cod As String
linha = 2
origem = novoPedido.origem.Text
plan = "combobox"
BDcombobox.Select

Do Until Sheets(plan).Cells(linha, 1) = ""
If Sheets(plan).Cells(linha, 2) = origem Then
novoPedido.codOrigem = Sheets(plan).Cells(linha, 1)
Exit Sub

End If
linha = linha + 1
Loop

End Sub

Public Function lcTabela(plan As String, Optional tabela As String)
Dim vetor As Integer

lcTabela = vetor
End Function

Public Function PopularListView(lv As Object, tabela As String, Optional txb As Object)
'Declaração de variáveis
Dim wksOrigem As Worksheet
Dim rData As Range
Dim rCell As Range
Dim LstItem As ListItem
Dim lincont As Long
Dim colCont As Long
Dim viewCont As Long
Dim colTab As Integer
Dim plan As String
Dim totalItens As Integer
Dim i As Long
Dim j As Long

totalItens = 0
If plan = "estoque" Then
plan = "estoque"
lincont = contaRegistros(plan, 2, 1)
colTab = 1
End If
If tabela = "cliente" Then
plan = "cliente"
lincont = contaRegistros(plan, 3, 1)
colTab = 1
End If
If tabela = "fornecedor" Then
plan = "cliente"
lincont = contaRegistros(plan, 3, 24)
colTab = 24
End If
If tabela = "entregador" Then
plan = "cliente"
lincont = contaRegistros(plan, 3, 42)
colTab = 42
End If
'Definição da planilha de origem
Set wksOrigem = Worksheets(plan)

'Definição do range de origem
Set rData = wksOrigem.Range("A1").CurrentRegion
'Adicionar cabeçalho no listview com laço de repetição 'For'

'Alimentar variável linCont com número de linhas do intervalo fonte

'Alimentar variável colCont com número de linhas do intervalo fonte
colCont = colTab
'Popular o ListView
For i = 3 To lincont + 2
Set LstItem = lv.ListItems.Add(Text:=rData(i, colTab).Value)
totalItens = totalItens + 1
For j = colTab To colCont
LstItem.ListSubItems.Add Text:=rData(i, j + 1).Value
Next j
Next i
txb.Text = totalItens
End Function

Public Function popularLVPesquisa(pesq As String, lv As Object, tabela As String, Optional txb As Object)
   'Declaração de variáveis
Dim wksOrigem As Worksheet
Dim rData As Range
Dim rCell As Range
Dim LstItem As ListItem
Dim linha As Integer
Dim coluna As Integer
Dim lincont As Long
Dim colCont As Long
Dim viewCont As Long
Dim plan As String
Dim colTab As Integer
Dim totalItens As Integer
Dim i As Long
Dim j As Long

totalItens = 0
If plan = "estoque" Then
plan = "estoque"
lincont = contaRegistros(plan, 2, 1)
colTab = 1
linha = 2
End If
If tabela = "cliente" Then
plan = "cliente"
lincont = contaRegistros(plan, 3, 1)
colTab = 1
linha = 3
End If
If tabela = "fornecedor" Then
plan = "cliente"
lincont = contaRegistros(plan, 3, 24)
colTab = 24
linha = 3
End If
If tabela = "entregador" Then
plan = "cliente"
lincont = contaRegistros(plan, 3, 42)
colTab = 42
linha = 3
End If
'Definição da planilha de origem
Set wksOrigem = Worksheets(plan)
 coluna = colTab
'Definição do range de origem
Set rData = wksOrigem.Range("A2").CurrentRegion
 
'Adicionar cabeçalho no listview com laço de repetição 'For'

lv.ListItems.Clear
 
'Alimentar variável colCont com número de linhas do intervalo fonte
colCont = colTab + 1
'Popular o ListView
Sheets(plan).Select
     
    With wksOrigem
           While .Cells(linha, coluna).Value <> Empty
            Valor_Celula = .Cells(linha, coluna + 1).Value
            
            If UCase(Left(Valor_Celula, Len(pesq))) = UCase(pesq) Then
                
                Set LstItem = lv.ListItems.Add(Text:=rData(linha, colTab).Value)
                totalItens = totalItens + 1
                For j = colTab To colCont
                    LstItem.ListSubItems.Add Text:=rData(linha, j + 1).Value
                Next j
            End If
            linha = linha + 1
        Wend
    End With
    txb.Text = totalItens
End Function
