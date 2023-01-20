Attribute VB_Name = "salvaVenda"
Function salvaVendaPDV(dados As Variant)
    
    Dim plan As String
    Dim linha As Integer
    Dim coluna As Integer
    
    coluna = 1
    plan = "vendas"
    linha = ultimaLinha(plan, 1)
    
    Sheets(plan).Cells(linha, coluna) = dados(0)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(1)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(2)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(3)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(4)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(5)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(6)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(7)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(8)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(9)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(10)
    
End Function

Function salvaPagamento(dados As Variant)
    Dim plan As String
    Dim linha As Integer
    Dim coluna As Integer
    
    coluna = 1
    plan = "parcelado"
    linha = ultimaLinha(plan, 1)
    
    Sheets(plan).Cells(linha, coluna) = dados(0)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(1)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(2)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(3)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(4)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(5)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(6)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(7)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(8)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(9)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(10)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(11)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(12)
    
End Function

Function salvaParcelas(dados As Variant)
    Dim plan As String
    Dim linha As Integer
    Dim coluna As Integer
    plan = "parcelado"
    coluna = 19
    linha = ultimaLinha(plan, 19)
        
        Sheets(plan).Cells(linha, coluna) = dados(0)
        coluna = coluna + 1
        Sheets(plan).Cells(linha, coluna) = dados(1)
        coluna = coluna + 1
        Sheets(plan).Cells(linha, coluna) = dados(2)
        coluna = coluna + 1
        Sheets(plan).Cells(linha, coluna) = dados(3)
        coluna = coluna + 1
        Sheets(plan).Cells(linha, coluna) = dados(4)
        coluna = coluna + 1
        Sheets(plan).Cells(linha, coluna) = dados(5)
        coluna = coluna + 1
        Sheets(plan).Cells(linha, coluna) = dados(6)
        coluna = coluna + 1
        Sheets(plan).Cells(linha, coluna) = dados(7)
        coluna = coluna + 1
        Sheets(plan).Cells(linha, coluna) = dados(8)
        
        linha = linha + 1
    
    
End Function

Function salvaProdutosVenda(dados As Variant)
    Dim plan As String
    Dim linha As Integer
    Dim coluna As Integer
    
    coluna = 1
    plan = "vendaProdutos"
    linha = ultimaLinha(plan, 1)
    
    Sheets(plan).Cells(linha, coluna) = dados(0)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(1)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(2)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(3)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(4)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(5)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(6)
    coluna = coluna + 1
    Sheets(plan).Cells(linha, coluna) = dados(7)
    
End Function

