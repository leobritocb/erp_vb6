Attribute VB_Name = "pagamentoCheck"
Sub iniciaPagamento()
    Dim desconto As Double
    Dim acrescimo As Double
    Dim colunas As Variant
    Dim dados As Variant
    
    limpaPagamento
    pagamento.lblSubtotalGeral.Caption = pdv.totalVenda.Caption
    
    limpaAvista
    limpaAprazo
    limpaCartao
    limpaOutros
    
    calculaPagGeral

End Sub

Sub desabilitaBtnPagamento()
    pagamento.btnAvista.Enabled = False
    pagamento.btnCartao.Enabled = False
    pagamento.btnAprazo.Enabled = False
    pagamento.btnOutros.Enabled = False
    
    pagamento.btnPagarAvista.Enabled = False
    pagamento.btnPagarCartao.Enabled = False
    pagamento.btnGerarParcelasPrazo.Enabled = False
    pagamento.btnAdiconarPagOutros.Enabled = False
    
End Sub

Sub habilitaBtnPagamento()
    
    pagamento.btnAvista.Enabled = True
    pagamento.btnCartao.Enabled = True
    pagamento.btnAprazo.Enabled = True
    pagamento.btnOutros.Enabled = True
    
    pagamento.btnPagarAvista.Enabled = True
    pagamento.btnPagarCartao.Enabled = True
    pagamento.btnGerarParcelasPrazo.Enabled = True
    pagamento.btnAdiconarPagOutros.Enabled = True
    
End Sub

Sub limpaLV(lv As Object, colunas As Variant)
    lv.ListItems.Clear
    lv.ColumnHeaders.Clear
    lv.View = lvwReport
    lv.GridLines = True
    lv.ColumnHeaders.Add , , colunas(0)
    lv.ColumnHeaders.Add , , colunas(1)
    lv.ColumnHeaders.Add , , colunas(2)
    lv.ColumnHeaders.Add , , colunas(3)
    lv.ColumnHeaders.Add , , colunas(4)
    lv.ColumnHeaders(1).Width = 25
    lv.ColumnHeaders(2).Width = 100
    lv.ColumnHeaders(3).Width = 65
    lv.ColumnHeaders(4).Width = 75
    lv.ColumnHeaders(5).Width = 70
End Sub

Sub geraParcelas()
    Dim parcela As Double
    Dim dados As Variant
    Dim Data As Date
    Dim parcelas(0 To 9) As Variant
    
    Data = pagamento.txbDT1Vencimento
    parcela = CDbl(pagamento.lblValorParceladoPrazo.Caption) / CInt(pagamento.txbNumParcelasPrazo)
    pagamento.lvPrazo.ListItems.Clear
    For i = 1 To CInt(pagamento.txbNumParcelasPrazo)
        dados = Array(i, pagamento.cbxFormaPagamentoPrazo, i, Data, Format(parcela, "#0.00"))
        Call lvPag(pagamento.lvPrazo, dados)
        
        'parcelas(0) = ultimaVenda()
        parcelas(1) = i
        parcelas(2) = Format(parcela, "#0.00")
        parcelas(3) = Data
        parcelas(4) = ""
        parcelas(5) = ""
        parcelas(6) = ""
        parcelas(7) = ""
        parcelas(8) = ""
        
        'Call salvaParcelas(parcelas)
        
        Data = DateAdd("d", pagamento.txbPrazoDias, Data)
        
    Next
    
    
End Sub


Function lvPag(lv As Object, dados As Variant)
    'Declaração de variáveis
    Dim LstItem As ListItem
    
    'Popular o ListView
    Set LstItem = lv.ListItems.Add(Text:=dados(0))
    LstItem.ListSubItems.Add Text:=dados(1)
    LstItem.ListSubItems.Add Text:=dados(2)
    LstItem.ListSubItems.Add Text:=dados(3)
    LstItem.ListSubItems.Add Text:=dados(4)
End Function

Sub efetivaAvista()
    
    If CDbl(pagamento.txbRecebidoAvista) >= CDbl(pagamento.lblAreceberGeral.Caption) Then desabilitaBtnPagamento
    
    pagamento.lblRecebidoGeral.Caption = Format(CDbl(pagamento.lblRecebidoGeral) + CDbl(pagamento.txbRecebidoAvista), "#0.00")
    calculaPagGeral
    
    Dim dados(0 To 12) As Variant

    'dados(0) = ultimaVenda()
    dados(1) = ""
    dados(2) = pagamento.lblDescontoGeral.Caption
    dados(3) = pagamento.lblAcrescimoGeral.Caption
    dados(4) = pagamento.lblSubtotalGeral.Caption
    dados(5) = pagamento.lblRecebidoGeral.Caption
    dados(6) = "A VISTA"
    dados(7) = 1
    dados(8) = ""
    dados(9) = pagamento.cbxFormaPagamentoAvista
    dados(10) = "A VISTA"
    dados(11) = "Usuario"
    dados(12) = "Obs"
    'salvaPagamento (dados)
    
    limpaAvista
    
End Sub

Sub calculaPagGeral()
    pagamento.lblTotalGeral.Caption = Format(CDbl(pagamento.lblSubtotalGeral.Caption) + CDbl(pagamento.lblAcrescimoGeral) - CDbl(pagamento.lblDescontoGeral), "#0.00")
    pagamento.lblAreceberGeral.Caption = Format(CDbl(pagamento.lblTotalGeral.Caption) - CDbl(pagamento.lblRecebidoGeral.Caption), "#0.00")
    pagamento.lblTroco.Caption = Format(CDbl(pagamento.lblRecebidoGeral.Caption) - CDbl(pagamento.lblTotalGeral.Caption), "#0.00")
    
    If CDbl(pagamento.lblTroco.Caption) < 0 Then pagamento.lblTroco.Caption = "0,00"
    If CDbl(pagamento.lblRecebidoGeral.Caption) < 0 Then pagamento.lblRecebidoGeral.Caption = "0,00"
    If CDbl(pagamento.lblAreceberGeral.Caption) < 0 Then pagamento.lblAreceberGeral.Caption = "0,00"
    
End Sub

Sub calculaParcelas()
    pagamento.cbxParcelasCartao.Clear
    For i = 1 To 12
        pagamento.cbxParcelasCartao.AddItem i & " x " & Format(CDbl(pagamento.lblTotalCartao.Caption) / i, "#0.00")
    Next
    pagamento.cbxParcelasCartao = ""
    pagamento.cbxParcelasCartao.ListIndex = 0
End Sub

Sub autorizaDesconto()
Dim senha As String
    senha = 123 'Sheets("usuario").Cells(linhaGerente("gerente"), 5)

    If desconto.txbSenhaAdmin = senha Then
    pagamento.lblDescontoGeral.Caption = Format(CDbl(pagamento.lblDescontoGeral.Caption) + CDbl(desconto.txbDesconto), "#0.00")
    Unload desconto
    End If
    calculaPagGeral
End Sub

Sub efetivaOutros()
    
    If CDbl(pagamento.txbValorOutros) <= 0 Then
    MsgBox "Digite o valor para receber", vbOKOnly, "Pagamento"
    Exit Sub
    End If
    dadosLV = Array(pagamento.lblNumPagOutros.Caption + 1, pagamento.cbxFormaPagamentoOutros, pagamento.cbxMetodo, pagamento.txbDataOutros, pagamento.txbValorOutros)
    
    pagamento.lblNumPagOutros.Caption = pagamento.lblNumPagOutros.Caption + 1
    
    Call lvPag(pagamento.lvOutros, dadosLV)
    
    If CDbl(pagamento.txbValorOutros) >= CDbl(pagamento.lblAreceberGeral.Caption) Then desabilitaBtnPagamento
    
    pagamento.lblRecebidoGeral.Caption = Format(CDbl(pagamento.lblRecebidoGeral) + CDbl(pagamento.txbValorOutros), "#0.00")
    calculaPagGeral
    Dim dados(0 To 12) As Variant

    'dados(0) = ultimaVenda()
    dados(1) = ""
    dados(2) = pagamento.lblDescontoGeral.Caption
    dados(3) = pagamento.lblAcrescimoGeral.Caption
    dados(4) = pagamento.lblSubtotalGeral.Caption
    dados(5) = pagamento.lblRecebidoGeral.Caption
    dados(6) = "OUTROS"
    dados(7) = pagamento.lblNumPagOutros
    dados(8) = ""
    dados(9) = pagamento.cbxFormaPagamentoOutros
    dados(10) = pagamento.cbxMetodo
    dados(11) = "Usuario"
    dados(12) = "Obs"
    'salvaPagamento (dados)
    
    limpaAdOutros
    
End Sub

Sub efetivaCartao()
    pagamento.lblAcrescimoGeral.Caption = Format(CDbl(pagamento.lblAcrescimoGeral) + CDbl(pagamento.txbAcrescimoCartao), "#0.00")
    
    If CDbl(pagamento.lblTotalCartao) >= CDbl(pagamento.lblAreceberGeral.Caption) Then desabilitaBtnPagamento
    pagamento.lblRecebidoGeral.Caption = Format(CDbl(pagamento.lblRecebidoGeral) + CDbl(pagamento.lblTotalCartao), "#0.00")
    calculaPagGeral
    Dim dados(0 To 12) As Variant

    'dados(0) = ultimaVenda()
    dados(1) = ""
    dados(2) = pagamento.lblDescontoGeral.Caption
    dados(3) = pagamento.lblAcrescimoGeral.Caption
    dados(4) = pagamento.lblSubtotalGeral.Caption
    dados(5) = pagamento.lblRecebidoGeral.Caption
    dados(6) = "CARTÃO"
    dados(7) = pagamento.lblNParcelas.Caption
    dados(8) = ""
    dados(9) = pagamento.cbxTipoCartao
    dados(10) = pagamento.cbxMetodo
    dados(11) = "Usuario"
    dados(12) = "Obs"
    'salvaPagamento (dados)
    limpaCartao
End Sub

Sub efetivaAprazo()
        
    pagamento.lblAcrescimoGeral.Caption = Format(CDbl(pagamento.lblAcrescimoGeral) + CDbl(pagamento.txbAcrescimoCartao), "#0.00")
    
    If CDbl(pagamento.lblTotalPrazo.Caption) >= CDbl(pagamento.lblAreceberGeral.Caption) Then desabilitaBtnPagamento
    pagamento.lblRecebidoGeral.Caption = Format(CDbl(pagamento.lblRecebidoGeral) + CDbl(pagamento.txbEntrada), "#0.00")
    calculaPagGeral
    'geraParcelas
    
    Dim dados(0 To 12) As Variant
    

    'dados(0) = ultimaVenda()
    dados(1) = pagamento.txbEntrada
    dados(2) = pagamento.lblDescontoGeral.Caption
    dados(3) = pagamento.lblAcrescimoGeral.Caption
    dados(4) = pagamento.lblSubtotalGeral.Caption
    dados(5) = pagamento.lblRecebidoGeral.Caption
    dados(6) = "A PRAZO"
    dados(7) = pagamento.txbNumParcelasPrazo
    dados(8) = pagamento.cbxFormaEntrada
    dados(9) = pagamento.cbxFormaPagamentoPrazo
    dados(10) = "A PRAZO"
    dados(11) = "Usuario"
    dados(12) = "Obs"
    
    
    'salvaPagamento (dados)
    
    limpaCartao
End Sub

Sub finalizaVenda()
    Dim dados(0 To 11) As Variant

    'dados(0) = ultimaVenda()
    dados(1) = pdv.lblCliente.Caption
    dados(2) = pdv.lblEntregador.Caption
    dados(3) = pdv.lblVendedor.Caption
    dados(4) = Date
    dados(5) = Time
    dados(6) = ""
    dados(7) = ""
    dados(8) = "PDV"
    dados(9) = "OBS"
    dados(10) = "Usuario"
    'salvaVendaPDV (dados)
End Sub




