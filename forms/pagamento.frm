VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pagamento 
   Caption         =   "Pagamento"
   ClientHeight    =   9345.001
   ClientLeft      =   240
   ClientTop       =   390
   ClientWidth     =   15360
   OleObjectBlob   =   "pagamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "pagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim caminho As String
Dim acrescimoRS As String

Private Sub btnAvista_Click()
Me.mpPagamento.Value = 0

Me.imgTipoVenda.Picture = LoadPicture(caminho & "\barraverde.bmp")
Me.lblTipoPag.ForeColor = &HFFFFFF
Me.lblTipoPag.Caption = "VENDA A VISTA"
Me.imgBtnFoco.Visible = True
Me.imgBtnFoco.Top = 76

Me.cbxFormaPagamentoAvista.Value = "DINHEIRO"

Me.txbRecebidoAvista.SetFocus

limpaAvista

End Sub

Private Sub btnAPrazo_Click()
Me.mpPagamento.Value = 1

Me.lvPrazo.Top = 228
Me.imgTipoVenda.Picture = LoadPicture(caminho & "\barraamarela.bmp")
Me.lblTipoPag.ForeColor = &H0&
Me.lblTipoPag.Caption = "VENDA A PRAZO"
Me.imgBtnFoco.Visible = True
Me.imgBtnFoco.Top = 231

limpaAprazo
End Sub

Private Sub btnCartao_Click()
Me.mpPagamento.Value = 2

Me.cbxMetodo.AddItem "A VISTA"
Me.cbxMetodo.AddItem "A PRAZO"

Me.imgTipoVenda.Picture = LoadPicture(caminho & "\barraazulescuro.bmp")
Me.lblTipoPag.ForeColor = &HFFFFFF
Me.lblTipoPag.Caption = "VENDA A CRÉDITO"
Me.imgBtnFoco.Visible = True
Me.imgBtnFoco.Top = 153

limpaCartao

End Sub

Private Sub btnOutros_Click()
Me.mpPagamento.Value = 3

Me.imgTipoVenda.Picture = LoadPicture(caminho & "\barracinza.bmp")
Me.lblTipoPag.ForeColor = &H0&
Me.lblTipoPag.Caption = "OUTROS MEIOS DE PAGAMENTO"
Me.imgBtnFoco.Visible = True
Me.imgBtnFoco.Top = 309

limpaOutros
End Sub

'================================================================================
                        'Atalhos de teclhado
'================================================================================





'================================================================================
                        'A VISTA
'================================================================================
Private Sub btnPagarAvista_Click()
'Call lvPag(Me.lvPagamentoGeral, 1)
    Call efetivaAvista
End Sub

Private Sub cbxParcelasCartao_Click()
 Me.lblNParcelas.Caption = Me.cbxParcelasCartao.ListIndex + 1
End Sub

Private Sub CommandButton11_Click()
desconto.Show
calculaPagGeral

End Sub

Private Sub lblDescontoGeral_Click()
Me.lblDescontoGeral = Format(0, "#0.00")
calculaPagGeral
limpaAvista

End Sub

Private Sub txbRecebidoAvista_Change()
Me.txbRecebidoAvista = formataMoeda(Me.txbRecebidoAvista)
End Sub



'================================================================================
                        'A PRAZO
'================================================================================

Private Sub txbEntrada_Change()

Me.txbEntrada = formataMoeda(Me.txbEntrada)
Me.lblValorParceladoPrazo.Caption = Format(CDbl(Me.lblTotalPrazo.Caption) - CDbl(Me.txbEntrada), "#0.00")
End Sub

Private Sub txbAcrescimoPrazo_Change()

If Me.optValorPrazo.Value = True Then
Me.txbAcrescimoPrazo = formataMoeda(Me.txbAcrescimoPrazo)
Me.lblValorParceladoPrazo.Caption = Format(CDbl(Me.lblTotalPrazo.Caption) - CDbl(Me.txbEntrada) + CDbl(Me.txbAcrescimoPrazo), "#0.00")
Else:
Me.txbAcrescimoPrazo = formataPorcento(Me.txbAcrescimoPrazo)
acrescimoRS = calculaPorcentagem(Me.lblTotalPrazo.Caption, Me.txbAcrescimoPrazo)

Me.lblValorParceladoPrazo.Caption = Format(CDbl(Me.lblTotalPrazo.Caption) - CDbl(Me.txbEntrada) + acrescimoRS, "0.00")
End If

End Sub

Private Sub optPorcentoPrazo_Click()
Me.txbAcrescimoPrazo = "0,00"
Me.lblTotalPrazo.Caption = Format(CDbl(pagamento.lblAreceberGeral.Caption) + CDbl(pagamento.txbAcrescimoPrazo), "#0.00")
Me.lblValorParceladoPrazo.Caption = Format(CDbl(pagamento.lblAreceberGeral.Caption) + CDbl(pagamento.txbAcrescimoPrazo) - CDbl(pagamento.txbEntrada), "#0.00")

End Sub

Private Sub optValorPrazo_Click()
Me.txbAcrescimoPrazo = "0,00"
Me.lblTotalPrazo.Caption = Format(CDbl(pagamento.lblAreceberGeral.Caption) + CDbl(pagamento.txbAcrescimoPrazo), "#0.00")
calculaParcelas
End Sub

Private Sub txbAcrescimoPrazo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Me.optPorcentoPrazo.Value = True Then
acrescimoRS = calculaPorcentagem(Me.lblTotalPrazo.Caption, Me.txbAcrescimoPrazo)
Me.optValorPrazo.Value = True

Me.txbAcrescimoPrazo = Format(acrescimoRS, "0.00")
End If

Me.lblTotalPrazo = Format(CDbl(pagamento.lblAreceberGeral.Caption) + CDbl(Me.txbAcrescimoPrazo), "#0.00")

End Sub

Private Sub btnGerarParcelasPrazo_Click()
    Call efetivaAprazo
End Sub



'================================================================================
                        'CARTAO
'================================================================================
Private Sub txbAcrescimoCartao_Change()

If Me.optValorCartao.Value = True Then
Me.txbAcrescimoCartao = formataMoeda(Me.txbAcrescimoCartao)
Else: Me.txbAcrescimoCartao = formataPorcento(Me.txbAcrescimoCartao)
End If
End Sub

Private Sub optPorcentoCartao_Click()
Me.txbAcrescimoCartao = Format(0, "#0.00")
Me.lblTotalCartao.Caption = Format(CDbl(pagamento.lblAreceberGeral.Caption) + CDbl(pagamento.txbAcrescimoCartao), "#0.00")
calculaParcelas
End Sub

Private Sub optValorCartao_Click()
Me.txbAcrescimoCartao = Format(0, "#0.00")
Me.lblTotalCartao.Caption = Format(CDbl(pagamento.lblAreceberGeral.Caption) + CDbl(pagamento.txbAcrescimoCartao), "#0.00")
calculaParcelas
End Sub

Private Sub txbAcrescimoCartao_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Me.optPorcentoCartao.Value = True Then
acrescimoRS = calculaPorcentagem(Me.lblTotalCartao.Caption, Me.txbAcrescimoCartao)
Me.optValorCartao.Value = True

Me.txbAcrescimoCartao = Format(acrescimoRS, "0.00")
End If

Me.lblTotalCartao = Format(CDbl(pagamento.lblAreceberGeral.Caption) + CDbl(Me.txbAcrescimoCartao), "#0.00")
calculaParcelas
End Sub

Private Sub btnPagarCartao_Click()
   Call efetivaCartao
End Sub
'================================================================================
                        'Outros
'================================================================================

Private Sub txbValorOutros_Change()
    Me.txbValorOutros = formataMoeda(Me.txbValorOutros)
End Sub


Private Sub btnAdiconarPagOutros_Click()
    efetivaOutros
End Sub





'================================================================================
                        'Inicializar Form
'================================================================================


Private Sub UserForm_initialize()

caminho = retornaDiretorio() & "\setup"
Me.imgBtnFoco.Visible = False

Me.lvOutros.ListItems.Clear
Me.lvPrazo.ListItems.Clear

Call popularCbxFormaPagamento(Me.cbxFormaPagamentoPrazo, 2)
Call popularCbxFormaPagamento(Me.cbxFormaEntrada, 5)
Call popularCbxFormaPagamento(Me.cbxFormaPagamentoOutros, 4)
Call popularCbxFormaPagamento(Me.cbxFormaPagamentoAvista, 1)
Call popularCbxFormaPagamento(Me.cbxTipoCartao, 3)


habilitaBtnPagamento
Call iniciaPagamento
Me.lblSubtotalGeral = "120,00"
calculaPagGeral
Call btnAvista_Click
Call btnAPrazo_Click
Call btnCartao_Click
Call btnOutros_Click
Call btnAvista_Click

End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = vbKeyF2 Then
formQuantidade.Show
End If

If KeyCode = vbKeyF1 Then
desconto.Show
End If
End Sub
