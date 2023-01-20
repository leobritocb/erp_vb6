VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} descontoItem 
   Caption         =   "Desconto"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   OleObjectBlob   =   "descontoItem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "descontoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim descontoRS As Double

Private Sub btnAplicarDesconto_Click()
'autorizaDescontoItem

End Sub

Private Sub UserForm_initialize()
    Me.optValorDesconto.Value = True
    Me.lblValorProduto.Caption = "15,99"
End Sub

Private Sub optPorcentoDesconto_Click()
Me.txbDesconto = Format(0, "#0.00")
Me.lblValorcomDesconto.Caption = Format(CDbl(Me.lblValorProduto.Caption) - CDbl(Me.txbDesconto), "#0.00")
End Sub

Private Sub optValorDesconto_Click()

Me.txbDesconto = Format(0, "#0.00")
Me.lblValorcomDesconto.Caption = Format(CDbl(Me.lblValorProduto.Caption) - CDbl(Me.txbDesconto), "#0.00")

End Sub

Private Sub txbDesconto_Change()
If Me.optValorDesconto.Value = True Then
Me.txbDesconto = formataMoeda(Me.txbDesconto)
Else: Me.txbDesconto = formataPorcento(Me.txbDesconto)
End If
End Sub

Private Sub txbDesconto_Exit(ByVal Cancel As MSForms.ReturnBoolean)

If Me.optPorcentoDesconto.Value = True Then
descontoRS = calculaPorcentagem(Me.lblValorProduto.Caption, Me.txbDesconto)
Me.optValorDesconto.Value = True

Me.txbDesconto = Format(descontoRS, "0.00")
End If

Me.lblValorcomDesconto.Caption = Format(CDbl(Me.lblValorProduto.Caption) - CDbl(Me.txbDesconto), "#0.00")

End Sub
