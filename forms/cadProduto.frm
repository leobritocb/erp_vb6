VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadProduto 
   Caption         =   "Cadastrar Produto"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10290
   OleObjectBlob   =   "cadProduto.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cadProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dirImagem_Click()
cmDialog.CancelError = True
 On Error GoTo ErrHandler

 ' Set filters
 cmDialog.Filter = "Fotos (*.jpg)|*jpg|" & _
"Bitmaps (*.bmp)|*bmp|" & _
  "Icones (*.ico)|*ico|" & _
 "All Files (*.*)|*.*| "
 
 ' Specify default filter to *.txt
 cmDialog.FilterIndex = 1

 ' Display the Open dialog box, and
 ' save the selected file in the
 ' variable strFileName
 cmDialog.ShowOpen
 strFileName = cmDialog.FileName
Me.dirImagem.Caption = cmDialog.FileName
 
ErrHandler:
 'User pressed the Cancel button


Call carregaImg(Me.imgProduto, Me.dirImagem.Caption)
End Sub

Private Sub txbPeso_Change()
    Me.txbPeso = formataPeso(Me.txbPeso)
End Sub

Private Sub txbPrecoCusto_Change()
Me.txbPrecoCusto = formataMoeda(Me.txbPrecoCusto)
End Sub

Private Sub txbPrecoCusto_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Me.txbPrecoCusto = "" Then Me.txbPrecoCusto = "0,00"
Me.txbLucroPorc = Format(margemLucro(Me.txbPrecoCusto, Me.txbPrecoVenda), "#0.00")
Me.txbLucroRS = Format(Me.txbPrecoVenda - Me.txbPrecoCusto, "#0.00")
End Sub

Private Sub txbPrecoVenda_Change()
Me.txbPrecoVenda = formataMoeda(Me.txbPrecoVenda)
End Sub

Private Sub txbPrecoVenda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Me.txbPrecoVenda = "" Then Me.txbPrecoVenda = "0,00"
Me.txbLucroPorc = Format(margemLucro(Me.txbPrecoCusto, Me.txbPrecoVenda), "#0.00")
Me.txbLucroRS = Format(Me.txbPrecoVenda - Me.txbPrecoCusto, "#0.00")
End Sub

Private Sub UserForm_initialize()
txbCodProduto = ultimoRegistroN("TB_Produtos")

Me.txbPrecoCusto = Format(0, "#0.00")
Me.txbPrecoVenda = Format(0, "#0.00")
'Me.txbCodProduto.Enabled = False
Me.txbDescricao.SetFocus
Call populaUN(Me.cbxUn)
Call populaCategorias(Me.cbxCategoria)
Call populaMarcas(Me.cbxMarca)
End Sub

Private Sub btnSair_Click()
Unload Me
End Sub

Private Sub btnSalvar_Click()
Dim tela As String
Dim texto As String
tela = "Cadastro de produtos"
texto = Me.txbCodProduto.Text + "-" + Me.txbDescricao
If Me.txbDescricao = "" Then
MsgBox "Digite a descrição do Produto!", , tela
produto.SetFocus
Exit Sub
End If
Call salvarImg(Me.imgProduto, Me.dirImagem.Caption, "cadProduto", Me.txbCodProduto)
Call salvarProduto
limpaCadProdutos
Unload Me
End Sub

Private Sub txbDescricao_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbEstoque_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = soNumeros(KeyAscii)
End Sub

