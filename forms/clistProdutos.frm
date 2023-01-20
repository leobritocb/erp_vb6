VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} clistProdutos 
   Caption         =   "Produtos"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13920
   OleObjectBlob   =   "clistProdutos.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "clistProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Private Sub nome_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
''  Me.nome.SelStart = 0
''  Me.nome.SelLength = Len(Me.nome.Text)
''End Sub

Private Sub lv_Click()
    If lv.ListItems.Count > 0 Then
        X = lv.SelectedItem.Index
        Me.txbCodProduto = lv.ListItems.item(X)
        dadosProduto
    End If
End Sub

Private Sub lv_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    
    If lv.ListItems.Count > 0 Then
        X = lv.SelectedItem.Index
        Me.txbCodProduto = lv.ListItems.item(X)
    End If
    dadosProduto
End Sub

Private Sub pesqCliente_Change()
    'Call populaProdutosPesq(lv, Me.pesqCliente.Text, Me.pb)
End Sub

Private Sub pesqCliente_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 13 Then
    Call populaProdutosPesq(lv, Me.pesqCliente.Text, Me.pb)
    KeyCode = 0

    pesqCliente.SetFocus
End If

End Sub


Private Sub UserForm_Activate()
    Call populaProdutos(lv, Me.pb)
    Me.lv.FullRowSelect = True
End Sub

Private Sub UserForm_initialize()

Me.pb.Visible = False
titulo = "Produtos"

Me.lv.ListItems.Clear
    Me.lv.ColumnHeaders.Clear
    Me.lv.View = lvwReport
    Me.lv.GridLines = True
    Me.lv.ColumnHeaders.Add , , "Cod"
    Me.lv.ColumnHeaders.Add , , "Produto"
    Me.lv.ColumnHeaders.Add , , "Estoque"
    Me.lv.ColumnHeaders(1).Width = 56
    Me.lv.ColumnHeaders(2).Width = 225
    Me.lv.ColumnHeaders(3).Width = 56
    
    Call populaCategorias(Me.cbxCategoria)
    Call populaMarcas(Me.cbxMarca)
    Call populaUN(Me.cbxUn)


End Sub



Private Sub btnNovo_Click()

cadProduto.Show vbModal
Call UserForm_Activate
End Sub
'
'Private Sub btnSalvar_Click()
''Call editarCliente("", Me.cod)
'Call UserForm_initialize
'End Sub


