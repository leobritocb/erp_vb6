VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} clistClientes 
   Caption         =   "Clientes"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13320
   OleObjectBlob   =   "clistClientes.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "clistClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_initialize()
    titulo = "Cliente"
    Me.lv.ListItems.Clear
    Me.lv.ColumnHeaders.Clear
    Me.lv.View = lvwReport
    Me.lv.GridLines = True
    Me.lv.ColumnHeaders.Add , , "Cod"
    Me.lv.ColumnHeaders.Add , , "Cliente"
    Me.lv.ColumnHeaders(1).Width = 40
    Me.lv.ColumnHeaders(2).Width = 242
   Call populaClientes(Me.lv, Me.pb)
   Call populaEstCivil(Me.cbxEstCivil)
   Call populaSexo(Me.cbxSexo)
   Call populaUF(Me.cbxUf)
    Me.lv.FullRowSelect = True
End Sub

Private Sub lv_Click()
If lv.ListItems.Count > 0 Then
X = lv.SelectedItem.Index
Me.txbCod = lv.ListItems.item(X)
dadosCliente
End If
End Sub

Private Sub txbnome_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Me.txbNome.SelStart = 0
  Me.txbNome.SelLength = Len(Me.txbNome.Text)
End Sub

Private Sub txbpesqCliente_Change()
Call popularLVPesquisa(Me.txbPesqCliente.Text, lv, "cliente", Me.txbTotalClientes)
End Sub

Private Sub optBtnJuridico_Click()
Me.lblNome = "Razão Social"
Me.lblApelido = "Nome Fantasia"
Me.lblCnpj = "CNPJ"
Me.lblNascimento = "Data de abertura"
Me.lblRG = "IE"
Me.cbxEstCivil.Visible = False
Me.cbxSexo.Visible = False
Me.Label16.Visible = False
Me.Label18.Visible = False
Me.imgEstCivil.Visible = False
Me.imgSexo.Visible = False
End Sub
'
Private Sub optBtnFisico_Click()

Me.lblNome = "Nome"
Me.lblApelido = "Apelido"
Me.lblCnpj = "CPF"
Me.lblNascimento = "Data de nascimento"
Me.lblRG = "RG"
Me.cbxEstCivil.Visible = True
Me.cbxSexo.Visible = True
Me.Label16.Visible = True
Me.Label18.Visible = True
Me.imgEstCivil.Visible = True
Me.imgSexo.Visible = True

End Sub

Private Sub btnSair_Click()
Unload Me
End Sub

Private Sub btnExcluir_Click()

'Call excluirCliente("", cod)
Call UserForm_initialize

End Sub

Private Sub btnNovo_Click()
cadCliente.Show vbModal
Call UserForm_initialize
End Sub

Private Sub btnSalvar_Click()
'Call editarCliente("", Me.cod)
Call UserForm_initialize
End Sub
