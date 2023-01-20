VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de  clientes"
   ClientHeight    =   9090
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   8175
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "cadCliente.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cadCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub btnSair_Click()
'Unload Me
'End Sub
''
''Private Sub frame4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
''    With btnSair
''        .BackColor = RGB(240, 240, 240)
''        .ForeColor = RGB(0, 0, 0)
''    End With
''
''End Sub

Private Sub btnSalvar_Click()
Dim tela As String
Dim texto As String
tela = "Cadastro de clientes"
texto = Me.txbCod.Text + "-" + Me.txbNome
If Me.txbNome = "" Then
MsgBox "Digite o nome do cliente!", , tela
Me.txbNome.SetFocus
Exit Sub
End If
Call salvarCliente
limpaCadCliente
End Sub

Private Sub cbxUf_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.lblCodUf.Caption = codUF(Me.cbxUf)
End Sub

Private Sub txbApelido_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbBairro_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbCelular_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Me.txbCelular.Text = formataCelular(KeyAscii, Me.txbCelular.Text)
End Sub

Private Sub txbCidade_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbComplemento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbDtNascimento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Me.txbDtNascimento.Text = formataData(KeyAscii, Me.txbDtNascimento.Text)
End Sub

Private Sub txblimiteCredito_Change()
Me.txbLimiteCredito.Value = formataMoeda(Me.txbLimiteCredito.Value)
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
Me.optBtnFisico.Value = False
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
Me.optBtnJuridico.Value = False

End Sub

Private Sub txbNome_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'
'Private Sub userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    With btnSair
' '       .BackColor = RGB(255, 255, 255)
' '      .ForeColor = RGB(0, 0, 0)
'    End With
'
'End Sub
'
Private Sub txbnumeroend_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = soNumeros(KeyAscii)
End Sub

Private Sub txbcep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Me.txbCep.Text = formataCEP(KeyAscii, Me.txbCep.Text)
End Sub
'
Private Sub txbcpf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If Me.optBtnFisico.Value = True Then
Me.txbCpf = formataCPF(KeyAscii, Me.txbCpf.Text)
End If
If Me.optBtnJuridico.Value = True Then
Me.txbCpf = formataCNPJ(KeyAscii, Me.txbCpf.Text)
End If
End Sub

Private Sub txbRua_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbtelefone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Me.txbTelefone.Text = formataTelefone(KeyAscii, Me.txbTelefone.Text)
End Sub

Private Sub UserForm_initialize()
Me.optBtnFisico.Value = True
Me.txbCod.Text = ultimoRegistroN("TB_Clientes")
Me.txbData.Text = Date
Call populaUF(Me.cbxUf)
Call populaEstCivil(Me.cbxEstCivil)
Call populaSexo(Me.cbxSexo)
End Sub

'Private Sub btnSair_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    With btnSair
'        .BackColor = RGB(240, 240, 240)
'        .ForeColor = RGB(255, 255, 255)
'    End With
'End Sub
