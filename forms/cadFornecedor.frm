VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadFornecedor 
   Caption         =   "Cadastro de  Fornecedor"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8190
   OleObjectBlob   =   "cadFornecedor.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cadFornecedor"
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
tela = "Cadastro de Fornecedor"
texto = Me.txbCod.Text + "-" + Me.txbNome
If Me.txbNome = "" Then
MsgBox "Digite o nome do Fornecedor!", , tela
Me.txbNome.SetFocus
Exit Sub
End If
Call salvarFornecedor
limpaCadFornecedor
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

Private Sub optBtnJuridico_Click()
Me.lblNome = "Razão Social"
Me.lblApelido = "Nome Fantasia"
Me.lblCnpj = "CNPJ"
Me.optBtnFisico.Value = False
End Sub
'
Private Sub optBtnFisico_Click()

Me.lblNome = "Nome"
Me.lblApelido = "Apelido"
Me.lblCnpj = "CPF"
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
Me.txbCod.Text = ultimoRegistroN("TB_Fornecedor")
Me.txbData.Text = Date
Call populaUF(Me.cbxUf)
End Sub

'Private Sub btnSair_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    With btnSair
'        .BackColor = RGB(240, 240, 240)
'        .ForeColor = RGB(255, 255, 255)
'    End With
'End Sub
