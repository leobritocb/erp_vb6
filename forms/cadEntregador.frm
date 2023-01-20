VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cadEntregador 
   Caption         =   "Cadastro de  Entregador"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7935
   OleObjectBlob   =   "cadEntregador.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cadEntregador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalvar_Click()
Dim tela As String
Dim texto As String
tela = "Cadastro de Entregador"
texto = Me.txbCod.Text + "-" + Me.txbNome
If Me.txbNome = "" Then
MsgBox "Digite o nome do cliente!", , tela
Me.txbNome.SetFocus
Exit Sub
End If
Call salvarEntregador(tela, texto)
limpaCadEntregador
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

Private Sub txbMarca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbModelo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbNome_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbnumeroend_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = soNumeros(KeyAscii)
End Sub

Private Sub txbcep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Me.txbCep.Text = formataCEP(KeyAscii, Me.txbCep.Text)
End Sub
'
Private Sub txbcpf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Me.txbCpf.Text = formataCPF(KeyAscii, Me.txbCpf.Text)
End Sub

Private Sub txbPlaca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Me.txbPlaca.Text = formataPlaca(KeyAscii, Me.txbPlaca.Text)
End Sub

Private Sub txbRua_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txbtelefone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Me.txbTelefone.Text = formataTelefone(KeyAscii, Me.txbTelefone.Text)
End Sub

Private Sub txbVeiculo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub UserForm_initialize()

Me.txbCod.Text = ultimoEntregador()
Me.txbData.Text = Date
Application.ScreenUpdating = True
Application.EnableEvents = True
Call popularCbxUF(Me.cbxUf)
End Sub
