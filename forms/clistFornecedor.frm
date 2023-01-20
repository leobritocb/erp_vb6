VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} clistFornecedor 
   Caption         =   "Fornecedores"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13605
   OleObjectBlob   =   "clistFornecedor.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "clistFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub nome_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'  Me.nome.SelStart = 0
'  Me.nome.SelLength = Len(Me.nome.Text)
'End Sub

Private Sub pesqCliente_Change()
'Call popularLVPesquisa(Me.pesqCliente.Text, lv, "fornecedor")
End Sub

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
    'Call PopularListView(lv, "fornecedor")

End Sub

Private Sub btnSair_Click()
Unload Me
End Sub

Private Sub numero_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

KeyAscii = soNumeros(KeyAscii)
End Sub

Private Sub cep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
cep.MaxLength = 9 '12345-678
Select Case KeyAscii
Case 8 'Aceita o BACK SPACE
Case 13: SendKeys "{TAB}" 'Emula o TAB
Case 48 To 57
If cep.SelStart = 5 Then cep.SelText = "-"
Case Else: KeyAscii = 0 'Ignora os outros caracteres
End Select
End Sub

Private Sub cnpj_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

KeyAscii = soNumeros(KeyAscii)
End Sub

Private Sub tel_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
tel.MaxLength = 14 '(99)9999-99999
Select Case KeyAscii
Case 8 'Aceita o BACK SPACE
Case 13: SendKeys "{TAB}" 'Emula o TAB
Case 48 To 57
If tel.SelStart = 0 Then tel.SelText = "("
If tel.SelStart = 3 Then tel.SelText = ")"
If tel.SelStart = 8 Then tel.SelText = "-"
Case Else: KeyAscii = 0 'Ignora os outros caracteres
End Select
End Sub

Private Sub cnpj_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Len(cnpj) <= 11 Then
cnpj = Format(cnpj, "000"".""000"".""000-00")
Else
cnpj = Format(cnpj, "00"".""000"".""000""/""0000-00")
End If
End Sub

Private Sub btnExcluir_Click()

'Call excluirCliente("", cod)
Call UserForm_initialize

End Sub

Private Sub btnNovo_Click()
cadCliente.Show
Call UserForm_initialize
End Sub

Private Sub btnSalvar_Click()
'Call editarCliente("", Me.cod)
Call UserForm_initialize
End Sub
