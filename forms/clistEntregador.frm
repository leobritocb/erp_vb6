VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} clistEntregador 
   Caption         =   "Entregadores"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13335
   OleObjectBlob   =   "clistEntregador.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "clistEntregador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'Dim lngCursor As Long
'On Error Resume Next
'
'With ListView1 ' Alterar nome do Listview
'
'lngCursor = .MousePointer
'.MousePointer = xlWait
'Dim L As Long
'Dim strFormat As String
'Dim strData() As String
'
'Dim lngIndex As Long
'lngIndex = ColumnHeader.Index - 1
'
'If ColumnHeader.Index - 1 = 0 Then ' Alterar a coluna que desejam classificar por ordem alfabética
'
'.SortOrder = (.SortOrder + 1) Mod 2
'.SortKey = ColumnHeader.Index - 1
'.Sorted = True
'
'With .ListItems
'If (lngIndex > 0) Then
'For L = 1 To .Count
'With .item(L).ListSubItems(lngIndex)
'strData = Split(.Tag, Chr$(0))
'.Text = strData(0)
'.Tag = strData(1)
'End With
'Next L
'Else
'
'For L = 1 To .Count
'With .item(L)
'strData = Split(.Tag, Chr$(0))
'.Text = strData(0)
'.Tag = strData(1)
'End With
'Next L
'End If
'End With
'
'
'End If
'
'
'' Classificar por Número
'
'strFormat = String(30, "0") & "." & String(30, "0")
'
'With .ListItems
'
'If (lngIndex > 0) Then
'For L = 1 To .Count
'With .item(L).ListSubItems(lngIndex)
'.Tag = .Text & Chr$(0) & .Tag
'If IsNumeric(.Text) Then
'If CDbl(.Text) >= 0 Then
'.Text = Format(CDbl(.Text), strFormat)
'Else
'.Text = "&" & InvNumber( _
'Format(0 - CDbl(.Text), strFormat))
'End If
'Else
'.Text = ""
'End If
'End With
'Next L
'Else
'For L = 1 To .Count
'With .item(L)
'.Tag = .Text & Chr$(0) & .Tag
'If IsNumeric(.Text) Then
'If CDbl(.Text) >= 0 Then
'.Text = Format(CDbl(.Text), strFormat)
'Else
'.Text = "&" & InvNumber(Format(0 - CDbl(.Text), strFormat))
' End If
'Else
'.Text = ""
'End If
'End With
'Next L
'End If
'End With
'
'.SortOrder = (.SortOrder + 1) Mod 2
'.SortKey = ColumnHeader.Index - 1
'.Sorted = True
'
'With .ListItems
'If (lngIndex > 0) Then
'For L = 1 To .Count
'With .item(L).ListSubItems(lngIndex)
'strData = Split(.Tag, Chr$(0))
'.Text = strData(0)
'.Tag = strData(1)
'End With
'Next L
'Else
'For L = 1 To .Count
'With .item(L)
'strData = Split(.Tag, Chr$(0))
'.Text = strData(0)
'.Tag = strData(1)
'End With
'Next L
'End If
'End With
'
'.MousePointer = lngCursor
'
'End With


'End Sub

'Private Sub nome_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'  Me.nome.SelStart = 0
'  Me.nome.SelLength = Len(Me.nome.Text)
'End Sub

Private Sub pesqCliente_Change()
Dim valor_pesq As String
valor_pesq = pesqCliente.Text
   'Declaração de variáveis
Dim wksOrigem As Worksheet
Dim rData As Range
Dim rCell As Range
Dim LstItem As ListItem
Dim linha As Integer
Dim coluna As Integer
Dim lincont As Long
Dim colCont As Long
Dim viewCont As Long
Dim i As Long
Dim j As Long

coluna = 2
linha = 3
'Definição da planilha de origem
Set wksOrigem = Worksheets("cliente")
 
'Definição do range de origem
Set rData = wksOrigem.Range("A2").CurrentRegion
 
'Adicionar cabeçalho no listview com laço de repetição 'For'

Me.lv.ListItems.Clear
    Me.lv.ColumnHeaders.Clear
    Me.lv.View = lvwReport
    Me.lv.GridLines = True
    Me.lv.ColumnHeaders.Add , , "Cod"
    Me.lv.ColumnHeaders.Add , , "Cliente"
    Me.lv.ColumnHeaders(1).Width = 40
    Me.lv.ColumnHeaders(2).Width = 238
'Alimentar variável linCont com número de linhas do intervalo fonte
lincont = clientesConta()
 
'Alimentar variável colCont com número de linhas do intervalo fonte
colCont = 4
'Popular o ListView
lv.ListItems.Clear
Sheets("cliente").Select
     
    With wksOrigem
           While .Cells(linha, coluna).Value <> Empty
            Valor_Celula = .Cells(linha, coluna).Value
            
            If UCase(Left(Valor_Celula, Len(valor_pesq))) = UCase(valor_pesq) Then
                
             Set LstItem = lv.ListItems.Add(Text:=rData(linha, 1).Value)
             
             For j = 2 To colCont
                LstItem.ListSubItems.Add Text:=rData(linha, j).Value
             Next j
             
            End If
            linha = linha + 1
        Wend
    End With

End Sub

Private Sub txbpesqEntregador_Change()
Call popularLVPesquisa(Me.txbpesqEntregador.Text, lv, "entregador")
End Sub

Private Sub UserForm_initialize()

Me.lv.ListItems.Clear
    Me.lv.ColumnHeaders.Clear
    Me.lv.View = lvwReport
    Me.lv.GridLines = True
    Me.lv.ColumnHeaders.Add , , "Cod"
    Me.lv.ColumnHeaders.Add , , "Cliente"
    Me.lv.ColumnHeaders(1).Width = 40
    Me.lv.ColumnHeaders(2).Width = 242
   ' Call PopularListView(lv, "entregador")

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
cadCliente.Show vbModal
Call UserForm_initialize
End Sub

Private Sub btnSalvar_Click()
'Call editarCliente("", Me.cod)
Call UserForm_initialize
End Sub
