VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_FECHAMENTOS 
   Caption         =   "(Fechamento)"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7080
   OleObjectBlob   =   "FORM_FECHAMENTOS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_FECHAMENTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btn_dataINI_Click()
On Error Resume Next
Me.txt_data = GetCalendario

End Sub

Private Sub CommandButton1_Click()
On Error Resume Next

Application.Calculation = xlAutomatic


If txt_nome = TextBox2 Then


If txt_nome.Text = "" Then
        MsgBox "Informe o usuário responsável"
        txt_nome.SetFocus
        Exit Sub
    End If

Sheets("fechamento").Select

Dim DATAS As Date
Dim valorr As Double
DATAS = txt_data
valorr = TextBox1

Range("B6").Select
ActiveCell.Value = txt_nome
ActiveCell.Offset(0, 0).Value = txt_nome

Range("B9").Select
ActiveCell.Value = DATAS
ActiveCell.Offset(0, 0).Value = DATAS

Range("B8").Select
ActiveCell.Value = valorr
ActiveCell.Offset(0, 0).Value = valorr




Call BT_excluir_Click


Exit Sub
End If
MsgBox "Ops: Algo está errado, você não pode fechar o caixa com esse usuário, verifique!", vbQuestion, "Alerta"


'FORM_FECHAMENTO_CUPOM.Show
End Sub


Sub GerarPDF()
On Error Resume Next
Application.ScreenUpdating = False

Dim SvInput As String
Dim Data As String
Dim var_MENSAGEM
Dim nome As String
     

'Para determinar o fim da planilha com o nome "pdff", e "Banco" o nome da planilha
pdff = Worksheets("fechamento").UsedRange.Rows.Count
    
'Selecionar o inicio e o fim da planilha
Range("A1:L" & pdff).Activate
    
'Nome = InputBox("Digite o nome para o relatório", "Gerar Relatório PDF")
Data = VBA.Format(VBA.Date, "dd-mm-yyyy")
SvInput = ThisWorkbook.Path & Application.PathSeparator & nome & "_" & Data & ".pdf"

With ActiveSheet
.ExportAsFixedFormat _
Type:=x1TypePDF, _
FileName:=SvInput, _
OpenAfterPublish:=True
End With
        
Application.ScreenUpdating = True
End Sub


Private Sub ListBox1_Click()
Application.ScreenUpdating = False

On Error Resume Next
txt_id.Text = ListBox1.List(ListBox1.ListIndex, 0)
txt_nome.Text = ListBox1.List(ListBox1.ListIndex, 1)
TextBox1.Text = ListBox1.List(ListBox1.ListIndex, 2)
txt_data.Text = ListBox1.List(ListBox1.ListIndex, 3)
CommandButton1.Enabled = True
End Sub

Private Sub UserForm_initialize()
'Sheets("fechamento").Select
'TXT_INICIAL.Text = Range("a10").Value

'If TXT_INICIAL.Value = vbNullString Then
   ' With CommandButton1
        '.Enabled = False
        '.Locked = True
   ' End With
'Else
    'With CommandButton1
        '.Enabled = True
       ' .Locked = False
   ' End With
'End If



'****Carrega os dados na listbox
ListBox1.Clear
linha = 2
linhalistbox = 0
contador_lista = 0

Do Until Sheets("HISTORICO_CAIXA").Cells(linha, 1) = ""
    
With Me.ListBox1
.AddItem
.List(linhalistbox, 0) = Sheets("HISTORICO_CAIXA").Cells(linha, 1)
.List(linhalistbox, 1) = Sheets("HISTORICO_CAIXA").Cells(linha, 2)
.List(linhalistbox, 2) = Format(Sheets("HISTORICO_CAIXA").Cells(linha, 3).Value, "#,###0.00")
.List(linhalistbox, 3) = Sheets("HISTORICO_CAIXA").Cells(linha, 4)
.List(linhalistbox, 4) = Sheets("HISTORICO_CAIXA").Cells(linha, 5)



End With
        
linha = linha + 1
linhalistbox = linhalistbox + 1
contador_lista = contador_lista + 1
'lbl_registros = contador_lista
Loop

End Sub

Private Sub UserForm_Activate()
On Error Resume Next

Sheets("pedidos").Select
TextBox2 = Range("e3")


CommandButton1.Enabled = False
End Sub


Private Sub BT_excluir_Click()
On Error Resume Next
Application.ScreenUpdating = False

Sheets("HISTORICO_CAIXA").Select

Dim CODIGO As String

linha = 2
CODIGO = txt_id

Sheets("HISTORICO_CAIXA").Select
Do Until Sheets("HISTORICO_CAIXA").Cells(linha, 1) = ""
  'condicção para localizar o código
  If Sheets("HISTORICO_CAIXA").Cells(linha, 1) = CODIGO Then
     Sheets("HISTORICO_CAIXA").Cells(linha, 1).Select
     
     
     Dim resposta As String 'cria a variável resposta
     resposta = MsgBox("Deseja realmente fazer o fechameto de caixa?", vbYesNo) 'cria a mensagem para determinar qual ação será executada
        
        If resposta = vbYes Then ' se a resposta for sim então
     
       'comando para deletar toda a linha
        ActiveCell.Rows("1:1").EntireRow.Select
        Selection.Delete Shift:=xlUp
        ActiveCell.Select
        
        'MsgBox ("Excluído com sucesso!!!")
        'Unload Me
        'FORM_PRODUTO.Show
        'txt_tipos = ""
        
        'Unload Me

        
        Else
        End If
        
   End If
   
linha = linha + 1

Loop




FORM_FECHAMENTO_CUPOM.Show

Sheets("fechamento").Select
Range("B8") = ""
'Call UserForm_INITIALIZE

Application.ScreenUpdating = True
End Sub


