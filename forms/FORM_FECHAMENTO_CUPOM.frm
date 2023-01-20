VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_FECHAMENTO_CUPOM 
   Caption         =   "(Fechamento de Caixa)"
   ClientHeight    =   8490.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
   OleObjectBlob   =   "FORM_FECHAMENTO_CUPOM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_FECHAMENTO_CUPOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub FOR_FECHAMENTO_Click()
On Error Resume Next
Sheets("fechamento").Select
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
IgnorePrintAreas:=False
End Sub

Private Sub TexTbox11_change()
On Error Resume Next
TextBox11 = Format(TextBox11, "R$ #,###0.00")
End Sub

Private Sub TextBox14_Change()
On Error Resume Next
TextBox14 = Format(TextBox14, "h:mm;@")
End Sub

Private Sub TextBox25_Change()
On Error Resume Next
TextBox25 = Format(TextBox25, "R$ #,###0.00")
End Sub

Private Sub TextBox26_Change()
On Error Resume Next
TextBox26 = Format(TextBox26, "R$ #,###0.00")
End Sub

Private Sub TextBox27_Change()
On Error Resume Next
TextBox27 = Format(TextBox27, "R$ #,###0.00")
End Sub

Private Sub TextBox28_Change()
On Error Resume Next
TextBox28 = Format(TextBox28, "R$ #,###0.00")
End Sub

Private Sub TextBox29_Change()
On Error Resume Next
TextBox29 = Format(TextBox29, "R$ #,###0.00")
End Sub

Private Sub TextBox30_Change()
On Error Resume Next
TextBox30 = Format(TextBox30, "R$ #,###0.00")
End Sub

Private Sub TextBox31_Change()
On Error Resume Next
TextBox31 = Format(TextBox31, "R$ #,###0.00")
End Sub

Private Sub TextBox32_Change()
On Error Resume Next
TextBox32 = Format(TextBox32, "R$ #,###0.00")
End Sub

Private Sub TextBox33_Change()
On Error Resume Next
TextBox33 = Format(TextBox33, "R$ #,###0.00")
End Sub

Private Sub TextBox35_Change()
On Error Resume Next
TextBox35 = Format(TextBox35, "R$ #,###0.00")
End Sub

Private Sub UserForm_initialize()
Sheets("fechamento").Select

TextBox1 = Range("A1")
TextBox2 = Range("A2")
TextBox3 = Range("A3")
TextBox4 = Range("A4")
TextBox5 = Range("A5")
TextBox6 = Range("A6")
TextBox7 = Range("A7")
TextBox8 = Range("A8")
TextBox9 = Range("A9")

TextBox10 = Range("B6")
TextBox11 = Range("B8")
TextBox12 = Range("B9")
TextBox13 = Range("A10")
TextBox14 = Range("B10")
TextBox15 = Range("A12")


TextBox17 = Range("A14")
TextBox18 = Range("A15")
TextBox19 = Range("A16")
TextBox20 = Range("A17")
TextBox21 = Range("A18")
TextBox22 = Range("A19")
TextBox23 = Range("A20")
TextBox24 = Range("A21")

TextBox25 = Range("B14")
TextBox26 = Range("B15")
TextBox27 = Range("B16")
TextBox28 = Range("B17")
TextBox29 = Range("B18")
TextBox30 = Range("B19")
TextBox31 = Range("B20")
TextBox32 = Range("B21")

TextBox33 = Range("B22")
TextBox34 = Range("A22")

TextBox35 = Range("B24")
TextBox36 = Range("A24")


End Sub


