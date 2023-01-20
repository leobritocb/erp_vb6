VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_FILTRO 
   Caption         =   "Filtro"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   OleObjectBlob   =   "FORM_FILTRO.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_FILTRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btExecutar_Click()
On Error Resume Next
    Plan42.Range("A5:R10000").ClearContents
    lin = 5
    linha = 5
    
    If cdDataINI = "" Or cdDataFIM = "" Then Exit Sub
    
    Do Until Plan7.Cells(lin, 1) = ""
        If Plan7.Cells(lin, 12) >= CDate(cdDataINI) And _
            Plan7.Cells(lin, 12) <= CDate(cdDataFIM) Then
            
            Plan42.Cells(linha, 1) = Plan7.Cells(lin, 1)
            Plan42.Cells(linha, 2) = Plan7.Cells(lin, 2)
            Plan42.Cells(linha, 3) = Plan7.Cells(lin, 3)
            Plan42.Cells(linha, 4) = Plan7.Cells(lin, 4)
            Plan42.Cells(linha, 5) = Plan7.Cells(lin, 5)
            Plan42.Cells(linha, 6) = Plan7.Cells(lin, 17)
            Plan42.Cells(linha, 7) = Plan7.Cells(lin, 6)
            Plan42.Cells(linha, 8) = Plan7.Cells(lin, 7)
            Plan42.Cells(linha, 9) = Plan7.Cells(lin, 8)
            Plan42.Cells(linha, 10) = Plan7.Cells(lin, 9)
            Plan42.Cells(linha, 11) = Plan7.Cells(lin, 10)
            Plan42.Cells(linha, 12) = CDate(Plan7.Cells(lin, 12))
                                     
            Plan42.Cells(linha, 13) = Plan7.Cells(lin, 13)
            Plan42.Cells(linha, 14) = Plan7.Cells(lin, 14)
            Plan42.Cells(linha, 15) = Plan7.Cells(lin, 15)
            Plan42.Cells(linha, 16) = Plan7.Cells(lin, 16)
            'Plan42.Cells(linha, 17) = Plan7.Cells(lin, 17)
            
            linha = linha + 1
        End If
        lin = lin + 1
    Loop
    MsgBox "Processo concluído - " & cdDataINI & " à " & cdDataFIM
    
    'Call PDF_VENDAS_PERIODO
End Sub

Private Sub Image23_Click()
On Error Resume Next
Me.cdDataINI = GetCalendario
End Sub

Private Sub Image24_Click()
On Error Resume Next
Me.cdDataFIM = GetCalendario
End Sub


