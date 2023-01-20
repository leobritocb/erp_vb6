VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendário 
   Caption         =   "Calendário"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12690
   OleObjectBlob   =   "frmCalendário.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalendário"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDateSelectedVar As Date

Public Property Get SelectDate() As Date
    SelectDate = vDateSelectedVar
End Property

Private Sub UserForm_initialize()
    'A data inicial é a atual:
    lblHoje = "Hoje: " & Format(Date, "dd/mm/yyyy")
    sb = Year(Date) * 12 + Month(Date)
End Sub

Private Sub UserForm_queryclose(Cancel As Integer, CloseMode As Integer)
    'Impede que se dê Unload no formulário, caso contrário a linha que testa
    'frm.Tag na linha seguinte do módulo mdlCalendário dará erro, pois o objeto
    'deixará de existir. Ao invés de dar Unload, usa-se Hide para o objeto
    'continuar a existir na memória.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Hide
    End If
End Sub

Private Sub lblHoje_Click()
    'Quando se clica no Label do dia atual, o calendário atualiza-se
    'para o mês atual.
    
    'O modo de cálculo do mês em questão é o número de meses.
    'Como um ano possui 12 meses, o valor da ScrollBar é o número
    'total de meses:
    sb = Year(Date) * 12 + Month(Date)
    Me.Hide
End Sub

Private Sub sb_Change()
    'Deve-se atualizar o calendário ao alterar a ScrollBar.
    'O valor do calendário é uma divisão inteira (observe o símbolo \)
    'de anos e o resto do valor por 12 como mês:
    ATUALIZAR DateSerial(sb \ 12, sb Mod 12, 1)
End Sub

Private Sub ATUALIZAR(dt As Date)
    'Rotina que atualiza todos os Label do calendário
    
    Dim L As Long
    Dim C As Long
    Dim cInício As Long
    Dim DtDia As Date
    Dim Ctrl As Control
    Dim strLC As String
    
    lblMêsAno = Format(dt, "mmmm yyyy")
    
    For L = 1 To 6 'Linhas do calendário
        For C = 1 To 7 'Colunas do calendário
            Set Ctrl = Controls("l" & L & "c" & C)
            'O entendimento da linha abaixo é fundamental para entender como todos os
            'labels foram povoados:
            DtDia = DateSerial(Year(dt), Month(dt), (L - 1) * 7 + C - Weekday(dt) + 1)
            Ctrl.Caption = Format(Day(DtDia), "00")
            Ctrl.Tag = DtDia
            'Dias de um mês diferente do mês visualizado ficarão na cor cinza claro:
            If Month(DtDia) <> Month(dt) Then
                Ctrl.ForeColor = &HFFFFFF
            End If
            'Realçar dia atual presente, caso esteja visível no calendário:
            If DtDia = Date Then
                Ctrl.ForeColor = &HC00000
                Ctrl.BackColor = &HFFFF&
                Ctrl.BackStyle = fmBackStyleOpaque
            End If
        Next C
    Next L

End Sub
