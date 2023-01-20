VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalend�rio 
   Caption         =   "Calend�rio"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12690
   OleObjectBlob   =   "frmCalend�rio.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalend�rio"
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
    'A data inicial � a atual:
    lblHoje = "Hoje: " & Format(Date, "dd/mm/yyyy")
    sb = Year(Date) * 12 + Month(Date)
End Sub

Private Sub UserForm_queryclose(Cancel As Integer, CloseMode As Integer)
    'Impede que se d� Unload no formul�rio, caso contr�rio a linha que testa
    'frm.Tag na linha seguinte do m�dulo mdlCalend�rio dar� erro, pois o objeto
    'deixar� de existir. Ao inv�s de dar Unload, usa-se Hide para o objeto
    'continuar a existir na mem�ria.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Hide
    End If
End Sub

Private Sub lblHoje_Click()
    'Quando se clica no Label do dia atual, o calend�rio atualiza-se
    'para o m�s atual.
    
    'O modo de c�lculo do m�s em quest�o � o n�mero de meses.
    'Como um ano possui 12 meses, o valor da ScrollBar � o n�mero
    'total de meses:
    sb = Year(Date) * 12 + Month(Date)
    Me.Hide
End Sub

Private Sub sb_Change()
    'Deve-se atualizar o calend�rio ao alterar a ScrollBar.
    'O valor do calend�rio � uma divis�o inteira (observe o s�mbolo \)
    'de anos e o resto do valor por 12 como m�s:
    ATUALIZAR DateSerial(sb \ 12, sb Mod 12, 1)
End Sub

Private Sub ATUALIZAR(dt As Date)
    'Rotina que atualiza todos os Label do calend�rio
    
    Dim L As Long
    Dim C As Long
    Dim cIn�cio As Long
    Dim DtDia As Date
    Dim Ctrl As Control
    Dim strLC As String
    
    lblM�sAno = Format(dt, "mmmm yyyy")
    
    For L = 1 To 6 'Linhas do calend�rio
        For C = 1 To 7 'Colunas do calend�rio
            Set Ctrl = Controls("l" & L & "c" & C)
            'O entendimento da linha abaixo � fundamental para entender como todos os
            'labels foram povoados:
            DtDia = DateSerial(Year(dt), Month(dt), (L - 1) * 7 + C - Weekday(dt) + 1)
            Ctrl.Caption = Format(Day(DtDia), "00")
            Ctrl.Tag = DtDia
            'Dias de um m�s diferente do m�s visualizado ficar�o na cor cinza claro:
            If Month(DtDia) <> Month(dt) Then
                Ctrl.ForeColor = &HFFFFFF
            End If
            'Real�ar dia atual presente, caso esteja vis�vel no calend�rio:
            If DtDia = Date Then
                Ctrl.ForeColor = &HC00000
                Ctrl.BackColor = &HFFFF&
                Ctrl.BackStyle = fmBackStyleOpaque
            End If
        Next C
    Next L

End Sub
