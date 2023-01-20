Attribute VB_Name = "geral"
Option Explicit
Dim Tempo As Date
Public pasta As String


Sub inicio_contagem()
End Sub

Sub atualizahora()
'novoPedido.hora.Text = Format(Now, "hh:mm:ss")
'Call inicio_contagem
End Sub
Sub para_contagem()
End Sub

Public Function margemLucro(custo, venda)
Dim margem As Double
Dim res As Double
If custo > 0 And venda <> "" Then
margem = (venda - custo) / custo * 100
res = margem
Else: res = 100
End If
If custo = venda Then res = 0
margemLucro = res
End Function

Public Function calculaPorcentagem(valor, porcento)
Dim res As Double
valor = CDbl(valor)
porcento = CDbl(porcento)
If porcento > 0 Then
porcento = porcento / 100
End If
If valor > 0 And porcento > 0 Then
res = valor * porcento
Else: res = 0
End If
calculaPorcentagem = res
End Function


Function InvNumber(ByVal Number As String) As String

   
    Static i As Integer
    
    For i = 1 To Len(Number)
    
        Select Case Mid$(Number, i, 1)
        Case "-": Mid$(Number, i, 1) = " "
        Case "0": Mid$(Number, i, 1) = "9"
        Case "1": Mid$(Number, i, 1) = "8"
        Case "2": Mid$(Number, i, 1) = "7"
        Case "3": Mid$(Number, i, 1) = "6"
        Case "4": Mid$(Number, i, 1) = "5"
        Case "5": Mid$(Number, i, 1) = "4"
        Case "6": Mid$(Number, i, 1) = "3"
        Case "7": Mid$(Number, i, 1) = "2"
        Case "8": Mid$(Number, i, 1) = "1"
        Case "9": Mid$(Number, i, 1) = "0"
        End Select
        
    Next
    
    InvNumber = Number
    
    
  End Function




Public Function criarPasta()
Dim ConferePasta As String
Dim caminho As String
 
caminho = retornaDiretorio

ConferePasta = caminho & "\impressao"
If Dir(ConferePasta, vbDirectory) = "" Then
MkDir ConferePasta
MsgBox "O diretório: " & ConferePasta & " foi criado!", vbInformation, "AVISO"
End If

ConferePasta = caminho & "\clientes"
If Dir(ConferePasta, vbDirectory) = "" Then
MkDir ConferePasta
MsgBox "O diretório: " & ConferePasta & " foi criado!", vbInformation, "AVISO"
End If

ConferePasta = caminho & "\clientes\imagens"
If Dir(ConferePasta, vbDirectory) = "" Then
MkDir ConferePasta
MsgBox "O diretório: " & ConferePasta & " foi criado!", vbInformation, "AVISO"
End If

ConferePasta = caminho & "\produtos"
If Dir(ConferePasta, vbDirectory) = "" Then
MkDir ConferePasta
MsgBox "O diretório: " & ConferePasta & " foi criado!", vbInformation, "AVISO"
End If

ConferePasta = caminho & "\produtos\imagens"
If Dir(ConferePasta, vbDirectory) = "" Then
MkDir ConferePasta
MsgBox "O diretório: " & ConferePasta & " foi criado!", vbInformation, "AVISO"
End If

ConferePasta = caminho & "\temp"
If Dir(ConferePasta, vbDirectory) = "" Then
MkDir ConferePasta
MsgBox "O diretório: " & ConferePasta & " foi criado!", vbInformation, "AVISO"
End If

ConferePasta = caminho & "\setup"
If Dir(ConferePasta, vbDirectory) = "" Then
MkDir ConferePasta
MsgBox "O diretório: " & ConferePasta & " foi criado!", vbInformation, "AVISO"
End If

'Day (Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now)
End Function
Public Function retornaDiretorio()
    retornaDiretorio = App.Path
End Function


Sub txbSelect(txb As Object)
    txb.SelStart = 0
    txb.SelLength = Len(txb.Text)
End Sub
