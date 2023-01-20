Attribute VB_Name = "formata"
'função que aceita somente numeros
Public Function soNumeros(L As IReturnInteger)
    Select Case L
        Case Asc("0") To Asc("9")
            soNumeros = L
        Case Else
            soNumeros = 0
    End Select
End Function

Public Function soInt(valor As String)
    If IsNumeric(valor) Then
        If Left(valor, 1) = "0" Then
        soInt = Replace(valor, "0", "", 1, 1)
        Else: soInt = valor
        End If
     Else
        If valor = "" Then
        Exit Function
        Else
        soInt = Left(valor, (Len(valor) - 1))
        End If
        Exit Function
    End If
End Function



Function formataMoeda(valor)
    
    If IsNumeric(valor) Then
        If InStr(1, valor, "-") >= 1 Then valor = Replace(valor, "-", "") 'retira sinal negativo
        If InStr(1, valor, ",") >= 1 Then valor = CDbl(Replace(valor, ",", "")) 'retirar a virgula
        If InStr(1, valor, ".") >= 1 Then valor = Replace(valor, ".", "") 'para trabalhar melhor retiramos ponto
        Select Case Len(valor) 'verifica casas para inserção de ponto
            Case 1
            numponto = "00" & valor
            Case 2
            numponto = "0" & valor
            Case 6 To 8
            numponto = Left(valor, Len(valor) - 5) & "." & Right(valor, 5)
            Case 9 To 11
            numponto = inseriPonto(8, valor)
            Case 12 To 14
            numponto = inseriPonto(11, valor)
            Case Else
            numponto = valor
        End Select
        numvirgula = Left(numponto, Len(numponto) - 2) & "," & Right(numponto, 2)
        formataMoeda = numvirgula
    Else
        If valor = "" Then
        Exit Function
        Else
        formataMoeda = Left(valor, (Len(valor) - 1))
        End If
        Exit Function
    End If
End Function

Public Function formataPorcento(valor As String)
    If IsNumeric(valor) And Len(valor) < 6 Then
        If InStr(1, valor, "-") >= 1 Then valor = Replace(valor, "-", "") 'retira sinal negativo
        If InStr(1, valor, ",") >= 1 Then valor = CDbl(Replace(valor, ",", "")) 'retirar a virgula
        If InStr(1, valor, ".") >= 1 Then valor = Replace(valor, ".", "") 'para trabalhar melhor retiramos ponto
        Select Case Len(valor) 'verifica casas para inserção de ponto
            Case 1
            numponto = "00" & valor
            Case 2
            numponto = "0" & valor
            Case 6 To 8
            numponto = Left(valor, Len(valor) - 5) & "." & Right(valor, 5)
            Case 9 To 11
            numponto = inseriPonto(8, valor)
            Case 12 To 14
            numponto = inseriPonto(11, valor)
            Case Else
            numponto = valor
        End Select
        numvirgula = Left(numponto, Len(numponto) - 2) & "," & Right(numponto, 2)
        formataPorcento = numvirgula
    Else
        If valor = "" Then
        Exit Function
        Else
        formataPorcento = Left(valor, (Len(valor) - 1))
        End If
        Exit Function
    End If
End Function

Function formataPeso(valor)
    
    If IsNumeric(valor) Then
        If InStr(1, valor, "-") >= 1 Then valor = Replace(valor, "-", "") 'retira sinal negativo
        If InStr(1, valor, ",") >= 1 Then valor = CDbl(Replace(valor, ",", "")) 'retirar a virgula
        If InStr(1, valor, ".") >= 1 Then valor = Replace(valor, ".", "") 'para trabalhar melhor retiramos ponto
        Select Case Len(valor) 'verifica casas para inserção de ponto
            Case 1
            numponto = "000" & valor
            Case 2
            numponto = "00" & valor
            Case 3
            numponto = "0" & valor
            Case 7 To 9
            numponto = Left(valor, Len(valor) - 6) & "." & Right(valor, 6)
            Case 10 To 12
            numponto = inseriPonto(9, valor)
            Case 13 To 15
            numponto = inseriPonto(12, valor)
            Case Else
            numponto = valor
        End Select
        numvirgula = Left(numponto, Len(numponto) - 2) & "," & Right(numponto, 3)
        formataPeso = numvirgula
    Else
        If valor = "" Then
        Exit Function
        Else
        formataPeso = Left(valor, (Len(valor) - 1))
        End If
        Exit Function
    End If
End Function

 
Function inseriPonto(Inicio, valor)
    i = Left(valor, Len(valor) - Inicio)
    M1 = Left(Right(valor, Inicio), 3)
    M2 = Left(Right(valor, 8), 3)
    F = Right(valor, 5)
    If (M2 = M1) And (Len(valor) < 12) Then
    inseriPonto = i & "." & M1 & "." & F
    Else
    inseriPonto = i & "." & M1 & "." & M2 & "." & F
    End If
End Function
 
Public Function formataCPF(ByVal KeyAscii As MSForms.ReturnInteger, texto As String) As String
Select Case KeyAscii
Case 8, 48 To 57
If Len(texto) = 14 Then KeyAscii = 0
If Len(texto) = 3 Then texto = texto & "."
If Len(texto) = 7 Then texto = texto & "."
If Len(texto) = 11 Then texto = texto & "-"
Case Else
KeyAscii = 0
End Select
formataCPF = texto
End Function

Public Function formataCNPJ(ByVal KeyAscii As MSForms.ReturnInteger, texto As String) As String
Select Case KeyAscii
Case 8, 48 To 57
If Len(texto) = 18 Then KeyAscii = 0
If Len(texto) = 2 Then texto = texto & "."
If Len(texto) = 6 Then texto = texto & "."
If Len(texto) = 10 Then texto = texto & "/"
If Len(texto) = 15 Then texto = texto & "-"
Case Else
KeyAscii = 0
End Select
formataCNPJ = texto
End Function

Public Function formataTelefone(ByVal KeyAscii As MSForms.ReturnInteger, texto As String) As String
Select Case KeyAscii
Case 8, 48 To 57
If Len(texto) = 14 Then KeyAscii = 0
If Len(texto) = 0 Then texto = texto & "("
If Len(texto) = 3 Then texto = texto & ")"
If Len(texto) = 4 Then texto = texto & " "
If Len(texto) = 9 Then texto = texto & "-"
Case Else
KeyAscii = 0
End Select
formataTelefone = texto
End Function

Public Function formataCelular(ByVal KeyAscii As MSForms.ReturnInteger, texto As String) As String
Select Case KeyAscii
Case 8, 48 To 57
If Len(texto) = 15 Then KeyAscii = 0
If Len(texto) = 0 Then texto = texto & "("
If Len(texto) = 3 Then texto = texto & ")"
If Len(texto) = 4 Then texto = texto & " "
If Len(texto) = 10 Then texto = texto & "-"
Case Else
KeyAscii = 0
End Select
formataCelular = texto
End Function

Public Function formataCEP(ByVal KeyAscii As MSForms.ReturnInteger, texto As String) As String
Select Case KeyAscii
Case 8, 48 To 57
If Len(texto) = 9 Then KeyAscii = 0
If Len(texto) = 5 Then texto = texto & "-"
Case Else
KeyAscii = 0
End Select
formataCEP = texto
End Function

Public Function formataData(ByVal KeyAscii As MSForms.ReturnInteger, texto As String) As String
Select Case KeyAscii
Case 8, 48 To 57
If Len(texto) = 10 Then KeyAscii = 0
If Len(texto) = 2 Then texto = texto & "/"
If Len(texto) = 5 Then texto = texto & "/"
Case Else
KeyAscii = 0
End Select
formataData = texto
End Function

Public Function formataPlaca(ByVal KeyAscii As MSForms.ReturnInteger, texto As String) As String
Select Case KeyAscii
Case 8, 48 To 57, 97 To 122, 65 To 190
If Len(texto) = 8 Then KeyAscii = 0
If Len(texto) = 3 Then texto = texto & "-"
Case Else
KeyAscii = 0
End Select
formataPlaca = texto
End Function

Function virgulaPonto(valor)
    If InStr(1, valor, ".") >= 1 Then valor = Replace(valor, ".", "")
    If InStr(1, valor, ",") >= 1 Then valor = Replace(valor, ",", ".")
    virgulaPonto = valor
End Function

Function removeAspaSimples(valor)
    If InStr(1, valor, "'") >= 1 Then valor = Replace(valor, "'", "")
    removeAspaSimples = valor
End Function

Function numToBd(valor)
    If InStr(1, valor, ".") >= 1 Then valor = Replace(valor, ".", "")
    If InStr(1, valor, ",") >= 1 Then valor = Replace(valor, ",", "")
    If InStr(1, valor, "-") >= 1 Then valor = Replace(valor, "-", "")
    If InStr(1, valor, "(") >= 1 Then valor = Replace(valor, "(", "")
    If InStr(1, valor, ")") >= 1 Then valor = Replace(valor, ")", "")
    If InStr(1, valor, " ") >= 1 Then valor = Replace(valor, " ", "")
    numToBd = valor
End Function
