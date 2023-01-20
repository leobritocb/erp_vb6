Attribute VB_Name = "imagem"
Sub salvarImg(img As Object, localImagem As String, tela As String, cod As String, Optional descricao As String)
    
    Dim nomePasta, nomeArquivo As String
    If cadProduto.dirImagem.Caption <> "Carregar Imagem" Then
    On Error GoTo erro
    If tela = "cadProduto" Then nomePasta = "C:\GettingTec\produtos\imagens"
    If tela = "cadCliente" Then nomePasta = "C:\GettingTec\clientes\imagens"
    
    nomeArquivo = descricao & cod & ".bmp"
            
            SavePicture img.Picture, nomePasta & "\" & nomeArquivo
        Exit Sub
erro:      MsgBox "Erro ao salvar imagem!"
    Resume Next
    Else: cadProduto.dirImagem.Caption = ""
    End If

End Sub

Sub carregaImg(img As Object, localImagem)
    
    
    On Error GoTo erro
    
    img.Picture = LoadPicture(localImagem)
    img.PictureSizeMode = fmPictureSizeModeStretch
    
    Exit Sub
erro:      MsgBox "Imagem inválida"
    Resume Next
    
End Sub

Sub saveImg()
'Dim tmpSheet As Worksheet
'Dim tmpChart As Chart
'Dim tmpImg As Object
'Dim img As String
'
'On Error GoTo erro
'ActiveSheet.Shapes.Range(Array("img")).Select
'Selection.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
'
'application.ScreenUpdating = False
'  Set tmpSheet = Worksheets.Add
'  Charts.Add
'  ActiveChart.Location Where:=xlLocationAsObject, Name:=tmpSheet.Name
'  Set tmpChart = ActiveChart
'  With tmpChart
'    .Paste
'    Set tmpImg = Selection
'    With .Parent
'      .Height = 500
'      .Width = 700
'    End With
'  End With
'
'img = "C:\Tectel\" & _
'      "romaneio" & ".bmp"
'
'tmpChart.Export FileName:=img, FilterName:="bmp"
'
'application.DisplayAlerts = False
'tmpSheet.Delete
'application.DisplayAlerts = True
'
'application.ScreenUpdating = True
'
'GoTo Fim
'
'erro:
'MsgBox "Erro: " & Err.Description, _
'vbCritical, _
'"Erro: " & Err.Number
'
'Fim:
'Set tmpSheet = Nothing
'Set tmpChart = Nothing
'Set tmpImg = Nothing
'Image1.Picture = LoadPicture("C:\Tectel\romaneio.bmp")
End Sub


