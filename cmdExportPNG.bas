Attribute VB_Name = "cmdExportPNG"

Public Sub cTextEdit(Jogador As String, Tamanho As String, Numero As String)
Dim s As Shape
Dim sr As ShapeRange

Set sr = ActivePage.Shapes.FindShapes("Tamanho")
For Each s In sr
    s.Text.Story = Tamanho
Next s

Set sr = ActivePage.Shapes.FindShapes("Jogador")
For Each s In sr
    s.Text.Story = Jogador
Next s

Set sr = ActivePage.Shapes.FindShapes("Numero")
For Each s In sr
    s.Text.Story = Numero
Next s

End Sub

Public Sub exportPNG(layerNameShort As String, playerName As String, numberID As String, sizeN As String)
Dim l As Layer
Dim lr As Layers
Dim fileName As String
Dim opt As New StructExportOptions
With opt
    .AntiAliasingType = cdrNormalAntiAliasing
    .ImageType = cdrCMYKColorImage
    .Overwrite = True
End With

cTextEdit playerName, sizeN, numberID

Set lr = ActivePage.Layers
For Each l In lr
    If l.TreeNode.Children.Count <> 0 And _
       InStr(l.Name, layerNameShort) > 0 Then
        
        cLayerShapesSelect l.Name
    
        fileName = ActiveDocument.FilePath & Replace(ActiveDocument.Name, ".cdr", "") & l.Name & _
                   "-" & playerName & "-" & sizeN & ".jpg"
    
        ActiveDocument.Export fileName, cdrJPEG, cdrSelection, opt

    End If
Next l
End Sub


Public Sub cLayerShapesSelect(layerName As String)
ActivePage.Layers.Find(layerName).Shapes.All.CreateSelection
End Sub


Public Sub cExportPNG()
fExportPNG.Show
End Sub
