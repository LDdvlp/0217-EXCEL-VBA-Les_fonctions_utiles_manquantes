Attribute VB_Name = "Module1"
Sub AddBigBlueTriangleOnCommentIndicator()

    Dim pWs As Worksheet
    Dim pComment As Comment
    Dim pRng As Range
    Dim pShape As Shape
    Set pWs = Application.ActiveSheet
    wShp = 20
    hShp = 10
    For Each pComment In pWs.Comments
        Set pRng = pComment.Parent
        Set pShape = pWs.Shapes.AddShape(msoShapeRightTriangle, pRng.Offset(0, 1).Left - wShp, pRng.Top, wShp, hShp)
        With pShape
            .Flip msoFlipVertical
            .Flip msoFlipHorizontal
            .Fill.ForeColor.SchemeColor = 12
            .Fill.Visible = msoTrue
            .Fill.Solid
            .Line.Visible = msoFalse
        End With
    Next
    
End Sub

Sub RemoveBigBlueTriangleOnCommentIndicator()

    Dim pWs As Worksheet
    Dim pShape As Shape
    Set pWs = Application.ActiveSheet
    For Each pShape In pWs.Shapes
        If Not pShape.TopLeftCell.Comment Is Nothing Then
            If pShape.AutoShapeType = msoShapeRightTriangle Then
                pShape.Delete
            End If
        End If
    Next
    
End Sub


