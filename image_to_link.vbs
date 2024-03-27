Sub ReplacePicturesWithHyperlinks()
    Dim ws As Worksheet
    Dim shp As Shape 
    Dim hl As Hyperlink As String
    Dim anchorCell As Range
    Dim targetCell As Range
    
    Dim linkToSide As Boolean
    linkToSide = False
    'Specify the worksheet'
    Set ws = ThisWorkbook.Sheets("Sheet1")

    'Loop through all shapes in the worksheet'
    For Each shp In ws.shapes
        'Check if the shape is a picture'
        If shp.Type = msoPicture Then
            'Store the top left cell reference'
            Set anchorCell = shp.TopLeftCell

            'Extract the hyperling from the picture'
            On Error Resume Next 'In case the picture doesn't have a Hyperlink
            link = shp.Hyperlink.Address
            On Error GoTo 0

            'If there is a link found'
            If link <> "" Then
                If linkToSide Then 'Add cell to the side'
                    If NotIsEmpty(anchorCell.Offset(0,1).Value Then
                        Set targetCell = anchorCell.Offset(0,1).End(x1ToRight).Offset(0,1)
                    Else
                        Set targetCell = anchorCell.Offset(0,1)
                    End If
                Else 'Replace cell with picture'
                    Set targetCell = anchorCell
                    shp.Delete
                End If
                ws.Hyperlinks.Add Anchor:= targetCell, Adress:=link, TextToDisplay:=targetCell
            End If
        End If
    Next shp
End Sub
