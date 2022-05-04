Attribute VB_Name = "AddCommentFromCellBelow"
Sub AddCommentFromCellBelow()

    'Adds comments to cell(s) using text from the cell directly underneath the selected cell(s)
    Dim TempString As String
    Dim CRange As Range, CCell As Range
    
    Set CRange = Selection
    
    'Move through selected range and add comments to each cell using the text in the cell below
    For Each CCell In CRange
        TempString = CCell.Offset(1, 0)
            If TempString <> "" Then
                CCell.AddComment TempString
            End If
    Next CCell

End Sub
