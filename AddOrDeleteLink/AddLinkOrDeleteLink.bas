Attribute VB_Name = "AddLinkOrDeleteLink"
Sub ActivateLink()

    'Add hyperlinks to cells in selected range
    '   using the contents of the cell as the address
    
    Dim CurrRange As Range, CurrCell As Range

    Set CurrRange = Selection

    For Each CurrCell In CurrRange
        CurrCell.Hyperlinks.Add CurrCell, Address:=CurrCell.Value
    Next CurrCell

End Sub

Sub DeleteLink()
    'Remove hyperlinks from cells in selected range
    
    Dim CurrRange As Range, CurrCell As Range
    
    Set CurrRange = Selection


    For Each CurrCell In CurrRange
        CurrCell.Hyperlinks.Delete
    Next CurrCell

End Sub
