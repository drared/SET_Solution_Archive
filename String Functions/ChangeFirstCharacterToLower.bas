Attribute VB_Name = "ChangeFirstCharacterToLower"

Sub ChangeFirstCharacterToLower()

    Dim CRange As Range, CurrCell As Range
    Dim StrLen As Long
    
    Set CRange = Selection

    
    'For each cell in the selected range, replace the first character in
    '   the string with the lower case version
    For Each CurrCell In CRange
        If CurrCell <> "" Then
            StrLen = Len(CurrCell) - 1
            CurrCell = LCase(Mid(CurrCell, 1, 1)) & Mid(CurrCell, 2, StrLen)
        End If
    Next CurrCell
End Sub
