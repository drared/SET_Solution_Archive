Attribute VB_Name = "ChangeWholeToPercentLeaveBlanks"
Sub ChangeWholeToPercentLeaveBlanks()
    
    Dim SelectedRange As Range
    For Each SelectedRange In Selection
        If SelectedRange <> "" Then
            SelectedRange.NumberFormat = "0.0" & Chr(34) & "%" & Chr(34)
            SelectedRange.HorizontalAlignment = xlCenter
        End If
    Next SelectedRange


End Sub
