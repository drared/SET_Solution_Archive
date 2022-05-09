Function ConcatenateRange(SelectionRange As Range, Optional InsertStr As String)
    'Concatenates a range of values
    '  Optional 2nd argument can be used to insert characters between each value ie comma, space etc
    
    Dim CurrCell As Range
    Dim TmpStr As String
    Dim CountCell As Integer
    
    TmpStr = ""
    CountCell = 0
    For Each CurrCell In SelectionRange
        CountCell = CountCell + 1
        If CountCell < SelectionRange.Count Then
            If CurrCell <> "" Then TmpStr = TmpStr & CurrCell & InsertStr
        Else
            If CurrCell <> "" Then TmpStr = TmpStr & CurrCell
        End If
        
    Next CurrCell
    

  ConcatenateRange = TmpStr

End Function