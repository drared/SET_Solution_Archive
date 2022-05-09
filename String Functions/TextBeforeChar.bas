Function TextBeforeChar(ByVal InString As String, DefChar As String, Optional InstNum As Integer) As String
    'Returns the string before the specified char(s)
    '   Allows an optional argument to specify the instance of the specified character.
    '   if the character is not found in the string, the original string is returned.
    
    Dim FirstPos As Integer, CurrPos As Integer
    Dim RetString As String
    Dim DefCharCount As Integer
    Dim CharInd As Integer
    
    If InstNum = 0 Then
        FirstPos = Application.WorksheetFunction.Search(DefChar, InString, 1) - 1
        TextBeforeChar = Mid(InString, 1, FirstPos)
        Exit Function
    Else
        DefCharCount = (Len(InString) - Len(Replace(InString, DefChar, ""))) / Len(DefChar)
        If InstNum > DefCharCount Then
            TextBeforeChar = InString
            Exit Function
        End If
        CurrPos = 1
        For CharInd = 1 To InstNum
            CurrPos = Application.WorksheetFunction.Search(DefChar, InString, CurrPos + 1)
        Next CharInd
        CurrPos = CurrPos - 1
        TextBeforeChar = Mid(InString, 1, CurrPos)
    End If

End Function