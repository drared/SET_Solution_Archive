Function TextAfterChar(ByVal InString As String, DefChar As String, Optional InstNum As Integer) As String
    'Returns the string after the specified char(s)
    '   Allows an optional argument to specify the instance of the specified character.
    '   If the character is not found then a null string is returned
    
    Dim FirstPos As Integer, CurrPos As Integer
    Dim DefCharCount As Integer
    Dim CharInd As Integer
    
    If InstNum = 0 Then
        FirstPos = Application.WorksheetFunction.Search(DefChar, InString, 1) + Len(DefChar)
        TextAfterChar = Mid(InString, FirstPos, Len(InString))
        Exit Function
    Else
        DefCharCount = (Len(InString) - Len(Replace(InString, DefChar, ""))) / Len(DefChar)
        If InstNum > DefCharCount Then
            TextAfterChar = ""
            Exit Function
        End If
        CurrPos = 1
        For CharInd = 1 To InstNum
            CurrPos = Application.WorksheetFunction.Search(DefChar, InString, CurrPos + Len(DefChar))
        Next CharInd
        CurrPos = CurrPos + Len(DefChar)
        TextAfterChar = Mid(InString, CurrPos, Len(InString))
    End If

End Function