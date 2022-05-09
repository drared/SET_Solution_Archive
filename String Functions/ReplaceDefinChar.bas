
Function ReplaceDefinChar(ByVal InString As String, CharToRep As String, ReplacementChar As String)
    'Replaces a group of a defined character by a single instance of a defined character
    '   For example, the string "sdfdes------sdfhd-----kghgh----hsdfh--", using the arguments CharToRep as "-" &
    '   ReplacementChar as " " will return "sdfdes sdfhd kghgh hsdfh"

    Dim StrLen As Integer, CurrCharInd As Integer, NextCharInd As Integer, TmpInd As Integer
    Dim WholeStr As String, CurrChar As String, NextChar As String, ResultString As String
    Dim LoopOne As Boolean

    'Initialise
    ResultString = ""
    WholeStr = Trim(InString)
    StrLen = Len(WholeStr)

    'Check the suitability of the Input string (InString)
    'Exit if string is too short
    If StrLen <= 1 OR Len(CharToRep) > 1 Then
        Exit Function
    End If

    'Exit if character to replace (CharToRep) is not contained within the string
    If Not IsNumeric(InStr(CharToRep, WholeStr)) Then
        Exit Function
    End If

    'Only applies if more than 1 character matches so only count
    '   upto the second last character
    'Logic: Move along the input string, one character at a time. Put each character
    '   not matching character to replace in the return string. Instances of character
    '   to replace are not put in return string but are instead replaces with a single
    '   instance of the replacement character.
    TmpInd = 1
    'Loop through the input string
    Do While TmpInd <= (StrLen - 1)
        CurrChar = Mid(WholeStr, TmpInd, 1)
        LoopOne = True
        Do While CurrChar = CharToRep
            'Add replacement character only for the first instance
            If LoopOne = True Then
                ResultString = ResultString & ReplacementChar
                LoopOne = False
            End If
            TmpInd = TmpInd + 1
            'Case where char to replace is at the end of the string
            If TmpInd > StrLen Then
                CurrChar = ""
                Exit Do
            End If
            CurrChar = Mid(WholeStr, TmpInd, 1)
        Loop
        ResultString = ResultString & CurrChar
        TmpInd = TmpInd + 1
    Loop

    ReplaceDefinChar = ResultString

End Function