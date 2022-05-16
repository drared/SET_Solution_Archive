Attribute VB_Name = "Module5"


Sub ListWorksheetsInNamedRanges()
    'Procedure to list all worksheets in the named range list for easy access
    
    Dim CurrWS As Worksheet
    Dim ThisWorkBook As Workbook
    
    Set ThisWorkBook = ActiveWorkbook
    
    For Each CurrWS In ThisWorkBook.Worksheets
        ThisWorkBook.Names.Add _
            Name:="WS_" & CurrWS.Name, _
            RefersTo:="=" & CurrWS.Name & "!" & CurrWS.Range("A1").Address
    Next CurrWS
    
End Sub
