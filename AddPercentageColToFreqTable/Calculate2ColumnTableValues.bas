Attribute VB_Name = "CalculateTableValues"
Sub CalculateTableValues()
    'Author: Dale Anderson
    'Description: Using a selected range of two columns and n number of rows
    ' with the first row containing frequencies to be summed, this macro totals
    ' the frequencies at the bottom row, calculates the proportion of frequencies,
    ' displays them in the second column and shows their total ie 100% in the
    ' bottom of the second column.
    ' All values are formated as centered and percentages are shown with
    ' 1 decimal place, totals are shown in bold
    
    
    Dim WorkingRange As Range
    Dim RowsInArea As Integer, n As Integer
    Dim ValuesToBeSummed As Range
    Dim RangeString1 As String, RangeString2 As String
    Dim RangeString3 As String, RangeString4 As String
    Dim SumOfFreq As Range, SumOfPercent As Range
    Dim CheckRange As Range
    Set WorkingRange = Selection
    RowsInArea = WorkingRange.Rows.count
    
    
    '----------------Error Checking------------------------------------------------
        '----------------Check the selected area is the correct size-------------
        If WorkingRange.Columns.count <> 2 Or Not WorkingRange.Rows.count > 1 Then
            MsgBox "The selected area must contain two columns with the first" & _
                " column containing the values to be summed and an empty row " & _
                "for the totals" & Chr(10) & Chr(10) & "Please select a different area"
            Exit Sub
        End If
        '------------------------------------------------------------------------
        '----------------Check that no text is selected-------------------------
        For Each CheckRange In WorkingRange
            If IsNumeric(CheckRange) = False Then
                MsgBox "The selected area should only contain blank cells" & _
                    " or cells with numeric values"
                Exit Sub
            End If
        Next CheckRange
        '------------------------------------------------------------------------
    '-----------------------------------------------------------------------------



    '----------------Define ranges to work with---------------------
    With WorkingRange
        RangeString1 = .Cells(1).Address
        RangeString2 = .Cells(RowsInArea - 1, 1).Address
        RangeString3 = .Cells(1, 2).Address
        RangeString4 = .Cells(RowsInArea - 1, 2).Address
        Set SumOfFreq = .Cells(RowsInArea, 1)
        Set SumOfPercent = .Cells(RowsInArea, 2)
    End With
    '----------------------------------------------------------------
    '-----------Define summation formulas to total required values -----
    SumOfFreq.Formula = _
        "=sum(" & RangeString1 & ":" & RangeString2 & ")"
    SumOfPercent.Formula = _
        "=sum(" & RangeString3 & ":" & RangeString4 & ")"
    '------------------------------------------------------------------
    '------------Define formulas to calculate percentages--------------
    For n = 1 To RowsInArea - 1
        With WorkingRange.Cells(n, 2)
            .Formula = "=" & .Offset(0, -1).Address & "/" & SumOfFreq.Address
        End With
    Next n
    '-------------------------------------------------------------------
    '------------------Format Selected area as required------------------
    With WorkingRange
        .HorizontalAlignment = xlCenter
        .Range(Cells(RowsInArea, 1), Cells(RowsInArea, 2)).Font.Bold = True
        .Range(Cells(1, 2), Cells(RowsInArea, 2)).Style = "Percent"
        .Range(Cells(1, 2), Cells(RowsInArea - 1, 2)).NumberFormat = _
            "0.0" & "%"
    End With
    '-------------------------------------------------------------------
    
End Sub



