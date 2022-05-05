Attribute VB_Name = "Module1"
Option Explicit

Sub CombineSheets()
    'Routine to select data received from confirmit and to cut and paste that data into a single worksheet
    '   Historically, some survey software, faced with limited columns, would split data into multiple worksheets.
    '   This routine was used to collate that data into a single worksheet
    'Overview: This routine is run from the 'master' workbook where the data is to be moved to
    '   The user is prompted to select the workbook where the desired output is. The routine then moves through the
    '   worksheets, copying the data into the single 'master' worksheet.
    '   NOTE: A lot of assumptions are used... Records are aligned, no extra sheets have been added etc.
    
    Dim LastRow As Integer, CurrWS_LastCol As Integer, MasterFirstCol As Integer, MasterLastCol As Integer
    Dim LastCell As Range
    Dim WS_Count As Integer
    Dim Counter As Integer
    
    Dim CurrWB As Workbook, MasterWB As Workbook
    Dim CurrWS As Worksheet, MasterWS As Worksheet
    
    'Master workbook is the workbook to put the data into (ie the collated data)
    Set MasterWB = ThisWorkbook
    
    'Select the workbook to collate (ie the data received from confirmit)
    Application.Dialogs(xlDialogOpen).Show
    
    'CurrWB is the workbook with the original data received from confirmit
    Set CurrWB = ActiveWorkbook
    
    'MasterWS is the worksheet to put the data into
    Set MasterWS = MasterWB.Worksheets("Original Data")
    
    'Clear current contents of master worksheet
    MasterWB.Activate
    MasterWS.Activate
    MasterWS.Cells.Select
    Selection.Clear
    
    'Go through each worksheet of the received data and cut and paste into the MasterWS (Original Data)
    For WS_Count = 1 To CurrWB.Worksheets.Count
        '
        Set CurrWS = CurrWB.Worksheets(WS_Count)
        Set LastCell = CurrWS.Cells(1, 1).SpecialCells(xlCellTypeLastCell)
        LastRow = LastCell.Row
        CurrWS_LastCol = LastCell.Column

        If WS_Count = 1 Then
            MasterFirstCol = 1
            MasterLastCol = CurrWS_LastCol
        End If

        CurrWS.Activate
        CurrWB.Activate
        CurrWS.Select
        CurrWS.Range(CurrWS.Cells(1, 1), CurrWS.Cells(LastRow, CurrWS_LastCol)).Copy
        MasterWS.Activate
        MasterWS.Range(MasterWS.Cells(1, MasterFirstCol), MasterWS.Cells(LastRow, MasterLastCol)).Select
        MasterWS.Paste Destination:=Selection
        MasterFirstCol = MasterLastCol + 1
        MasterLastCol = MasterLastCol + CurrWS_LastCol



    Next WS_Count
    
    MasterWS.[A1].Select
    CurrWB.Close


End Sub



