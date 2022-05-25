Attribute VB_Name = "ListFilesModule"
Option Explicit

Sub ListFiles()
    'Routine to move through specified folder and
    '   list all files in that folder
    Dim FldPth As String, TmpStr As String
    Dim RowCnt As Integer
    
    Main.Range("FolderList").Clear
    FldPth = Main.Range("SelectedFolder")
    RowCnt = 6
    
    'Validate specified folder path - and this would be there too....
    If FldPth = "" Then FldPth = ThisWorkbook.Path & "\"
    If Right(FldPth, 1) <> "\" Then FldPth = FldPth & "\"
    
    'Initiate TmpStr
    TmpStr = Dir(FldPth)

    Do Until TmpStr = "" Or RowCnt > 20000
        RowCnt = RowCnt + 1
        Main.Cells(RowCnt, 4) = TmpStr
        TmpStr = Dir(, vbNormal)
    Loop
    
    If RowCnt > 20000 Then
        MsgBox "There were a lot of files. An internal safety limit" & _
            " was reached. Not all files will be listed"
    End If
    
End Sub
