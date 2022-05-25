Attribute VB_Name = "RenameModule"
Option Explicit

Sub RenameFiles()
    'Renames fils in selected folder from file name listed in first column to file name listed
    '   in second column
    
    'Dim OldName As Name, NewName As Name
    Dim OldName, NewName
    Dim FoldPath As String
    Dim ACell As Range, ColA As Range
    
    FoldPath = RenameSht.Range("FileLocation")
    Set ColA = RenameSht.Range("D4:D1048576")
    Set ACell = ColA.Cells(1)
    OldName = ACell.Value
    
    Do Until OldName = ""
        Name FoldPath & ACell.Value As FoldPath & ACell.Offset(0, 2).Value
        Set ACell = ACell.Offset(1, 0)
        OldName = ACell.Value
    Loop

End Sub
