Attribute VB_Name = "PickFolderModule"
Option Explicit

Sub ShowFileDialog()
    ' Show dialog box to allow folder selection - test for git
    Dim FolderPath As String, FileName As String, TmpPathStr As String
    Dim JustPath As String

        With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = ThisWorkbook.Path
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If
        TmpPathStr = .SelectedItems(1)
        Main.Range("SelectedFolder") = TmpPathStr & "\"
        RenameSht.Range("FileLocation") = TmpPathStr & "\"

    End With
End Sub
