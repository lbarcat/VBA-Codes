Sub Example1()
'Update 20150831
    Dim xFSO As Object
    Dim xFolder As Object
    Dim xFile As Object
    Dim xFiDialog As FileDialog
    Dim xPath As String
    Dim I As Integer
    Set xFiDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If xFiDialog.Show = -1 Then
        xPath = xFiDialog.SelectedItems(1)
    End If
    Set xFiDialog = Nothing
    If xPath = "" Then Exit Sub
    Set xFSO = CreateObject("Scripting.FileSystemObject")
    Set xFolder = xFSO.GetFolder(xPath)
    For Each xFile In xFolder.Files
        I = I + 1
        ActiveSheet.Hyperlinks.Add Cells(I, 1), xFile.Path, , , xFile.Name
    Next
End Sub