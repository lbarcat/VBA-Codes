'Em caso de erro ir em ferramentas, referencias e habilitar (Microsoft Scripting Runtime)
Sub Lista_Arquivos_nas_pastas()
   Dim RootFolder$
   RootFolder = Localiza_Dir
      If RootFolder = "" Then Exit Sub
      Workbooks.Add
         With Range("A1")
            .Formula = "Arquivos do Diretório: " & RootFolder
            .Font.Bold = True
            .Font.Size = 12
         End With
     Range("A3").Formula = "Caminho: "
     Range("B3").Formula = "Nome: "
     Range("C3").Formula = "Data Criação: "
     Range("D3").Formula = "Data último Acesso: "
     Range("E3").Formula = "Data última Modificação: "
         With Range("A3:E3")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
         End With
   ListFilesInFolder RootFolder, True
   Columns("A:H").AutoFit
End Sub

Sub ListFilesInFolder(SourceFolderName As String, IncludeSubfolders As Boolean)
   Dim FSO As Scripting.FileSystemObject
   Dim SourceFolder As Scripting.Folder
   Dim SubFolder As Scripting.Folder
   Dim FileItem As Scripting.File
   Dim r As Long
   Set FSO = CreateObject("Scripting.FileSystemObject")
   Set SourceFolder = FSO.GetFolder(SourceFolderName)
   r = Range("A65536").End(xlUp).Row + 1
   For Each FileItem In SourceFolder.Files
   Cells(r, 1).Formula = FileItem.ParentFolder
   Cells(r, 2).Formula = FileItem.Name
   Cells(r, 3).Formula = FileItem.DateCreated
   Cells(r, 3).NumberFormatLocal = "dd / mm / aaaa"
   Cells(r, 4).Formula = FileItem.DateLastAccessed
   Cells(r, 5).Formula = FileItem.DateLastModified
   Cells(r, 5).NumberFormatLocal = "dd / mm / aaaa"
   r = r + 1
   Next FileItem
   If IncludeSubfolders Then
      For Each SubFolder In SourceFolder.SubFolders
         ListFilesInFolder SubFolder.Path, True
         Next SubFolder
    End If
   Set FileItem = Nothing
   Set SourceFolder = Nothing
   Set FSO = Nothing
   ActiveWorkbook.Saved = True
End Sub

Private Function Localiza_Dir()
   Dim objShell, objFolder, chemin, SecuriteSlash
   Set objShell = CreateObject("Shell.Application")
   Set objFolder = _
   objShell.BrowseForFolder(&H0&, "Procurar por um Diretório", &H1&)
   On Error Resume Next
   chemin = objFolder.ParentFolder.ParseName(objFolder.Title).Path & ""
   If objFolder.Title = "Bureau" Then
      chemin = "C:WindowsBureau"
   End If
   If objFolder.Title = "" Then
      chemin = ""
   End If
   SecuriteSlash = InStr(objFolder.Title, ":")
   If SecuriteSlash > 0 Then
      chemin = Mid(objFolder.Title, SecuriteSlash - 1, 2) & ""
   End If
Localiza_Dir = chemin
End Function