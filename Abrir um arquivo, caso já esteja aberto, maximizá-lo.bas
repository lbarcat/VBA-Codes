Sub AbrirMaximizar()
   
    Dim Arquivo As String, Diretorio As String, Extensao As String
Diretorio = "C:\Users\usuarioteste\Documents"'Editar de acordo com o seu diretório
Arquivo = "Planilha 1" 'Editar de acordo com o nome do arquivo
Extensao = "xlsx" 'Editar de acordo com a extensão do arquivo
Dim wkb As Workbook
On Error Resume Next
Set wkb = application.Workbooks(Arquivo)

If Err <> 0 Then 
Set wkb = application.Workbooks.Open(Diretorio + "\" + Arquivo + "." & Extensao, UpdateLinks:=0)

End If

On Error GoTo 0
    Workbooks(Arquivo).Activate

End Sub