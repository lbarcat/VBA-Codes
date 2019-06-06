Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
 
Dim Senha As String
Senha = "123"
 
If InputBox("Digite a senha para Salvar, ou em branco apenas fecha.", "Proteção") = Senha Then
   Exit Sub
Else
   If SaveAsUI = True Then
      MsgBox "Não é permitido 'Salvar Como'"
      Cancel = True
      Exit Sub
   End If
 
   If SaveAsUI = False Then
      MsgBox "Não é permitido 'Salvar'"
      Cancel = True
      Exit Sub
   End If
End If
 
End Sub
