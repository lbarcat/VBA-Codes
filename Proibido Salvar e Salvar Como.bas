Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
 
Dim Senha As String
Senha = "123"
 
If InputBox("Digite a senha para Salvar, ou em branco apenas fecha.", "Prote��o") = Senha Then
   Exit Sub
Else
   If SaveAsUI = True Then
      MsgBox "N�o � permitido 'Salvar Como'"
      Cancel = True
      Exit Sub
   End If
 
   If SaveAsUI = False Then
      MsgBox "N�o � permitido 'Salvar'"
      Cancel = True
      Exit Sub
   End If
End If
 
End Sub
