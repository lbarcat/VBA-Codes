Public Sub Verificar()
      Dim CompName As String

      CompName = Environ$("ComputerName")
         If CompName <> "PC_Max" Then 'Aqui você irá colocar o nome da máquina autorizada, CMD - digitar hostname
              MsgBox "Este computador não tem direito de executar esta aplicação." 'Mensagem de erro exibida se o nome não bater
              ActiveWorkbook.Close SaveChanges:=False
         End If
End Sub