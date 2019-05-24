Sub gravar()
     ThisWorkbook.Save
     Call timer
End Sub

Sub timer()
     Application.OnTime Now + TimeValue("00:01:00"), "gravar"
End Sub