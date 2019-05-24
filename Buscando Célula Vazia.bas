Sub teste()
    contaLinha = 1
    verificaCel = Cells(contaLinha, 1).Value
    Do While verificaCel <> “”
    contaLinha = contaLinha + 1
    verificaCel = Cells(contaLinha, 1).Value
  Loop
  
  MsgBox "A linha vazia é " + CStr(contaLinha)
  
End Sub
