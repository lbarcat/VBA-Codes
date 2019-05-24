Sub reexibir() 

Dim i As Integer 
For i = 1 To 5 ' Quantidade de planilhas a ser exibida, Tem que ter no minimo uma célula ativa 
Sheets(i).Visible = -1 
Next i 

End Sub
