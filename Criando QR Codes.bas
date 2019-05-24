Attribute VB_Name = "Módulo7"
Function QrCode(codetext As String)
'O QR Code é um código de barras bi-dimensional utilizado cada vez mais para leitura de informação e aplicações e web sites.

'O Google Code dispõe de uma API (Google Chart API) muito útil para gerar gráficos e nela está incluída a possibilidade de _
criar QR Codes: https://developers.google.com/chart/infographics/docs/qr_codes

'A API é composta por três parâmetros principais:

'   chs = <width>x<height>: se refere á dimensão da imagem do QrCode, ou seja largura x altura

'   cht = qr: se refere ao tipo de gráfico a utilizar, que neste caso será o QR code. O texto do código é o texto que _
    vamos querer incluir no código gerado, e o tipo de encoding será o UTF pretendido, Shift_JIS, UTF-8, ou ISO-8859-1 _
    (valor opcional).

'   chl= <texto do código> : Os dados a serem codificados. Os dados podem ser dígitos (0-9), caracteres alfanuméricos, _
    bytes binários de dados ou Kanji. Você não pode misturar tipos de dados dentro de um código QR. _
    Os dados devem ser UTF-8 codificados por URL. Note que os URLs tem um comprimento máximo de 2K, portanto, se você quiser _
    codificar mais de 2K bytes (menos os outros caracteres de URL), você terá que enviar seus dados usando POST.

' Esta Função gera código QR por meio de uma função personalizada, utilizando-se de uma API do Google, que pode ser endereço _
  de uma célula
'FORMAS DE UTILIZAÇÃO:
' 1) Pode ser feita referência a uma célula, p. ex:  =QrCode(A1);
' 2) O texto pode ser inserido diretamente entre aspas na função: ex: =QrCode("TEXTO"). O valor deve ser uma string

Dim URL As String, MyCell As Range

Set MyCell = Application.Caller

'SINTAXE E PARÂMETROS DA URL:
'
'                                              chs= <width>x<height>: define o tamanho da imagem, largura x altura _
'                                              |     (ideal que seja de no mínimo 100x100 pixels)
'                                              |
'                                              |         cht=qr: especifica o tipo QR code
'                                              |        |
'                        URL base (fixa)       |        |            chl=<data>: dado, texto que será codificado
'      _________________|________________  ____|____  __|__   ______|_______
'     |                                  ||         ||     | |              |
URL = "https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=" & codetext
On Error Resume Next
'Apaga a imagem anterior, se houver
  ActiveSheet.Pictures("QR_" & MyCell.Address(False, False)).Delete
On Error GoTo 0
ActiveSheet.Pictures.Insert(URL).Select
With Selection.ShapeRange(1)
 .PictureFormat.CropLeft = 15
 .PictureFormat.CropRight = 15
 .PictureFormat.CropTop = 15
 .PictureFormat.CropBottom = 15
 .Name = "QR_" & MyCell.Address(False, False)
 .Left = MyCell.Left + 2
 .Top = MyCell.Top + 2
End With
QrCode = ""
End Function


