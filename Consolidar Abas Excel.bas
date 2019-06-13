Sub Combine()
'UpdatebyExtendoffice20180205
    Dim I As Long
    Dim xRg As Range
    On Error Resume Next
    Worksheets.Add Sheets(1)
    ActiveSheet.Name = "Combined"
   For I = 2 To Sheets.Count
        Set xRg = Sheets(1).UsedRange
        If I > 2 Then
            Set xRg = Sheets(1).Cells(xRg.Rows.Count + 1, 1)
        End If
        Sheets(I).Activate
        ActiveSheet.UsedRange.Copy xRg
    Next
End Sub