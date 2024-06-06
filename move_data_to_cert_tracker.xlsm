Sub TransferDataAndFillDifficulty()
    Dim srcWs As Worksheet, destWs As Worksheet
    Dim srcLastRow As Long, destLastRow As Long
    Dim i As Long, j As Long
    Dim name As String, cert As String, dateVal As Date
    
    ' Set references to the source and destination sheets
    Set srcWs = ThisWorkbook.Sheets("SourceSheetName") ' Replace "SourceSheetName" with the actual name of your source sheet
    Set destWs = ThisWorkbook.Sheets("DestinationSheetName") ' Replace "DestinationSheetName" with the actual name of your destination sheet
    
    ' Find the last row in each sheet
    srcLastRow = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row
    destLastRow = destWs.Cells(destWs.Rows.Count, "B").End(xlUp).Row
    
    ' Loop through each row in the source sheet
    For i = 2 To srcLastRow ' Assuming row 1 is header
        ' Get data from the source sheet
        name = srcWs.Cells(i, 1).Value
        cert = srcWs.Cells(i, 2).Value
        dateVal = srcWs.Cells(i, 3).Value

        destWs.Cells(destLastRow + 1, 2).Value = name
        destWs.Cells(destLastRow + 1, 8).Value = cert
        destWs.Cells(destLastRow + 1, 10).Value = dateVal ' Assuming the date goes to column J
        destLastRow = destLastRow + 1
        End If
    Next i
End Sub
