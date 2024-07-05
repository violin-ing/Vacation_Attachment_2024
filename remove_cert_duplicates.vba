Sub RemoveDuplicates()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long, j As Long
    
    ' Set references to Sheet1 and Sheet2
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sheet2")
    
    ' Find the last row in each sheet
    lastRow1 = ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row in Sheet2
    For i = lastRow2 To 2 Step -1
        ' Get the name and cert combination from Sheet2
        Dim nameCert As String
        nameCert = ws2.Cells(i, 1).Value & ", " & ws2.Cells(i, 2).Value
        
        ' Check if the combination exists in Sheet1
        For j = 2 To lastRow1
            If nameCert = ws1.Cells(j, 2).Value & ", " & ws1.Cells(j, 8).Value Then
                ' Remove the row if the combination is found in Sheet1
                ws2.Rows(i).Delete
                Exit For
            End If
        Next j
    Next i
End Sub
