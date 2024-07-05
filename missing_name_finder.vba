Sub CopyMissingNames()
    Dim ws1 As Worksheet, ws2 As Worksheet, wsNew As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long, i As Long, j As Long
    Dim newRow As Long
    Dim name1 As String, name2 As String
    
    ' Set worksheets
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sheet2")
    
    ' Create new sheet
    Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsNew.Name = "FilteredData"
    
    ' Get last rows
    lastRow1 = ws1.Cells(ws1.Rows.Count, "D").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "B").End(xlUp).Row
    
    ' Set header for new sheet
    wsNew.Range("A1:E1").Value = Array("Column C", "Column D", "Column H", "Column I", "Column J")
    newRow = 2
    
    ' Loop through names in Sheet1
    For i = 2 To lastRow1
        ' Normalize the name from Sheet1
        name1 = NormalizeName(ws1.Cells(i, "D").Value)
        Dim found As Boolean
        found = False
        
        ' Loop through names in Sheet2 to find a match
        For j = 2 To lastRow2
            ' Normalize the name from Sheet2
            name2 = NormalizeName(ws2.Cells(j, "B").Value)
            If name1 = name2 Then
                found = True
                Exit For
            End If
        Next j
        
        ' If name not found in Sheet2, copy data to new sheet
        If Not found Then
            wsNew.Cells(newRow, 1).Value = ws1.Cells(i, "C").Value
            wsNew.Cells(newRow, 2).Value = ws1.Cells(i, "D").Value
            wsNew.Cells(newRow, 3).Value = ws1.Cells(i, "H").Value
            wsNew.Cells(newRow, 4).Value = ws1.Cells(i, "I").Value
            wsNew.Cells(newRow, 5).Value = ws1.Cells(i, "J").Value
            newRow = newRow + 1
        End If
    Next i
    
    ' Auto-fit columns in the new sheet
    wsNew.Columns("A:E").AutoFit
    
    MsgBox "Data copied to new sheet 'FilteredData'", vbInformation
End Sub

Function NormalizeName(name As String) As String
    ' Convert to lowercase
    name = LCase(name)
    ' Remove commas
    name = Replace(name, ",", "")
    ' Remove parts in brackets
    Dim pos As Long
    pos = InStr(name, "(")
    If pos > 0 Then
        name = Trim(Left(name, pos - 1))
    End If
    ' Remove leading and trailing spaces
    name = Trim(name)
    NormalizeName = name
End Function
