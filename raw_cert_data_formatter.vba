Sub ReformatData()
    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim name As String
    Dim data As String
    Dim X As String
    Dim y As String
    Dim newRow As Long
    Dim skipColumn As Boolean
    
    ' Set source sheet
    Set srcSheet = ThisWorkbook.Sheets("Sheet1")
    
    ' Define destination sheet name
    sheetName = "Reformatted_Data"
    
    ' Check if the destination sheet already exists, and delete if it does
    On Error Resume Next
    Set destSheet = ThisWorkbook.Sheets(sheetName)
    If Not destSheet Is Nothing Then
        Application.DisplayAlerts = False
        destSheet.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    
    ' Create a new destination sheet
    Set destSheet = ThisWorkbook.Sheets.Add
    destSheet.Name = sheetName
    
    ' Find the last row in the source sheet
    lastRow = srcSheet.Cells(srcSheet.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row in the source sheet
    newRow = 1
    For i = 1 To lastRow
        ' Get the name
        name = srcSheet.Cells(i, 1).Value
        
        ' Find the last column with data in the current row
        lastCol = srcSheet.Cells(i, srcSheet.Columns.Count).End(xlToLeft).Column
        
        ' Loop through each data column starting from column 6 (assuming data starts from column F)
        For j = 6 To lastCol
            ' Get the data value
            data = srcSheet.Cells(i, j).Value
            
            ' Check if the data contains any unwanted values
            If InStr(data, ";") > 0 Then
                X = Split(data, ";")(0)
                Y = Split(data, ";")(1)
                
                skipColumn = False
                If UCase(X) = "NIL" Or UCase(X) = "N/A" Or UCase(X) = "NA" Or X = "-" Or UCase(X) = "N.A" Or UCase(X) = "N.A." Or UCase(X) = "N. A" Then
                    skipColumn = True
                End If
                If UCase(Y) = "NIL" Or UCase(Y) = "N/A" Or UCase(Y) = "NA" Or Y = "-" Or UCase(Y) = "N.A" Or UCase(Y) = "N.A." Or UCase(Y) = "N. A" Then
                    skipColumn = True
                End If
                
                ' Skip the column if it contains any unwanted values
                If Not skipColumn Then
                    ' Write the reformatted data to the destination sheet
                    destSheet.Cells(newRow, 1).Value = name
                    destSheet.Cells(newRow, 2).Value = X
                    destSheet.Cells(newRow, 3).Value = Y
                    
                    ' Move to the next row in the destination sheet
                    newRow = newRow + 1
                End If
            End If
        Next j
    Next i
    
    MsgBox "Data reformatting complete.", vbInformation
End Sub
