Sub MapUnits()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim nameUnitDict As Object
    Dim lastRowWs1 As Long
    Dim lastRowWs2 As Long
    Dim i As Long
    Dim name As String
    Dim unit As String
    
    Set ws1 = ThisWorkbook.Sheets("Sheet1") ' Adjust source sheet name accordingly
    Set ws2 = ThisWorkbook.Sheets("Sheet2") ' Adjust destination sheet name accordingly
    Set nameUnitDict = CreateObject("Scripting.Dictionary")
    
    lastRowWs1 = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRowWs1
        name = ws1.Cells(i, 1).Value ' Adjust the "1" depending on the col
        unit = ws1.Cells(i, 2).Value ' Adjust the "2" depending on the col
        If Not nameUnitDict.exists(name) Then
            nameUnitDict.Add name, unit
        End If
    Next i
    
    lastRowWs2 = ws2.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRowWs2
        name = ws2.Cells(i, 1).Value ' Edit the "1" depending on the col
        If nameUnitDict.exists(name) Then
            ws2.Cells(i, 3).Value = nameUnitDict(name) ' Edit the "3" depending on the col
        End If
    Next i
    
    Set nameUnitDict = Nothing
    Set ws1 = Nothing
    Set ws2 = Nothing
End Sub
