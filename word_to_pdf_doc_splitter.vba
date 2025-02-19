Sub SplitDocumentToPDFs()
    Dim doc As Document
    Dim pageCount As Long
    Dim i As Long
    Dim startPage As Long
    Dim endPage As Long
    Dim pdfFileName As String
    Dim folderPath As String
    Dim nameLine As String
    Dim namePart As String
    Dim filePath As String

    filePath = "C:\Path\To\Your\Document.docx"
    Set doc = Documents.Open(filePath)
    
    pageCount = doc.ComputeStatistics(wdStatisticPages)
    folderPath = ThisDocument.Path & "\PDFs\"
    
    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0

    For i = 1 To pageCount Step 5
        startPage = i
        endPage = i + 4
        
        If endPage > pageCount Then
            endPage = pageCount
        End If
        
        nameLine = ExtractNameFromRange(doc, startPage, endPage)
        If nameLine <> "" Then
            namePart = ExtractName(nameLine)
            pdfFileName = folderPath & namePart & ".pdf"
        Else
            pdfFileName = folderPath & "Document_Pages_" & startPage & "_to_" & endPage & ".pdf"
        End If

        ExportRangeToPDF doc, startPage, endPage, pdfFileName
    Next i
End Sub

Sub ExportRangeToPDF(doc As Document, startPage As Long, endPage As Long, pdfFileName As String)
    Dim tempDoc As Document
    Dim range As Range
    
    Set range = doc.Range(doc.GoTo(wdGoToPage, wdGoToAbsolute, startPage).Start, _
                          doc.GoTo(wdGoToPage, wdGoToAbsolute, endPage + 1).Start - 1)
    
    Set tempDoc = Documents.Add
    tempDoc.Content.FormattedText = range.FormattedText
  
    tempDoc.ExportAsFixedFormat OutputFileName:=pdfFileName, _
                                ExportFormat:=wdExportFormatPDF, _
                                OpenAfterExport:=False, _
                                OptimizeFor:=wdExportOptimizeForPrint, _
                                Range:=wdExportAllDocument, _
                                Item:=wdExportDocumentContent, _
                                IncludeDocProps:=True, _
                                KeepIRM:=True, _
                                CreateBookmarks:=wdExportCreateNoBookmarks, _
                                DocStructureTags:=True, _
                                BitmapMissingFonts:=True, _
                                UseISO19005_1:=False

    tempDoc.Close SaveChanges:=False
End Sub

Function ExtractNameFromRange(doc As Document, startPage As Long, endPage As Long) As String
    Dim range As Range
    Dim line As String
    Dim paragraphs As Paragraphs
    Dim i As Integer

    Set range = doc.Range(doc.GoTo(wdGoToPage, wdGoToAbsolute, startPage).Start, _
                          doc.GoTo(wdGoToPage, wdGoToAbsolute, endPage + 1).Start - 1)
    
    Set paragraphs = range.Paragraphs
    For i = 1 To paragraphs.Count
        line = paragraphs(i).Range.Text
        If InStr(line, "NRIC, ADDRESS") > 0 Then
            ExtractNameFromRange = line
            Exit Function
        End If
    Next i

    ExtractNameFromRange = ""
End Function

Function ExtractName(line As String) As String
    Dim parts() As String

    parts = Split(line, ",")
    If UBound(parts) >= 1 Then
        ExtractName = Trim(parts(1))
    Else
        ExtractName = "Unknown_Name"
    End If
End Function
