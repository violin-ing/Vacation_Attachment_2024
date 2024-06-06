# Prompt the user for the date
$date = Read-Host "Enter the date in the format '2024-06-05'"

# Path to the input CSV file and the output Excel file
$inputCsv = "C:\path\to\your\datafile.csv" # Replace with the path to your CSV file
$outputExcel = "C:\path\to\filtered_data.xlsx" # Replace with the desired path for the Excel file

# Function to normalize date format in email
function Get-DatePartFromEmail($email) {
    if ($email -match "\d{8}") {
        return $matches[0]
    }
    return $null
}

# Read the CSV file
$data = Import-Csv -Path $inputCsv

# Filter the rows based on the date and extract necessary columns
$filteredData = @()
foreach ($row in $data) {
    $email = $row.Column1 # Adjust the column index based on your CSV structure
    $datePart = Get-DatePartFromEmail $email
    if ($datePart -and $datePart -eq $date.Replace("-", "")) {
        $filteredData += [PSCustomObject]@{
            Email = $email
            TestType = $row.Column3 # Adjust the column index based on your CSV structure
            Score = $row.Column5 # Adjust the column index based on your CSV structure
            TotalScore = $row.Column7 # Adjust the column index based on your CSV structure
        }
    }
}

# Create an Excel application object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Set header values
$headers = @("Email", "TestType", "Score", "TotalScore")
for ($i = 0; $i -lt $headers.Length; $i++) {
    $worksheet.Cells.Item(1, $i + 1).Value2 = $headers[$i]
}

# Add filtered data to the worksheet
$rowIndex = 2
foreach ($item in $filteredData) {
    $worksheet.Cells.Item($rowIndex, 1).Value2 = $item.Email
    $worksheet.Cells.Item($rowIndex, 2).Value2 = $item.TestType
    $worksheet.Cells.Item($rowIndex, 3).Value2 = $item.Score
    $worksheet.Cells.Item($rowIndex, 4).Value2 = $item.TotalScore
    $rowIndex++
}

# Auto-fit the columns
$worksheet.Columns.AutoFit()

# Save the workbook
$workbook.SaveAs($outputExcel)
$excel.Quit()

# Release the COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "Filtered data has been exported to $outputExcel"
