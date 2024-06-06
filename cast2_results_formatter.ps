# Ensure ImportExcel module is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force
}

# Prompt the user for the date
$date = Read-Host "Enter the date in the format '2024-06-05'"

# Path to the input CSV file and the output Excel file
$inputCsv = "path\to\your\datafile.csv" # Replace with the path to your CSV file
$outputExcel = "path\to\filtered_data.xlsx" # Replace with the desired path for the Excel file

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

# Export the filtered data to an Excel file
$filteredData | Export-Excel -Path $outputExcel -WorksheetName "FilteredData" -AutoSize

Write-Host "Filtered data has been exported to $outputExcel"
