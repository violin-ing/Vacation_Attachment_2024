# Function to format and filter the CSV
function Format-Csv {
    param (
        [string]$inputCsv,
        [string]$outputTxt,
        [string]$date
    )

    # Import the CSV
    $csvData = Import-Csv -Path $inputCsv

    # Filter and extract relevant data
    $filteredData = $csvData | Where-Object { $_.Date -eq $date } | Select-Object @{Name="Name";Expression={$_.Column1}}, @{Name="TestName";Expression={$_.Column3}}, @{Name="Score";Expression={$_.Column4}}

    # Create output string
    $output = $filteredData | ForEach-Object {
        "Name: $($_.Name) - Test: $($_.TestName) - Score: $($_.Score)"
    }

    # Write to the output text file
    $output | Out-File -FilePath $outputTxt -Encoding utf8
}

# Input parameters
$inputCsv = "path\to\your\input.csv"
$outputTxt = "path\to\your\output.txt"

# Take date as user input
$date = Read-Host "Enter the date (format: YYYY-MM-DD)"

# Call the function with the provided parameters
Format-Csv -inputCsv $inputCsv -outputTxt $outputTxt -date $date