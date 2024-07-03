# Input and output file paths
$inputFile = "path\to\your\input.txt"
$outputFile = "path\to\your\output.txt"

# Take date as user input
$date = Read-Host "Enter the date (format: YYYY-MM-DD)"

# Read the input file
$lines = Get-Content -Path $inputFile

# Initialize an array to hold the filtered and formatted output
$outputLines = @()

# Process each line
foreach ($line in $lines) {
    # Split the line by comma to get columns
    $columns = $line -split ","
    
    # Check if the line contains the specified date
    if ($columns[2] -eq $date) {
        # Extract Name (Col 1), Test Name (Col 3), Score (Col 4)
        $name = $columns[0]
        $testName = $columns[2]
        $score = $columns[3]
        
        # Format the output line
        $formattedLine = "Name: $name - Test: $testName - Score: $score"
        
        # Add the formatted line to the output array
        $outputLines += $formattedLine
    }
}

# Write the output to the specified file
$outputLines | Out-File -FilePath $outputFile -Encoding utf8

# Notify the user that the operation is complete
Write-Host "Filtered and formatted data has been written to $outputFile"