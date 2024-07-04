# Input and output file paths (EDIT ACCORDINGLY)
$inputFile = "C:\Users\user\Desktop\part_results.txt"
$outputFile = "C:\Users\user\Desktop\out.txt"

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
        $name = $columns[0]
        $testName = $columns[3]
        $score = $columns[5]
        $total = $columns[7]
        
        # Format the output line
        $formattedLine = "$name --- $testName --- $score/$total"
        
        # Add the formatted line to the output array
        $outputLines += $formattedLine
    }
}

# Write the output to the specified file
$outputLines | Out-File -FilePath $outputFile -Encoding utf8

# Notify the user that the operation is complete
Write-Host "Filtered and formatted data has been written to $outputFile"
