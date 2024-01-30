# Function to convert IP range to the desired format
function ConvertToRangeFormat($ipRange) {
    $ipStart, $ipEnd = $ipRange -split '-'
    $range = "range,$ipStart,$ipEnd,Default"
    return $range
}

# Prompt user for the input file path
$inputFilePath = Read-Host "Enter the path to the input text file (e.g., C:\Path\To\Your\Input\File.txt)"

# Check if the specified file exists
if (-not (Test-Path $inputFilePath -PathType Leaf)) {
    Write-Host "Error: The specified file does not exist."
    exit
}

# Construct the default output file path based on the input file name
$outputFilePath = [System.IO.Path]::ChangeExtension($inputFilePath, 'csv')

# Read IP ranges from the input file
$ipRanges = Get-Content -Path $inputFilePath

# Initialize an array to store formatted IP ranges
$formattedRanges = @()

# Convert each IP range to the desired format
foreach ($ipRange in $ipRanges) {
    $formattedRange = ConvertToRangeFormat $ipRange
    $formattedRanges += $formattedRange
}

# Export the formatted IP ranges to the output file
$formattedRanges | Out-File -FilePath $outputFilePath -Encoding UTF8

Write-Host "Conversion completed. Formatted IP ranges exported to $outputFilePath"
