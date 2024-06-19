# Replace with the location of your input file
$inputFile = "path\to\your\input\file.txt" 

# Replace with the path of the script you need to distribute
$scriptPath = "path\to\your\script.ps1" 

# Replace with the flags needed to execute your script
$scriptFlags = "-flag1 -flag2"

# Local folder where the output will be stored
$localOutputFolder = "path\to\your\local\output\folder"

# Reads the file line by line. Each line represents a hostname
Get-Content $inputFile | ForEach-Object {
    $hostname = $_

    # Copies the script to the target machine
    Copy-Item -Path $scriptPath -Destination "\\$hostname\c$\scripts"

    # Invokes the copied script remotely with specified flags
    $output = Invoke-Command -ComputerName $hostname -ScriptBlock { param($path,$flags) & "$path" $flags } -ArgumentList "c:\scripts\script.ps1",$scriptFlags

    # Defines the filename of the output file
    $outputFileName = "$hostname.html"

    # Saves the output to a file locally
    $output | Set-Content -Path "$localOutputFolder\$outputFileName"
}
