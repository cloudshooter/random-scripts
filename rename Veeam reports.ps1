# Get all HTML files in the current directory
$htmlFiles = Get-ChildItem -Filter "*.html"

foreach ($file in $htmlFiles) {
    # Read the content of the HTML file
    $content = Get-Content $file.FullName -Raw

    # Use regular expressions to extract job type and name
    $match = [regex]::Match($content, '(Backup job|Backup Copy job|Replication job|Replication Copy job): (.+?)<div class="jobDescription"')
    
    if ($match.Success) {
        $jobType = $match.Groups[1].Value
        $jobName = $match.Groups[2].Value

        # Remove invalid characters from the job type and name
        $jobType = $jobType -replace '[\\/:*?"<>|]', ''
        $jobName = $jobName -replace '[\\/:*?"<>|]', ''

        # Remove leading and trailing whitespaces from the job type and name
        $jobType = $jobType.Trim()
        $jobName = $jobName.Trim()

        # Create a new file name
        $newFileName = "{0} - {1}.html" -f $jobType, $jobName

        # Rename the file
        Rename-Item $file.FullName -NewName $newFileName
        Write-Host "Renamed $($file.Name) to $($newFileName)"
    } else {
        Write-Host "No matching pattern found in $($file.Name)"
    }
}
