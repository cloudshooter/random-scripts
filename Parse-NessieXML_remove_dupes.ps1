#CUSTOM VARIABLES
Param(
    [Parameter(Mandatory=$True)]
    $XmlPath
)

#GLOBAL VARIABLES
$TableCellColorHeaderMain = 0x00C0FF
$TableCellColorHeaderChild = 0x595959
$TableCellColorContent = 0xD9D9D9
$TableFontColorAccent = 0xFFFFFF
$TableCellColorCritical = 0x0000C0
$TableCellColorHigh = 0x1771F7
$TableCellColorMedium = 0x00A3FF
$TableCellColorLow = 0x50B000

Function Run-Main {

# Import XML
[xml]$RawXml = Get-Content $XmlPath -ErrorAction Stop

# Sort based on severity then plugin ID
$SortedData = $RawXml.NewDataSet.Table | Sort-Object -Property @{Expression = {$_.severity}; Descending=$true}, @{Expression = {$_.pluginId}}

Write-Host "Creating new report..."
# create Word object
$Word = New-Object -ComObject Word.Application
$Document = $Word.Documents.Add()
$Selection = $Word.Selection

Write-Host "Generating content..."
Set-DocumentStyles
Create-Report

Write-Host "`nReport is complete! Make sure you save it."
$Word.Visible = $True
}

Function Set-DocumentStyles {
    $Document.Styles.Item("Heading 1").Font.Name = "Calibri (Body)"
    $Document.Styles.Item("Heading 1").Font.Bold = $True
    $Document.Styles.Item("Heading 1").Font.Color  = 0x000000
    $Document.Styles.Item("Heading 1").Font.Size = 14
    $Document.Styles.Item("Heading 1").Font.SmallCaps = $True

    $Document.Styles.Item("Heading 2").Font.Name = "Calibri (Body)"
    $Document.Styles.Item("Heading 2").Font.Bold = $True
    $Document.Styles.Item("Heading 2").Font.Color  = 0x000000
    $Document.Styles.Item("Heading 2").Font.Size = 11
    $Document.Styles.Item("Heading 2").Font.SmallCaps = $True
}

Function Create-Table {

# create table
$Table = $Selection.Tables.Add($Selection.Range,7,3, [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior, [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow)

# Merge Cells
$Table.Cell(1,1).Merge($Table.Cell(1,2))
$Table.Cell(2,1).Merge($Table.Cell(2,2))

$Table.Cell(3,1).Merge($Table.Cell(3,2))
$Table.Cell(3,1).Merge($Table.Cell(3,2))

$Table.Cell(4,1).Merge($Table.Cell(4,2))
$Table.Cell(4,1).Merge($Table.Cell(4,2))

$Table.Cell(5,1).Merge($Table.Cell(5,2))
$Table.Cell(5,1).Merge($Table.Cell(5,2))

$Table.Cell(6,1).Merge($Table.Cell(6,2))
$Table.Cell(6,1).Merge($Table.Cell(6,2))    

# Set Cell Spacing and Border
$Table.Borders.Enable = $false
$Table.Spacing = 1

# Fill Titles
$Table.Cell(1,1).Range.Text = 'Severity'
$Table.Cell(1,2).Range.Text = 'Total Affected'

$Table.Cell(4,1).Range.Text = 'Recommendation'

$Table.Cell(6,1).Range.Text = 'Affected Asset(s)'

# Format Titles
$Table.Cell(1,1).Range.Font.Bold = $true
$Table.Cell(1,2).Range.Font.Bold = $true

$Table.Cell(2,1).Range.Font.Color = $TableFontColorAccent

$Table.Cell(4,1).Range.Font.Bold = $true
$Table.Cell(4,1).Range.Font.Color = $TableFontColorAccent

$Table.Cell(6,1).Range.Font.Bold = $true
$Table.Cell(6,1).Range.Font.Color = $TableFontColorAccent

# Apply Shading
$Table.Cell(1,1).Shading.BackgroundPatternColor = $TableCellColorHeaderMain
$Table.Cell(1,2).Shading.BackgroundPatternColor = $TableCellColorHeaderMain

$Table.Cell(2,2).Shading.BackgroundPatternColor = $TableCellColorContent

$Table.Cell(3,1).Shading.BackgroundPatternColor = $TableCellColorContent

$Table.Cell(4,1).Shading.BackgroundPatternColor = $TableCellColorHeaderChild

$Table.Cell(5,1).Shading.BackgroundPatternColor = $TableCellColorContent

$Table.Cell(6,1).Shading.BackgroundPatternColor = $TableCellColorHeaderChild

$Table.Cell(7,1).Shading.BackgroundPatternColor = $TableCellColorContent
$Table.Cell(7,2).Shading.BackgroundPatternColor = $TableCellColorContent
$Table.Cell(7,3).Shading.BackgroundPatternColor = $TableCellColorContent

$Selection.EndKey(6) | Out-Null
$Selection.TypeParagraph()
}

Function Display-Progress {
    param(
        [int]$Index,
        [int]$Total
    )

$CurrentPercentage = [int](($Index / $Total) * 100)
Write-Host "`r              `r" -NoNewLine
Write-Host "$CurrentPercentage% processed" -NoNewLine
}

Function Create-Report {
    $TotalCritical = 0
    $TotalHigh = 0
    $TotalMedium = 0
    $TotalLow = 0

    $PreviousPluginID = 0
    $CurrentVulnCount = 0
    $CurrentAssetCount = 0
    $CurrentAssetsCol1 = [System.Collections.ArrayList]::new()
    $CurrentAssetsCol2 = [System.Collections.ArrayList]::new()
    $CurrentAssetsCol3 = [System.Collections.ArrayList]::new()
    $CurrentHosts = New-Object System.Collections.Generic.HashSet[string]

    ForEach($Item in $SortedData) {

    # If different than previous vuln, finish previous table and create new
    if($Item.pluginID -ne $PreviousPluginID) {
        # FINISH PREVIOUS TABLE (IF NOT FIRST VULN)
        if($CurrentVulnCount -ne 0) {
            $CurrentTable = $Document.Tables[$Document.Tables.Count]
            $CurrentTable.Cell(2,2).Range.Text = [String]$CurrentAssetCount
            $CurrentTable.Cell(7,1).Range.Text = $CurrentAssetsCol1 -Join "`v"
            $CurrentTable.Cell(7,2).Range.Text = $CurrentAssetsCol2 -Join "`v"
            $CurrentTable.Cell(7,3).Range.Text = $CurrentAssetsCol3 -Join "`v"
            
            $CurrentAssetsCol1.Clear()
            $CurrentAssetsCol2.Clear()
            $CurrentAssetsCol3.Clear()
            $CurrentAssetCount = 0
            $CurrentHosts.Clear()
        }
        
        # CREATE NEW TABLE
        $CurrentVulnCount++
        $Selection.Style = "Heading 2"
        $Selection.Range.ListFormat.ApplyNumberDefault()
        $Selection.Range.ListFormat.ListOutdent()
        $Selection.TypeText($Item.pluginName)
        $Selection.TypeParagraph()
        
        # Create table and get object        
        Create-Table
        $CurrentTable = $Document.Tables[$Document.Tables.Count]
        
        # Add content to table and format
        $CurrentTable.Cell(2,1).Range.Text = $Item.risk_factor
        $CurrentTable.Cell(3,1).Range.Text = $Item.description
        $CurrentTable.Cell(5,1).Range.Text = $Item.Solution
        
        # Format table based on severity
        switch ($Item.risk_factor) {
            'Low'      { $CurrentTable.Cell(2,1).Shading.BackgroundPatternColor = $TableCellColorLow; break}
            'Medium'   { $CurrentTable.Cell(2,1).Shading.BackgroundPatternColor = $TableCellColorMedium; break }
            'High'     { $CurrentTable.Cell(2,1).Shading.BackgroundPatternColor = $TableCellColorHigh; break }
            'Critical' { $CurrentTable.Cell(2,1).Shading.BackgroundPatternColor = $TableCellColorCritical; break }
        }
        
        # Update PluginID flag for next round
        $PreviousPluginID = $Item.pluginID
    }
    
    # Add Asset to Column array to later populate table
    $name = $Item.HostIP

    if ($CurrentHosts.Add($name)) {
        if($CurrentAssetCount % 3 -eq 1) {
            $CurrentAssets
                            $CurrentAssetsCol1.Add($name) | Out-Null
                $CurrentAssetCount++
            }
            elseif($CurrentAssetCount % 3 -eq 2) {
                $CurrentAssetsCol2.Add($name) | Out-Null
                $CurrentAssetCount++
            }
            else {
                $CurrentAssetsCol3.Add($name) | Out-Null
                $CurrentAssetCount++
            }
        }
        
        # Update Total counts for summary report
        switch ($Item.risk_factor) {
            'Low'      { $TotalLow++; break}
            'Medium'   { $TotalMedium++; break }
            'High'     { $TotalHigh++; break }
            'Critical' { $TotalCritical++; break }
        }
        
        Display-Progress -Index $SortedData.IndexOf($Item) -Total $SortedData.Count
    }
    
    # COMPLETE LAST TABLE
    $CurrentTable = $Document.Tables[$Document.Tables.Count]
    $CurrentTable.Cell(2,2).Range.Text = [String]$CurrentAssetCount
    $CurrentTable.Cell(7,1).Range.Text = $CurrentAssetsCol1 -Join "`v"
    $CurrentTable.Cell(7,2).Range.Text = $CurrentAssetsCol2 -Join "`v"
    $CurrentTable.Cell(7,3).Range.Text = $CurrentAssetsCol3 -Join "`v"
    
    $Selection.Style = "Heading 2"
    $Selection.TypeText("Totals")
    $Selection.TypeParagraph()
    $Selection.TypeText("Critical: $TotalCritical")
    $Selection.TypeParagraph()
    $Selection.TypeText("High: $TotalHigh")
    $Selection.TypeParagraph()
    $Selection.TypeText("Medium: $TotalMedium")
    $Selection.TypeParagraph()
    $Selection.TypeText("Low: $TotalLow")
    $Selection.TypeParagraph()
    $Selection.TypeText("Unique Vulns: $CurrentVulnCount")
}

Try {
    Run-Main
}
Catch {
    Write-Host "Uh-oh! Something went wrong, here's the error message:" -ForegroundColor Red -BackgroundColor Black
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host ''
    
    exit
}
