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
    # Style setup as before
}

Function Create-Table {
    # Table creation and setup as before
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
    $TotalCritical = $TotalHigh = $TotalMedium = $TotalLow = 0
    $UniqueCritical = $UniqueHigh = $UniqueMedium = $UniqueLow = 0

    $ProcessedPluginIDsCritical = New-Object System.Collections.Generic.HashSet[int]
    $ProcessedPluginIDsHigh = New-Object System.Collections.Generic.HashSet[int]
    $ProcessedPluginIDsMedium = New-Object System.Collections.Generic.HashSet[int]
    $ProcessedPluginIDsLow = New-Object System.Collections.Generic.HashSet[int]

    $PreviousPluginID = 0
    $CurrentVulnCount = 0
    $CurrentAssetCount = 0
    $CurrentAssetsCol1 = [System.Collections.ArrayList]::new()
    $CurrentAssetsCol2 = [System.Collections.ArrayList]::new()
    $CurrentAssetsCol3 = [System.Collections.ArrayList]::new()
    $CurrentHosts = New-Object System.Collections.Generic.HashSet[string]

    ForEach($Item in $SortedData) {
        if($Item.pluginID -ne $PreviousPluginID) {
            if($CurrentVulnCount -ne 0) {
                # Finalize previous table
                $CurrentTable = $Document.Tables[$Document.Tables.Count]
                $CurrentTable.Cell(2, 2).Range.Text = [String]$CurrentAssetCount
                $CurrentTable.Cell(7, 1).Range.Text = $CurrentAssetsCol1 -join "`v"
                $CurrentTable.Cell(7, 2).Range.Text = $CurrentAssetsCol2 -join "`v"
                $CurrentTable.Cell(7, 3).Range.Text = $CurrentAssetsCol3 -join "`v"
                
                $CurrentAssetsCol1.Clear()
                $CurrentAssetsCol2.Clear()
                $CurrentAssetsCol3.Clear()
                $CurrentAssetCount = 0
                $CurrentHosts.Clear()
            }

            # Setup new table
            $CurrentVulnCount++
            $Selection.Style = "Heading 2"
            $Selection.TypeText($Item.pluginName)
            $Selection.TypeParagraph()
            Create-Table
            $CurrentTable = $Document.Tables[$Document.Tables.Count]
            $PreviousPluginID = $Item.pluginID
        }

        # Update asset columns
        $name = $Item.HostIP
        if($CurrentHosts.Add($name)) {
            $columnIndex = $CurrentAssetCount % 3
            switch ($columnIndex) {
                0 { $CurrentAssetsCol1.Add($name) | Out-Null }
                1 { $CurrentAssetsCol2.Add($name) | Out-Null }
                2 { $CurrentAssetsCol3.Add($name) | Out-Null }
            }
            $CurrentAssetCount++

            switch ($Item.risk_factor) {
                'Low' {
                    $TotalLow++
                    if ($ProcessedPluginIDsLow.Add($Item.pluginID)) { $UniqueLow++ }
                    break
                }
                'Medium' {
                    $TotalMedium++
                    if ($ProcessedPluginIDsMedium.Add($Item.pluginID)) { $UniqueMedium++ }
                    break
                }
                'High' {
                    $TotalHigh++
                    if ($ProcessedPluginIDsHigh.Add($Item.pluginID)) { $UniqueHigh++ }
                    break
                }
                'Critical' {
                    $TotalCritical++
                    if ($ProcessedPluginIDsCritical.Add($Item.pluginID)) { $UniqueCritical++ }
                    break
                }
            }
        }
    }

    # Finalize last table if any
    if ($CurrentVulnCount -ne 0) {
        $CurrentTable = $Document.Tables[$Document.Tables.Count]
        $CurrentTable.Cell(2, 2).Range.Text = [String]$CurrentAssetCount
        $CurrentTable.Cell(7, 1).Range.Text = $CurrentAssetsCol1 -join "`v"
        $CurrentTable.Cell(7, 2).Range.Text = $CurrentAssetsCol2 -join "`v"
        $CurrentTable.Cell(7, 3).Range.Text = $CurrentAssetsCol3 -join "`v"
    }

    # Display totals and unique counts
    $Selection.TypeText("Totals")
    $Selection.TypeParagraph()
    $Selection.TypeText("Total Critical: $TotalCritical`nUnique Critical: $UniqueCritical")
    $Selection.TypeParagraph()
    $Selection.TypeText("Total High: $TotalHigh`nUnique High: $UniqueHigh")
    $Selection.TypeParagraph()
    $Selection.TypeText("Total Medium: $TotalMedium`nUnique Medium: $UniqueMedium")
    $Selection.TypeParagraph()
    $Selection.TypeText("Total Low: $TotalLow`nUnique Low: $UniqueLow")
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
