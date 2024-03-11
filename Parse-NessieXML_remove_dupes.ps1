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
    $Document.Styles.Item("Heading 1").Font.Color = 0x000000
    $Document.Styles.Item("Heading 1").Font.Size = 14
    $Document.Styles.Item("Heading 1").Font.SmallCaps = $True

    $Document.Styles.Item("Heading 2").Font.Name = "Calibri (Body)"
    $Document.Styles.Item("Heading 2").Font.Bold = $True
    $Document.Styles.Item("Heading 2").Font.Color = 0x000000
    $Document.Styles.Item("Heading 2").Font.Size = 11
    $Document.Styles.Item("Heading 2").Font.SmallCaps = $True
}

Function Create-Table {
    # create table
    $Table = $Selection.Tables.Add($Selection.Range, 7, 3, [Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior, [Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow)

    # Merged cells setup and other table configurations as per original script...

    # Apply Shading and other formatting as per original script...

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
                # Finalize previous table if not first vuln
                Finalize-CurrentTable
            }

            # New vuln; setup
            $CurrentVulnCount++
            Setup-NewTable $Item
            $PreviousPluginID = $Item.pluginID
        }

        # Process current item
        $name = $Item.HostIP
        if($CurrentHosts.Add($name)) {
            Update-AssetColumns $name
            Update-SeverityCounts $Item
        }
    }

    # Complete processing of last item
    if($CurrentVulnCount -ne 0) {
        Finalize-CurrentTable
    }

    # Display totals
    Display-Totals
}

Function Finalize-CurrentTable {
    $CurrentTable = $Document.Tables[$Document.Tables.Count]
    $CurrentTable.Cell(2, 2).Range.Text = [String]$CurrentAssetCount
    Populate-AssetColumns
    Clear-CurrentAssetData
}

Function Setup-NewTable($Item) {
    $Selection.Style = "Heading 2"
    $Selection.TypeText($Item.pluginName)
    $Selection.TypeParagraph()
    Create-Table
    $CurrentTable = $Document.Tables[$Document.Tables.Count]
    Configure-CurrentTable $Item
}

Function Update-AssetColumns($name) {
    $columnIndex = $CurrentAssetCount % 3
    switch($columnIndex) {
        0 { $CurrentAssetsCol1.Add($name) | Out-Null }
        1 { $CurrentAssetsCol2.Add($name) | Out-Null }
        2 { $CurrentAssetsCol3.Add($name) | Out-Null }
    }
    $CurrentAssetCount++
}

Function Update-SeverityCounts($Item) {
    # Update Total counts
    switch ($Item.risk_factor) {
        'Low'      { $TotalLow++ if ($ProcessedPluginIDsLow.Add($Item.pluginID)) { $UniqueLow++ } }
        'Medium'   { $TotalMedium++ if ($ProcessedPluginIDsMedium.Add($Item.pluginID)) { $UniqueMedium++ } }
        'High'     { $TotalHigh++ if ($ProcessedPluginIDsHigh.Add($Item.pluginID)) { $UniqueHigh++ } }
        'Critical' { $TotalCritical++ if ($ProcessedPluginIDsCritical.Add($Item.pluginID)) { $UniqueCritical++ } }
    }
}

Function Clear-CurrentAssetData {
    $CurrentAssetsCol1.Clear()
    $CurrentAssetsCol2.Clear()
    $CurrentAssetsCol3.Clear()
    $CurrentAssetCount = 0
    $CurrentHosts.Clear()
}

Function Populate-AssetColumns {
    $CurrentTable.Cell(7, 1).Range.Text = $CurrentAssetsCol1 -join "`v"
    $CurrentTable.Cell(7, 2).Range.Text = $CurrentAssetsCol2 -join "`v"
    $CurrentTable.Cell(7, 3).Range.Text = $CurrentAssetsCol3 -join "`v"
}

Function Configure-CurrentTable($Item) {
    # Configure table based on severity and fill in fixed fields
}

Function Display-Totals {
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
