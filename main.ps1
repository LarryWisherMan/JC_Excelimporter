$RawDataFolder = ".\Raw"
$ProcessedDataFolder = ".\Processed"

Remove-Excel

Get-Module ExcelHelpers -All | Remove-Module -Force -ErrorAction SilentlyContinue

Import-Module  ".\ExcelHelpers.psm1"



#region Todo - Process raw data files

$rawDataFiles = Get-ChildItem -Path $RawDataFolder -File

Function Convert-DATtoCSV {
    Param()
}

foreach ($file in $rawDataFiles) {
    $fileExtension = $file.Extension
    $fileName = $file.BaseName
    $fileFullName = $file.FullName

    $FileNameParts = $fileName -split "_"

    $PNPart = $FileNameParts[0]
    $SNPart = $FileNameParts[1]
    $TestPart = $FileNameParts[2]
    $DatePart = $FileNameParts[3]

    $CsvContent = Import-Csv -Path $fileFullName

    # Check if CsvContent has rows
    if ($CsvContent.Count -eq 0) {
        Write-Host "No data found in $fileFullName"
        continue
    }

    # Rename columns
    $newHeaders = @("PN", "SN", "Test", "Date") # Add new headers first
    $originalHeaders = $CsvContent[0].PSObject.Properties.Name
    foreach ($column in $originalHeaders) {
        if ($column -match "^Ch\d") {
            $newHeader = $column -replace " ", "-" # Replace spaces with hyphens for CH columns
        } else {
            $newHeader = $column -replace " ", "" # Remove spaces for other columns
        }
        $newHeaders += $newHeader
    }
    # Create new CSV content with updated headers and append new columns to each record
    $newCsvContent = foreach ($row in $CsvContent) {
        $newRow = [PSCustomObject]@{
            PN = $PNPart
            SN = $SNPart
            Test = $TestPart
            Date = $DatePart
        }
        foreach ($column in $originalHeaders) {
            $newHeader = if ($column -match "^Ch\d") {
                $column -replace " ", "-"
            } else {
                $column -replace " ", ""
            }
            $newRow | Add-Member -MemberType NoteProperty -Name $newHeader -Value $row."$column"
        }
        $newRow
    }


    # Save the updated CSV
    $NewFileName = $filename + "_Processed.csv"
    $newFilePath = "$ProcessedDataFolder\$NewFileName"
    $newCsvContent | Export-Csv -Path $newFilePath -NoTypeInformation
}
#endregion

<#
foreach ($file in $rawDataFiles) {
    $fileExtension = $file.Extension
    $fileName = $file.BaseName
    $fileFullName = $file.FullName

    switch ($fileExtension) {
        ".dat" {
            $csvFileName = $fileName + ".csv"
            $csvFileFullName = $ProcessedDataFolder + "\" + $csvFileName
            Convert-DATtoCSV -InputFile $fileFullName -OutputFile $csvFileFullName
        }
        ".csv" {
            $csvFileName = $fileName + ".csv"
            $csvFileFullName = $ProcessedDataFolder + "\" + $csvFileName
            Copy-Item -Path $fileFullName -Destination $csvFileFullName
        }
        default {
            Write-Host "Unsupported file type: $fileExtension"
        }
    }
}

#>

#endregion


#region import processed data files from csv
$processedDataFiles = Get-ChildItem -Path $ProcessedDataFolder -File

$processedData = foreach ($file in $processedDataFiles) {
    Import-Csv -Path $file.FullName
}

#endregion



#load type Library
Add-Type -AssemblyName "Microsoft.Office.Interop.Excel"
$xlConditionValues = [Microsoft.Office.Interop.Excel.XLConditionValueTypes]
$xlTheme = [Microsoft.Office.Interop.Excel.XLThemeColor]
$xlChart = [Microsoft.Office.Interop.Excel.XLChartType]
$xlIconSet = [Microsoft.Office.Interop.Excel.XLIconSet]
$xlDirection = [Microsoft.Office.Interop.Excel.XLDirection]


$Columns = (
"Oventemp",
"Manifoldpressure",
"Ch1-pressure",
"CH1-pressure-error",
"Ch1-temperature",
"Ch1-temp-error"
)


#$Data = $processedData |select $columns | ConvertTo-CSV -NoTypeInformation -Delimiter "`t" 

$Data = $processedData | ConvertTo-CSV -NoTypeInformation -Delimiter "`t" 

$tempSettings = $processedData.Oventemp |ForEach-Object {[math]::Round([int]$_,[MidpointRounding]::toEven)} | Sort-Object -Unique


function Group-ByOventempRange {
    param (
        [array]$Data,
        [int]$RangeSize = 10
    )

    $Data | ForEach-Object {
        $rangeStart = [math]::Floor($_.Oventemp / $RangeSize) * $RangeSize
        $rangeEnd = $rangeStart + $RangeSize - 1
        [PSCustomObject]@{
            Range = "$rangeStart-$rangeEnd"
            Data = $_
        }
    } | Group-Object -Property Range
}


$GroupedByTempRange = Group-ByOventempRange -Data $processedData -RangeSize 10



$ExcelObject = New-ExcelWorkbook

$WS = $ExcelObject.Worksheets[0]
Set-WorksheetData -ws $ws -Data $Data
$Table = New-InsertTable -ws $ws -TableName "Data"


$chart = New-SheetChart -ws $worksheet -ChartTitle "Channel #1 Pressure Curvefit Accuracy" -XAxisTitle "Desired Pressure [%FS]" -YAxisTitle "Pressure Error [PSIA]" -SeriesNames @("Series1", "Series2", "Series3", "Series4", "Series5", "Series6", "Series7", "Series8", "Series9", "Series10") -TableName "Data"

Remove-Excel


Create-Plot -excel $ExcelObject.Excel -SheetName "Data" -ChartTitle "Channel #1 Pressure Curvefit Accuracy" -XAxisTitle "Desired Pressure [%FS]" -YAxisTitle "Pressure Error [PSIA]" -TableName "Data" -SeriesColumnName "Oventemp" -XValuesColumnName "Manifoldpressure" -YValuesColumnName "CH1-pressure-error" -ChartTop 10 -ChartLeft 10

<#Old Code 


function Create-Plot {
    param (
        [Excell.Application]$excel,
        [string]$SheetName,
        [string]$ChartTitle,
        [string]$XAxisTitle,
        [string]$YAxisTitle,
        [string]$DataRange,
        [string]$SeriesColumn,
        [string]$XValuesColumn,
        [string]$YValuesColumn,
        [int]$ChartTop,
        [int]$ChartLeft
    )

    $sheet = $excel.Worksheets.Item($SheetName)
    $chartObject = $sheet.ChartObjects().Add($ChartLeft, $ChartTop, 500, 300)
    $chart = $chartObject.Chart
    $chart.ChartType = [Microsoft.Office.Interop.Excel.XlChartType]::xlLineMarkers

    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $ChartTitle

    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlCategory).HasTitle = $true
    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlCategory).AxisTitle.Text = $XAxisTitle

    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlValue).HasTitle = $true
    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlValue).AxisTitle.Text = $YAxisTitle

    $tempSettings = $sheet.Range($SeriesColumn + "2:" + $SeriesColumn + $row).Value2 | Sort-Object -Unique
    $seriesIndex = 1

    foreach ($tempSetting in $tempSettings) {
        if ($tempSetting -eq $null) { continue }

        $seriesData = $sheet.Range($DataRange).Cells | Where-Object { $_.Value2 -eq $tempSetting }
        if ($seriesData.Count -eq 0) { continue }

        $series = $chart.SeriesCollection().NewSeries()
        $series.Name = "Temperature: $tempSetting"
        $series.XValues = $sheet.Range("$XValuesColumn" + "2:" + "$XValuesColumn" + $row)
        $series.Values = $sheet.Range("$YValuesColumn" + "2:" + "$YValuesColumn" + $row)
        $seriesIndex++
    }
}

#>
