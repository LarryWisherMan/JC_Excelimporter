$RawDataFolder = ".\Raw"
$ProcessedDataFolder = ".\Processed"

Import-Module  ".\ExcelHelpers.psm1"

Get-Module ExcelHelpers -All | Remove-Module -Force -ErrorAction SilentlyContinue

#region Todo - Process raw data files

$rawDataFiles = Get-ChildItem -Path $RawDataFolder -File

Function Convert-DATtoCSV {
    Param()
}

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


$Data = $processedData | ConvertTo-CSV -NoTypeInformation -Delimiter "`t" 


$ExcelObject = New-ExcelWorkbook

$WS = $ExcelObject.Worksheets[0]
Set-WorksheetData -ws $ws -Data $Data
$Table = New-InsertTable -ws $ws -TableName "Data"
$Chart = New-SheetChart -ws $ws -ChartTitle "Relative Error (Pressure)" -XAxisTitle "Pressure Settings" -YAxisTitle "Relative Error"

Remove-Excel


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
