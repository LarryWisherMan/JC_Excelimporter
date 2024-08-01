Add-Type -AssemblyName "Microsoft.Office.Interop.Excel"


$script:ExcelObjectHolder = @{
    Excel      = $null
    Workbook   = $null
    Worksheets = @()
    Charts     = @()
    Tables     = @()
}


Function Get-ExcelObject {
    return $script:ExcelObjectHolder
}

Function Set-ExcelObject {
    param(
        $ExcelObject
    )

    $script:ExcelObjectHolder = $ExcelObject
}

Function New-ExcelSheet {
    param(
        $Workbook = $script:ExcelObjectHolder.Workbook,
        $SheetName
    )

    $sheet = $Workbook.Worksheets.Add()
    $sheet.Name = $SheetName
    return $sheet
}

Function New-ExcelWorkbook {
    param(
        [string]$defaultSheetName = "Data"
    ) 

    $excel = new-object -ComObject Excel.Application
    $wb = $excel.workbooks.add()
    $ws = $wb.activesheet
    $ws.Name = $defaultSheetName
    $excel.Visible = $true

    $script:ExcelObjectHolder.Excel = $excel
    $script:ExcelObjectHolder.Workbook = $wb
    $script:ExcelObjectHolder.Worksheets += $ws

    return $script:ExcelObjectHolder
}



Function New-SheetChart {
    param(
        $ws,
        [Microsoft.Office.Interop.Excel.XlChartType]$ChartType = [Microsoft.Office.Interop.Excel.XlChartType]::xlLineMarkers,
        [string]$ChartTitle,
        [string]$XAxisTitle,
        [string]$YAxisTitle,
        [string[]]$SeriesNames,
        [string]$TableName
    )

    # Create Chart
    $chart = $ws.Shapes.AddChart().Chart
    $chart.ChartType = $ChartType

    # Set the data range for the chart using the table name
    $table = $ws.ListObjects.Item($TableName)
    $chart.SetSourceData($table.Range)

    # Modify the chart title
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $ChartTitle

    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlCategory).HasTitle = $true
    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlCategory).AxisTitle.Text = $XAxisTitle

    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlValue).HasTitle = $true
    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlValue).AxisTitle.Text = $YAxisTitle

    # Set the series names if provided
    if ($SeriesNames) {
        for ($i = 0; $i -lt $SeriesNames.Length; $i++) {
            $chart.SeriesCollection($i + 1).Name = $SeriesNames[$i]
        }
    }

    # Customize series formatting
    foreach ($series in $chart.SeriesCollection()) {
        $series.MarkerStyle = [Microsoft.Office.Interop.Excel.XlMarkerStyle]::xlMarkerStyleCircle
        $series.MarkerSize = 5
    }

    # Save the chart object
    $script:ExcelObjectHolder.Charts += $chart

    return $chart
}


function Create-Plot {
    param (
        $excel, 
        [string]$SheetName,
        [string]$ChartTitle,
        [string]$XAxisTitle,
        [string]$YAxisTitle,
        [string]$TableName,
        [string]$SeriesColumnName,
        [string]$XValuesColumnName,
        [string]$YValuesColumnName,
        [int]$ChartTop = 10,
        [int]$ChartLeft = 10
    )

    $sheet = $excel.Worksheets.Item($SheetName)
    $table = $sheet.ListObjects.Item($TableName)

    $chartObject = $sheet.ChartObjects().Add($ChartLeft, $ChartTop, 500, 300)
    $chart = $chartObject.Chart
    $chart.ChartType = [Microsoft.Office.Interop.Excel.XlChartType]::xlLineMarkers

    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $ChartTitle

    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlCategory).HasTitle = $true
    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlCategory).AxisTitle.Text = $XAxisTitle

    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlValue).HasTitle = $true
    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlValue).AxisTitle.Text = $YAxisTitle

    # Get unique series values from the SeriesColumnName
    $seriesValues = @()
    foreach ($row in $table.ListRows) {
        $value = $row.Range.Columns.Item($table.ListColumns[$SeriesColumnName].Index).Value2
        if ($value -and -not $seriesValues.Contains($value)) {
            $seriesValues += $value
        }
    }

    $seriesValues = $seriesValues | Sort-Object -Unique
    # Create series for each unique value in the SeriesColumnName
    foreach ($seriesValue in $seriesValues) {
        $series = $chart.SeriesCollection().NewSeries()

        $SeriesName = "Series: $seriesValue"
        $series.Name = $SeriesName

        # Filter rows matching the series value
        $xValues = @()
        $yValues = @()
        foreach ($row in $table.ListRows) {
            if ($row.Range.Columns.Item($table.ListColumns[$SeriesColumnName].Index).Value2 -eq $seriesValue) {
                $xValues += $row.Range.Columns.Item($table.ListColumns[$XValuesColumnName].Index).Value2
                $yValues += $row.Range.Columns.Item($table.ListColumns[$YValuesColumnName].Index).Value2
            }
        }

        # Set X and Y values for the series
        $chart.SeriesCollection($SeriesName).XValues = @($xValues)
        $chart.SeriesCollection($SeriesName).Values = @($yValues)

    }
}

Function Set-WorksheetData {
    param(
        $ws,
        $Data
    )

    #Populate test data onto worksheet
    $Data | c:\windows\system32\clip.exe
    $ws.Range("A1").Select | Out-Null
    $ws.paste()
    $ws.UsedRange.Columns.AutoFit() | Out-Null
}

Function New-InsertTable {
    param(
        $ws,
        $Data,
        $TableName
    )

    $table = $ws.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $ws.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null).Name = $TableName

    $script:ExcelObjectHolder.Tables += $table

    return $table
}

Function Save-Excel {
    param(
        [Microsoft.Office.Interop.Excel.Application]$excel,
        [string]$fileName
    )

    $excel.ActiveWorkbook.SaveAs($fileName)
}

Function Remove-Excel {
    param(
        $ExcelObjectHolder = $script:ExcelObjectHolder
    )

    $excel = $ExcelObjectHolder.Excel
    $workbook = $ExcelObjectHolder.Workbook
    $workbook.Close($false) | Out-Null

    $excel.Quit() | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect() | Out-Null
    [System.GC]::WaitForPendingFinalizers() | Out-Null

    $script:ExcelObjectHolder = @{
        Excel      = $null
        Workbook   = $null
        Worksheets = @()
        Charts     = @()
        Tables     = @()
    }
}

function New-ExcelNamedRanges {
    param (
        $excel = $script:ExcelObjectHolder.Excel,
        [string]$SheetName,
        [string]$TableName,
        [string]$OvenTempColumn,
        [string]$ManifoldpressureColumn,
        [string]$PressureErrorColumn
    )

    # Load the Excel COM object
    $worksheet = $excel.Worksheets.Item($SheetName)

    # Get the table and unique values from the specified column
    $listObject = $worksheet.ListObjects.Item($TableName)
    $ovenTempRange = $listObject.ListColumns.Item($OvenTempColumn).DataBodyRange
    $uniqueOvenTemps = @($ovenTempRange.Value2 | Select-Object -Unique)

    # Create named ranges for each unique OvenTemp value
    foreach ($ovenTemp in $uniqueOvenTemps) {
        if ($null -ne $ovenTemp) {
            # Apply filter to the table
            $listObject.Range.AutoFilter($listObject.ListColumns.Item($OvenTempColumn).Index, $ovenTemp)
            
            # Get the visible rows for Manifoldpressure and CH1-pressure-error
            $visibleRows = $worksheet.UsedRange.SpecialCells(12)  # xlCellTypeVisible
            $manifoldpressureRange = @()
            $pressureErrorRange = @()
            foreach ($row in $visibleRows.Rows) {
                $manifoldpressureRange += $row.Cells.Item(1, $listObject.ListColumns.Item($ManifoldpressureColumn).Index).Address(0, 0)
                $pressureErrorRange += $row.Cells.Item(1, $listObject.ListColumns.Item($PressureErrorColumn).Index).Address(0, 0)
            }

            # Create named ranges
            $xRange = @($manifoldpressureRange)
            $yRange = @($pressureErrorRange)
            $xName = "$TableName" + "_" + "$ovenTemp" + "_X"
            $yName = "$TableName" + "_" + "$ovenTemp" + "_Y"
            $excel.Names.Add($xName, "=$xRange")
            $excel.Names.Add($yName, "=$yRange")
            
            # Clear filter
            $listObject.Range.AutoFilter()
        }
    }
}
# Example usage
#New-ExcelNamedRanges -SheetName "Data" -TableName "Data" -OvenTempColumn "OvenTemp" -ManifoldpressureColumn "Manifoldpressure" -PressureErrorColumn "CH1-pressure-error"
