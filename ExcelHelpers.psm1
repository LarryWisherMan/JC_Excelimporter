Add-Type -AssemblyName "Microsoft.Office.Interop.Excel"


$script:ExcelObjectHolder = @{
    Excel = $null
    Workbook = $null
    Worksheets = @()
    Charts = @()
    Tables = @()
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
        [Microsoft.Office.Interop.Excel.XlChartType]$ChartType = ([Microsoft.Office.Interop.Excel.XLChartType]::xlLineMarkers),
        [string]$ChartTitle,
        [string]$XAxisTitle,
        [string]$YAxisTitle
    )

    #Create Chart
    $chart = $ws.Shapes.AddChart().Chart
    $chart.chartType = $ChartType

    #modify the chart title
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $ChartTitle

    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlCategory).HasTitle = $true
    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlCategory).AxisTitle.Text = $XAxisTitle

    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlValue).HasTitle = $true
    $chart.Axes([Microsoft.Office.Interop.Excel.XlAxisType]::xlValue).AxisTitle.Text = $YAxisTitle

    $script:ExcelObjectHolder.Charts += $chart

    return $chart
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
    $workbook.Close($false) |Out-Null

    $excel.Quit() |Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) |Out-Null
    [System.GC]::Collect() |Out-Null
    [System.GC]::WaitForPendingFinalizers() |Out-Null
}