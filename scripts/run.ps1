param([switch]$Open)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$output = Join-Path $root 'QualityTracker.xlsm'
if (Test-Path -LiteralPath $output) { throw "Output already exists: $output. Move it before rebuilding." }
$excel = New-Object -ComObject Excel.Application
$book = $null
try {
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $project = $book.VBProject
    if ($null -eq $project) {
        throw 'Excel blocks VBA project access. Use the manual import steps in README.md, or enable Trust access to the VBA project object model in Excel Trust Center if permitted, then rerun.'
    }
    Get-ChildItem (Join-Path $root 'vba') -Filter '*.bas' | ForEach-Object {
        [void]$project.VBComponents.Import($_.FullName)
    }
    $book.SaveAs($output, 52)
    $prefix = "'" + $book.Name + "'!"
    $excel.Run($prefix + 'RunAllTests', $false)
    $rows = $excel.Run($prefix + 'ImportCsv', (Join-Path $root 'samples\inspections.csv'))
    $issues = $excel.Run($prefix + 'ValidateData')
    $excel.Run($prefix + 'BuildReport')
    if ($rows -ne 273 -or $issues -ne 5) { throw "Unexpected sample totals: rows=$rows, issues=$issues" }
    $clean = $book.Worksheets.Item('Clean')
    if ($clean.Cells.Item($clean.Rows.Count, 1).End(-4162).Row -ne 269) { throw 'Expected 268 clean rows.' }
    $expected = @{
        Etch = @(52,0,100,50.115,48.701,51.814)
        Deposition = @(53,0,100,100.029,96.851,102.952)
        Lithography = @(48,0,100,1.009,0.94,1.072)
        Implant = @(54,0,100,199.937,195.957,203.856)
        Anneal = @(61,5,91.8,1000.544,977.71,1026.722)
    }
    $report = $book.Worksheets.Item('Report')
    for ($r = 2; $r -le 6; $r++) {
        $step = [string]$report.Cells.Item($r,1).Value2
        if (-not $expected.ContainsKey($step)) { throw "Unexpected process step: $step" }
        for ($c = 2; $c -le 7; $c++) {
            if ([Math]::Abs([double]$report.Cells.Item($r,$c).Value2 - $expected[$step][$c-2]) -gt 0.000001) {
                throw "Sample mismatch at Report row $r column $c"
            }
        }
        $expected.Remove($step)
    }
    if ($expected.Count -ne 0 -or $report.ChartObjects().Count -ne 1) { throw 'Missing process step or yield chart.' }
    # Regression: a final record with an empty LotID must still be validated.
    $data = $book.Worksheets.Item('Data')
    $data.Cells.Item(275,2).Value2 = 'Etch'
    if ($excel.Run($prefix + 'ValidateData') -ne 6) { throw 'Trailing missing LotID regression failed.' }
    $data.Rows.Item(275).Clear() | Out-Null
    [void]$excel.Run($prefix + 'ValidateData')
    $report.Activate()
    $book.Save()
    Write-Output "PASS: unit tests, sample totals, all report metrics, chart, trailing missing LotID. Saved $output"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
if ($Open) { Start-Process -FilePath $output }
