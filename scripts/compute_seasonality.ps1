$path = "C:\Users\Leandro Piedra\OneDrive - McKinsey & Company\Desktop\Test Cursor\CSP Reforecast\Movements as of April.xlsx"
$ex = New-Object -ComObject Excel.Application
$ex.Visible = $false
$wb = $ex.Workbooks.Open($path)

function Get-MonthWeightsFromPeriodCol {
    param($sheetName, $periodCol)
    $sh = $wb.Worksheets.Item($sheetName)
    $rc = [int]$sh.UsedRange.Rows.Count
    $agg = @{}
    for ($r = 2; $r -le $rc; $r++) {
        $p = [string]$sh.Cells.Item($r, $periodCol).Text
        if ($p.Length -lt 6) { continue }
        $mo = [int]$p.Substring(4, 2)
        if (-not $agg.ContainsKey($mo)) { $agg[$mo] = 0 }
        $agg[$mo] += 1
    }
    $tot = ($agg.Values | Measure-Object -Sum).Sum
    Write-Output "--- $sheetName (count=$tot) ---"
    1..12 | ForEach-Object {
        $m = $_
        $v = if ($agg.ContainsKey($m)) { $agg[$m] } else { 0 }
        $pct = if ($tot -gt 0) { [math]::Round(100 * $v / $tot, 2) } else { 0 }
        Write-Output ("Month {0,2}: {1,4}  ({2}%)" -f $m, $v, $pct)
    }
}

Get-MonthWeightsFromPeriodCol "Arrivals" 1
Get-MonthWeightsFromPeriodCol "Departures" 1

$sh = $wb.Worksheets.Item("Promotions")
$rc = [int]$sh.UsedRange.Rows.Count
$agg = @{}
for ($r = 2; $r -le $rc; $r++) {
    $mo = [int]$sh.Cells.Item($r, 5).Value2
    $v = [double]$sh.Cells.Item($r, 6).Value2
    if (-not $agg.ContainsKey($mo)) { $agg[$mo] = 0 }
    $agg[$mo] += $v
}
$tot = ($agg.Values | Measure-Object -Sum).Sum
Write-Output "--- Promotions (sum VALUE=$tot) ---"
1..12 | ForEach-Object {
    $m = $_
    $v = if ($agg.ContainsKey($m)) { $agg[$m] } else { 0 }
    $pct = if ($tot -gt 0) { [math]::Round(100 * $v / $tot, 2) } else { 0 }
    Write-Output ("Month {0,2}: {1,4}  ({2}%)" -f $m, $v, $pct)
}

Write-Output ""
Write-Output "=== Pivot rows mentioning Seasonality / % ==="
$sh = $wb.Worksheets.Item("Pivot")
$rc = [int]$sh.UsedRange.Rows.Count
for ($r = 1; $r -le $rc; $r++) {
    $a = [string]$sh.Cells.Item($r, 1).Text
    if ($a -match "Season|Depart|Promo|Arrival|%") {
        $cells = @()
        for ($c = 1; $c -le 14; $c++) { $cells += [string]$sh.Cells.Item($r, $c).Text }
        Write-Output ("R{0}: {1}" -f $r, ($cells -join " | "))
    }
}

$wb.Close($false)
$ex.Quit()
