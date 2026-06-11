@echo off
setlocal
set "EXCEL_PATH=C:\Users\Leandro Piedra\OneDrive - McKinsey & Company\Desktop\Test Cursor\CSP Reforecast\Movements as of April.xlsx"
powershell.exe -NoLogo -NoProfile -Command ^
  "& { $path=$env:EXCEL_PATH; $ex=New-Object -ComObject Excel.Application; $ex.Visible=$false; $wb=$ex.Workbooks.Open($path); foreach ($sn in @('Pivot')) { $sh=$wb.Worksheets.Item($sn); $rc=[Math]::Min(25,[int]$sh.UsedRange.Rows.Count); $cc=[Math]::Min(20,[int]$sh.UsedRange.Columns.Count); Write-Output ('=== ' + $sn + ' rows=' + $sh.UsedRange.Rows.Count + ' cols=' + $sh.UsedRange.Columns.Count + ' ==='); for ($r=1; $r -le $rc; $r++) { $cells=@(); for ($c=1; $c -le $cc; $c++) { $cells += [string]$sh.Cells.Item($r,$c).Text }; Write-Output ($cells -join ' | ') } Write-Output '' }; $wb.Close($false); $ex.Quit() }"
