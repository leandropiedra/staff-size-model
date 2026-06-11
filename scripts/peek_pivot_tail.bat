@echo off
setlocal
set "EXCEL_PATH=C:\Users\Leandro Piedra\OneDrive - McKinsey & Company\Desktop\Test Cursor\CSP Reforecast\Movements as of April.xlsx"
powershell.exe -NoLogo -NoProfile -Command "& { $path=$env:EXCEL_PATH; $ex=New-Object -ComObject Excel.Application; $ex.Visible=$false; $wb=$ex.Workbooks.Open($path); $sh=$wb.Worksheets.Item('Pivot'); $rc=[int]$sh.UsedRange.Rows.Count; for ($r=35; $r -le $rc; $r++) { $cells=@(); for ($c=1; $c -le 14; $c++) { $cells += [string]$sh.Cells.Item($r,$c).Text }; Write-Output ('R'+$r+': '+($cells -join ' | ')) }; $wb.Close($false); $ex.Quit() }"
endlocal
