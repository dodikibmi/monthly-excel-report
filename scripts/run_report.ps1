# Unattended monthly run (Windows Task Scheduler): SQL extract -> Excel macro -> .xlsx + .pdf
#   powershell -File scripts/run_report.ps1 [-Period 2026-10]
# Without -Period it reports the previous calendar month.
param([string]$Period = (Get-Date).AddMonths(-1).ToString('yyyy-MM'))
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
python (Join-Path $root 'scripts\extract.py') $Period
if ($LASTEXITCODE -ne 0) { throw "SQL extract failed for $Period" }
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open((Join-Path $root 'MonthlyReport.xlsm'))
    $path = $xl.Run('BuildMonthlyReport', $Period)
    $wb.Save(); $wb.Close($false)          # keeps the Log sheet history
    if ($path -like 'FAILED:*') { Write-Error "Monthly report ${Period}: $path"; exit 1 }
    "report: $path"
}
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
