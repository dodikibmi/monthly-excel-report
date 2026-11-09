# Builds MonthlyReport.xlsm from vba/ReportBuilder.bas: Settings sheet (named cells), Log sheet, "Build report" button.
# The VBA module is imported once by hand (no security setting is changed): open the workbook, Alt+F11,
# File > Import File > vba\ReportBuilder.bas, save. Pass -WithModule only if your Excel already trusts VBA project access.
#   powershell -File scripts/build_xlsm.ps1
param([switch]$WithModule)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add()
    $s = $wb.Worksheets.Item(1); $s.Name = 'Settings'
    $rows = @(
        @('Monthly sales report', '', ''),
        @('', '', ''),
        @('Setting', 'Value', 'What it means'),
        @('Period', '2026-10', 'Month to report, yyyy-mm (the scheduler passes it automatically)'),
        @('DataFolder', '.\data', 'Where the SQL extract CSV files are (.\ = next to this workbook)'),
        @('OutputFolder', '.\out', 'Where the .xlsx and .pdf report are saved'),
        @('EmailMode', 'none', 'none = only save files; draft = also create an Outlook draft with both files attached'),
        @('EmailTo', 'finance@example.com', 'Recipient of the draft')
    )
    for ($r = 0; $r -lt $rows.Count; $r++) { for ($c = 0; $c -lt 3; $c++) { $s.Cells.Item($r + 1, $c + 1).Value2 = $rows[$r][$c] } }
    $s.Range('B4').NumberFormat = '@'; $s.Range('B4').Value2 = '2026-10'
    foreach ($n in 'Period', 'DataFolder', 'OutputFolder', 'EmailMode', 'EmailTo') {
        $row = 4 + [array]::IndexOf(@('Period', 'DataFolder', 'OutputFolder', 'EmailMode', 'EmailTo'), $n)
        $null = $wb.Names.Add($n, "=Settings!`$B`$$row")
    }
    $s.Range('A1').Font.Size = 16; $s.Range('A1').Font.Bold = $true
    $h = $s.Range('A3:C3'); $h.Font.Bold = $true; $h.Font.Color = 0xFFFFFF; $h.Interior.Color = 0x5F3A1F
    $s.Range('B4:B8').Interior.Color = 0xCCF2FF
    $s.Columns.Item(1).ColumnWidth = 16; $s.Columns.Item(2).ColumnWidth = 22; $s.Columns.Item(3).ColumnWidth = 80
    $btn = $s.Buttons().Add($s.Range('A10').Left, $s.Range('A10').Top, 160, 30)
    $btn.Caption = 'Build report'; $btn.OnAction = 'BuildFromSettings'
    $log = $wb.Worksheets.Add([Type]::Missing, $s); $log.Name = 'Log'
    $log.Range('A1:D1').Value2 = @('Run at', 'Period', 'Status', 'Detail')
    $log.Range('A1:D1').Font.Bold = $true; $log.Columns.Item(1).ColumnWidth = 18; $log.Columns.Item(4).ColumnWidth = 70
    $log.Columns.Item(1).NumberFormat = 'dd mmm yyyy hh:mm'
    if ($WithModule) { $null = $wb.VBProject.VBComponents.Import((Join-Path $root 'vba\ReportBuilder.bas')) }
    $s.Activate()
    $out = Join-Path $root 'MonthlyReport.xlsm'
    if (Test-Path $out) { Remove-Item $out }
    $wb.SaveAs($out, 52)
    $wb.Close($false)
    "wrote $out"
}
finally {
    $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
