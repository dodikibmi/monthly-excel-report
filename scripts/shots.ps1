# Renders docs/*.png from the generated report (each sheet -> PDF page -> PNG) and the Settings sheet of the macro workbook.
param([string]$Period = '2026-10')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$tmp = Join-Path $env:TEMP 'mer-shots'; New-Item -ItemType Directory -Force $tmp | Out-Null
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open((Join-Path $root "out\Sales-Report-$Period.xlsx"), 0, $true)
    foreach ($s in @('Summary', 'By region', 'Top customers', 'Data')) {
        $ws = $wb.Worksheets.Item($s)
        $ws.PageSetup.Orientation = 2; $ws.PageSetup.Zoom = $false
        $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1
        if ($s -eq 'Data') { $ws.PageSetup.PrintArea = 'A1:J25' }
        $ws.ExportAsFixedFormat(0, (Join-Path $tmp (($s -replace ' ', '-') + '.pdf')))
    }
    $wb.Close($false)
    $m = $xl.Workbooks.Open((Join-Path $root 'MonthlyReport.xlsm'), 0, $true)
    foreach ($s in @('Settings', 'Log')) {
        $ws = $m.Worksheets.Item($s); $ws.PageSetup.Orientation = 2; $ws.PageSetup.Zoom = $false
        $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1
        $ws.ExportAsFixedFormat(0, (Join-Path $tmp "$s.pdf"))
    }
    $m.Close($false)
}
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
python -c @"
import fitz, pathlib
tmp = pathlib.Path(r'$tmp'); docs = pathlib.Path(r'$root') / 'docs'
for pdf in tmp.glob('*.pdf'):
    page = fitz.open(pdf)[0]
    r = fitz.Rect(page.rect); bb = None
    for b in page.get_drawings(): bb = b['rect'] if bb is None else bb | b['rect']
    for b in page.get_text('blocks'): bb = fitz.Rect(b[:4]) if bb is None else bb | fitz.Rect(b[:4])
    bb = (bb or r) + (-12, -12, 12, 12)
    page.get_pixmap(dpi=150, clip=bb & r).save(docs / (pdf.stem.lower() + '.png'))
    print('docs/' + pdf.stem.lower() + '.png')
"@
