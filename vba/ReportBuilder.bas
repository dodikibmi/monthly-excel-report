Attribute VB_Name = "ReportBuilder"
' Monthly sales report: SQL extract (CSV) -> KPI summary, pivots, chart -> .xlsx + PDF (+ optional Outlook draft).
' Entry points:
'   BuildFromSettings            button on the Settings sheet (reads the named cells)
'   BuildMonthlyReport "2026-10" for a scheduler / script (Application.Run); returns the report path,
'                                or "FAILED: <reason>" (never raises: an error thrown back to Application.Run
'                                closes a hidden Excel instead of reaching the caller)
Option Explicit

Private Const XL_DATABASE As Long = 1
Private Const XL_ROW_FIELD As Long = 1
Private Const XL_COLUMN_FIELD As Long = 2
Private Const XL_SUM As Long = -4157
Private Const XL_OPENXML As Long = 51
Private Const XL_PDF As Long = 0
Private Const XL_COLUMN_CLUSTERED As Long = 51

Public Sub BuildFromSettings()
    Dim result As String
    result = BuildMonthlyReport(CStr(Setting("Period")))
    If Left$(result, 7) = "FAILED:" Then
        MsgBox result & vbLf & vbLf & "Details are on the Log sheet.", vbCritical, "Monthly sales report"
    Else
        MsgBox "Report saved:" & vbLf & result, vbInformation, "Monthly sales report"
    End If
End Sub

Public Function BuildMonthlyReport(ByVal period As String) As String
    Dim t0 As Double: t0 = Timer
    Dim dataDir As String, outDir As String, wb As Workbook, data As Worksheet, lo As ListObject
    Dim cur As Variant, prev As Variant, outFile As String, lineCount As Long

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    dataDir = FolderOf(Setting("DataFolder"))
    outDir = FolderOf(Setting("OutputFolder"))
    If Dir(dataDir & "extract-" & period & ".csv") = "" Then Err.Raise vbObjectError + 1, , "No extract for " & period & " in " & dataDir

    Set wb = Workbooks.Add
    Set data = wb.Worksheets(1)
    data.Name = "Data"
    Set lo = LoadCsv(dataDir & "extract-" & period & ".csv", data)
    cur = Totals(lo)
    lineCount = lo.ListRows.Count
    prev = TotalsFromCsv(dataDir & "extract-" & PreviousPeriod(period) & ".csv")

    WriteSummary wb, period, cur, prev
    AddPivots wb, lo
    wb.Worksheets("Summary").Activate

    If Dir(outDir, vbDirectory) = "" Then MkDir outDir
    outFile = outDir & "Sales-Report-" & period & ".xlsx"
    wb.SaveAs outFile, XL_OPENXML
    wb.Worksheets("Summary").ExportAsFixedFormat XL_PDF, outDir & "Sales-Report-" & period & ".pdf"
    If LCase$(Setting("EmailMode")) = "draft" Then OutlookDraft period, cur, outFile, outDir & "Sales-Report-" & period & ".pdf"
    Set lo = Nothing: Set data = Nothing          ' release references before the workbook goes away
    wb.Close False
    Set wb = Nothing

    LogLine period, "OK", Format(Timer - t0, "0.0") & " s, " & lineCount & " lines, revenue " & Format(cur(0), "#,##0.00")
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    BuildMonthlyReport = outFile
    Exit Function

Fail:
    Dim msg As String: msg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    LogLine period, "FAILED", msg
    BuildMonthlyReport = "FAILED: " & msg
End Function

' ---------- data ----------

Private Function LoadCsv(ByVal file As String, ByVal target As Worksheet) As ListObject
    Dim src As Workbook, n As Long
    Workbooks.OpenText Filename:=file, Origin:=65001, DataType:=1, Comma:=True, Local:=False
    Set src = ActiveWorkbook
    src.Worksheets(1).UsedRange.Copy target.Range("A1")
    src.Close False
    n = target.Cells(target.Rows.Count, 1).End(xlUp).Row
    Set LoadCsv = target.ListObjects.Add(1, target.Range("A1").Resize(n, 10), , 1)
    LoadCsv.Name = "Sales"
    LoadCsv.TableStyle = "TableStyleMedium2"
    LoadCsv.ListColumns("Date").DataBodyRange.NumberFormat = "dd mmm yyyy"
    LoadCsv.ListColumns("Revenue").DataBodyRange.NumberFormat = "#,##0.00"
    LoadCsv.ListColumns("Margin").DataBodyRange.NumberFormat = "#,##0.00"
    target.Columns("A:J").AutoFit
End Function

' Returns Array(revenue, margin, distinct orders)
Private Function Totals(ByVal lo As ListObject) As Variant
    Dim seen As Object, r As Range, rev As Double, mar As Double
    Set seen = CreateObject("Scripting.Dictionary")
    For Each r In lo.DataBodyRange.Rows
        rev = rev + r.Cells(1, 9).Value
        mar = mar + r.Cells(1, 10).Value
        seen(CStr(r.Cells(1, 1).Value)) = True
    Next
    Totals = Array(Round(rev, 2), Round(mar, 2), seen.Count)
End Function

Private Function TotalsFromCsv(ByVal file As String) As Variant
    Dim tmp As Workbook, lo As ListObject
    If Dir(file) = "" Then TotalsFromCsv = Empty: Exit Function
    Set tmp = Workbooks.Add
    Set lo = LoadCsv(file, tmp.Worksheets(1))
    TotalsFromCsv = Totals(lo)
    tmp.Close False
End Function

' ---------- report sheets ----------

Private Sub WriteSummary(ByVal wb As Workbook, ByVal period As String, ByVal cur As Variant, ByVal prev As Variant)
    Dim ws As Worksheet, i As Long, labels As Variant, nowV As Variant, prevV As Variant, hasPrev As Boolean
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = "Summary"
    hasPrev = Not IsEmpty(prev)
    ActiveWindow.DisplayGridlines = False

    ws.Range("A1").Value = "Harbor Supplies - monthly sales report"
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Period " & period & IIf(hasPrev, "  |  compared with " & PreviousPeriod(period), "") & _
                           "  |  generated " & Format(Now, "dd mmm yyyy hh:nn")
    ws.Range("A2").Font.Color = RGB(90, 90, 90)

    ws.Range("A4:D4").Value = Array("Measure", period, PreviousPeriod(period), "Change")
    labels = Array("Revenue", "Gross margin", "Margin %", "Orders", "Average order value")
    nowV = Array(cur(0), cur(1), SafeDiv(cur(1), cur(0)), cur(2), SafeDiv(cur(0), cur(2)))
    If hasPrev Then prevV = Array(prev(0), prev(1), SafeDiv(prev(1), prev(0)), prev(2), SafeDiv(prev(0), prev(2)))
    For i = 0 To 4
        ws.Cells(5 + i, 1).Value = labels(i)
        ws.Cells(5 + i, 2).Value = nowV(i)
        If hasPrev Then
            ws.Cells(5 + i, 3).Value = prevV(i)
            If i = 2 Then
                ws.Cells(5 + i, 4).Value = nowV(i) - prevV(i)          ' percentage points
                ws.Cells(5 + i, 4).NumberFormat = "+0.0"" pp"";-0.0"" pp"""
            Else
                ws.Cells(5 + i, 4).Value = SafeDiv(nowV(i) - prevV(i), prevV(i))
                ws.Cells(5 + i, 4).NumberFormat = "+0.0%;-0.0%"
            End If
            ws.Cells(5 + i, 4).Font.Color = IIf(ws.Cells(5 + i, 4).Value < 0, RGB(192, 0, 0), RGB(0, 128, 0))
        End If
    Next
    ws.Range("B5:C6,B9:C9").NumberFormat = "#,##0.00"
    ws.Range("B7:C7").NumberFormat = "0.0%"
    ws.Range("B8:C8").NumberFormat = "#,##0"
    If hasPrev Then ws.Range("D7").Value = (nowV(2) - prevV(2)) * 100

    With ws.Range("A4:D4")
        .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(31, 58, 95)
    End With
    ws.Range("A4:D9").Borders.Color = RGB(200, 200, 200)
    ws.Range("A11").Value = "Returns are negative revenue in the month they come back; cancelled orders are excluded."
    ws.Range("A11").Font.Italic = True: ws.Range("A11").Font.Color = RGB(90, 90, 90)
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B:D").ColumnWidth = 16
End Sub

Private Sub AddPivots(ByVal wb As Workbook, ByVal lo As ListObject)
    Dim pc As PivotCache, ws As Worksheet, pt As PivotTable, ch As ChartObject, top As Worksheet, pt2 As PivotTable

    Set pc = wb.PivotCaches.Create(XL_DATABASE, lo.Range)

    ' Region x category
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets("Summary"))
    ws.Name = "By region"
    Set pt = pc.CreatePivotTable(ws.Range("A3"), "RegionCategory")
    pt.PivotFields("Region").Orientation = XL_ROW_FIELD
    pt.PivotFields("Category").Orientation = XL_COLUMN_FIELD
    pt.AddDataField(pt.PivotFields("Revenue"), "Revenue ", XL_SUM).NumberFormat = "#,##0"
    pt.TableStyle2 = "PivotStyleMedium2"
    ws.Range("A1").Value = "Revenue by region and product category"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 13

    ' Chart on the summary sheet, fed by a small region table (a chart on the pivot would move with it)
    Dim sumWs As Worksheet, r As Long, region As Variant
    Set sumWs = wb.Worksheets("Summary")
    sumWs.Range("F4:G4").Value = Array("Region", "Revenue")
    r = 5
    For Each region In Array("North", "South", "East", "West")
        sumWs.Cells(r, 6).Value = region
        sumWs.Cells(r, 7).Value = Application.WorksheetFunction.SumIfs(lo.ListColumns("Revenue").DataBodyRange, _
                                                                     lo.ListColumns("Region").DataBodyRange, region)
        r = r + 1
    Next
    sumWs.Range("G5:G8").NumberFormat = "#,##0"
    sumWs.Range("F4:G4").Font.Bold = True
    Set ch = sumWs.ChartObjects.Add(sumWs.Range("A13").Left, sumWs.Range("A13").Top, 520, 260)
    ch.Name = "RevenueByRegion"
    With ch.Chart
        .ChartType = XL_COLUMN_CLUSTERED
        .SetSourceData sumWs.Range("F4:G8")
        .HasTitle = True
        .ChartTitle.Text = "Revenue by region"
        .HasLegend = False
    End With

    ' Top 10 customers
    Set top = wb.Worksheets.Add(After:=ws)
    top.Name = "Top customers"
    Set pt2 = pc.CreatePivotTable(top.Range("A3"), "TopCustomers")
    pt2.PivotFields("Customer").Orientation = XL_ROW_FIELD
    pt2.AddDataField(pt2.PivotFields("Revenue"), "Revenue ", XL_SUM).NumberFormat = "#,##0"
    pt2.AddDataField(pt2.PivotFields("Margin"), "Margin ", XL_SUM).NumberFormat = "#,##0"
    pt2.PivotFields("Customer").AutoSort 2, "Revenue "                 ' descending
    pt2.PivotFields("Customer").PivotFilters.Add2 Type:=1, DataField:=pt2.DataFields("Revenue "), Value1:=10  ' top 10
    pt2.TableStyle2 = "PivotStyleMedium2"
    top.Range("A1").Value = "Top 10 customers by revenue"
    top.Range("A1").Font.Bold = True: top.Range("A1").Font.Size = 13
End Sub

' ---------- email ----------

Private Sub OutlookDraft(ByVal period As String, ByVal cur As Variant, ByVal xlsx As String, ByVal pdf As String)
    Dim ol As Object, m As Object
    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then LogLine period, "NOTE", "Outlook not available, no email draft": Exit Sub
    Set m = ol.CreateItem(0)
    m.To = Setting("EmailTo")
    m.Subject = "Sales report " & period & ": revenue " & Format(cur(0), "#,##0")
    m.HTMLBody = "<p>Hello,</p><p>Attached is the sales report for " & period & ": revenue <b>" & Format(cur(0), "#,##0.00") & _
                 "</b>, gross margin " & Format(cur(1), "#,##0.00") & ", " & cur(2) & " orders.</p><p>Regards</p>"
    m.Attachments.Add xlsx
    m.Attachments.Add pdf
    m.Save                          ' a draft: a person reads it and presses Send
End Sub

' ---------- helpers ----------

Private Function Setting(ByVal name As String) As Variant
    Setting = ThisWorkbook.Names(name).RefersToRange.Value
End Function

Private Function FolderOf(ByVal p As String) As String
    If p = "" Or p = "." Then p = ThisWorkbook.Path
    If Left$(p, 2) = ".\" Then p = ThisWorkbook.Path & Mid$(p, 2)
    If Right$(p, 1) <> "\" Then p = p & "\"
    FolderOf = p
End Function

Private Function PreviousPeriod(ByVal period As String) As String
    PreviousPeriod = Format(DateSerial(CInt(Left$(period, 4)), CInt(Right$(period, 2)) - 1, 1), "yyyy-mm")
End Function

Private Function SafeDiv(ByVal a As Double, ByVal b As Double) As Double
    If b = 0 Then SafeDiv = 0 Else SafeDiv = a / b
End Function

Private Sub LogLine(ByVal period As String, ByVal status As String, ByVal detail As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Log")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Resize(1, 4).Value = Array(Now, period, status, detail)
End Sub
