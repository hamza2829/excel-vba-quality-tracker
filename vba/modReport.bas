Attribute VB_Name = "modReport"
Option Explicit

Public Const SHEET_REPORT As String = "Report"
Public Const YIELD_ALERT As Double = 95   ' yield below this value is marked red

' Builds one line per process step from the sheet "Clean":
' count, out of spec, yield percent, mean, min, max. Adds a chart.
Public Sub BuildReport()
    Dim wsClean As Worksheet, wsRep As Worksheet
    Dim lastRow As Long, r As Long, k As Variant, outRow As Long
    Dim stepName As String, v As Double
    Dim cnt As Object, oos As Object, sums As Object, mins As Object, maxs As Object

    Set wsClean = ThisWorkbook.Worksheets(SHEET_CLEAN)
    Set wsRep = GetOrCreateSheet(SHEET_REPORT)
    ClearSheet wsRep

    Set cnt = CreateObject("Scripting.Dictionary")
    Set oos = CreateObject("Scripting.Dictionary")
    Set sums = CreateObject("Scripting.Dictionary")
    Set mins = CreateObject("Scripting.Dictionary")
    Set maxs = CreateObject("Scripting.Dictionary")

    lastRow = wsClean.Cells(wsClean.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        stepName = CStr(wsClean.Cells(r, 2).Value)
        v = CDbl(wsClean.Cells(r, 5).Value)
        If Not cnt.Exists(stepName) Then
            cnt.Add stepName, 0: oos.Add stepName, 0: sums.Add stepName, 0#
            mins.Add stepName, v: maxs.Add stepName, v
        End If
        cnt(stepName) = cnt(stepName) + 1
        sums(stepName) = sums(stepName) + v
        If v < mins(stepName) Then mins(stepName) = v
        If v > maxs(stepName) Then maxs(stepName) = v
        If IsOutOfSpec(v, CDbl(wsClean.Cells(r, 6).Value), CDbl(wsClean.Cells(r, 7).Value)) Then
            oos(stepName) = oos(stepName) + 1
        End If
    Next r

    wsRep.Range("A1:G1").Value = Array("ProcessStep", "Measurements", "Out of spec", "Yield %", "Mean", "Min", "Max")
    outRow = 1
    For Each k In cnt.Keys
        outRow = outRow + 1
        wsRep.Cells(outRow, 1).Value = k
        wsRep.Cells(outRow, 2).Value = cnt(k)
        wsRep.Cells(outRow, 3).Value = oos(k)
        wsRep.Cells(outRow, 4).Value = YieldPercent(CLng(cnt(k)), CLng(oos(k)))
        wsRep.Cells(outRow, 5).Value = Round(sums(k) / cnt(k), 3)
        wsRep.Cells(outRow, 6).Value = mins(k)
        wsRep.Cells(outRow, 7).Value = maxs(k)
    Next k

    wsRep.Rows(1).Font.Bold = True
    wsRep.Columns("A:G").AutoFit

    If outRow >= 2 Then
        Dim rng As Range
        Set rng = wsRep.Range(wsRep.Cells(2, 4), wsRep.Cells(outRow, 4))
        rng.FormatConditions.Delete
        With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=" & YIELD_ALERT)
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
        End With
        AddYieldChart wsRep, outRow
    End If
End Sub

Private Sub AddYieldChart(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=ws.Range("I2").Left, Top:=ws.Range("I2").Top, Width:=380, Height:=220)
    With co.Chart
        .ChartType = xlColumnClustered
        With .SeriesCollection.NewSeries
            .Name = "Yield %"
            .Values = ws.Range(ws.Cells(2, 4), ws.Cells(lastRow, 4))
            .XValues = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))
        End With
        .HasTitle = True
        .ChartTitle.Text = "Yield per process step"
        .HasLegend = False
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScale = 100
    End With
End Sub

' Saves the sheet "Report" as a PDF next to the workbook.
Public Sub ExportReportPdf()
    Dim p As String
    p = ThisWorkbook.Path
    If p = "" Then p = CurDir$
    ThisWorkbook.Worksheets(SHEET_REPORT).ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=p & Application.PathSeparator & "quality_report.pdf", Quality:=xlQualityStandard
End Sub
