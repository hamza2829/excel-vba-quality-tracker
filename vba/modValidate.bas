Attribute VB_Name = "modValidate"
Option Explicit

Public Const SHEET_ISSUES As String = "Issues"
Public Const SHEET_CLEAN As String = "Clean"

' Checks every row of "Data" in a fixed order. A row gets at most one issue (the first one found).
' Valid rows are copied to the sheet "Clean" with real numbers and dates.
' Returns the number of issues.
Public Function ValidateData() As Long
    Dim wsData As Worksheet, wsIssue As Worksheet, wsClean As Worksheet
    Dim lastRow As Long, r As Long, issues As Long, cleanRow As Long
    Dim lot As String, stepName As String, key As String
    Dim measured As Double, specMin As Double, specMax As Double
    Dim d As Date, reason As String
    Dim seen As Object

    Set wsData = ThisWorkbook.Worksheets(SHEET_DATA)
    Set wsIssue = GetOrCreateSheet(SHEET_ISSUES)
    Set wsClean = GetOrCreateSheet(SHEET_CLEAN)
    ClearSheet wsIssue
    ClearSheet wsClean

    Set seen = CreateObject("Scripting.Dictionary")

    wsIssue.Range("A1:C1").Value = Array("Row in Data", "Problem", "Row content")
    wsClean.Range("A1:H1").Value = Array("LotID", "ProcessStep", "Date", "Tool", "Measured", "SpecMin", "SpecMax", "Operator")

    Dim lastCell As Range
    Set lastCell = wsData.Cells.Find(What:="*", After:=wsData.Cells(1, 1), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    lastRow = 1
    If Not lastCell Is Nothing Then lastRow = lastCell.Row
    issues = 0: cleanRow = 1

    For r = 2 To lastRow
        reason = ""
        lot = Trim$(CStr(wsData.Cells(r, 1).Value))
        stepName = Trim$(CStr(wsData.Cells(r, 2).Value))

        If lot = "" Then
            reason = "Missing LotID"
        ElseIf Not TryParseNumber(CStr(wsData.Cells(r, 5).Value), measured) Then
            reason = "Measured value is not a number"
        ElseIf Not TryParseNumber(CStr(wsData.Cells(r, 6).Value), specMin) _
            Or Not TryParseNumber(CStr(wsData.Cells(r, 7).Value), specMax) Then
            reason = "Spec limit is not a number"
        ElseIf specMin > specMax Then
            reason = "SpecMin is greater than SpecMax"
        ElseIf Not TryParseIsoDate(CStr(wsData.Cells(r, 3).Value), d) Then
            reason = "Invalid date"
        Else
            key = lot & "|" & stepName
            If seen.Exists(key) Then
                reason = "Duplicate LotID and ProcessStep"
            Else
                seen.Add key, True
            End If
        End If

        If reason <> "" Then
            issues = issues + 1
            wsIssue.Cells(issues + 1, 1).Value = r
            wsIssue.Cells(issues + 1, 2).Value = reason
            wsIssue.Cells(issues + 1, 3).Value = RowAsText(wsData, r)
        Else
            cleanRow = cleanRow + 1
            wsClean.Cells(cleanRow, 1).Value = lot
            wsClean.Cells(cleanRow, 2).Value = stepName
            wsClean.Cells(cleanRow, 3).Value = d
            wsClean.Cells(cleanRow, 3).NumberFormat = "yyyy-mm-dd"
            wsClean.Cells(cleanRow, 4).Value = CStr(wsData.Cells(r, 4).Value)
            wsClean.Cells(cleanRow, 5).Value = measured
            wsClean.Cells(cleanRow, 6).Value = specMin
            wsClean.Cells(cleanRow, 7).Value = specMax
            wsClean.Cells(cleanRow, 8).Value = CStr(wsData.Cells(r, 8).Value)
        End If
    Next r

    wsIssue.Rows(1).Font.Bold = True
    wsClean.Rows(1).Font.Bold = True
    wsIssue.Columns("A:C").AutoFit
    wsClean.Columns("A:H").AutoFit
    ValidateData = issues
End Function

Private Function RowAsText(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim c As Long, s As String
    For c = 1 To 8
        If c > 1 Then s = s & ", "
        s = s & CStr(ws.Cells(r, c).Value)
    Next c
    RowAsText = s
End Function
