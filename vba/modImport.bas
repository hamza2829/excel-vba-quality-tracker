Attribute VB_Name = "modImport"
Option Explicit

Public Const SHEET_DATA As String = "Data"

Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Public Sub ClearSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co
End Sub

' Reads a comma separated file into the sheet "Data". Returns the number of data rows.
Public Function ImportCsv(ByVal filePath As String) As Long
    Dim ws As Worksheet, ff As Integer, lineText As String
    Dim parts() As String, r As Long, c As Long, n As Long
    Dim errorNumber As Long, errorText As String

    If Dir(filePath) = "" Then
        Err.Raise vbObjectError + 1000, "ImportCsv", "File not found: " & filePath
    End If

    Set ws = GetOrCreateSheet(SHEET_DATA)
    ClearSheet ws
    ws.Cells.NumberFormat = "@"          ' keep everything as text, validation decides later

    ff = FreeFile
    On Error GoTo ImportFailed
    Open filePath For Input As #ff
    r = 0
    Do While Not EOF(ff)
        Line Input #ff, lineText
        If Len(Trim$(lineText)) > 0 Then
            r = r + 1
            parts = Split(lineText, ",")
            If UBound(parts) <> 7 Or InStr(lineText, Chr$(34)) > 0 Then
                Err.Raise vbObjectError + 1001, "ImportCsv", "Expected 8 unquoted columns at record " & r
            End If
            If r = 1 Then
                If lineText <> "LotID,ProcessStep,Date,Tool,Measured,SpecMin,SpecMax,Operator" Then
                    Err.Raise vbObjectError + 1002, "ImportCsv", "CSV header does not match the expected columns."
                End If
            End If
            For c = 0 To UBound(parts)
                ws.Cells(r, c + 1).Value = parts(c)
            Next c
        End If
    Loop
    Close #ff
    ff = 0
    If r = 0 Then Err.Raise vbObjectError + 1003, "ImportCsv", "CSV file is empty."

    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Columns("A:H").AutoFit
    ImportCsv = IIf(r > 0, r - 1, 0)
    Exit Function
ImportFailed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If ff <> 0 Then Close #ff
    On Error GoTo 0
    Err.Raise errorNumber, "ImportCsv", errorText
End Function
