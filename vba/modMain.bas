Attribute VB_Name = "modMain"
Option Explicit

' Run this macro: pick the CSV, then Data, Issues, Clean and Report sheets are created.
Public Sub RunQualityTracker()
    Dim path As Variant, rowCount As Long, issues As Long

    path = Application.GetOpenFilename("CSV files (*.csv), *.csv", , "Choose the inspection CSV")
    If VarType(path) = vbBoolean Then Exit Sub        ' user cancelled

    Application.ScreenUpdating = False
    On Error GoTo Fail

    rowCount = ImportCsv(CStr(path))
    issues = ValidateData()
    BuildReport

    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(SHEET_REPORT).Activate
    MsgBox "Imported " & rowCount & " rows." & vbCrLf & _
           "Found " & issues & " problem rows (see sheet Issues)." & vbCrLf & _
           "The report is on the sheet Report.", vbInformation, "Quality Tracker"
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "Something went wrong: " & Err.Description, vbCritical, "Quality Tracker"
End Sub
