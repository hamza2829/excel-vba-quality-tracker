Attribute VB_Name = "modTests"
Option Explicit

Private passed As Long, failed As Long

Private Sub AssertTrue(ByVal condition As Boolean, ByVal name As String)
    If condition Then
        passed = passed + 1
    Else
        failed = failed + 1
        Debug.Print "FAIL: " & name
    End If
End Sub

' Run RunAllTests, then open the Immediate window (Ctrl+G) to see the result.
Public Sub RunAllTests(Optional ByVal showMessage As Boolean = True)
    passed = 0: failed = 0
    Test_TryParseNumber
    Test_TryParseIsoDate
    Test_IsOutOfSpec
    Test_YieldPercent
    Debug.Print "Tests passed: " & passed & ", failed: " & failed
    If showMessage Then MsgBox "Tests passed: " & passed & ", failed: " & failed, IIf(failed = 0, vbInformation, vbExclamation), "Unit tests"
    If failed > 0 Then Err.Raise vbObjectError + 1100, "RunAllTests", CStr(failed) & " tests failed"
End Sub

Private Sub Test_TryParseNumber()
    Dim v As Double
    AssertTrue TryParseNumber("12.5", v) And v = 12.5, "12.5 is parsed"
    AssertTrue TryParseNumber("-3", v) And v = -3, "negative integer"
    AssertTrue Not TryParseNumber("n/a", v), "text is rejected"
    AssertTrue Not TryParseNumber("", v), "empty is rejected"
    AssertTrue Not TryParseNumber("1.2.3", v), "two dots are rejected"
    AssertTrue Not TryParseNumber("-", v), "only a minus is rejected"
    AssertTrue Not TryParseNumber(String$(400, "9"), v), "overflow is rejected"
End Sub

Private Sub Test_TryParseIsoDate()
    Dim d As Date
    AssertTrue TryParseIsoDate("2026-09-01", d) And d = DateSerial(2026, 9, 1), "valid date"
    AssertTrue TryParseIsoDate("2028-02-29", d), "leap day is valid"
    AssertTrue Not TryParseIsoDate("2027-02-29", d), "no leap day in 2027"
    AssertTrue Not TryParseIsoDate("2026-13-45", d), "month 13 is rejected"
    AssertTrue Not TryParseIsoDate("01.09.2026", d), "wrong format is rejected"
    AssertTrue Not TryParseIsoDate("0000-01-01", d), "year zero is rejected"
    AssertTrue Not TryParseIsoDate("0099-01-01", d), "two digit year conversion is rejected"
    AssertTrue TryParseIsoDate("9999-12-31", d), "last supported date"
    AssertTrue Not TryParseIsoDate("1900-02-29", d), "century is not a leap year"
    AssertTrue TryParseIsoDate("2000-02-29", d), "400 year leap day"
End Sub

Private Sub Test_IsOutOfSpec()
    AssertTrue Not IsOutOfSpec(50, 48, 52), "inside spec"
    AssertTrue Not IsOutOfSpec(48, 48, 52), "lower limit is inside"
    AssertTrue IsOutOfSpec(52.1, 48, 52), "above spec"
    AssertTrue IsOutOfSpec(47.9, 48, 52), "below spec"
End Sub

Private Sub Test_YieldPercent()
    AssertTrue YieldPercent(100, 0) = 100, "no failures"
    AssertTrue YieldPercent(61, 5) = 91.8, "61 measurements, 5 out of spec"
    AssertTrue YieldPercent(0, 0) = 0, "no measurements"
End Sub
