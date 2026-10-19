Attribute VB_Name = "modLogic"
Option Explicit

' Pure helper functions. No worksheet access, so they can be tested easily.

Public Function TryParseNumber(ByVal text As String, ByRef result As Double) As Boolean
    ' Accepts digits, one decimal point and an optional leading minus.
    ' Always uses the dot as decimal separator, independent of the Windows region.
    Dim t As String, i As Long, ch As String, dots As Long, digits As Long
    t = Trim$(text)
    If Len(t) = 0 Then Exit Function
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch = "-" And i = 1 Then
            ' sign is fine at the start
        ElseIf ch = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        ElseIf ch >= "0" And ch <= "9" Then
            digits = digits + 1
        Else
            Exit Function
        End If
    Next i
    If digits = 0 Then Exit Function
    On Error GoTo InvalidNumber
    result = Val(t)          ' Val always uses the dot
    TryParseNumber = True
InvalidNumber:
End Function

Public Function TryParseIsoDate(ByVal text As String, ByRef result As Date) As Boolean
    ' Expects yyyy-mm-dd and checks that the date really exists.
    Dim t As String, y As Long, m As Long, d As Long
    t = Trim$(text)
    If Len(t) <> 10 Then Exit Function
    If Mid$(t, 5, 1) <> "-" Or Mid$(t, 8, 1) <> "-" Then Exit Function
    If Not IsDigits(Mid$(t, 1, 4)) Or Not IsDigits(Mid$(t, 6, 2)) Or Not IsDigits(Mid$(t, 9, 2)) Then Exit Function
    y = CLng(Mid$(t, 1, 4)): m = CLng(Mid$(t, 6, 2)): d = CLng(Mid$(t, 9, 2))
    If y < 100 Or y > 9999 Then Exit Function
    If m < 1 Or m > 12 Or d < 1 Then Exit Function
    If d > DaysInMonth(y, m) Then Exit Function
    result = DateSerial(y, m, d)
    TryParseIsoDate = True
End Function

Public Function DaysInMonth(ByVal y As Long, ByVal m As Long) As Long
    Select Case m
        Case 4, 6, 9, 11: DaysInMonth = 30
        Case 2
            DaysInMonth = 28
            If y Mod 400 = 0 Or (y Mod 4 = 0 And y Mod 100 <> 0) Then DaysInMonth = 29
        Case 1, 3, 5, 7, 8, 10, 12: DaysInMonth = 31
    End Select
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsDigits = True
End Function

Public Function IsOutOfSpec(ByVal value As Double, ByVal specMin As Double, ByVal specMax As Double) As Boolean
    IsOutOfSpec = (value < specMin) Or (value > specMax)
End Function

Public Function YieldPercent(ByVal total As Long, ByVal outOfSpec As Long) As Double
    If total <= 0 Then
        YieldPercent = 0
    Else
        YieldPercent = Round(100# * (total - outOfSpec) / total, 1)
    End If
End Function
