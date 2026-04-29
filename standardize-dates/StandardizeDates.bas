Attribute VB_Name = "StandardizeDates"
Option Explicit

' Standardize Mixed Date Formats
' Source: https://excelmacros.net/tools/standardize-dates
' Offline. No API calls. No external dependencies.

' Asks the user which format their text dates are currently in (US M/D/Y,
' EU/UK D/M/Y, or ISO Y/M/D), parses each cell accordingly, and writes back
' as a real Excel date with YYYY-MM-DD display format.

Public Sub StandardizeDates()
    Dim r As Range
    Dim cell As Range
    Dim formatAns As String
    Dim formatType As Long
    Dim originalText As String
    Dim parts() As String
    Dim sep As String
    Dim parsedDate As Date
    Dim convertedCount As Long
    Dim skippedCount As Long
    Dim y As Long, m As Long, d As Long
    Dim parseOk As Boolean

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    formatAns = InputBox( _
        "Which format are your dates currently in?" & vbCrLf & vbCrLf & _
        "1 = M/D/Y or M-D-Y (US, like 04/28/2026)" & vbCrLf & _
        "2 = D/M/Y or D-M-Y (UK, EU, India, like 28/04/2026)" & vbCrLf & _
        "3 = Y/M/D or Y-M-D (ISO, like 2026-04-28)" & vbCrLf & vbCrLf & _
        "Type 1, 2, or 3:", _
        "Standardize Dates", "1")
    If StrPtr(formatAns) = 0 Then Exit Sub
    If Not IsNumeric(Trim$(formatAns)) Then
        MsgBox "Type 1, 2, or 3.", vbExclamation
        Exit Sub
    End If
    formatType = CLng(Trim$(formatAns))
    If formatType < 1 Or formatType > 3 Then
        MsgBox "Type 1, 2, or 3.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    convertedCount = 0
    skippedCount = 0

    For Each cell In r.Cells
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            originalText = Trim$(CStr(cell.Value))
            sep = DetectSeparator(originalText)
            parseOk = False
            If Len(sep) > 0 Then
                parts = Split(originalText, sep)
                If UBound(parts) - LBound(parts) = 2 Then
                    Select Case formatType
                        Case 1: m = SafeLong(parts(0)): d = SafeLong(parts(1)): y = SafeLong(parts(2))
                        Case 2: d = SafeLong(parts(0)): m = SafeLong(parts(1)): y = SafeLong(parts(2))
                        Case 3: y = SafeLong(parts(0)): m = SafeLong(parts(1)): d = SafeLong(parts(2))
                    End Select
                    If y > 0 And y < 100 Then y = y + 2000
                    If y >= 1900 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
                        On Error Resume Next
                        parsedDate = DateSerial(y, m, d)
                        If Err.Number = 0 Then parseOk = True
                        Err.Clear
                        On Error GoTo CleanFail
                    End If
                End If
            End If
            If parseOk Then
                cell.NumberFormat = "yyyy-mm-dd"
                cell.Value = parsedDate
                convertedCount = convertedCount + 1
            Else
                skippedCount = skippedCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "Standardized " & convertedCount & " date(s) to YYYY-MM-DD." & vbCrLf & _
           "Skipped " & skippedCount & " unparseable cell(s).", _
           vbInformation, "Standardize Dates"
    Exit Sub

NoSelection:
    MsgBox "Select the range of date text to standardize first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function DetectSeparator(ByVal s As String) As String
    If InStr(s, "/") > 0 Then DetectSeparator = "/": Exit Function
    If InStr(s, "-") > 0 Then DetectSeparator = "-": Exit Function
    If InStr(s, ".") > 0 Then DetectSeparator = ".": Exit Function
    DetectSeparator = ""
End Function

Private Function SafeLong(ByVal s As String) As Long
    If IsNumeric(Trim$(s)) Then SafeLong = CLng(Trim$(s)) Else SafeLong = 0
End Function
