Attribute VB_Name = "HighlightTopBottomValues"
Option Explicit

' Highlight Top and Bottom Values
' Source: https://excelmacros.net/tools/highlight-top-bottom-values
' Offline. No API calls. No external dependencies.

' Asks for a number N, then colors the top N values green and the bottom N
' values red in the selection. Uses Excel's LARGE and SMALL functions to find
' the threshold values, then paints any cell at or above/below the threshold.

Public Sub HighlightTopBottomValues()
    Dim r As Range
    Dim cell As Range
    Dim nAns As String
    Dim n As Long
    Dim topThreshold As Double
    Dim bottomThreshold As Double
    Dim numericCount As Long
    Dim topPainted As Long
    Dim bottomPainted As Long
    Dim v As Double

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    nAns = InputBox( _
        "How many top and bottom values to highlight?" & vbCrLf & _
        "(e.g. 5 = top 5 green + bottom 5 red)", _
        "Highlight Top and Bottom Values", "5")
    If StrPtr(nAns) = 0 Then Exit Sub
    If Not IsNumeric(Trim$(nAns)) Then
        MsgBox "Type a whole number like 5.", vbExclamation
        Exit Sub
    End If
    n = CLng(Trim$(nAns))
    If n < 1 Then
        MsgBox "N must be at least 1.", vbExclamation
        Exit Sub
    End If

    ' Count numeric cells.
    numericCount = 0
    For Each cell In r.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            numericCount = numericCount + 1
        End If
    Next cell

    If numericCount < 2 Then
        MsgBox "Need at least 2 numeric cells in the selection.", vbExclamation
        Exit Sub
    End If

    If n > numericCount Then n = numericCount

    On Error Resume Next
    topThreshold = Application.WorksheetFunction.Large(r, n)
    bottomThreshold = Application.WorksheetFunction.Small(r, n)
    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    topPainted = 0
    bottomPainted = 0

    For Each cell In r.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            v = CDbl(cell.Value)
            If v >= topThreshold Then
                cell.Interior.Color = RGB(198, 239, 206) ' light green
                cell.Font.Color = RGB(0, 97, 0)
                topPainted = topPainted + 1
            ElseIf v <= bottomThreshold Then
                cell.Interior.Color = RGB(255, 199, 206) ' light red
                cell.Font.Color = RGB(156, 0, 6)
                bottomPainted = bottomPainted + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Highlighted " & topPainted & " top value(s) green and " & _
           bottomPainted & " bottom value(s) red.", _
           vbInformation, "Highlight Top and Bottom Values"
    Exit Sub

NoSelection:
    MsgBox "Select the range of numbers to scan first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
