Attribute VB_Name = "LetterGradeFromScores"
Option Explicit

' Letter Grade From Numeric Scores
' Source: https://excelmacros.net/tools/letter-grade-from-scores
' Offline. No API calls. No external dependencies.

Public Sub LetterGradeFromScores()
    Dim r As Range
    Dim cell As Range
    Dim spec As String
    Dim parts() As String
    Dim aMin As Double, bMin As Double, cMin As Double, dMin As Double
    Dim score As Double
    Dim grade As String
    Dim destCol As Long
    Dim ws As Worksheet
    Dim countAdded As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection
    If r.Columns.Count > 1 Then
        MsgBox "Select a single column of scores.", vbExclamation
        Exit Sub
    End If

    spec = InputBox( _
        "Enter four cutoffs separated by commas: A min, B min, C min, D min." & vbCrLf & _
        "Anything below the D minimum is F." & vbCrLf & vbCrLf & _
        "Default scale: 90,80,70,60", _
        "Letter Grade Calculator", "90,80,70,60")
    If Len(Trim$(spec)) = 0 Then Exit Sub

    parts = Split(spec, ",")
    If UBound(parts) - LBound(parts) <> 3 Then
        MsgBox "Enter exactly four numbers separated by commas.", vbExclamation
        Exit Sub
    End If
    If Not (IsNumeric(Trim$(parts(0))) And IsNumeric(Trim$(parts(1))) _
            And IsNumeric(Trim$(parts(2))) And IsNumeric(Trim$(parts(3)))) Then
        MsgBox "All four cutoffs must be numbers.", vbExclamation
        Exit Sub
    End If
    aMin = CDbl(Trim$(parts(0)))
    bMin = CDbl(Trim$(parts(1)))
    cMin = CDbl(Trim$(parts(2)))
    dMin = CDbl(Trim$(parts(3)))
    If Not (aMin > bMin And bMin > cMin And cMin > dMin) Then
        MsgBox "Cutoffs must be strictly decreasing (e.g. 90 > 80 > 70 > 60).", vbExclamation
        Exit Sub
    End If

    Set ws = r.Worksheet
    destCol = r.Column + 1

    Application.ScreenUpdating = False
    countAdded = 0

    For Each cell In r.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            score = CDbl(cell.Value)
            If score >= aMin Then
                grade = "A"
            ElseIf score >= bMin Then
                grade = "B"
            ElseIf score >= cMin Then
                grade = "C"
            ElseIf score >= dMin Then
                grade = "D"
            Else
                grade = "F"
            End If
            ws.Cells(cell.Row, destCol).Value = grade
            countAdded = countAdded + 1
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "Wrote " & countAdded & " letter grade(s) to column " & destCol & ".", vbInformation, _
           "Letter Grade Calculator"
    Exit Sub

NoSelection:
    MsgBox "Select a single column of numeric scores first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
