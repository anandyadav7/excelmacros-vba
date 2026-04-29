Attribute VB_Name = "AddRunningTotalColumn"
Option Explicit

' Add Running Total Column
' Source: https://excelmacros.net/tools/add-running-total-column
' Offline. No API calls. No external dependencies.

' For a single column of numbers, writes a cumulative running total in the
' adjacent column to the right. Skips non-numeric cells (the running total
' carries forward but isn't written for those rows).

Public Sub AddRunningTotalColumn()
    Dim r As Range
    Dim cell As Range
    Dim ws As Worksheet
    Dim runningTotal As Double
    Dim countWritten As Long
    Dim destCol As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection
    If r.Columns.Count > 1 Then
        MsgBox "Select a single column of numbers.", vbExclamation
        Exit Sub
    End If

    Set ws = r.Worksheet
    destCol = r.Column + 1

    Application.ScreenUpdating = False
    runningTotal = 0
    countWritten = 0

    ' Header in the destination column.
    ws.Cells(r.Row, destCol).Value = "Running Total"
    ws.Cells(r.Row, destCol).Font.Bold = True

    For Each cell In r.Cells
        If cell.Row = r.Row Then
            ' header row - skip
        ElseIf IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            runningTotal = runningTotal + CDbl(cell.Value)
            ws.Cells(cell.Row, destCol).Value = runningTotal
            countWritten = countWritten + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Wrote running total to " & countWritten & " cell(s) in column " & _
           ws.Columns(destCol).Address(False, False) & ".", _
           vbInformation, "Add Running Total Column"
    Exit Sub

NoSelection:
    MsgBox "Select a single column of numbers first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
