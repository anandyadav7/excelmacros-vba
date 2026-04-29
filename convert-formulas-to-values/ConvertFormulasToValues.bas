Attribute VB_Name = "ConvertFormulasToValues"
Option Explicit

' Convert Formulas to Values
' Source: https://excelmacros.net/tools/convert-formulas-to-values
' Offline. No API calls. No external dependencies.

' Walks the selection and replaces every formula cell with its current value
' in place, preserving the cell's display format (dates stay dates, currency
' stays currency). Array formula cells that can't be assigned per-cell are
' skipped and reported.

Public Sub ConvertFormulasToValues()
    Dim r As Range
    Dim cell As Range
    Dim convertedCount As Long
    Dim skippedCount As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    convertedCount = 0
    skippedCount = 0

    For Each cell In r.Cells
        If cell.HasFormula Then
            On Error Resume Next
            cell.Value = cell.Value
            If Err.Number = 0 Then
                convertedCount = convertedCount + 1
            Else
                skippedCount = skippedCount + 1
                Err.Clear
            End If
            On Error GoTo CleanFail
        End If
    Next cell

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Converted " & convertedCount & " formula cell(s) to plain values." & vbCrLf & _
           "Skipped " & skippedCount & " cell(s) that couldn't be assigned (typically array formulas).", _
           vbInformation, "Convert Formulas to Values"
    Exit Sub

NoSelection:
    MsgBox "Select the range of formulas to freeze first.", vbExclamation
    Exit Sub

CleanFail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
