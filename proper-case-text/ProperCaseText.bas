Attribute VB_Name = "ProperCaseText"
Option Explicit

' Convert Text to Proper Case
' Source: https://excelmacros.net/tools/proper-case-text
' Offline. No API calls. No external dependencies.

' Capitalizes the first letter of each word and lowercases the rest. Skips
' formula cells and non-text cells. Result is written in place.

Public Sub ProperCaseText()
    Dim r As Range
    Dim cell As Range
    Dim convertedCount As Long
    Dim skippedFormulas As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    convertedCount = 0
    skippedFormulas = 0

    For Each cell In r.Cells
        If cell.HasFormula Then
            skippedFormulas = skippedFormulas + 1
        ElseIf Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            cell.Value = Application.WorksheetFunction.Proper(CStr(cell.Value))
            convertedCount = convertedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Converted " & convertedCount & " text cell(s) to Proper Case." & vbCrLf & _
           "Skipped " & skippedFormulas & " formula cell(s).", _
           vbInformation, "Proper Case Text"
    Exit Sub

NoSelection:
    MsgBox "Select the range of text to convert first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
