Attribute VB_Name = "HighlightFormulaCells"
Option Explicit

' Highlight All Formula Cells
' Source: https://excelmacros.net/tools/highlight-formula-cells
' Offline. No API calls. No external dependencies.

' Walks the selection and paints every cell containing a formula in light
' yellow with dark gold text. Useful for auditing which cells are calculated
' versus hardcoded.

Public Sub HighlightFormulaCells()
    Dim r As Range
    Dim cell As Range
    Dim highlightedCount As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    highlightedCount = 0

    For Each cell In r.Cells
        If cell.HasFormula Then
            cell.Interior.Color = RGB(255, 235, 156)
            cell.Font.Color = RGB(120, 80, 0)
            highlightedCount = highlightedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Highlighted " & highlightedCount & " formula cell(s) in your selection.", _
           vbInformation, "Highlight Formula Cells"
    Exit Sub

NoSelection:
    MsgBox "Select the range to scan for formulas first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
