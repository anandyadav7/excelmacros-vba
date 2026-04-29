Attribute VB_Name = "HighlightBlankCells"
Option Explicit

' Highlight Blank Cells
' Source: https://excelmacros.net/tools/highlight-blank-cells
' Offline. No API calls. No external dependencies.

' Walks the selection and paints every blank or whitespace-only cell with a
' light red fill. Useful for visually auditing where data is missing in a
' table before deciding whether to fill it or delete the row.

Public Sub HighlightBlankCells()
    Dim r As Range
    Dim cell As Range
    Dim isBlank As Boolean
    Dim highlightedCount As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    highlightedCount = 0

    For Each cell In r.Cells
        isBlank = False
        If IsEmpty(cell.Value) Then
            isBlank = True
        ElseIf VarType(cell.Value) = vbString Then
            If Trim$(CStr(cell.Value)) = "" Then isBlank = True
        End If

        If isBlank Then
            cell.Interior.Color = RGB(255, 199, 206)
            cell.Font.Color = RGB(156, 0, 6)
            highlightedCount = highlightedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Highlighted " & highlightedCount & " blank cell(s) in your selection.", _
           vbInformation, "Highlight Blank Cells"
    Exit Sub

NoSelection:
    MsgBox "Select the range to scan for blanks first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
