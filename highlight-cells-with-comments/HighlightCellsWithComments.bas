Attribute VB_Name = "HighlightCellsWithComments"
Option Explicit

' Highlight Cells With Comments
' Source: https://excelmacros.net/tools/highlight-cells-with-comments
' Offline. No API calls. No external dependencies.

' Walks the selection and paints every cell that has a Comment (the legacy
' yellow-triangle kind) with a light green fill. Useful for spotting which
' cells in a model carry annotations from the author.

Public Sub HighlightCellsWithComments()
    Dim r As Range
    Dim cell As Range
    Dim count As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    count = 0

    For Each cell In r.Cells
        If Not cell.Comment Is Nothing Then
            cell.Interior.Color = RGB(198, 224, 180)
            cell.Font.Color = RGB(40, 80, 30)
            count = count + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Highlighted " & count & " cell(s) that have a comment in your selection.", _
           vbInformation, "Highlight Cells With Comments"
    Exit Sub

NoSelection:
    MsgBox "Select the range to scan for cell comments first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
