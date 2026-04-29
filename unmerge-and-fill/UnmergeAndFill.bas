Attribute VB_Name = "UnmergeAndFill"
Option Explicit

' Unmerge Cells and Fill Down
' Source: https://excelmacros.net/tools/unmerge-and-fill
' Offline. No API calls. No external dependencies.

' Walks the selection. For each merged area, copies the top-left value into a
' variable, unmerges the area, and writes the value into every formerly-merged
' cell. Result: data that pivots, sorts, and filters cleanly.

Public Sub UnmergeAndFill()
    Dim r As Range
    Dim cell As Range
    Dim mergedArea As Range
    Dim val As Variant
    Dim mergeCount As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    mergeCount = 0

    For Each cell In r.Cells
        If cell.MergeCells Then
            Set mergedArea = cell.MergeArea
            ' Process each merged area only once, when we hit its top-left cell.
            If cell.Address = mergedArea.Cells(1, 1).Address Then
                val = mergedArea.Cells(1, 1).Value
                mergedArea.UnMerge
                mergedArea.Value = val
                mergeCount = mergeCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Unmerged " & mergeCount & " merged area(s) and copied the value into every cell.", _
           vbInformation, "Unmerge and Fill"
    Exit Sub

NoSelection:
    MsgBox "Select the range with merged cells first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
