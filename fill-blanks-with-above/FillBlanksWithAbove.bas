Attribute VB_Name = "FillBlanksWithAbove"
Option Explicit

' Fill Blank Cells With Value Above
' Source: https://excelmacros.net/tools/fill-blanks-with-above
' Offline. No API calls. No external dependencies.

' For each column in the selection, walks top to bottom. When a cell is blank,
' fills it with the most recent non-blank value seen in that column. Resets at
' the top of each column.

Public Sub FillBlanksWithAbove()
    Dim r As Range
    Dim col As Long
    Dim row As Long
    Dim lastVal As Variant
    Dim haveLast As Boolean
    Dim filledCount As Long
    Dim cellValue As Variant

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    filledCount = 0

    For col = 1 To r.Columns.Count
        haveLast = False
        lastVal = Empty
        For row = 1 To r.Rows.Count
            cellValue = r.Cells(row, col).Value
            If IsEmpty(cellValue) Or (VarType(cellValue) = vbString And Trim$(CStr(cellValue)) = "") Then
                If haveLast Then
                    r.Cells(row, col).Value = lastVal
                    filledCount = filledCount + 1
                End If
            Else
                lastVal = cellValue
                haveLast = True
            End If
        Next row
    Next col

    Application.ScreenUpdating = True

    MsgBox "Filled " & filledCount & " blank cell(s) with the value from above.", _
           vbInformation, "Fill Blanks With Above"
    Exit Sub

NoSelection:
    MsgBox "Select the range with blanks to fill first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
