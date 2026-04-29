Attribute VB_Name = "ReverseRowOrder"
Option Explicit

' Reverse Row Order
' Source: https://excelmacros.net/tools/reverse-row-order
' Offline. No API calls. No external dependencies.

' Reverses the order of rows in the selection top-to-bottom. The first row
' becomes the last and vice versa. Operates on every column in the selection
' simultaneously so rows stay aligned.

Public Sub ReverseRowOrder()
    Dim r As Range
    Dim n As Long
    Dim cols As Long
    Dim i As Long, j As Long
    Dim values As Variant
    Dim reversed As Variant

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection
    If r.Rows.Count < 2 Then
        MsgBox "Select at least 2 rows to reverse.", vbExclamation
        Exit Sub
    End If

    n = r.Rows.Count
    cols = r.Columns.Count

    Application.ScreenUpdating = False

    values = r.Value
    ReDim reversed(1 To n, 1 To cols)
    For i = 1 To n
        For j = 1 To cols
            reversed(i, j) = values(n - i + 1, j)
        Next j
    Next i

    r.Value = reversed

    Application.ScreenUpdating = True

    MsgBox "Reversed " & n & " row(s) in place.", _
           vbInformation, "Reverse Row Order"
    Exit Sub

NoSelection:
    MsgBox "Select the range of rows to reverse first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
