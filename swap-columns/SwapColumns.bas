Attribute VB_Name = "SwapColumns"
Option Explicit

' Swap Two Columns
' Source: https://excelmacros.net/tools/swap-columns
' Offline. No API calls. No external dependencies.

' Swaps the values of two columns in the selection. Works with two adjacent
' columns or two non-contiguous columns selected with Ctrl+click. Both columns
' must have the same number of rows.

Public Sub SwapColumns()
    Dim col1 As Range, col2 As Range
    Dim values1 As Variant, values2 As Variant
    Dim rowCount As Long

    On Error GoTo CleanFail

    If Selection.Areas.Count = 2 Then
        ' Two non-contiguous selections
        Set col1 = Selection.Areas(1)
        Set col2 = Selection.Areas(2)
        If col1.Columns.Count <> 1 Or col2.Columns.Count <> 1 Then
            MsgBox "Each non-contiguous selection must be exactly one column.", vbExclamation
            Exit Sub
        End If
        If col1.Rows.Count <> col2.Rows.Count Then
            MsgBox "Both columns must have the same number of rows.", vbExclamation
            Exit Sub
        End If
    ElseIf Selection.Columns.Count = 2 Then
        ' Two adjacent columns in a single selection
        Set col1 = Selection.Columns(1)
        Set col2 = Selection.Columns(2)
    Else
        MsgBox "Select exactly 2 columns to swap." & vbCrLf & _
               "Either two adjacent columns, or two non-contiguous columns with Ctrl+click.", _
               vbExclamation
        Exit Sub
    End If

    rowCount = col1.Rows.Count

    Application.ScreenUpdating = False

    values1 = col1.Value
    values2 = col2.Value
    col1.Value = values2
    col2.Value = values1

    Application.ScreenUpdating = True

    MsgBox "Swapped 2 columns of " & rowCount & " row(s) each.", _
           vbInformation, "Swap Columns"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
