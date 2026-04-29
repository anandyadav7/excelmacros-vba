Attribute VB_Name = "TransposeRowsAndColumns"
Option Explicit

' Transpose Rows and Columns
' Source: https://excelmacros.net/tools/transpose-rows-and-columns
' Offline. No API calls. No external dependencies.

' Takes the user's selection and writes a transposed copy at the destination
' cell they pick. An R-by-C selection becomes a C-by-R block. Original data
' is not modified, so you can keep both layouts side by side.

Public Sub TransposeRowsAndColumns()
    Dim r As Range
    Dim destCell As Range

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    On Error Resume Next
    Set destCell = Application.InputBox( _
        Prompt:="Click the top-left cell where the transposed result should appear." & vbCrLf & _
                "(The original data is left in place.)", _
        Title:="Transpose Rows and Columns", _
        Type:=8)
    On Error GoTo CleanFail
    If destCell Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    r.Copy
    destCell.Cells(1, 1).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, _
                                       SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False

    Application.ScreenUpdating = True

    MsgBox "Transposed " & r.Rows.Count & " row(s) x " & r.Columns.Count & " column(s) into " & _
           r.Columns.Count & " row(s) x " & r.Rows.Count & " column(s) at " & _
           destCell.Cells(1, 1).Address(False, False) & ".", _
           vbInformation, "Transpose Rows and Columns"
    Exit Sub

NoSelection:
    MsgBox "Select the range you want to transpose first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
