Attribute VB_Name = "InsertRowNumbers"
Option Explicit

' Insert Row Numbers
' Source: https://excelmacros.net/tools/insert-row-numbers
' Offline. No API calls. No external dependencies.

' Fills the selected single column with sequential integers starting at the
' number you specify. Useful for adding an "ID" or "row number" column to a
' table without dragging the fill handle.

Public Sub InsertRowNumbers()
    Dim r As Range
    Dim startAns As String
    Dim startNum As Long
    Dim i As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection
    If r.Columns.Count > 1 Then
        MsgBox "Select a single column.", vbExclamation
        Exit Sub
    End If

    startAns = InputBox( _
        "Start numbering from?" & vbCrLf & _
        "(Type 1 for 1,2,3..., type 0 for 0,1,2..., type 100 for 100,101,102...)", _
        "Insert Row Numbers", "1")
    If StrPtr(startAns) = 0 Then Exit Sub
    If Not IsNumeric(Trim$(startAns)) Then
        MsgBox "Type a whole number like 1 or 100.", vbExclamation
        Exit Sub
    End If
    startNum = CLng(Trim$(startAns))

    Application.ScreenUpdating = False
    For i = 1 To r.Rows.Count
        r.Cells(i, 1).Value = startNum + i - 1
    Next i
    Application.ScreenUpdating = True

    MsgBox "Filled " & r.Rows.Count & " cell(s) with sequential numbers " & _
           "starting at " & startNum & ".", _
           vbInformation, "Insert Row Numbers"
    Exit Sub

NoSelection:
    MsgBox "Select the column where the numbers should go first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
