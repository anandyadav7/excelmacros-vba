Attribute VB_Name = "SplitCellOnDelimiter"
Option Explicit

' Split Cell on Delimiter
' Source: https://excelmacros.net/tools/split-cell-on-delimiter
' Offline. No API calls. No external dependencies.

' Asks for a delimiter, then splits each cell in the selection on that
' delimiter, writing each part to a column to the right of the source cell.
' Useful for splitting comma-separated lists, semicolon-separated tags, etc.

Public Sub SplitCellOnDelimiter()
    Dim r As Range
    Dim cell As Range
    Dim delim As String
    Dim ws As Worksheet
    Dim parts() As String
    Dim i As Long
    Dim splitCount As Long
    Dim text As String

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection
    If r.Columns.Count > 1 Then
        MsgBox "Select a single column to split.", vbExclamation
        Exit Sub
    End If

    delim = InputBox( _
        "Enter the delimiter to split on." & vbCrLf & _
        "Examples: ,  |  ;  /  -  Tab" & vbCrLf & vbCrLf & _
        "For Tab, type the word Tab.", _
        "Split Cell on Delimiter", ",")
    If StrPtr(delim) = 0 Then Exit Sub
    If Len(delim) = 0 Then
        MsgBox "Delimiter cannot be blank.", vbExclamation
        Exit Sub
    End If
    If LCase$(Trim$(delim)) = "tab" Then delim = vbTab

    Set ws = r.Worksheet

    Application.ScreenUpdating = False
    splitCount = 0

    For Each cell In r.Cells
        If Not IsEmpty(cell.Value) And Not cell.HasFormula Then
            text = CStr(cell.Value)
            parts = Split(text, delim)
            For i = LBound(parts) To UBound(parts)
                ws.Cells(cell.Row, cell.Column + 1 + (i - LBound(parts))).Value = Trim$(parts(i))
            Next i
            splitCount = splitCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Split " & splitCount & " cell(s) on '" & _
           IIf(delim = vbTab, "Tab", delim) & "'.", _
           vbInformation, "Split Cell on Delimiter"
    Exit Sub

NoSelection:
    MsgBox "Select the column to split first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
