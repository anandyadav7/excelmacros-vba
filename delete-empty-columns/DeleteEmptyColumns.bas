Attribute VB_Name = "DeleteEmptyColumns"
Option Explicit

' Delete Empty Columns
' Source: https://excelmacros.net/tools/delete-empty-columns
' Offline. No API calls. No external dependencies.

' Scans each column inside the selection. If every cell in that column (within
' the selected row range) is empty or whitespace-only, marks the column for
' deletion. Deletes from right to left so column indices stay valid.

Public Sub DeleteEmptyColumns()
    Dim r As Range
    Dim col As Long
    Dim row As Long
    Dim colIsEmpty As Boolean
    Dim deleteCols() As Long
    Dim deleteCount As Long
    Dim ws As Worksheet
    Dim startCol As Long
    Dim i As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    ReDim deleteCols(1 To r.Columns.Count)
    deleteCount = 0

    For col = 1 To r.Columns.Count
        colIsEmpty = True
        For row = 1 To r.Rows.Count
            If Not IsEmpty(r.Cells(row, col).Value) Then
                If Trim$(CStr(r.Cells(row, col).Value)) <> "" Then
                    colIsEmpty = False
                    Exit For
                End If
            End If
        Next row
        If colIsEmpty Then
            deleteCount = deleteCount + 1
            deleteCols(deleteCount) = col
        End If
    Next col

    Set ws = r.Worksheet
    startCol = r.Column

    Application.ScreenUpdating = False
    For i = deleteCount To 1 Step -1
        ws.Columns(startCol + deleteCols(i) - 1).Delete Shift:=xlShiftToLeft
    Next i
    Application.ScreenUpdating = True

    MsgBox "Removed " & deleteCount & " empty column(s).", _
           vbInformation, "Delete Empty Columns"
    Exit Sub

NoSelection:
    MsgBox "Select the range to scan for empty columns first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
