Attribute VB_Name = "RemoveBlankRows"
Option Explicit

' Remove Blank Rows
' Source: https://excelmacros.net/tools/remove-blank-rows
' Offline. No API calls. No external dependencies.

' Deletes any row inside the selection where every cell is empty (or contains
' only whitespace). Deletes from the bottom up so the remaining row indices
' stay valid during the loop.

Public Sub RemoveBlankRows()
    Dim r As Range
    Dim row As Long
    Dim col As Long
    Dim isBlank As Boolean
    Dim deleteRows() As Long
    Dim deleteCount As Long
    Dim ws As Worksheet
    Dim startRow As Long
    Dim i As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    ReDim deleteRows(1 To r.Rows.Count)
    deleteCount = 0

    For row = 1 To r.Rows.Count
        isBlank = True
        For col = 1 To r.Columns.Count
            If Not IsEmpty(r.Cells(row, col).Value) Then
                If Trim$(CStr(r.Cells(row, col).Value)) <> "" Then
                    isBlank = False
                    Exit For
                End If
            End If
        Next col
        If isBlank Then
            deleteCount = deleteCount + 1
            deleteRows(deleteCount) = row
        End If
    Next row

    Set ws = r.Worksheet
    startRow = r.Row

    Application.ScreenUpdating = False
    For i = deleteCount To 1 Step -1
        ws.Rows(startRow + deleteRows(i) - 1).Delete Shift:=xlShiftUp
    Next i
    Application.ScreenUpdating = True

    MsgBox "Removed " & deleteCount & " blank row(s).", _
           vbInformation, "Remove Blank Rows"
    Exit Sub

NoSelection:
    MsgBox "Select the range to scan for blank rows first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
