Attribute VB_Name = "RemoveDuplicatesMultiColumn"
Option Explicit

' Remove Duplicates by Multiple Columns
' Source: https://excelmacros.net/tools/remove-duplicates-multi-column
' Offline. No API calls. No external dependencies.

Public Sub RemoveDuplicatesMultiColumn()
    Dim selectedRange As Range
    Dim colSpec As String
    Dim colParts() As String
    Dim colIndices() As Long
    Dim i As Long, r As Long
    Dim seen As Object
    Dim key As String
    Dim deleteRows() As Long
    Dim deleteCount As Long
    Dim startRow As Long
    Dim ws As Worksheet

    On Error GoTo CleanFail

    Set selectedRange = Selection
    If selectedRange Is Nothing Then GoTo NoSelection
    If selectedRange.Rows.Count < 2 Then GoTo NoSelection
    If selectedRange.Columns.Count < 1 Then GoTo NoSelection

    colSpec = InputBox( _
        "Enter the column numbers to use for the duplicate check, separated by commas." & vbCrLf & _
        "Example: 1,3 means columns 1 and 3 of your selection." & vbCrLf & vbCrLf & _
        "Rows are duplicates only when ALL chosen column values match an earlier row.", _
        "Remove Duplicates by Multiple Columns")

    If Len(Trim$(colSpec)) = 0 Then Exit Sub

    colParts = Split(colSpec, ",")
    ReDim colIndices(LBound(colParts) To UBound(colParts))
    For i = LBound(colParts) To UBound(colParts)
        If Not IsNumeric(Trim$(colParts(i))) Then
            MsgBox "Column list must be comma-separated numbers, like 1,3.", vbExclamation
            Exit Sub
        End If
        colIndices(i) = CLng(Trim$(colParts(i)))
        If colIndices(i) < 1 Or colIndices(i) > selectedRange.Columns.Count Then
            MsgBox "Column " & colIndices(i) & " is outside your selection (which has " & _
                   selectedRange.Columns.Count & " columns).", vbExclamation
            Exit Sub
        End If
    Next i

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1 ' vbTextCompare = case-insensitive

    ReDim deleteRows(1 To selectedRange.Rows.Count)
    deleteCount = 0

    ' Row 1 is treated as the header and skipped.
    For r = 2 To selectedRange.Rows.Count
        key = ""
        For i = LBound(colIndices) To UBound(colIndices)
            key = key & "||" & CStr(selectedRange.Cells(r, colIndices(i)).Value)
        Next i

        If seen.Exists(key) Then
            deleteCount = deleteCount + 1
            deleteRows(deleteCount) = r
        Else
            seen.Add key, True
        End If
    Next r

    ' Delete from bottom up using absolute worksheet row numbers, so the
    ' selection's relative indexing doesn't drift as rows are removed.
    Set ws = selectedRange.Worksheet
    startRow = selectedRange.Row

    Application.ScreenUpdating = False
    For i = deleteCount To 1 Step -1
        ws.Rows(startRow + deleteRows(i) - 1).Delete Shift:=xlShiftUp
    Next i
    Application.ScreenUpdating = True

    MsgBox "Removed " & deleteCount & " duplicate row(s).", vbInformation, _
           "Remove Duplicates by Multiple Columns"
    Exit Sub

NoSelection:
    MsgBox "Select your data range first (including the header row).", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
