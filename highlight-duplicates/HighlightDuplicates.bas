Attribute VB_Name = "HighlightDuplicates"
Option Explicit

' Highlight Duplicate Values
' Source: https://excelmacros.net/tools/highlight-duplicates
' Offline. No API calls. No external dependencies.

Public Sub HighlightDuplicates()
    Dim r As Range
    Dim cell As Range
    Dim seen As Object
    Dim val As String
    Dim highlightedCount As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Then GoTo NoSelection
    If r.Cells.CountLarge < 2 Then GoTo NoSelection

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1 ' vbTextCompare = case-insensitive

    Application.ScreenUpdating = False

    ' First pass: count occurrences of each non-empty value.
    For Each cell In r.Cells
        If Not IsEmpty(cell.Value) Then
            val = CStr(cell.Value)
            If seen.Exists(val) Then
                seen(val) = seen(val) + 1
            Else
                seen.Add val, 1
            End If
        End If
    Next cell

    ' Second pass: paint cells whose value count >= 2.
    highlightedCount = 0
    For Each cell In r.Cells
        If Not IsEmpty(cell.Value) Then
            val = CStr(cell.Value)
            If seen.Exists(val) Then
                If seen(val) >= 2 Then
                    cell.Interior.Color = RGB(255, 199, 206) ' light red fill
                    cell.Font.Color = RGB(156, 0, 6)         ' dark red text
                    highlightedCount = highlightedCount + 1
                End If
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Highlighted " & highlightedCount & " duplicate cell(s).", vbInformation, _
           "Highlight Duplicate Values"
    Exit Sub

NoSelection:
    MsgBox "Select at least 2 cells to check for duplicates.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
