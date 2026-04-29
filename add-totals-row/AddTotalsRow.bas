Attribute VB_Name = "AddTotalsRow"
Option Explicit

' Add Totals Row to Numeric Columns
' Source: https://excelmacros.net/tools/add-totals-row
' Offline. No API calls. No external dependencies.

' Inserts a totals row immediately below the selection. For every column with
' at least one numeric value, writes a SUM formula. The first non-numeric
' column gets a "Total" label. Top border + bold are applied to the row.

Public Sub AddTotalsRow()
    Dim r As Range
    Dim ws As Worksheet
    Dim totalsRow As Long
    Dim col As Long
    Dim sourceCol As Long
    Dim row As Long
    Dim numericCol() As Boolean
    Dim hasAnyNumeric As Boolean
    Dim sumCount As Long
    Dim labelDone As Boolean

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Set ws = r.Worksheet
    totalsRow = r.Row + r.Rows.Count

    ' First pass: identify which selection columns hold numeric data.
    ReDim numericCol(1 To r.Columns.Count)
    hasAnyNumeric = False
    For col = 1 To r.Columns.Count
        sourceCol = r.Column + col - 1
        For row = r.Row To r.Row + r.Rows.Count - 1
            If IsNumeric(ws.Cells(row, sourceCol).Value) And _
               Not IsEmpty(ws.Cells(row, sourceCol).Value) And _
               VarType(ws.Cells(row, sourceCol).Value) <> vbString Then
                numericCol(col) = True
                hasAnyNumeric = True
                Exit For
            End If
        Next row
    Next col

    If Not hasAnyNumeric Then
        MsgBox "No numeric columns found in your selection. Nothing to total.", _
               vbInformation, "Add Totals Row"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    sumCount = 0
    labelDone = False

    For col = 1 To r.Columns.Count
        sourceCol = r.Column + col - 1
        ws.Cells(totalsRow, sourceCol).Borders(xlEdgeTop).LineStyle = xlContinuous
        ws.Cells(totalsRow, sourceCol).Font.Bold = True

        If numericCol(col) Then
            ws.Cells(totalsRow, sourceCol).Formula = _
                "=SUM(" & ws.Cells(r.Row, sourceCol).Address(False, False) & ":" & _
                ws.Cells(r.Row + r.Rows.Count - 1, sourceCol).Address(False, False) & ")"
            sumCount = sumCount + 1
        ElseIf Not labelDone Then
            ws.Cells(totalsRow, sourceCol).Value = "Total"
            labelDone = True
        End If
    Next col

    Application.ScreenUpdating = True
    MsgBox "Added totals row at row " & totalsRow & " with " & sumCount & " SUM formula(s).", _
           vbInformation, "Add Totals Row"
    Exit Sub

NoSelection:
    MsgBox "Select the data range first (without an existing totals row).", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
