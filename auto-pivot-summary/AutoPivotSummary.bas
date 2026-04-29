Attribute VB_Name = "AutoPivotSummary"
Option Explicit

' Auto Pivot Summary
' Source: https://excelmacros.net/tools/auto-pivot-summary
' Offline. No API calls. No external dependencies.

Public Sub AutoPivotSummary()
    Dim r As Range
    Dim data As Variant
    Dim totalRows As Long, totalCols As Long
    Dim spec As String
    Dim parts() As String
    Dim groupCol As Long, valueCol As Long
    Dim rIdx As Long
    Dim key As String
    Dim val As Variant
    Dim sums As Object
    Dim counts As Object
    Dim wb As Workbook
    Dim summary As Worksheet
    Dim outRow As Long
    Dim k As Variant

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Then GoTo NoSelection
    If r.Rows.Count < 2 Or r.Columns.Count < 2 Then
        MsgBox "Select at least 2 rows (header + data) and 2 columns.", vbExclamation
        Exit Sub
    End If

    spec = InputBox( _
        "Enter the GROUP-BY column number, comma, then the column to SUM." & vbCrLf & _
        "Example: 1,3 groups by column 1 and sums column 3.", _
        "Auto Pivot Summary")
    If Len(Trim$(spec)) = 0 Then Exit Sub

    parts = Split(spec, ",")
    If UBound(parts) - LBound(parts) <> 1 Then
        MsgBox "Enter exactly two numbers separated by a comma.", vbExclamation
        Exit Sub
    End If
    If Not IsNumeric(Trim$(parts(0))) Or Not IsNumeric(Trim$(parts(1))) Then
        MsgBox "Both values must be numbers.", vbExclamation
        Exit Sub
    End If
    groupCol = CLng(Trim$(parts(0)))
    valueCol = CLng(Trim$(parts(1)))
    If groupCol < 1 Or groupCol > r.Columns.Count Or valueCol < 1 Or valueCol > r.Columns.Count Then
        MsgBox "Column numbers must be within the selection (1 to " & r.Columns.Count & ").", vbExclamation
        Exit Sub
    End If

    data = r.Value
    totalRows = UBound(data, 1)
    totalCols = UBound(data, 2)

    Set sums = CreateObject("Scripting.Dictionary")
    sums.CompareMode = 1
    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = 1

    For rIdx = 2 To totalRows
        key = CStr(data(rIdx, groupCol))
        If Len(Trim$(key)) = 0 Then key = "(blank)"
        val = data(rIdx, valueCol)
        If IsNumeric(val) Then
            If sums.Exists(key) Then
                sums(key) = sums(key) + CDbl(val)
                counts(key) = counts(key) + 1
            Else
                sums.Add key, CDbl(val)
                counts.Add key, 1
            End If
        End If
    Next rIdx

    If sums.Count = 0 Then
        MsgBox "No numeric values found in column " & valueCol & ".", vbExclamation
        Exit Sub
    End If

    Set wb = ActiveWorkbook
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("Summary").Delete
    On Error GoTo CleanFail
    Application.DisplayAlerts = True

    Application.ScreenUpdating = False
    Set summary = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    summary.Name = "Summary"

    summary.Cells(1, 1).Value = CStr(data(1, groupCol))
    summary.Cells(1, 2).Value = "Sum of " & CStr(data(1, valueCol))
    summary.Cells(1, 3).Value = "Count"
    summary.Range(summary.Cells(1, 1), summary.Cells(1, 3)).Font.Bold = True

    outRow = 2
    For Each k In sums.Keys
        summary.Cells(outRow, 1).Value = k
        summary.Cells(outRow, 2).Value = sums(k)
        summary.Cells(outRow, 3).Value = counts(k)
        outRow = outRow + 1
    Next k

    summary.Columns("A:C").AutoFit
    Application.ScreenUpdating = True
    summary.Activate

    MsgBox "Created Summary sheet with " & sums.Count & " group(s).", vbInformation, _
           "Auto Pivot Summary"
    Exit Sub

NoSelection:
    MsgBox "Select your data range first (including the header row).", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
