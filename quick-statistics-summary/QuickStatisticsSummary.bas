Attribute VB_Name = "QuickStatisticsSummary"
Option Explicit

' Quick Statistics Summary
' Source: https://excelmacros.net/tools/quick-statistics-summary
' Offline. No API calls. No external dependencies.

Public Sub QuickStatisticsSummary()
    Dim r As Range
    Dim cell As Range
    Dim values() As Double
    Dim count As Long
    Dim i As Long
    Dim sum As Double, mean As Double, median As Double
    Dim variance As Double, stdev As Double
    Dim minV As Double, maxV As Double
    Dim modeV As Variant
    Dim ws As Worksheet
    Dim destCol As Long, destRow As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection
    If r.Columns.Count > 1 Then
        MsgBox "Select a single column of numbers.", vbExclamation
        Exit Sub
    End If

    ReDim values(1 To r.Cells.CountLarge)
    count = 0
    For Each cell In r.Cells
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            count = count + 1
            values(count) = CDbl(cell.Value)
        End If
    Next cell

    If count = 0 Then
        MsgBox "No numeric values found in the selection.", vbExclamation
        Exit Sub
    End If

    ' Truncate the array so unused trailing zeros don't pollute Mode and other stats.
    If count < r.Cells.CountLarge Then
        ReDim Preserve values(1 To count)
    End If

    ' Mean, min, max, sum.
    sum = 0
    minV = values(1)
    maxV = values(1)
    For i = 1 To count
        sum = sum + values(i)
        If values(i) < minV Then minV = values(i)
        If values(i) > maxV Then maxV = values(i)
    Next i
    mean = sum / count

    ' Median: sort and pick middle (or average of two middle for even count).
    QuickSort values, 1, count
    If count Mod 2 = 1 Then
        median = values((count + 1) \ 2)
    Else
        median = (values(count \ 2) + values(count \ 2 + 1)) / 2
    End If

    ' Sample standard deviation (divides by n-1).
    variance = 0
    For i = 1 To count
        variance = variance + (values(i) - mean) ^ 2
    Next i
    If count > 1 Then variance = variance / (count - 1) Else variance = 0
    stdev = Sqr(variance)

    ' Mode via Excel's worksheet function. Falls back to "(no mode)" if no value repeats.
    On Error Resume Next
    modeV = Application.WorksheetFunction.Mode(values)
    If Err.Number <> 0 Then modeV = "(no mode)"
    Err.Clear
    On Error GoTo CleanFail

    Set ws = r.Worksheet
    destCol = r.Column + 2
    destRow = r.Row

    Application.ScreenUpdating = False
    ws.Cells(destRow, destCol).Value = "Statistic"
    ws.Cells(destRow, destCol + 1).Value = "Value"
    ws.Cells(destRow, destCol).Font.Bold = True
    ws.Cells(destRow, destCol + 1).Font.Bold = True

    ws.Cells(destRow + 1, destCol).Value = "Count"
    ws.Cells(destRow + 1, destCol + 1).Value = count
    ws.Cells(destRow + 2, destCol).Value = "Mean"
    ws.Cells(destRow + 2, destCol + 1).Value = mean
    ws.Cells(destRow + 3, destCol).Value = "Median"
    ws.Cells(destRow + 3, destCol + 1).Value = median
    ws.Cells(destRow + 4, destCol).Value = "Mode"
    ws.Cells(destRow + 4, destCol + 1).Value = modeV
    ws.Cells(destRow + 5, destCol).Value = "Std Dev (sample)"
    ws.Cells(destRow + 5, destCol + 1).Value = stdev
    ws.Cells(destRow + 6, destCol).Value = "Min"
    ws.Cells(destRow + 6, destCol + 1).Value = minV
    ws.Cells(destRow + 7, destCol).Value = "Max"
    ws.Cells(destRow + 7, destCol + 1).Value = maxV
    ws.Cells(destRow + 8, destCol).Value = "Range"
    ws.Cells(destRow + 8, destCol + 1).Value = maxV - minV

    ws.Range(ws.Cells(destRow, destCol), ws.Cells(destRow + 8, destCol + 1)).Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "Stats written for " & count & " value(s).", vbInformation, _
           "Quick Statistics Summary"
    Exit Sub

NoSelection:
    MsgBox "Select a single column of numbers first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' In-place quicksort, ascending.
Private Sub QuickSort(ByRef a() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As Double, tmp As Double
    i = lo
    j = hi
    pivot = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < pivot
            i = i + 1
        Loop
        Do While a(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            tmp = a(i)
            a(i) = a(j)
            a(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then QuickSort a, lo, j
    If i < hi Then QuickSort a, i, hi
End Sub
