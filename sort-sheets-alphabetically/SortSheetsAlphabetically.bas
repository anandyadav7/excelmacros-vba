Attribute VB_Name = "SortSheetsAlphabetically"
Option Explicit

' Sort Sheets Alphabetically
' Source: https://excelmacros.net/tools/sort-sheets-alphabetically
' Offline. No API calls. No external dependencies.

' Reorders the tabs of the active workbook in A-to-Z or Z-to-A order using a
' simple bubble sort with case-insensitive name compare.

Public Sub SortSheetsAlphabetically()
    Dim i As Long, j As Long
    Dim n As Long
    Dim ans As String
    Dim ascending As Boolean

    On Error GoTo CleanFail

    n = ActiveWorkbook.Worksheets.Count
    If n < 2 Then
        MsgBox "Workbook only has " & n & " sheet(s); nothing to sort.", vbInformation
        Exit Sub
    End If

    ans = InputBox( _
        "Sort order:" & vbCrLf & _
        "1 = A to Z" & vbCrLf & _
        "2 = Z to A", _
        "Sort Sheets Alphabetically", "1")
    If StrPtr(ans) = 0 Then Exit Sub

    Select Case Trim$(ans)
        Case "1": ascending = True
        Case "2": ascending = False
        Case Else
            MsgBox "Type 1 or 2.", vbExclamation
            Exit Sub
    End Select

    Application.ScreenUpdating = False

    For i = 1 To n - 1
        For j = 1 To n - i
            If ascending Then
                If StrComp(ActiveWorkbook.Worksheets(j).Name, _
                           ActiveWorkbook.Worksheets(j + 1).Name, _
                           vbTextCompare) > 0 Then
                    ActiveWorkbook.Worksheets(j + 1).Move _
                        Before:=ActiveWorkbook.Worksheets(j)
                End If
            Else
                If StrComp(ActiveWorkbook.Worksheets(j).Name, _
                           ActiveWorkbook.Worksheets(j + 1).Name, _
                           vbTextCompare) < 0 Then
                    ActiveWorkbook.Worksheets(j + 1).Move _
                        Before:=ActiveWorkbook.Worksheets(j)
                End If
            End If
        Next j
    Next i

    Application.ScreenUpdating = True
    MsgBox "Sorted " & n & " sheet(s) " & IIf(ascending, "A to Z", "Z to A") & ".", _
           vbInformation, "Sort Sheets"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
