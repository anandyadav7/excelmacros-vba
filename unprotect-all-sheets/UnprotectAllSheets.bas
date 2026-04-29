Attribute VB_Name = "UnprotectAllSheets"
Option Explicit

' Unprotect All Sheets
' Source: https://excelmacros.net/tools/unprotect-all-sheets
' Offline. No API calls. No external dependencies.

' Prompts for a password (or blank if no password), then attempts to unprotect
' every sheet. Sheets that fail to unprotect (because the password is wrong)
' are reported in the summary count.

Public Sub UnprotectAllSheets()
    Dim password As String
    Dim ws As Worksheet
    Dim unprotectedCount As Long
    Dim failedCount As Long

    On Error GoTo CleanFail

    password = InputBox( _
        "Enter the password (leave blank if the sheets have no password):", _
        "Unprotect All Sheets")
    If StrPtr(password) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    unprotectedCount = 0
    failedCount = 0

    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=password
        If Err.Number = 0 Then
            unprotectedCount = unprotectedCount + 1
        Else
            failedCount = failedCount + 1
            Err.Clear
        End If
        On Error GoTo CleanFail
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Unprotected " & unprotectedCount & " sheet(s)." & vbCrLf & _
           IIf(failedCount > 0, "Failed on " & failedCount & " sheet(s) (wrong password?).", ""), _
           vbInformation, "Unprotect All Sheets"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
