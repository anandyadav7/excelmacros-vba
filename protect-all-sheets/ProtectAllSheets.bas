Attribute VB_Name = "ProtectAllSheets"
Option Explicit

' Protect All Sheets
' Source: https://excelmacros.net/tools/protect-all-sheets
' Offline. No API calls. No external dependencies.

' Prompts for a password, then applies sheet protection with that password to
' every sheet in the active workbook. WARNING: forgetting the password makes
' the sheets unrecoverable without external recovery tools.

Public Sub ProtectAllSheets()
    Dim password As String
    Dim confirm As String
    Dim ws As Worksheet
    Dim count As Long

    On Error GoTo CleanFail

    password = InputBox( _
        "Enter the password to apply to ALL sheets." & vbCrLf & vbCrLf & _
        "WARNING: forgetting this password makes the sheets unrecoverable." & vbCrLf & _
        "Leave blank if you want to protect without a password.", _
        "Protect All Sheets")
    If StrPtr(password) = 0 Then Exit Sub

    If Len(password) > 0 Then
        confirm = InputBox("Re-enter the password to confirm:", "Protect All Sheets")
        If StrPtr(confirm) = 0 Then Exit Sub
        If password <> confirm Then
            MsgBox "Passwords don't match. No changes made.", vbExclamation
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    count = 0

    For Each ws In ActiveWorkbook.Worksheets
        ws.Protect Password:=password
        count = count + 1
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Protected " & count & " sheet(s).", _
           vbInformation, "Protect All Sheets"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
