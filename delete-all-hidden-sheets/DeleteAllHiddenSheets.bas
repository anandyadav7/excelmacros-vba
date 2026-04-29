Attribute VB_Name = "DeleteAllHiddenSheets"
Option Explicit

' Delete All Hidden Sheets
' Source: https://excelmacros.net/tools/delete-all-hidden-sheets
' Offline. No API calls. No external dependencies.

' Asks for confirmation, then deletes every hidden and very-hidden sheet from
' the active workbook. The active workbook is modified in place. Save first
' if you want to keep a backup.

Public Sub DeleteAllHiddenSheets()
    Dim ws As Worksheet
    Dim sheetsToDelete As Collection
    Dim ans As VbMsgBoxResult
    Dim i As Long

    On Error GoTo CleanFail

    Set sheetsToDelete = New Collection
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then sheetsToDelete.Add ws
    Next ws

    If sheetsToDelete.Count = 0 Then
        MsgBox "No hidden sheets found in this workbook.", vbInformation
        Exit Sub
    End If

    ans = MsgBox("Delete " & sheetsToDelete.Count & " hidden sheet(s) from this workbook?" & vbCrLf & _
                 "This is irreversible. Save the workbook first if you want a backup.", _
                 vbYesNo + vbExclamation, "Delete All Hidden Sheets")
    If ans <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For i = 1 To sheetsToDelete.Count
        sheetsToDelete(i).Delete
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Deleted " & sheetsToDelete.Count & " hidden sheet(s).", _
           vbInformation, "Delete All Hidden Sheets"
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
