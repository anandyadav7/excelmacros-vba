Attribute VB_Name = "AutofitAllColumnsAllSheets"
Option Explicit

' AutoFit All Columns on All Sheets
' Source: https://excelmacros.net/tools/autofit-all-columns-all-sheets
' Offline. No API calls. No external dependencies.

' Walks every visible sheet in the active workbook and runs AutoFit on every
' column of the used range. Result: every sheet looks consistent without
' manually clicking through each tab.

Public Sub AutofitAllColumnsAllSheets()
    Dim ws As Worksheet
    Dim count As Long

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    count = 0

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.UsedRange.Columns.AutoFit
            count = count + 1
        End If
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Autofitted columns on " & count & " visible sheet(s).", _
           vbInformation, "AutoFit All Columns on All Sheets"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
