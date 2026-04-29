Attribute VB_Name = "UnhideAllSheets"
Option Explicit

' Unhide All Sheets
' Source: https://excelmacros.net/tools/unhide-all-sheets
' Offline. No API calls. No external dependencies.

' Sets every sheet (including very-hidden ones) to visible. Useful for
' inheriting workbooks where someone hid sheets you need to see, or for
' auditing what's actually inside a multi-tab file.

Public Sub UnhideAllSheets()
    Dim ws As Worksheet
    Dim count As Long

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    count = 0

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            count = count + 1
        End If
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Unhid " & count & " sheet(s).", _
           vbInformation, "Unhide All Sheets"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
