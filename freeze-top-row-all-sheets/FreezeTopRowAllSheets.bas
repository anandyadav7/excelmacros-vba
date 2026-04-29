Attribute VB_Name = "FreezeTopRowAllSheets"
Option Explicit

' Freeze Top Row on All Sheets
' Source: https://excelmacros.net/tools/freeze-top-row-all-sheets
' Offline. No API calls. No external dependencies.

' Walks every visible sheet in the active workbook and applies Freeze Panes at
' row 2 so the header row stays visible during scrolling. If a sheet already
' has freeze panes, they are reset before applying the new freeze.

Public Sub FreezeTopRowAllSheets()
    Dim ws As Worksheet
    Dim originalSheet As Worksheet
    Dim count As Long

    On Error GoTo CleanFail

    Set originalSheet = ActiveSheet

    Application.ScreenUpdating = False
    count = 0

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ActiveWindow.FreezePanes = False
            ws.Range("A2").Select
            ActiveWindow.FreezePanes = True
            count = count + 1
        End If
    Next ws

    originalSheet.Activate
    Application.ScreenUpdating = True

    MsgBox "Froze the top row on " & count & " visible sheet(s).", _
           vbInformation, "Freeze Top Row on All Sheets"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
