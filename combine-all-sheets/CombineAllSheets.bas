Attribute VB_Name = "CombineAllSheets"
Option Explicit

' Combine All Sheets Into One
' Source: https://excelmacros.net/tools/combine-all-sheets
' Offline. No API calls. No external dependencies.

Public Sub CombineAllSheets()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim combined As Worksheet
    Dim source As Range
    Dim destRow As Long
    Dim hasHeader As Boolean
    Dim ans As VbMsgBoxResult
    Dim sheetsCombined As Long
    Dim isFirstSheetWithData As Boolean

    On Error GoTo CleanFail

    Set wb = ActiveWorkbook
    If wb Is Nothing Then GoTo NoWorkbook
    If wb.Worksheets.Count < 2 Then
        MsgBox "This workbook has only one sheet. Nothing to combine.", vbExclamation
        Exit Sub
    End If

    ans = MsgBox( _
        "Does the first row of each sheet contain headers?" & vbCrLf & vbCrLf & _
        "Click Yes to keep one header row at the top of the Combined sheet." & vbCrLf & _
        "Click No if your sheets have no header row.", _
        vbYesNoCancel + vbQuestion, "Combine All Sheets")
    If ans = vbCancel Then Exit Sub
    hasHeader = (ans = vbYes)

    ' Replace any existing Combined sheet.
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("Combined").Delete
    On Error GoTo CleanFail
    Application.DisplayAlerts = True

    Set combined = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    combined.Name = "Combined"

    Application.ScreenUpdating = False
    destRow = 1
    sheetsCombined = 0
    isFirstSheetWithData = True

    For Each ws In wb.Worksheets
        If ws.Name <> "Combined" Then
            Set source = Nothing
            On Error Resume Next
            Set source = ws.UsedRange
            On Error GoTo CleanFail

            If Not source Is Nothing Then
                If source.Cells.CountLarge > 0 Then
                    If hasHeader And Not isFirstSheetWithData Then
                        ' Skip the header row from sheets after the first.
                        If source.Rows.Count > 1 Then
                            source.Offset(1, 0).Resize(source.Rows.Count - 1, source.Columns.Count) _
                                .Copy combined.Cells(destRow, 1)
                            destRow = destRow + source.Rows.Count - 1
                            sheetsCombined = sheetsCombined + 1
                        End If
                    Else
                        source.Copy combined.Cells(destRow, 1)
                        destRow = destRow + source.Rows.Count
                        sheetsCombined = sheetsCombined + 1
                        isFirstSheetWithData = False
                    End If
                End If
            End If
        End If
    Next ws

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    combined.Activate
    combined.Cells(1, 1).Select

    MsgBox "Combined " & sheetsCombined & " sheet(s) into 'Combined'." & vbCrLf & _
           "Total rows: " & (destRow - 1), vbInformation, "Combine All Sheets"
    Exit Sub

NoWorkbook:
    MsgBox "Open the workbook that has your sheets first, then click into it before running this macro.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
