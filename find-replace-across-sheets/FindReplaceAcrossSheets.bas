Attribute VB_Name = "FindReplaceAcrossSheets"
Option Explicit

' Find and Replace Across All Sheets
' Source: https://excelmacros.net/tools/find-replace-across-sheets
' Offline. No API calls. No external dependencies.

Public Sub FindReplaceAcrossSheets()
    Dim findText As String
    Dim replaceText As String
    Dim caseAns As String
    Dim wholeAns As String
    Dim matchCase As Boolean
    Dim wholeCell As Boolean
    Dim ws As Worksheet
    Dim found As Range
    Dim firstAddress As String
    Dim totalReplaced As Long
    Dim sheetReplaced As Long
    Dim lookAt As Long

    On Error GoTo CleanFail

    findText = InputBox("Find what (in every sheet of this workbook):", _
                        "Find and Replace Across All Sheets")
    If Len(findText) = 0 Then Exit Sub

    replaceText = InputBox("Replace with (leave blank to delete the matching text):", _
                           "Find and Replace Across All Sheets")
    ' replaceText may be empty by design; only Cancel exits.
    If StrPtr(replaceText) = 0 Then Exit Sub

    caseAns = InputBox("Match case? Type Y for yes, anything else for no.", _
                       "Find and Replace Across All Sheets", "N")
    If StrPtr(caseAns) = 0 Then Exit Sub
    matchCase = (UCase$(Trim$(caseAns)) = "Y")

    wholeAns = InputBox("Match whole cell only? Type Y for yes, anything else for no.", _
                        "Find and Replace Across All Sheets", "N")
    If StrPtr(wholeAns) = 0 Then Exit Sub
    wholeCell = (UCase$(Trim$(wholeAns)) = "Y")

    If wholeCell Then
        lookAt = xlWhole
    Else
        lookAt = xlPart
    End If

    Application.ScreenUpdating = False
    totalReplaced = 0

    For Each ws In ActiveWorkbook.Worksheets
        sheetReplaced = 0
        Set found = ws.Cells.Find(What:=findText, LookIn:=xlValues, LookAt:=lookAt, _
                                   MatchCase:=matchCase, SearchOrder:=xlByRows, _
                                   SearchDirection:=xlNext)
        If Not found Is Nothing Then
            firstAddress = found.Address
            Do
                sheetReplaced = sheetReplaced + 1
                Set found = ws.Cells.FindNext(After:=found)
                If found Is Nothing Then Exit Do
            Loop While found.Address <> firstAddress

            ws.Cells.Replace What:=findText, Replacement:=replaceText, _
                             LookAt:=lookAt, SearchOrder:=xlByRows, _
                             MatchCase:=matchCase
        End If
        totalReplaced = totalReplaced + sheetReplaced
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Replaced " & totalReplaced & " occurrence(s) across " & _
           ActiveWorkbook.Worksheets.Count & " sheet(s).", vbInformation, _
           "Find and Replace Across All Sheets"
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
