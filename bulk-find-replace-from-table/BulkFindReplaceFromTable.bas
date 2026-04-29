Attribute VB_Name = "BulkFindReplaceFromTable"
Option Explicit

' Bulk Find and Replace From Table
' Source: https://excelmacros.net/tools/bulk-find-replace-from-table
' Offline. No API calls. No external dependencies.

' Reads a 2-column lookup table where column 1 = "find", column 2 = "replace".
' Applies every pair to the user's data range in one pass. Useful for batch
' country-code conversions, vendor-name normalization, account-code mapping,
' and any other "replace these many things at once" job.

Public Sub BulkFindReplaceFromTable()
    Dim dataRange As Range
    Dim lookupRange As Range
    Dim findText As String
    Dim replaceText As String
    Dim i As Long
    Dim pairsApplied As Long

    On Error GoTo CleanFail

    Set dataRange = Selection
    If dataRange Is Nothing Or dataRange.Cells.CountLarge < 1 Then GoTo NoData

    On Error Resume Next
    Set lookupRange = Application.InputBox( _
        Prompt:="Click to select the 2-column lookup table." & vbCrLf & _
                "Column 1 = find, Column 2 = replace." & vbCrLf & _
                "Skip the header row.", _
        Title:="Bulk Find Replace From Table", _
        Type:=8)
    On Error GoTo CleanFail
    If lookupRange Is Nothing Then Exit Sub
    If lookupRange.Columns.Count < 2 Then
        MsgBox "Pick a range with at least 2 columns.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    pairsApplied = 0

    For i = 1 To lookupRange.Rows.Count
        findText = CStr(lookupRange.Cells(i, 1).Value)
        replaceText = CStr(lookupRange.Cells(i, 2).Value)
        If Len(findText) > 0 Then
            dataRange.Replace What:=findText, Replacement:=replaceText, _
                              LookAt:=xlPart, MatchCase:=False, _
                              SearchOrder:=xlByRows
            pairsApplied = pairsApplied + 1
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Applied " & pairsApplied & " find/replace pair(s) to your data range.", _
           vbInformation, "Bulk Find Replace From Table"
    Exit Sub

NoData:
    MsgBox "Select your data range first, then run the macro.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
