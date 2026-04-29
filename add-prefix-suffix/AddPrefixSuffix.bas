Attribute VB_Name = "AddPrefixSuffix"
Option Explicit

' Add Prefix and Suffix to Cells
' Source: https://excelmacros.net/tools/add-prefix-suffix
' Offline. No API calls. No external dependencies.

' Prompts for a prefix and a suffix, then prepends/appends them to every
' non-empty cell in the selection. Forces text format on each modified cell so
' results like "+14155550100" or "0044..." don't get auto-converted by Excel.
' Formula cells are skipped to avoid corrupting their formulas.

Public Sub AddPrefixSuffix()
    Dim r As Range
    Dim cell As Range
    Dim prefix As String
    Dim suffix As String
    Dim modifiedCount As Long
    Dim skippedFormulas As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    prefix = InputBox( _
        "Prefix to add to the START of each cell." & vbCrLf & _
        "Leave blank if you only want a suffix.", _
        "Add Prefix and Suffix", "")
    If StrPtr(prefix) = 0 Then Exit Sub

    suffix = InputBox( _
        "Suffix to add to the END of each cell." & vbCrLf & _
        "Leave blank if you only want a prefix.", _
        "Add Prefix and Suffix", "")
    If StrPtr(suffix) = 0 Then Exit Sub

    If Len(prefix) = 0 And Len(suffix) = 0 Then
        MsgBox "Both prefix and suffix are blank. Nothing to do.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    modifiedCount = 0
    skippedFormulas = 0

    For Each cell In r.Cells
        If cell.HasFormula Then
            skippedFormulas = skippedFormulas + 1
        ElseIf Not IsEmpty(cell.Value) Then
            cell.NumberFormat = "@"
            cell.Value = prefix & CStr(cell.Value) & suffix
            modifiedCount = modifiedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Modified " & modifiedCount & " cell(s)." & vbCrLf & _
           "Skipped " & skippedFormulas & " formula cell(s).", _
           vbInformation, "Add Prefix and Suffix"
    Exit Sub

NoSelection:
    MsgBox "Select the range of cells to modify first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
