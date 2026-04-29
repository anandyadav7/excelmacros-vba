Attribute VB_Name = "RemoveLineBreaksFromCells"
Option Explicit

' Remove Line Breaks from Cells
' Source: https://excelmacros.net/tools/remove-line-breaks-from-cells
' Offline. No API calls. No external dependencies.

' Replaces newlines (CR, LF, CR+LF) and tabs with single spaces in every text
' cell of the selection, then collapses runs of spaces. Useful when CSV
' imports or web copy-paste introduce line breaks that break sort and lookup.

Public Sub RemoveLineBreaksFromCells()
    Dim r As Range
    Dim cell As Range
    Dim original As String
    Dim cleaned As String
    Dim modifiedCount As Long
    Dim skippedFormulas As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    modifiedCount = 0
    skippedFormulas = 0

    For Each cell In r.Cells
        If cell.HasFormula Then
            skippedFormulas = skippedFormulas + 1
        ElseIf Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            original = CStr(cell.Value)
            cleaned = NormalizeLineBreaks(original)
            If cleaned <> original Then
                cell.Value = cleaned
                modifiedCount = modifiedCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Removed line breaks from " & modifiedCount & " cell(s)." & vbCrLf & _
           "Skipped " & skippedFormulas & " formula cell(s).", _
           vbInformation, "Remove Line Breaks"
    Exit Sub

NoSelection:
    MsgBox "Select the range of cells to clean first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function NormalizeLineBreaks(ByVal s As String) As String
    Dim result As String
    result = Replace(s, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, vbTab, " ")
    ' Collapse runs of spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    NormalizeLineBreaks = Trim$(result)
End Function
