Attribute VB_Name = "RemoveAllFormatting"
Option Explicit

' Remove All Formatting
' Source: https://excelmacros.net/tools/remove-all-formatting
' Offline. No API calls. No external dependencies.

' Strips all formatting from the selection: cell colors, font colors, borders,
' number formats, font size, bold, italic. Cell values and formulas are left
' intact.

Public Sub RemoveAllFormatting()
    Dim r As Range
    Dim cellCount As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    cellCount = r.Cells.CountLarge

    Application.ScreenUpdating = False
    r.ClearFormats
    Application.ScreenUpdating = True

    MsgBox "Stripped formatting from " & cellCount & " cell(s)." & vbCrLf & _
           "Cell values and formulas are unchanged.", _
           vbInformation, "Remove All Formatting"
    Exit Sub

NoSelection:
    MsgBox "Select the range to strip formatting from first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
