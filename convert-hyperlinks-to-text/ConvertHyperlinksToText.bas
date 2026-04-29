Attribute VB_Name = "ConvertHyperlinksToText"
Option Explicit

' Convert Hyperlinks to Text
' Source: https://excelmacros.net/tools/convert-hyperlinks-to-text
' Offline. No API calls. No external dependencies.

' For each cell with a hyperlink in the selection, extracts the URL, removes
' the hyperlink object, and writes the URL as plain text. Useful when you need
' to copy URLs into another system that doesn't accept clickable links.

Public Sub ConvertHyperlinksToText()
    Dim r As Range
    Dim cell As Range
    Dim url As String
    Dim convertedCount As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    convertedCount = 0

    For Each cell In r.Cells
        If cell.Hyperlinks.Count > 0 Then
            url = cell.Hyperlinks(1).Address
            cell.Hyperlinks.Delete
            cell.Value = url
            convertedCount = convertedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Converted " & convertedCount & " hyperlink(s) to plain URL text.", _
           vbInformation, "Convert Hyperlinks to Text"
    Exit Sub

NoSelection:
    MsgBox "Select the range with hyperlinks first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
