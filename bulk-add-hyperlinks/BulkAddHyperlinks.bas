Attribute VB_Name = "BulkAddHyperlinks"
Option Explicit

' Bulk Add Hyperlinks
' Source: https://excelmacros.net/tools/bulk-add-hyperlinks
' Offline. No API calls. No external dependencies.

' Walks the selection, finds cells whose text looks like a URL (starts with
' http://, https://, or www.), and turns each into a clickable hyperlink.
' Cells that don't look like URLs are skipped.

Public Sub BulkAddHyperlinks()
    Dim r As Range
    Dim cell As Range
    Dim text As String
    Dim url As String
    Dim addedCount As Long
    Dim skippedCount As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    addedCount = 0
    skippedCount = 0

    For Each cell In r.Cells
        If Not IsEmpty(cell.Value) And Not cell.HasFormula Then
            text = Trim$(CStr(cell.Value))
            url = NormalizeUrl(text)
            If Len(url) > 0 Then
                cell.Hyperlinks.Delete
                cell.Hyperlinks.Add Anchor:=cell, Address:=url, TextToDisplay:=text
                addedCount = addedCount + 1
            Else
                skippedCount = skippedCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Added hyperlinks to " & addedCount & " cell(s)." & vbCrLf & _
           "Skipped " & skippedCount & " cell(s) that didn't look like URLs.", _
           vbInformation, "Bulk Add Hyperlinks"
    Exit Sub

NoSelection:
    MsgBox "Select the range of URL text to convert first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function NormalizeUrl(ByVal s As String) As String
    Dim lower As String
    lower = LCase$(s)
    If Left$(lower, 7) = "http://" Or Left$(lower, 8) = "https://" Then
        NormalizeUrl = s
    ElseIf Left$(lower, 4) = "www." Then
        NormalizeUrl = "https://" & s
    Else
        NormalizeUrl = ""
    End If
End Function
