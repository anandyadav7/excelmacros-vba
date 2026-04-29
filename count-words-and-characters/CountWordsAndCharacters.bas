Attribute VB_Name = "CountWordsAndCharacters"
Option Explicit

' Count Words and Characters Per Cell
' Source: https://excelmacros.net/tools/count-words-and-characters
' Offline. No API calls. No external dependencies.

' For each non-empty cell in the selection, writes a word count to the
' adjacent column and a character count to the column after that. Words are
' split on whitespace; consecutive whitespace counts as one separator.

Public Sub CountWordsAndCharacters()
    Dim r As Range
    Dim cell As Range
    Dim ws As Worksheet
    Dim text As String
    Dim words As Long
    Dim chars As Long
    Dim parts() As String
    Dim i As Long
    Dim countWritten As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection
    If r.Columns.Count > 1 Then
        MsgBox "Select a single column of text.", vbExclamation
        Exit Sub
    End If

    Set ws = r.Worksheet

    Application.ScreenUpdating = False
    countWritten = 0

    ' Header in the two destination columns.
    ws.Cells(r.Row, r.Column + 1).Value = "Words"
    ws.Cells(r.Row, r.Column + 2).Value = "Characters"
    ws.Cells(r.Row, r.Column + 1).Font.Bold = True
    ws.Cells(r.Row, r.Column + 2).Font.Bold = True

    For Each cell In r.Cells
        If cell.Row = r.Row Then
            ' header row: skip
        ElseIf Not IsEmpty(cell.Value) Then
            text = CStr(cell.Value)
            chars = Len(text)
            words = 0
            parts = Split(NormalizeWhitespace(text), " ")
            For i = LBound(parts) To UBound(parts)
                If Len(Trim$(parts(i))) > 0 Then words = words + 1
            Next i
            ws.Cells(cell.Row, r.Column + 1).Value = words
            ws.Cells(cell.Row, r.Column + 2).Value = chars
            countWritten = countWritten + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Wrote word and character counts for " & countWritten & " cell(s).", _
           vbInformation, "Count Words and Characters"
    Exit Sub

NoSelection:
    MsgBox "Select the column of text to count first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Replace tabs and newlines with spaces, collapse runs of whitespace into
' single spaces. Result is a clean space-separated string for word splitting.
Private Function NormalizeWhitespace(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String
    Dim inSpace As Boolean

    inSpace = False
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            If Not inSpace Then
                out = out & " "
                inSpace = True
            End If
        Else
            out = out & ch
            inSpace = False
        End If
    Next i

    NormalizeWhitespace = Trim$(out)
End Function
