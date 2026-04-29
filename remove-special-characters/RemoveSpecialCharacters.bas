Attribute VB_Name = "RemoveSpecialCharacters"
Option Explicit

' Remove Special Characters
' Source: https://excelmacros.net/tools/remove-special-characters
' Offline. No API calls. No external dependencies.

' Asks which characters to keep, then strips everything else from each text
' cell in the selection. Three modes: alphanumeric only, alphanumeric + spaces,
' or alphanumeric + spaces + basic punctuation (., -, _).

Public Sub RemoveSpecialCharacters()
    Dim r As Range
    Dim cell As Range
    Dim modeAns As String
    Dim allowSpaces As Boolean
    Dim allowBasicPunct As Boolean
    Dim original As String
    Dim cleaned As String
    Dim modifiedCount As Long
    Dim skippedFormulas As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    modeAns = InputBox( _
        "What characters do you want to KEEP?" & vbCrLf & vbCrLf & _
        "1 = Letters and digits only (a-z, A-Z, 0-9)" & vbCrLf & _
        "2 = Letters, digits, and spaces" & vbCrLf & _
        "3 = Letters, digits, spaces, and basic punctuation (. - _)" & vbCrLf & vbCrLf & _
        "Type 1, 2, or 3:", _
        "Remove Special Characters", "2")
    If StrPtr(modeAns) = 0 Then Exit Sub

    Select Case Trim$(modeAns)
        Case "1": allowSpaces = False: allowBasicPunct = False
        Case "2": allowSpaces = True: allowBasicPunct = False
        Case "3": allowSpaces = True: allowBasicPunct = True
        Case Else
            MsgBox "Type 1, 2, or 3.", vbExclamation
            Exit Sub
    End Select

    Application.ScreenUpdating = False
    modifiedCount = 0
    skippedFormulas = 0

    For Each cell In r.Cells
        If cell.HasFormula Then
            skippedFormulas = skippedFormulas + 1
        ElseIf Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            original = CStr(cell.Value)
            cleaned = StripChars(original, allowSpaces, allowBasicPunct)
            If cleaned <> original Then
                cell.Value = cleaned
                modifiedCount = modifiedCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Modified " & modifiedCount & " cell(s)." & vbCrLf & _
           "Skipped " & skippedFormulas & " formula cell(s).", _
           vbInformation, "Remove Special Characters"
    Exit Sub

NoSelection:
    MsgBox "Select the range of text to clean first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function StripChars(ByVal s As String, ByVal allowSpaces As Boolean, _
                            ByVal allowBasicPunct As Boolean) As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or _
           (ch >= "0" And ch <= "9") Then
            out = out & ch
        ElseIf ch = " " And allowSpaces Then
            out = out & ch
        ElseIf (ch = "." Or ch = "-" Or ch = "_") And allowBasicPunct Then
            out = out & ch
        End If
    Next i

    StripChars = out
End Function
