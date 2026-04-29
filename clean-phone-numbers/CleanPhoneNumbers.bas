Attribute VB_Name = "CleanPhoneNumbers"
Option Explicit

' Clean Phone Numbers
' Source: https://excelmacros.net/tools/clean-phone-numbers
' Offline. No API calls. No external dependencies.

Public Sub CleanPhoneNumbers()
    Dim r As Range
    Dim cell As Range
    Dim original As String
    Dim cleaned As String
    Dim keepPlusAns As String
    Dim keepPlus As Boolean
    Dim cleanedCount As Long
    Dim skippedCount As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    keepPlusAns = InputBox( _
        "Keep a leading '+' for international numbers? Type Y for yes, anything else for no." & vbCrLf & vbCrLf & _
        "Y: '+1 (415) 555-0100' becomes '+14155550100'" & vbCrLf & _
        "N: '+1 (415) 555-0100' becomes '14155550100'", _
        "Clean Phone Numbers", "Y")
    If StrPtr(keepPlusAns) = 0 Then Exit Sub
    keepPlus = (UCase$(Trim$(keepPlusAns)) = "Y")

    Application.ScreenUpdating = False
    cleanedCount = 0
    skippedCount = 0

    For Each cell In r.Cells
        If Not IsEmpty(cell.Value) Then
            original = CStr(cell.Value)
            cleaned = StripToDigits(original, keepPlus)
            If Len(cleaned) > 0 And cleaned <> original Then
                ' Force text format so leading zeros and long numbers don't get mangled.
                cell.NumberFormat = "@"
                cell.Value = cleaned
                cleanedCount = cleanedCount + 1
            ElseIf Len(cleaned) = 0 Then
                skippedCount = skippedCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Cleaned " & cleanedCount & " phone number(s)." & vbCrLf & _
           "Skipped " & skippedCount & " cell(s) with no digits.", _
           vbInformation, "Clean Phone Numbers"
    Exit Sub

NoSelection:
    MsgBox "Select the range of phone numbers to clean.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function StripToDigits(ByVal s As String, ByVal keepPlus As Boolean) As String
    Dim i As Long
    Dim ch As String
    Dim out As String
    Dim sawPlus As Boolean

    s = Trim$(s)
    sawPlus = False

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            out = out & ch
        ElseIf ch = "+" And keepPlus And Not sawPlus And Len(out) = 0 Then
            out = "+"
            sawPlus = True
        End If
        ' Everything else (spaces, dashes, parens, dots, letters) is dropped.
    Next i

    StripToDigits = out
End Function
