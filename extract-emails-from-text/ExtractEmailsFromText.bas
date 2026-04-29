Attribute VB_Name = "ExtractEmailsFromText"
Option Explicit

' Extract Email Addresses From Text
' Source: https://excelmacros.net/tools/extract-emails-from-text
' Offline. No API calls. No external dependencies.

' Walks the selection cell by cell. For each cell, finds every email-shaped
' substring (local-part@domain.tld) and writes them, comma-separated, to the
' cell one column to the right of the source cell. Works without VBScript.RegExp
' so it runs on Mac Excel too.

Public Sub ExtractEmailsFromText()
    Dim r As Range
    Dim cell As Range
    Dim ws As Worksheet
    Dim destCol As Long
    Dim source As String
    Dim foundEmails As String
    Dim totalEmails As Long
    Dim cellsWithEmails As Long
    Dim cellsScanned As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Set ws = r.Worksheet

    Application.ScreenUpdating = False
    totalEmails = 0
    cellsWithEmails = 0
    cellsScanned = 0

    For Each cell In r.Cells
        cellsScanned = cellsScanned + 1
        If Not IsEmpty(cell.Value) Then
            source = CStr(cell.Value)
            foundEmails = ExtractEmails(source, totalEmails)
            destCol = cell.Column + 1
            If Len(foundEmails) > 0 Then
                ws.Cells(cell.Row, destCol).Value = foundEmails
                cellsWithEmails = cellsWithEmails + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Scanned " & cellsScanned & " cell(s)." & vbCrLf & _
           "Found " & totalEmails & " email(s) in " & cellsWithEmails & " cell(s)." & vbCrLf & _
           "Results written to the column on the right.", _
           vbInformation, "Extract Email Addresses"
    Exit Sub

NoSelection:
    MsgBox "Select the column or range of cells to scan for emails.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Returns a comma-separated string of all valid-looking emails found in s.
' Increments totalEmails byref by the number of emails appended.
Private Function ExtractEmails(ByVal s As String, ByRef totalEmails As Long) As String
    Dim i As Long
    Dim atPos As Long
    Dim startLocal As Long
    Dim endDomain As Long
    Dim emailCandidate As String
    Dim out As String

    i = 1
    Do While i <= Len(s)
        atPos = InStr(i, s, "@")
        If atPos = 0 Then Exit Do

        startLocal = WalkLocalLeft(s, atPos)
        endDomain = WalkDomainRight(s, atPos)

        If startLocal < atPos And endDomain > atPos Then
            emailCandidate = Mid$(s, startLocal, endDomain - startLocal + 1)
            If LooksLikeEmail(emailCandidate) Then
                If Len(out) > 0 Then out = out & ", "
                out = out & emailCandidate
                totalEmails = totalEmails + 1
            End If
        End If

        ' Continue scanning after this @ regardless of whether it parsed.
        i = atPos + 1
    Loop

    ExtractEmails = out
End Function

' Walk left from the @ to find the start of the local-part. Stops at a char
' that's not a valid local-part character. Returns the column index (1-based)
' where the local-part begins, or atPos if no valid local-part exists.
Private Function WalkLocalLeft(ByVal s As String, ByVal atPos As Long) As Long
    Dim i As Long
    Dim ch As String
    i = atPos - 1
    Do While i >= 1
        ch = Mid$(s, i, 1)
        If IsLocalChar(ch) Then
            i = i - 1
        Else
            Exit Do
        End If
    Loop
    WalkLocalLeft = i + 1
End Function

' Walk right from the @ to find the end of the domain. Stops at a char that's
' not a valid domain character.
Private Function WalkDomainRight(ByVal s As String, ByVal atPos As Long) As Long
    Dim i As Long
    Dim ch As String
    i = atPos + 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If IsDomainChar(ch) Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    ' Trim trailing dots and hyphens (common when an email ends a sentence: "name@example.com.")
    Do While i > atPos + 1
        ch = Mid$(s, i - 1, 1)
        If ch = "." Or ch = "-" Then
            i = i - 1
        Else
            Exit Do
        End If
    Loop
    WalkDomainRight = i - 1
End Function

Private Function IsLocalChar(ByVal ch As String) As Boolean
    If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or _
       (ch >= "0" And ch <= "9") Then
        IsLocalChar = True
    ElseIf ch = "." Or ch = "_" Or ch = "-" Or ch = "+" Or ch = "%" Then
        IsLocalChar = True
    Else
        IsLocalChar = False
    End If
End Function

Private Function IsDomainChar(ByVal ch As String) As Boolean
    If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or _
       (ch >= "0" And ch <= "9") Then
        IsDomainChar = True
    ElseIf ch = "." Or ch = "-" Then
        IsDomainChar = True
    Else
        IsDomainChar = False
    End If
End Function

' Final sanity check: must contain @, must have at least one '.' after the @,
' the domain must not end in '.' or '-', and the TLD must be at least 2 chars.
Private Function LooksLikeEmail(ByVal s As String) As Boolean
    Dim atPos As Long
    Dim domain As String
    Dim lastDot As Long
    Dim tld As String

    atPos = InStr(s, "@")
    If atPos < 2 Then
        LooksLikeEmail = False
        Exit Function
    End If
    domain = Mid$(s, atPos + 1)
    If Len(domain) < 3 Then
        LooksLikeEmail = False
        Exit Function
    End If
    lastDot = InStrRev(domain, ".")
    If lastDot = 0 Or lastDot = Len(domain) Then
        LooksLikeEmail = False
        Exit Function
    End If
    tld = Mid$(domain, lastDot + 1)
    If Len(tld) < 2 Then
        LooksLikeEmail = False
        Exit Function
    End If
    LooksLikeEmail = True
End Function
