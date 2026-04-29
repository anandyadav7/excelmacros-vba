Attribute VB_Name = "ConvertTextToNumbers"
Option Explicit

' Convert Text to Numbers
' Source: https://excelmacros.net/tools/convert-text-to-numbers
' Offline. No API calls. No external dependencies.

Public Sub ConvertTextToNumbers()
    Dim r As Range
    Dim cell As Range
    Dim original As String
    Dim cleaned As String
    Dim isNegative As Boolean
    Dim convertedCount As Long
    Dim skippedCount As Long
    Dim numericValue As Double

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    convertedCount = 0
    skippedCount = 0

    For Each cell In r.Cells
        If Not IsEmpty(cell.Value) Then
            ' Skip cells that are already real numbers without a leading apostrophe.
            If VarType(cell.Value) = vbString Or cell.PrefixCharacter = "'" Then
                original = CStr(cell.Value)
                cleaned = NormalizeNumberCandidate(original, isNegative)

                If Len(cleaned) > 0 And IsNumeric(cleaned) Then
                    numericValue = CDbl(cleaned)
                    If isNegative Then numericValue = -numericValue
                    cell.NumberFormat = "General"
                    cell.Value = numericValue
                    convertedCount = convertedCount + 1
                Else
                    skippedCount = skippedCount + 1
                End If
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Converted " & convertedCount & " text cell(s) to numbers." & vbCrLf & _
           "Skipped " & skippedCount & " cell(s) that were not numeric.", _
           vbInformation, "Convert Text to Numbers"
    Exit Sub

NoSelection:
    MsgBox "Select the range of cells to convert first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Strip currency symbols, thousands separators, and whitespace.
' Detect parentheses as accounting-style negatives. Returns the cleaned
' digit/decimal string and sets isNegative byref.
Private Function NormalizeNumberCandidate(ByVal s As String, ByRef isNegative As Boolean) As String
    Dim i As Long
    Dim ch As String
    Dim out As String
    Dim hasMinus As Boolean

    isNegative = False
    hasMinus = False
    s = Trim$(s)

    ' Accounting-style negative: "(1,234.56)"
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        isNegative = True
        s = Mid$(s, 2, Len(s) - 2)
    End If

    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."
                out = out & ch
            Case "-"
                If Not hasMinus Then
                    isNegative = Not isNegative
                    hasMinus = True
                End If
            Case "+"
                ' positive sign, ignore
            Case " ", ",", Chr$(160), "$", "£", "€", "¥", "₹"
                ' strip
            Case Else
                ' anything else means this is not a clean number; abandon.
                NormalizeNumberCandidate = ""
                Exit Function
        End Select
    Next i

    NormalizeNumberCandidate = out
End Function
