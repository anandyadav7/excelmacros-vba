Attribute VB_Name = "ChangeCaseToUpperOrLower"
Option Explicit

' Change Case to UPPERCASE or lowercase
' Source: https://excelmacros.net/tools/change-case-to-upper-or-lower
' Offline. No API calls. No external dependencies.

' Asks whether to UPPERCASE or lowercase, then applies to every text cell in
' the selection. Skips formula cells and non-text cells.

Public Sub ChangeCaseToUpperOrLower()
    Dim r As Range
    Dim cell As Range
    Dim modeAns As String
    Dim toUpper As Boolean
    Dim modifiedCount As Long
    Dim skippedFormulas As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    modeAns = InputBox( _
        "Convert text to:" & vbCrLf & _
        "1 = UPPERCASE" & vbCrLf & _
        "2 = lowercase", _
        "Change Case", "1")
    If StrPtr(modeAns) = 0 Then Exit Sub

    Select Case Trim$(modeAns)
        Case "1": toUpper = True
        Case "2": toUpper = False
        Case Else
            MsgBox "Type 1 or 2.", vbExclamation
            Exit Sub
    End Select

    Application.ScreenUpdating = False
    modifiedCount = 0
    skippedFormulas = 0

    For Each cell In r.Cells
        If cell.HasFormula Then
            skippedFormulas = skippedFormulas + 1
        ElseIf Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            If toUpper Then
                cell.Value = UCase$(CStr(cell.Value))
            Else
                cell.Value = LCase$(CStr(cell.Value))
            End If
            modifiedCount = modifiedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Modified " & modifiedCount & " text cell(s) to " & _
           IIf(toUpper, "UPPERCASE", "lowercase") & "." & vbCrLf & _
           "Skipped " & skippedFormulas & " formula cell(s).", _
           vbInformation, "Change Case"
    Exit Sub

NoSelection:
    MsgBox "Select the range of text to convert first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
