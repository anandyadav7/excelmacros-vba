Attribute VB_Name = "SplitFullNames"
Option Explicit

' Split Full Names Into First and Last
' Source: https://excelmacros.net/tools/split-full-names
' Offline. No API calls. No external dependencies.

' Splits a single column of full names into First Name and Last Name in the
' two columns immediately to the right. Supports "First Last" and "Last, First"
' formats. Middle names go to the first-name column.

Public Sub SplitFullNames()
    Dim r As Range
    Dim cell As Range
    Dim fmtAns As String
    Dim formatType As Long
    Dim ws As Worksheet
    Dim source As String
    Dim firstName As String
    Dim lastName As String
    Dim parts() As String
    Dim splitCount As Long
    Dim i As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection
    If r.Columns.Count > 1 Then
        MsgBox "Select a single column of names.", vbExclamation
        Exit Sub
    End If

    fmtAns = InputBox( _
        "What format are your names in?" & vbCrLf & vbCrLf & _
        "1 = First Last (e.g. 'Priya Sharma' or 'Ana Maria Lopez')" & vbCrLf & _
        "2 = Last, First (e.g. 'Sharma, Priya')" & vbCrLf & vbCrLf & _
        "Type 1 or 2:", _
        "Split Full Names", "1")
    If StrPtr(fmtAns) = 0 Then Exit Sub

    Select Case Trim$(fmtAns)
        Case "1": formatType = 1
        Case "2": formatType = 2
        Case Else
            MsgBox "Type 1 or 2.", vbExclamation
            Exit Sub
    End Select

    Set ws = r.Worksheet

    Application.ScreenUpdating = False
    splitCount = 0

    For Each cell In r.Cells
        If Not IsEmpty(cell.Value) Then
            source = Trim$(CStr(cell.Value))
            firstName = ""
            lastName = ""

            If formatType = 1 Then
                parts = Split(source, " ")
                If UBound(parts) >= 1 Then
                    lastName = Trim$(parts(UBound(parts)))
                    firstName = ""
                    For i = LBound(parts) To UBound(parts) - 1
                        If Len(firstName) > 0 Then firstName = firstName & " "
                        firstName = firstName & Trim$(parts(i))
                    Next i
                Else
                    firstName = source
                End If
            Else
                parts = Split(source, ",")
                If UBound(parts) >= 1 Then
                    lastName = Trim$(parts(0))
                    firstName = Trim$(parts(1))
                Else
                    lastName = source
                End If
            End If

            ws.Cells(cell.Row, r.Column + 1).Value = firstName
            ws.Cells(cell.Row, r.Column + 2).Value = lastName
            splitCount = splitCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "Split " & splitCount & " name(s) into First and Last columns.", _
           vbInformation, "Split Full Names"
    Exit Sub

NoSelection:
    MsgBox "Select the column of full names first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
