Attribute VB_Name = "ListAllSheetNames"
Option Explicit

' List All Sheet Names
' Source: https://excelmacros.net/tools/list-all-sheet-names
' Offline. No API calls. No external dependencies.

' Inserts a new "Sheet Index" worksheet at position 1 with a numbered list of
' every other sheet. Each sheet name is a hyperlink jumping to that sheet's
' A1, plus a column showing whether the sheet is hidden.

Public Sub ListAllSheetNames()
    Dim ws As Worksheet
    Dim indexSheet As Worksheet
    Dim baseName As String
    Dim sheetName As String
    Dim suffix As Long
    Dim row As Long
    Dim listed As Long

    On Error GoTo CleanFail

    baseName = "Sheet Index"
    sheetName = baseName
    suffix = 1
    Do While SheetExists(sheetName)
        suffix = suffix + 1
        sheetName = baseName & " " & suffix
    Loop

    Set indexSheet = ActiveWorkbook.Worksheets.Add(Before:=ActiveWorkbook.Worksheets(1))
    indexSheet.Name = sheetName

    indexSheet.Cells(1, 1).Value = "#"
    indexSheet.Cells(1, 2).Value = "Sheet Name"
    indexSheet.Cells(1, 3).Value = "Visible"
    indexSheet.Range("A1:C1").Font.Bold = True

    row = 2
    listed = 0
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> sheetName Then
            indexSheet.Cells(row, 1).Value = listed + 1
            indexSheet.Hyperlinks.Add _
                Anchor:=indexSheet.Cells(row, 2), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
            If ws.Visible = xlSheetVisible Then
                indexSheet.Cells(row, 3).Value = "yes"
            Else
                indexSheet.Cells(row, 3).Value = "hidden"
            End If
            row = row + 1
            listed = listed + 1
        End If
    Next ws

    indexSheet.Columns("A:C").AutoFit
    indexSheet.Activate
    indexSheet.Range("A1").Select

    MsgBox "Created '" & sheetName & "' with links to " & listed & " sheet(s).", _
           vbInformation, "List All Sheet Names"
    Exit Sub

CleanFail:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function SheetExists(ByVal n As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If StrComp(ws.Name, n, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    SheetExists = False
End Function
