Attribute VB_Name = "ListCellsWithErrors"
Option Explicit

' List Cells With Errors
' Source: https://excelmacros.net/tools/list-cells-with-errors
' Offline. No API calls. No external dependencies.

' Walks the selection, finds every cell with a formula error (#REF!, #VALUE!,
' #N/A, #DIV/0!, #NUM!, #NAME?, #NULL!), and writes a summary table to a new
' sheet listing the cell address, error type, and formula.

Public Sub ListCellsWithErrors()
    Dim r As Range
    Dim cell As Range
    Dim errorList As Collection
    Dim resultsSheet As Worksheet
    Dim baseName As String
    Dim sheetName As String
    Dim suffix As Long
    Dim row As Long
    Dim entry As Variant

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    Set errorList = New Collection

    Application.ScreenUpdating = False

    For Each cell In r.Cells
        If IsError(cell.Value) Then
            errorList.Add Array(cell.Address(False, False, xlA1, True), _
                                ErrorTypeName(cell.Value), _
                                CStr(cell.Formula))
        End If
    Next cell

    If errorList.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "No formula errors found in the selection.", vbInformation
        Exit Sub
    End If

    baseName = "Error Report"
    sheetName = baseName
    suffix = 1
    Do While SheetExists(sheetName)
        suffix = suffix + 1
        sheetName = baseName & " " & suffix
    Loop

    Set resultsSheet = ActiveWorkbook.Worksheets.Add( _
        After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    resultsSheet.Name = sheetName

    resultsSheet.Cells(1, 1).Value = "Cell"
    resultsSheet.Cells(1, 2).Value = "Error"
    resultsSheet.Cells(1, 3).Value = "Formula"
    resultsSheet.Range("A1:C1").Font.Bold = True

    row = 2
    For Each entry In errorList
        resultsSheet.Cells(row, 1).Value = entry(0)
        resultsSheet.Cells(row, 2).Value = entry(1)
        resultsSheet.Cells(row, 3).Value = "'" & entry(2)
        row = row + 1
    Next entry

    resultsSheet.Columns("A:C").AutoFit
    resultsSheet.Activate

    Application.ScreenUpdating = True

    MsgBox "Found " & errorList.Count & " error(s). See sheet '" & sheetName & "'.", _
           vbInformation, "List Cells With Errors"
    Exit Sub

NoSelection:
    MsgBox "Select the range to scan for errors first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function ErrorTypeName(ByVal v As Variant) As String
    Select Case CVErr(v)
        Case CVErr(xlErrRef): ErrorTypeName = "#REF!"
        Case CVErr(xlErrValue): ErrorTypeName = "#VALUE!"
        Case CVErr(xlErrName): ErrorTypeName = "#NAME?"
        Case CVErr(xlErrNA): ErrorTypeName = "#N/A"
        Case CVErr(xlErrDiv0): ErrorTypeName = "#DIV/0!"
        Case CVErr(xlErrNull): ErrorTypeName = "#NULL!"
        Case CVErr(xlErrNum): ErrorTypeName = "#NUM!"
        Case Else: ErrorTypeName = "Unknown error"
    End Select
End Function

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
