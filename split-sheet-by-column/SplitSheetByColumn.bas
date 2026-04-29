Attribute VB_Name = "SplitSheetByColumn"
Option Explicit

' Split Sheet by Column Value
' Source: https://excelmacros.net/tools/split-sheet-by-column
' Offline. No API calls. No external dependencies.

Public Sub SplitSheetByColumn()
    Dim wb As Workbook
    Dim sourceSheet As Worksheet
    Dim sourceRange As Range
    Dim data As Variant
    Dim totalRows As Long, totalCols As Long
    Dim r As Long, c As Long
    Dim colSpec As String
    Dim splitCol As Long
    Dim val As String
    Dim safeName As String
    Dim groups As Object   ' Dictionary: safeName -> Collection of source row indices
    Dim sheetName As Variant
    Dim rowList As Collection
    Dim i As Long
    Dim outArr() As Variant
    Dim newSheet As Worksheet
    Dim sheetsCreated As Long

    On Error GoTo CleanFail

    Set wb = ActiveWorkbook
    If wb Is Nothing Then GoTo NoActive
    Set sourceSheet = ActiveSheet
    If sourceSheet Is Nothing Then GoTo NoActive

    Set sourceRange = sourceSheet.UsedRange
    If sourceRange.Rows.Count < 2 Then
        MsgBox "The active sheet needs at least one header row plus one data row.", vbExclamation
        Exit Sub
    End If

    colSpec = InputBox( _
        "Which column number do you want to split by?" & vbCrLf & _
        "(Column A = 1, B = 2, C = 3...)", _
        "Split Sheet by Column Value")
    If Len(Trim$(colSpec)) = 0 Then Exit Sub
    If Not IsNumeric(Trim$(colSpec)) Then
        MsgBox "Enter a column number, like 3.", vbExclamation
        Exit Sub
    End If
    splitCol = CLng(Trim$(colSpec))
    If splitCol < 1 Or splitCol > sourceRange.Columns.Count Then
        MsgBox "Column " & splitCol & " is outside the data (which has " & _
               sourceRange.Columns.Count & " columns).", vbExclamation
        Exit Sub
    End If

    ' Read all data into memory once. Massively faster than per-row copy/paste.
    data = sourceRange.Value
    totalRows = UBound(data, 1)
    totalCols = UBound(data, 2)

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = 1 ' case-insensitive

    For r = 2 To totalRows
        val = CStr(data(r, splitCol))
        If Len(Trim$(val)) = 0 Then val = "(blank)"
        safeName = MakeSafeSheetName(val)

        If Not groups.Exists(safeName) Then
            Set groups(safeName) = New Collection
        End If
        groups(safeName).Add r
    Next r

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    sheetsCreated = 0

    For Each sheetName In groups.Keys
        ' Replace any existing sheet with the same name.
        On Error Resume Next
        wb.Worksheets(CStr(sheetName)).Delete
        On Error GoTo CleanFail

        Set newSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        newSheet.Name = CStr(sheetName)

        Set rowList = groups(CStr(sheetName))
        ReDim outArr(1 To rowList.Count + 1, 1 To totalCols)

        ' Header row.
        For c = 1 To totalCols
            outArr(1, c) = data(1, c)
        Next c

        ' Data rows.
        For i = 1 To rowList.Count
            For c = 1 To totalCols
                outArr(i + 1, c) = data(rowList(i), c)
            Next c
        Next i

        newSheet.Range(newSheet.Cells(1, 1), newSheet.Cells(rowList.Count + 1, totalCols)).Value = outArr
        sheetsCreated = sheetsCreated + 1
    Next sheetName

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    sourceSheet.Activate

    MsgBox "Created " & sheetsCreated & " split sheet(s).", vbInformation, _
           "Split Sheet by Column Value"
    Exit Sub

NoActive:
    MsgBox "Click into the workbook and sheet you want to split first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Excel forbids these characters in sheet names: : \ / ? * [ ]
' Sheet names are also limited to 31 characters.
Private Function MakeSafeSheetName(ByVal raw As String) As String
    Dim s As String
    s = raw
    s = Replace(s, ":", "-")
    s = Replace(s, "\", "-")
    s = Replace(s, "/", "-")
    s = Replace(s, "?", "")
    s = Replace(s, "*", "")
    s = Replace(s, "[", "(")
    s = Replace(s, "]", ")")
    If Len(s) > 31 Then s = Left$(s, 31)
    s = Trim$(s)
    If Len(s) = 0 Then s = "Other"
    MakeSafeSheetName = s
End Function
