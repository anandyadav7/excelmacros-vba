Attribute VB_Name = "ExportEachSheetToCSV"
Option Explicit

' Export Each Sheet to CSV
' Source: https://excelmacros.net/tools/export-each-sheet-to-csv
' Offline. No API calls. No external dependencies.

' For each visible sheet in the active workbook, saves a .csv copy in the
' same folder as the workbook. Filenames follow the pattern
' "<workbook>_<sheet>.csv" with characters that aren't legal in filenames
' replaced with underscores.

Public Sub ExportEachSheetToCSV()
    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim folderPath As String
    Dim wbName As String
    Dim safeName As String
    Dim outPath As String
    Dim csvCount As Long

    On Error GoTo CleanFail

    If ActiveWorkbook.Path = "" Then
        MsgBox "Save the workbook to disk first, then run this macro.", _
               vbExclamation, "Export Each Sheet to CSV"
        Exit Sub
    End If

    folderPath = ActiveWorkbook.Path
    wbName = ActiveWorkbook.Name
    If InStr(wbName, ".") > 0 Then wbName = Left$(wbName, InStrRev(wbName, ".") - 1)

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    csvCount = 0

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Copy
            Set newWb = ActiveWorkbook
            safeName = SanitizeFileName(ws.Name)
            outPath = folderPath & Application.PathSeparator & wbName & "_" & safeName & ".csv"
            newWb.SaveAs Filename:=outPath, FileFormat:=xlCSV
            newWb.Close SaveChanges:=False
            csvCount = csvCount + 1
        End If
    Next ws

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Exported " & csvCount & " sheet(s) to CSV in:" & vbCrLf & folderPath, _
           vbInformation, "Export Each Sheet to CSV"
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Function SanitizeFileName(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "/", "\", ":", "*", "?", """", "<", ">", "|"
                out = out & "_"
            Case Else
                out = out & ch
        End Select
    Next i

    SanitizeFileName = out
End Function
