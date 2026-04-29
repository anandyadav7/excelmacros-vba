Attribute VB_Name = "ClearAllComments"
Option Explicit

' Clear All Comments
' Source: https://excelmacros.net/tools/clear-all-comments
' Offline. No API calls. No external dependencies.

' Walks the selection and deletes every cell comment found. Useful before
' sharing a workbook externally, or when cleaning up a model that's grown
' messy with author notes from multiple people.

Public Sub ClearAllComments()
    Dim r As Range
    Dim cell As Range
    Dim deletedCount As Long
    Dim ans As VbMsgBoxResult

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    ans = MsgBox("Delete every cell comment in the selection?" & vbCrLf & _
                 "This is irreversible. Save first if you want a backup.", _
                 vbYesNo + vbExclamation, "Clear All Comments")
    If ans <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    deletedCount = 0

    For Each cell In r.Cells
        If Not cell.Comment Is Nothing Then
            cell.Comment.Delete
            deletedCount = deletedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Deleted " & deletedCount & " comment(s).", _
           vbInformation, "Clear All Comments"
    Exit Sub

NoSelection:
    MsgBox "Select the range with comments to clear first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
