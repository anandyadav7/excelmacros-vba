Attribute VB_Name = "TrimWhitespaceAllCells"
Option Explicit

' Trim Whitespace From All Cells
' Source: https://excelmacros.net/tools/trim-whitespace-all-cells
' Offline. No API calls. No external dependencies.

Public Sub TrimWhitespaceAllCells()
    Dim r As Range
    Dim cell As Range
    Dim original As String
    Dim cleaned As String
    Dim cleanedCount As Long

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Then GoTo NoSelection
    If r.Cells.CountLarge < 1 Then GoTo NoSelection

    Application.ScreenUpdating = False
    cleanedCount = 0

    For Each cell In r.Cells
        ' Skip empty cells, formulas, and non-string values.
        If Not cell.HasFormula And Not IsEmpty(cell.Value) Then
            If VarType(cell.Value) = vbString Then
                original = CStr(cell.Value)
                cleaned = original

                ' Replace non-breaking space (Chr 160) with regular space.
                cleaned = Replace(cleaned, Chr(160), " ")
                ' Replace tabs and newlines with a single space.
                cleaned = Replace(cleaned, vbTab, " ")
                cleaned = Replace(cleaned, vbCrLf, " ")
                cleaned = Replace(cleaned, vbCr, " ")
                cleaned = Replace(cleaned, vbLf, " ")

                ' Collapse runs of multiple spaces into one.
                Do While InStr(cleaned, "  ") > 0
                    cleaned = Replace(cleaned, "  ", " ")
                Loop

                ' Final trim of leading/trailing spaces.
                cleaned = Trim$(cleaned)

                If cleaned <> original Then
                    cell.Value = cleaned
                    cleanedCount = cleanedCount + 1
                End If
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Cleaned " & cleanedCount & " cell(s).", vbInformation, _
           "Trim Whitespace From All Cells"
    Exit Sub

NoSelection:
    MsgBox "Select the cells you want to clean first.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
