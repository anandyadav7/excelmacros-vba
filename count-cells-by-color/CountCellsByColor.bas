Attribute VB_Name = "CountCellsByColor"
Option Explicit

' Count and Sum Cells by Color
' Source: https://excelmacros.net/tools/count-cells-by-color
' Offline. No API calls. No external dependencies.

' User selects the range to scan, runs the macro, then clicks a single sample
' cell whose background color is the color to count. Macro reports count and
' (if the matching cells are numeric) sum.

Public Sub CountCellsByColor()
    Dim r As Range
    Dim sample As Range
    Dim cell As Range
    Dim targetColor As Long
    Dim matchCount As Long
    Dim numericMatchCount As Long
    Dim sumValue As Double
    Dim hexColor As String

    On Error GoTo CleanFail

    Set r = Selection
    If r Is Nothing Or r.Cells.CountLarge < 1 Then GoTo NoSelection

    On Error Resume Next
    Set sample = Application.InputBox( _
        Prompt:="Click a single cell whose background color is the color to count.", _
        Title:="Count and Sum Cells by Color", _
        Type:=8)
    On Error GoTo CleanFail

    If sample Is Nothing Then Exit Sub
    If sample.Cells.CountLarge <> 1 Then
        MsgBox "Pick exactly one sample cell.", vbExclamation
        Exit Sub
    End If

    targetColor = sample.Interior.Color
    hexColor = ColorToHex(targetColor)

    Application.ScreenUpdating = False
    matchCount = 0
    numericMatchCount = 0
    sumValue = 0

    For Each cell In r.Cells
        If cell.Interior.Color = targetColor Then
            matchCount = matchCount + 1
            If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
                sumValue = sumValue + CDbl(cell.Value)
                numericMatchCount = numericMatchCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "Color: #" & hexColor & vbCrLf & _
           "Cells matching color: " & matchCount & vbCrLf & _
           "Numeric cells in match: " & numericMatchCount & vbCrLf & _
           "Sum of numeric matches: " & sumValue, _
           vbInformation, "Count and Sum Cells by Color"
    Exit Sub

NoSelection:
    MsgBox "Select the range to scan first, then run the macro.", vbExclamation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' Excel stores Interior.Color as BGR: low byte = red, high byte = blue.
' Convert to standard #RRGGBB hex for display.
Private Function ColorToHex(ByVal c As Long) As String
    Dim r As Long, g As Long, b As Long
    r = c Mod 256
    g = (c \ 256) Mod 256
    b = (c \ 65536) Mod 256
    ColorToHex = Right$("00" & Hex$(r), 2) & _
                 Right$("00" & Hex$(g), 2) & _
                 Right$("00" & Hex$(b), 2)
End Function
