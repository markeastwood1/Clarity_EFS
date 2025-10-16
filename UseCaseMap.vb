'
' Macros for the sheet named "UseCase Map"
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025, 2026
'
Option Explicit
Option Compare Text

' a way to have a global range definition so I can edit in one place
Const RANGE_USE_CASE As String = "C25"

Sub FindUseCaseTabs()
Dim cellValue As String
cellValue = Worksheets("Clarity").Range(RANGE_USE_CASE).Value

Dim Msg, Title, Style

Msg = "Use Case: " & cellValue & vbCrLf & vbCrLf
Title = "Use Case"
Style = vbOKOnly Or vbExclamation Or vbSystemModal Or vbMsgBoxSetForeground

MsgBox Msg, Style, Title
End Sub

' Get the cells with an X for the selected row
Sub CheckRowForData()
    Dim cell As Range
    Dim targetRow As Long
    Dim ws As Worksheet

    ' Set the worksheet and target row
    Set ws = ActiveSheet
    targetRow = 5 ' Change this to the desired row number

    ' Loop through each cell in the specified row within the used range
    For Each cell In ws.Rows(targetRow).UsedRange
        ' Check if the cell is not empty
        If Not IsEmpty(cell) Then
            ' Perform an action, such as displaying the cell's value
            Debug.Print "Cell " & cell.Address & " contains: " & cell.Value
        End If
    Next cell
End Sub