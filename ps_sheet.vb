'
' Macros PS sheet
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'

Option Explicit
Const RANGE_RESET As String = "C8:C20"

Private Sub Worksheet_Activate()
    On Error Resume Next

    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.usedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next

    Range("A1").Select

    '
    ' I'm using a Sub at the ThisWorkbook level so I can toggle
    ' this on/off easily for mass edits.
    '
    ThisWorkbook.Protect_This_Sheet

End Sub
Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.usedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next

    Target.AutoFit
    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ResetButton()

    ThisWorkbook.UnProtect_This_Sheet

    ButtonOn button:=ActiveSheet.Shapes("ResetButton")

    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("SOME data on this tab will clear and some comes from another tab.", vbOKCancel, "Please Note!")

    If userResponse = vbCancel Then
        ButtonOff button:=ActiveSheet.Shapes("ResetButton")
        Exit Sub
    End If

    Application.EnableEvents = False
    Range(RANGE_RESET).ClearContents
    Application.EnableEvents = True

    ButtonOff button:=ActiveSheet.Shapes("ResetButton")

    ThisWorkbook.Protect_This_Sheet
    ThisWorkbook.Structure_Protection_On

End Sub
Sub ToggleButtonByName(buttonName As String)
    ' refactoring some code to do this in 1 place rather than in several places.
    ' this method now only toggles the button on/off status
    On Error Resume Next

    Dim theButton As Shape
    Set theButton = ActiveSheet.Shapes(buttonName)

        If Not ButtonStatus(theButton) Then
            ButtonOn button:=theButton
        Else
            ButtonOff button:=theButton
        End If
End Sub
Private Sub ButtonOff(button As Shape)
    On Error Resume Next

    Dim GrayColor As Long
    Dim BlackColor As Long
    GrayColor = RGB(192, 192, 192)
    BlackColor = RGB(0, 0, 0)

    button.Fill.ForeColor.RGB = GrayColor
    button.TextFrame.Characters.Font.Color = BlackColor
End Sub
Private Sub ButtonOn(button As Shape)
    On Error Resume Next

    Dim GreenColor As Long
    Dim WhiteColor As Long
    GreenColor = RGB(51, 153, 102)
    WhiteColor = RGB(255, 255, 255)

    button.Fill.ForeColor.RGB = GreenColor
    button.TextFrame.Characters.Font.Color = WhiteColor
End Sub
Private Function ButtonStatus(button As Shape) As Boolean
    On Error Resume Next

    Dim GrayColor As Long
    GrayColor = RGB(192, 192, 192)

    If button.Fill.ForeColor.RGB = GrayColor Then
        ButtonStatus = False
    Else
        ButtonStatus = True
    End If
End Function
