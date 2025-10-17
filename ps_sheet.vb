'
' Macros PS sheet
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025
'

Option Explicit
Const RANGE_PS_RESET_1 As String = "C9:D21"
Const RANGE_PS_RESET_2 As String = "C26:C37"

' used for multi-select dropdowns
Const RANGE_MULTI_SELECT_1 As String = "C10:C12"
Const RANGE_MULTI_SELECT_2 As String = "C10:C12"
Const RANGE_MULTI_SELECT_MODELS As String = "C18:C19"

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

    Dim oldValue As String
    Dim newValue As String

    ' we have to do this here because Excel doesn't allow Application.undo inside the dropdown change event
    ' and even moving to inside one of the IFs below causes issues.
    Application.EnableEvents = False
    With Target
        newValue = .Value2
        Application.Undo
        oldValue = .Value2
        .Value2 = newValue
    End With
    Application.EnableEvents = True

    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.usedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next

    Dim rngValidatedCells As Range

     'this gets a range that has ALL CELLS with VALIDATIONS (of any type)
    Set rngValidatedCells = Cells.SpecialCells(xlCellTypeAllValidation)

    ' if we have cells with validation on this sheet and the one that changed is a cell with validation
    If Not rngValidatedCells Is Nothing Then
        ' if cell that changed has a validation
        If Not Intersect(Target, rngValidatedCells) Is Nothing Then
             'if the validation type is list... then do the multi-select logic...
            If Target.Validation.Type = xlValidateList Then
                Call DropdownMultiSelect(Target, oldValue, newValue)
            End If
        End If
    End If

    'Target.AutoFit
    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ResetButton()
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    ButtonOn button:=ActiveSheet.Shapes("ResetButton")

    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("Clear data on the PS tab.", vbOKCancel, "Please Note!")

    If userResponse = vbCancel Then
        ButtonOff button:=ActiveSheet.Shapes("ResetButton")
        Exit Sub
    End If

    Application.EnableEvents = False
    Range(RANGE_PS_RESET_1).ClearContents
    Range(RANGE_PS_RESET_2).ClearContents

    Range("D12").Value2 = "You can make a list by selecting multiple items."
    Range("D14").Value2 = "You can make a list by selecting multiple items."
    Range("D18").Value2 = "You can make a list by selecting multiple items."
    Range("D19").Value2 = "You can make a list by selecting multiple items."

    Range("D15").Value2 = "What do we know about the client model(s)?"
    Range("D20").Value2 = "Document what we know about the strategy."
    Range("D21").Value2 = "Document what we know about the strategy."
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
Sub DropdownMultiSelect(Target As Range, oldValue As String, newValue As String)
    'On Error Resume Next

    'ThisWorkbook.UnProtect_This_Sheet
    Dim DelimiterType As String
    DelimiterType = ", "

    ' there are only certain cells that I want to allow multi-select...
    ' we've already filtered cells with validation (in general) and cells (more specifically) that have the "list validation" metadata
    ' but now we want to allow this only for specific cells... which is where the below is used
    Dim MULTI_FIRST As Range
    Dim MULTI_SECOND As Range
    Dim MULTI_THIRD As Range

    Dim MULTI_SELECT_CELLS As Range

    ' besides hoping to make maintenance simpler
    ' there are limits to the number of unions and the number of "line continuation"

    Set MULTI_SELECT_CELLS = Union(Range(RANGE_MULTI_SELECT_1), Range(RANGE_MULTI_SELECT_2), Range(RANGE_MULTI_SELECT_MODELS))

    On Error Resume Next

   'did the change happen where we are interested?
    Dim inMultiSelectCell As Range
    Set inMultiSelectCell = Intersect(Target, MULTI_SELECT_CELLS)

    If inMultiSelectCell Is Nothing Then
        Exit Sub
    End If

    If oldValue = "" Or newValue = "" Then
         Exit Sub
    End If

    Dim isDuplicate As Boolean
    isDuplicate = CheckAlreadyInList(oldValue, newValue)

    If Not isDuplicate Then
        Application.EnableEvents = False
        Target.Value2 = oldValue & DelimiterType & newValue
        Application.EnableEvents = True
    Else
        ' we have to pt back the prior value because otherwise it keeps the new and loses the list.
        Target.Value2 = oldValue
    End If

    'ThisWorkbook.Protect_This_Sheet
End Sub
Function CheckAlreadyInList(theList As String, newValue As String) As Boolean
    ' try to disallow duplicates in the multi-select list
    On Error Resume Next
    Dim items As Variant
    Dim item As Variant
    Dim trimmedItem As String
    Dim trimmedNewValue As String

    trimmedNewValue = Trim(newValue)
    items = Split(theList, ",")

    For Each item In items
        trimmedItem = Trim(item)
        If StrComp(trimmedItem, trimmedNewValue, vbTextCompare) = 0 Then
            CheckAlreadyInList = True
            Exit Function
        End If
    Next

    CheckAlreadyInList = False
End Function
