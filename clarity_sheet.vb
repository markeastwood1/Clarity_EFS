'
' Macros for the sheet named "Clarity" -- This is the main sheet of the whole thing.
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025, 2026
'
Option Explicit
Public PreviousActiveCell As Range ' a place to hold the last active cell so we can scroll back to that location

' a way to have a global range definition so I can edit in one place
Const RANGE_USE_CASE As String = "C25"
Const RANGE_USE_CASE_SHEET As String = "UseCase Map"
Const RANGE_USE_CASE_MAP As String = "A2:B47"

Const RANGE_MODELS_QUESTION As String = "C26"
Const RANGE_MODELS_PROVIDER As String = "C27"
Const RANGE_MODEL_QUESTIONS As String = "A27:A28"

Sub Worksheet_Activate()
    ' process the cells marked as required

    Application.EnableEvents = False
    ThisWorkbook.Structure_Protection_Off
    ThisWorkbook.UnProtect_This_Sheet

    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.UsedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next

    ThisWorkbook.Structure_Protection_On
    Application.EnableEvents = True

    ThisWorkbook.Protect_This_Sheet
    showCaption
End Sub
Sub Worksheet_Change(ByVal Target As Range)
' Worksheet_Change() gets in-cell changes and Worksheet_SelectionChange only fires when the selection changes
'
    On Error Resume Next

    ThisWorkbook.UnProtect_This_Sheet
    Dim Cell As Range
    For Each Cell In Application.ActiveSheet.UsedRange
        ThisWorkbook.CheckRequiredCell Cell
    Next

    ShowHideModelsQuestions Target

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ShowHideModelsQuestions(Target As Range)
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    Dim modelsAnswer As Range

    'did the change happen where we are interested?
    Set modelsAnswer = Intersect(Target, Range(RANGE_MODELS_QUESTION))

  ' if the change happens outside of the range then ignore it
    If modelsAnswer Is Nothing Then
        Exit Sub
    End If

    If modelsAnswer.Value2 = "No" Or modelsAnswer.Value2 = "" Or modelsAnswer.Value2 = "Uncertain" Then
    ' hide the rows
        Dim r As Range

        Application.EnableEvents = False
        Range(RANGE_MODELS_PROVIDER).Value2 = ""
        Application.EnableEvents = True

        For Each r In Range(RANGE_MODEL_QUESTIONS).Rows
            r.EntireRow.Hidden = True
        Next r
        'Worksheets("Analytics").Visible = False
    Else
        ' show the rows
        For Each r In Range(RANGE_MODEL_QUESTIONS).Rows
             r.EntireRow.Hidden = False
        Next r
    End If

    ThisWorkbook.Protect_This_Sheet
End Sub
Sub ShowHideModelsTab(Target As Range)

    On Error Resume Next
    ThisWorkbook.Structure_Protection_Off

    Select Case Range(RANGE_MODELS_PROVIDER).Value2

    Case "FICO Models"
        Worksheets("Analytics").Visible = xlSheetVisible
    '    Worksheets("Opti Models").Visible = xlVeryHidden
    'Case "FICO Models + FICO Optimization"
    '    Worksheets("Analytics").Visible = xlSheetVisible
    '    Worksheets("Opti Models").Visible = xlSheetVisible
    'Case "Optimization Only"
    '    Worksheets("Analytics").Visible = xlVeryHidden
    '    Worksheets("Opti Models").Visible = xlSheetVisible
    Case "AIID Analytic Services (add notes)"
        Worksheets("Analytics").Visible = xlSheetVisible
        'Worksheets("Opti Models").Visible = xlVeryHidden
    Case Else
        ' hide them
        Worksheets("Analytics").Visible = xlVeryHidden
    '    Worksheets("Opti Models").Visible = xlVeryHidden

    End Select

    ThisWorkbook.Structure_Protection_On
End Sub
Sub ShowAllTabs()
    ' this is a helper function I'm using as I test the other things I'm implementing here
    On Error Resume Next

    If ThisWorkbook.requestPassword = True Then
        ThisWorkbook.Structure_Protection_Off

        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next

        ThisWorkbook.Structure_Protection_On
    End If

    ThisWorkbook.Worksheets("CLARITY").Activate
End Sub
Sub ResetButton()

    ThisWorkbook.UnProtect_This_Sheet

    ButtonOn button:=ActiveSheet.Shapes("Reset")

    ' hide all the tabs except Clarity and Triggers
    On Error Resume Next
    ThisWorkbook.Structure_Protection_Off

    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("SOME data on this tab will and some tabs will hide.", vbOKCancel, "Please Note!")

    If userResponse = vbCancel Then
        ButtonOff button:=ActiveSheet.Shapes("Reset")
        Exit Sub
    End If

    ThisWorkbook.Worksheets("CLARITY").Visible = xlSheetVisible

    'reset the toggles too
    ButtonOff button:=ActiveSheet.Shapes("ProfessionalServices")
    Worksheets("PS").Visible = xlVeryHidden

    DeletePicture1
    DeletePicture2

    ButtonOff button:=ActiveSheet.Shapes("Reset")

    ThisWorkbook.Protect_This_Sheet
    ThisWorkbook.Structure_Protection_On

End Sub
Sub ShowProfessionalServices()
    On Error Resume Next
    ToggleButtonByName buttonName:="ProfessionalServices"
    ToggleTabByName tabname:="PS"
End Sub
Sub ShowReviewTabs()
    On Error Resume Next

    If ThisWorkbook.requestPassword = False Then
        Exit Sub
    End If

    ToggleButtonByName buttonName:="ReviewRequired"
    ToggleTabByName tabname:="Review"

    ThisWorkbook.Structure_Protection_Off
    ThisWorkbook.Structure_Protection_On
End Sub
Sub ShowExecutiveTab()
    On Error Resume Next

    If ThisWorkbook.requestPassword = False Then
        Exit Sub
    End If

    ToggleButtonByName buttonName:="ExecutiveReview"
    ToggleTabByName tabname:="Executive"

    ThisWorkbook.Structure_Protection_Off
    ThisWorkbook.Structure_Protection_On
End Sub
Sub showCaption()
    On Error Resume Next
    ' show the full name in the title bar
    ThisWorkbook.Structure_Protection_Off
    ActiveWindow.Caption = ActiveWorkbook.FullName
    ThisWorkbook.Structure_Protection_On
End Sub
' The worksheets are largely locked meaning you can't just import images and
' put them where you want... this is how we allow that in a controlled way
Sub ImportPictureFromFile1()
    On Error Resume Next
    addImageByName "A63", "Picture 1"
End Sub
Sub ImportPictureFromFile2()
    On Error Resume Next
    addImageByName "D63", "Picture 2"
End Sub
Sub DeletePicture1()
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet

    deleteImageByName "Picture 1"
    ThisWorkbook.Protect_This_Sheet
End Sub
Sub DeletePicture2()
    On Error Resume Next

    ThisWorkbook.UnProtect_This_Sheet
    deleteImageByName "Picture 2"
    ThisWorkbook.Protect_This_Sheet
End Sub
Private Sub addImageByName(ByVal ImageLocation As String, ByVal ImageName As String)
    On Error Resume Next

    ThisWorkbook.UnProtect_This_Sheet
    Dim fNameAndPath As Variant

    fNameAndPath = Application.GetOpenFilename(Title:="Select picture to be imported")

    If fNameAndPath = False Then
        Exit Sub
    End If

    Dim s As Shape
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Clarity")
    Set s = ws.Shapes.AddPicture2(fNameAndPath, msoFalse, msoTrue, ThisWorkbook.Sheets("Clarity").Range(ImageLocation).Left, ThisWorkbook.Sheets("Clarity").Range(ImageLocation).Top, -1, -1, msoPictureCompressDocDefault)

    s.name = ImageName
    s.LockAspectRatio = msoTrue
    s.Width = GetWidthABC()

    ThisWorkbook.Protect_This_Sheet
End Sub
Private Sub deleteImageByName(ByVal arg1 As String)
    On Error Resume Next

    Dim pic As Shape

    For Each pic In ThisWorkbook.Worksheets("Clarity").Shapes
        If InStr(1, pic.name, arg1, vbTextCompare) <> 0 Then
            pic.Delete
        End If
    Next pic
End Sub
Function GetWidthABC()
    On Error Resume Next

    Dim i As Integer

    For i = 1 To 3 ' Columns A to E
        GetWidthABC = GetWidthABC + Columns(i).Width
    Next i
End Function
Sub SuperResetButton()
    On Error Resume Next
    ThisWorkbook.UnProtect_This_Sheet
    ThisWorkbook.Structure_Protection_Off

    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("Secret close all tabs button", vbOKCancel, "Super Reset!")

    If userResponse = vbCancel Then
        Exit Sub
    End If

    If ThisWorkbook.requestPassword = False Then
        Exit Sub
    End If

    DeletePicture1
    DeletePicture2

    ThisWorkbook.Protect_This_Sheet
    ThisWorkbook.Structure_Protection_On
End Sub
Sub Toggle_Workbook_Protection()
    ' this is a helper function I'm using as I test the other things I'm implementing here
    On Error Resume Next

    ' turn on protection is ok without password
    If ActiveWorkbook.ProtectStructure = False Then
        ThisWorkbook.Structure_Protection_On
        MsgBox "Workbook Structure Protected"
        Exit Sub
    End If

    ' turning off protection requires password
    If ThisWorkbook.requestPassword = True Then
        ThisWorkbook.Structure_Protection_Off
        MsgBox "Workbook Structure Unprotected"
    Else
        MsgBox "Incorrect Password, status not changed."
    End If
End Sub
Sub AddReferenceArchitecture()
    On Error Resume Next

    If ThisWorkbook.requestPassword = False Then
        Exit Sub
    End If

    ToggleButtonByName buttonName:="ShowReference"
    ToggleTabByName tabname:="Reference Architecture"

    ThisWorkbook.Structure_Protection_Off
    ThisWorkbook.Structure_Protection_On
End Sub
Sub ToggleLookupTabs()
    On Error Resume Next

    If ThisWorkbook.requestPassword = False Then
        Exit Sub
    End If
    ToggleTabByName tabname:="Lookups Clarity"
    ToggleTabByName tabname:="Lookups PS"

    ThisWorkbook.Structure_Protection_Off
    ThisWorkbook.Structure_Protection_On
End Sub
Sub ToggleTabByName(tabname As String)
    ' refactoring some code to do this in 1 place rather than in several places.
    ' this method now only toggles the tab visibility
    On Error Resume Next
    ThisWorkbook.Structure_Protection_Off

    ' toggle viability
    If Worksheets(tabname).Visible = xlSheetVisible Then
        Worksheets(tabname).Visible = xlVeryHidden
    Else
        Worksheets(tabname).Visible = xlSheetVisible
    End If

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