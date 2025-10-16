'
' Macros for the sheet named "thisworkbook", this is a special GLOBAL
' set of VB for the whole workbook
' Copyright FICO (Fair Isaac Inc) 2023, 2024, 2025, 2026
'

Option Explicit
Option Compare Text

Const RANGE_LAST_USER As String = "C2"
Const RANGE_LAST_EDIT_DATE As String = "B2"

Private Sub Workbook_Open()
    ' this runs when the workbook opens
    On Error Resume Next

    ' this is all sheets in the active workbook because Excel doesn't offer this at the sheet level
    ActiveWindow.Zoom = 100

    ' disallow drag and drop
    Application.CellDragAndDrop = False

    ThisWorkbook.ProtectAllSheets
    ThisWorkbook.Structure_Protection_On
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' this runs just before the workbook closes
    On Error Resume Next

    ThisWorkbook.Structure_Protection_Off

    UpdateLastAuthor
    UpdateLastModifiedDate

    ThisWorkbook.Structure_Protection_On
End Sub
Sub CheckRequiredCell(cell As Range)
    ' check that cells with the style "Required Normal" have a value or mark
    ' them with a red border
    On Error Resume Next

    If cell.Style = "Required Normal" Then
        With cell.Borders
            .LineStyle = xlContinuous

            If IsEmpty(cell) Then
                .Color = vbRed
                .Weight = xlWide
            Else
                 .Color = vbBlack
                 .Weight = xlThin
            End If
        End With
    End If

End Sub
Function myPassword() As String
    ' this gets reused everywhere we want to lock and unlock so we can maintain in 1 place.
    myPassword = "12345"
End Function
Private Sub UpdateLastAuthor()
' The Constant string holds the cell reference but I haven't found a way to encode the sheet name in that expression.
' Therefore you see the extra code "ThisWorkbook.Worksheets("Clarity")."
    On Error Resume Next

    ThisWorkbook.Structure_Protection_Off
    ThisWorkbook.Unprotect_Named_Sheet("Clarity")

    ThisWorkbook.Worksheets("Clarity").Range(RANGE_LAST_USER).Value = ThisWorkbook.BuiltinDocumentProperties("Last Author")

    ThisWorkbook.Protect_Named_Sheet("Clarity")
    ThisWorkbook.Structure_Protection_On
End Sub
Private Sub UpdateLastModifiedDate()
    On Error Resume Next

    ThisWorkbook.Structure_Protection_Off
    ThisWorkbook.Unprotect_Named_Sheet("Clarity")

    ThisWorkbook.Worksheets("Clarity").Range(RANGE_LAST_EDIT_DATE).Value = Now()

    ThisWorkbook.Protect_Named_Sheet("Clarity")
    ThisWorkbook.Structure_Protection_On

End Sub
Sub ProtectAllSheets()
    ' as the name implies loop all the tabs and set protection
    On Error Resume Next

    ThisWorkbook.Structure_Protection_Off
    Dim ws As Worksheet

    'For Each ws In ActiveWorkbook.Worksheets
        'ws.Protect password:=ThisWorkbook.myPassword
    'Next ws

    ' this allows CSA to import an image here... while locking everything else
    'ThisWorkbook.Worksheets("Review Summary").Protect password:=ThisWorkbook.myPassword, DrawingObjects:=False

    ThisWorkbook.Structure_Protection_On

End Sub
Private Sub UnprotectAllSheets()
    ' as the name implies loop all the tabs and remove protection

    On Error Resume Next
    ThisWorkbook.Structure_Protection_Off

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect password:=ThisWorkbook.myPassword
    Next ws

    ThisWorkbook.Structure_Protection_On

End Sub
Function getOS() As String
    Dim OSname As String

    OSname = Application.OperatingSystem

    If InStr(1, OSname, "Windows", vbTextCompare) Then
        getOS = "Windows"
    Else
        getOS = "Other"
    End If
End Function
'
' This used to be done at the sheet level but when I want to make mass changes I have to
' change things in 50 places. If all sheets reference this Sub then I can temporarily tun this off
' in one place.
Sub Protect_This_Sheet()
    'ActiveSheet.Protect ThisWorkbook.myPassword()
End Sub
Sub Protect_Named_Sheet(wsName As String)
    'ThisWorkbook.Sheets(wsName).Protect ThisWorkbook.myPassword()
End Sub
Sub UnProtect_This_Sheet()
    ActiveSheet.Unprotect ThisWorkbook.myPassword()
End Sub
Sub UnProtect_Named_Sheet(wsName As String)
    ThisWorkbook.Sheets(wsName).Unprotect ThisWorkbook.myPassword()
End Sub
Sub AllSheet_Protection_Off()
    ' this is a helper function for when I'm making changes
    On Error Resume Next

    Dim password As String

    password = InputBox("Please enter password", "Password needed!")

    If password = ThisWorkbook.myPassword() Then
       UnprotectAllSheets
    Else
        MsgBox "Incorrect Password, status not changed."
    End If
End Sub
Sub AllSheet_Protection_On()
    ' this is a helper function for when I'm making changes
    On Error Resume Next

    Dim password As String

    password = InputBox("Please enter password", "Password needed!")

    If password = ThisWorkbook.myPassword() Then
       ThisWorkbook.ProtectAllSheets
    Else
        MsgBox "Incorrect Password, status not changed."
    End If
End Sub
' make this easy to turn on and off at the workbook level
Sub Structure_Protection_On()
    'ThisWorkbook.Protect Structure:=True
End Sub

' make this easy to turn on and off at the workbook level
Sub Structure_Protection_Off()
    ThisWorkbook.Protect Structure:=False
End Sub
Function requestPassword() as Boolean
        Dim password As String

        password = InputBox("Please enter password", "Password needed!")

        If password <> ThisWorkbook.myPassword() Then
           requestPassword = False
        Else
           requestPassword = True
        End If
End Function