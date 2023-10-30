VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uCalendar 
   Caption         =   "frmDatePicker"
   ClientHeight    =   7248
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9600.001
   OleObjectBlob   =   "uCalendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public gDate        As New clsDate

Public Function Datepicker(Optional DateInput As Object) As String
    '@AssignedModule uCalendar
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    '@INCLUDE DECLARATION gDate
    Dim str         As String
    If VBA.TypeName(DateInput) = "Textbox" Or VBA.TypeName(DateInput) = "Range" Then str = DateInput.value
    If VBA.TypeName(DateInput) = "CommandButton" Or VBA.TypeName(DateInput) = "Label" Then str = DateInput.Caption

    'If DatepInput <> "" Then <--- replaced with next line
    If str <> "" Then

        Dim curDate As String
        With uCalendar
            .txtYearName = year(DateInput)
            .txtMonthNumber = Format(DateInput, "mm")

        End With

        With gDate
            .createDates txtYearName, txtMonthNumber
            .SelectDate .dFrame.Controls("lblDate" & Day(DateInput))
        End With
    Else

        With uCalendar
            .lblSelectedDate = Day(Date)
            .lblSelectedMonth = Format(Date, "mmmm")
            .lblSelectedYear = year(Date)
            curDate = Day(Date) & "." & .txtMonthNumber Mod 12 & "." & .txtYearName
            .lblSelectedDateName = Format(curDate, "dddd")
            .txtSelectedDate = Format(curDate, "dd.mm.yyyy")
            .txtMonthNumber = Format(Date, "mm")
        End With

        With gDate.lblDateBack

        End With

    End If

    Me.Show

    Datepicker = Me.txtSelectedDate.value

    If VBA.TypeName(DateInput) = "TextBox" Or VBA.TypeName(DateInput) = "Range" Then
        DateInput.value = Me.txtSelectedDate.value
    ElseIf VBA.TypeName(DateInput) = "CommandButton" Or VBA.TypeName(DateInput) = "Label" Then
        DateInput.Caption = Me.txtSelectedDate.value
    Else
        'Datepicker = Me.txtSelectedDate.Value <--- put this before If to return the value anyway
    End If

End Function

Private Sub frameDate_Click()
    '@AssignedModule uCalendar
    '@INCLUDE USERFORM uCalendar
    frameMonth.Visible = False
    frameYear.Visible = False
End Sub

Private Sub lblChoose_Click()
    '@AssignedModule uCalendar
    '@INCLUDE USERFORM uCalendar
    Unload Me
End Sub

Private Sub lblChoose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    '@AssignedModule uCalendar
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    '@INCLUDE DECLARATION gDate
    gDate.dFrame.Controls("lblDateBack").Visible = False
    gDate.dayMouseOut

End Sub

Private Sub lblClose_Click()
    '@AssignedModule uCalendar
    '@INCLUDE USERFORM uCalendar
    txtSelectedDate = ""
    Unload Me

End Sub

Private Sub lblMonthName_Click()
    '@AssignedModule uCalendar
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    '@INCLUDE DECLARATION gDate
    frameYear.Visible = False
    With frameMonth
        .Visible = True
        .Left = lblMonthName.Left
        .Top = lblMonthName.Top + 20

    End With
    gDate.createMonth txtMonthNumber
End Sub

Private Sub lblNextMonth_Click()
    '@AssignedModule uCalendar
    '@INCLUDE PROCEDURE getMonth
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    '@INCLUDE DECLARATION gDate
    With txtMonthNumber
        .TEXT = .TEXT + 1
        lblMonthName = getMonth(.TEXT)

        If lblMonthName = "JANUARY" Then
            txtYearName = txtYearName + 1
        End If
        '        gDate.createDates txtYearName, .Text

    End With
End Sub

Private Sub lblPreviewMonth_Click()
    '@AssignedModule uCalendar
    '@INCLUDE PROCEDURE getMonth
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    '@INCLUDE DECLARATION gDate
    With txtMonthNumber
        .TEXT = .TEXT - 1

        lblMonthName = getMonth(.TEXT)
        If lblMonthName.Caption = "DECEMBER" Then
            txtYearName = txtYearName - 1
        End If
        '       gDate.createDates txtYearName, .Text
    End With
End Sub

Private Sub lblRightBar_Click()
    '@AssignedModule uCalendar
    '@INCLUDE USERFORM uCalendar

End Sub

Private Sub lblRightBar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    '@AssignedModule uCalendar
    '@INCLUDE PROCEDURE moverForm
    '@INCLUDE USERFORM uCalendar
    moverForm Me, Me, Button
End Sub

Private Sub lblToday_Click()
    '@AssignedModule uCalendar
    '@INCLUDE PROCEDURE getMonth
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    '@INCLUDE DECLARATION gDate
    txtYearName = Format(Date, "yyyy")
    txtMonthNumber = Format(Date, "m")
    gDate.createDates Format(Date, "yyyy"), Format(Date, "mm")
    gDate.SelectDate gDate.dFrame.Controls("lblDate" & Day(Date))
End Sub

Private Sub txtMonthNumber_Change()
    '@AssignedModule uCalendar
    '@INCLUDE PROCEDURE getMonth
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    '@INCLUDE DECLARATION gDate
    lblMonthName = getMonth(txtMonthNumber)
    gDate.createDates txtYearName, txtMonthNumber
End Sub

Private Sub txtSelectedYear_Change()
    '@AssignedModule uCalendar
    '@INCLUDE USERFORM uCalendar

End Sub

Private Sub txtYearName_Change()
    '@AssignedModule uCalendar
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    '@INCLUDE DECLARATION gDate
    If Len(txtYearName.TEXT) = 4 Then
        gDate.createDates txtYearName, txtMonthNumber
    End If
End Sub

Private Sub txtYearName_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    '@AssignedModule uCalendar
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    '@INCLUDE DECLARATION gDate
    frameMonth.Visible = False
    With frameYear
        .Left = txtYearName.Left
        .Top = txtYearName.Top + 20
        .Visible = True
    End With
    gDate.createYear txtYearName
End Sub

Private Sub UserForm_Activate()
    '@AssignedModule uCalendar
    '@INCLUDE USERFORM uCalendar
    lblToday_Click
End Sub

Private Sub UserForm_Click()
    '@AssignedModule uCalendar
    '@INCLUDE USERFORM uCalendar
    Me.frameMonth.Visible = False
    Me.frameYear.Visible = False
End Sub

Private Sub UserForm_Initialize()
    '@AssignedModule uCalendar
    '@INCLUDE PROCEDURE HideTitleBarAndBorder
    '@INCLUDE PROCEDURE removeTudo
    '@INCLUDE USERFORM uCalendar
    Dim sMonth      As Integer
    SelectedDay = ""
    removeTudo Me
    HideTitleBarAndBorder Me

    With Me
        .Height = 308
        .Width = 480
    End With

    IconDesign lblPreviewMonth, "&HE26C"
    IconDesign lblNextMonth, "&HE26B"

End Sub

Private Sub IconDesign(Ctrl As control, IconCode As String)
    '@AssignedModule uCalendar
    '@INCLUDE USERFORM uCalendar
    '@INCLUDE DECLARATION Ctrl
    With Ctrl
        .Font.Name = "Segoe MDL2 Assets"
        .Caption = ChrW(IconCode)
        .Font.Size = 12
        .ForeColor = RGB(191, 191, 191)
        .TextAlign = fmTextAlignLeft
        .BorderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    '    moverForm Me, Me, Button
    '@AssignedModule uCalendar
    '@INCLUDE PROCEDURE moverForm
    '@INCLUDE USERFORM uCalendar
End Sub

