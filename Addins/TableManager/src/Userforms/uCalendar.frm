VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uCalendar 
   Caption         =   "frmDatePicker"
   ClientHeight    =   6168
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9372.001
   OleObjectBlob   =   "uCalendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Public gDate As New clsDate

Public Function Datepicker(Optional DateInput As Object) As String
    Dim str As String
    If VBA.TypeName(DateInput) = "Textbox" Or VBA.TypeName(DateInput) = "Range" Then str = DateInput.Value
    If VBA.TypeName(DateInput) = "CommandButton" Or VBA.TypeName(DateInput) = "Label" Then str = DateInput.Caption

    'If DatepInput <> "" Then <--- replaced with next line
    If str <> "" Then

        Dim curDate As String
        With uCalendar
            .txtYearName = Year(DateInput)
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
            .lblSelectedYear = Year(Date)
            curDate = Day(Date) & "." & .txtMonthNumber Mod 12 & "." & .txtYearName
            .lblSelectedDateName = Format(curDate, "dddd")
            .txtSelectedDate = Format(curDate, "dd.mm.yyyy")
            .txtMonthNumber = Format(Date, "mm")
        End With

        With gDate.lblDateBack

        End With

    End If

    Me.Show

    Datepicker = Me.txtSelectedDate.Value

    If VBA.TypeName(DateInput) = "TextBox" Or VBA.TypeName(DateInput) = "Range" Then
        DateInput.Value = Me.txtSelectedDate.Value
    ElseIf VBA.TypeName(DateInput) = "CommandButton" Or VBA.TypeName(DateInput) = "Label" Then
        DateInput.Caption = Me.txtSelectedDate.Value
    Else
        'Datepicker = Me.txtSelectedDate.Value <--- put this before If to return the value anyway
    End If

End Function

Private Sub frameDate_Click()
    frameMonth.Visible = False
    frameYear.Visible = False
End Sub

Private Sub lblChoose_Click()
    Unload Me
End Sub

Private Sub lblChoose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    gDate.dFrame.Controls("lblDateBack").Visible = False
    gDate.dayMouseOut

End Sub

Private Sub lblClose_Click()
    txtSelectedDate = ""
    Unload Me

End Sub

Private Sub lblMonthName_Click()
    frameYear.Visible = False
    With frameMonth
        .Visible = True
        .Left = lblMonthName.Left
        .Top = lblMonthName.Top + 20

    End With
    gDate.createMonth txtMonthNumber
End Sub

Private Sub lblNextMonth_Click()
    With txtMonthNumber
        .Text = .Text + 1
        lblMonthName = getMonth(.Text)

        If lblMonthName = "OCAK" Then
            txtYearName = txtYearName + 1
        End If
        '        gDate.createDates txtYearName, .Text

    End With
End Sub

Private Sub lblPreviewMonth_Click()
    With txtMonthNumber
        .Text = .Text - 1

        lblMonthName = getMonth(.Text)
        If lblMonthName.Caption = "ARALIK" Then
            txtYearName = txtYearName - 1
        End If
        '       gDate.createDates txtYearName, .Text
    End With
End Sub

Private Sub lblRightBar_Click()

End Sub

Private Sub lblRightBar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    moverForm Me, Me, Button
End Sub

Private Sub lblToday_Click()
    lblMonthName = getMonth(Format(Date, "mm"))
    txtYearName = Format(Date, "yyyy")
    txtMonthNumber = Format(Date, "m")
    gDate.createDates Format(Date, "yyyy"), Format(Date, "mm")
    
    gDate.SelectDate Controls("lblDate" & Day(Now)) 'added by alex
End Sub

Private Sub txtMonthNumber_Change()
    lblMonthName = getMonth(txtMonthNumber)
    gDate.createDates txtYearName, txtMonthNumber
End Sub

Private Sub txtSelectedYear_Change()

End Sub

Private Sub txtYearName_Change()
    If Len(txtYearName.Text) = 4 Then
        gDate.createDates txtYearName, txtMonthNumber
    End If
End Sub

Private Sub txtYearName_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    frameMonth.Visible = False
    With frameYear
        .Left = txtYearName.Left
        .Top = txtYearName.Top + 20
        .Visible = True
    End With
    gDate.createYear txtYearName
End Sub

Private Sub UserForm_Activate()
    lblToday_Click
End Sub

Private Sub UserForm_Click()
    Me.frameMonth.Visible = False
    Me.frameYear.Visible = False
End Sub

Private Sub UserForm_Initialize()
    Dim sMonth As Integer
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
End Sub

