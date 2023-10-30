VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents dForm As MSForms.UserForm
Attribute dForm.VB_VarHelpID = -1
Public WithEvents dFrame As MSForms.Frame
Attribute dFrame.VB_VarHelpID = -1
Public WithEvents mFrame As MSForms.Frame
Attribute mFrame.VB_VarHelpID = -1
Public WithEvents yFrame As MSForms.Frame
Attribute yFrame.VB_VarHelpID = -1
Public WithEvents lblDate As MSForms.Label
Attribute lblDate.VB_VarHelpID = -1
Public WithEvents lblMonth As MSForms.Label
Attribute lblMonth.VB_VarHelpID = -1
Public WithEvents lblYear As MSForms.Label
Attribute lblYear.VB_VarHelpID = -1
Public WithEvents lblTodayBack As MSForms.Label
Attribute lblTodayBack.VB_VarHelpID = -1
Public WithEvents lblSelectedDateBack As MSForms.Label
Attribute lblSelectedDateBack.VB_VarHelpID = -1
Public WithEvents lblDateBack As MSForms.Label
Attribute lblDateBack.VB_VarHelpID = -1
Public WithEvents lblMonthBack As MSForms.Label
Attribute lblMonthBack.VB_VarHelpID = -1
Public WithEvents lblYearBack As MSForms.Label
Attribute lblYearBack.VB_VarHelpID = -1

Public colDate      As New Collection
Public aDate        As New clsDate
Public colMonth     As New Collection
Public aMonth       As New clsDate
Public colYear      As New Collection
Public aYear        As New clsDate

Public Sub createDates(cYear As Integer, cMonth As Integer)
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    Dim lLeft As Integer, lTop As Integer
    Dim fDate       As Integer
    Set dForm = uCalendar
    Set dFrame = dForm.frameDate
    PassiveDay = "": SelectedDay = ""
    '  MsgBox cYear & "-" & cMonth
    '************************************
    '***********//frame date start//*****
    With dFrame
        .clear
    End With
    '***********//frame date start//*****
    '************************************
    Dim dt          As Date
    If cMonth Mod 12 = 0 Then cMonth = 12 Else cMonth = cMonth Mod 12
    dt = DateSerial(cYear, cMonth, 1)
    Dim firstDate   As Long
    firstDate = Weekday(dt, 2)
    Dim ilk         As Date
    ilk = DateSerial(year(dt), month(dt), 1 - (firstDate - 1))

    lLeft = 6: lTop = 6
    Dim gunsayi     As String
    Dim i           As Long
    For i = ilk To ilk + 41
        gunsayi = Format(Day(i), "#0")

        Set lblDate = dFrame.Controls.Add("Forms.Label.1", "lblDate" & gunsayi)
        With lblDate
            '            Debug.Print lblDate.Name
            .Left = lLeft
            .Top = lTop
            .Width = 32
            .Font.Name = "MontSerrat Medium"
            .Font.Size = 9
            .ForeColor = &H48372C
            .TextAlign = fmTextAlignCenter
            .BackStyle = fmBackStyleTransparent
            .Caption = gunsayi
            .Tag = i

            If CInt(year(CDate(i))) = CInt(year(CDate(dt))) And CInt(month(CDate(i))) = CInt(month(CDate(dt))) Then
                .Enabled = True
            Else
                .Enabled = False
            End If

            If Weekday(i, 2) = 7 Or Weekday(i, 2) = 6 Then
                .ForeColor = &H2144FF
            End If

            If i = Date And .Enabled = True Then

                Set lblTodayBack = dFrame.Controls.Add("Forms.Label.1", "lblTodayBack")
                With lblTodayBack
                    .Height = 1.2
                    .Width = 22
                    .Top = lblDate.Top + 16
                    .Left = lblDate.Left + 5
                    .BackColor = &H2144FF
                    .ZOrder (1)
                End With
            End If

            Set aDate = New clsDate
            Set aDate.lblDate = lblDate
            Set aDate.dFrame = dFrame
            Set aDate.dForm = dForm
            colDate.Add aDate

        End With

        If lLeft <= 238 Then
            lLeft = lLeft + 44

        Else
            lLeft = 6
            lTop = lTop + 30
        End If

    Next i
    Set lblSelectedDateBack = dFrame.Controls.Add("Forms.Label.1", "lblSelectedDateBack")
    With lblSelectedDateBack
        .Height = 26
        .Width = 30
        .Picture = dForm.lblSelectedDateBack.Picture

        .Visible = False
        .ZOrder (1)
    End With

    Set lblDateBack = dFrame.Controls.Add("Forms.Label.1", "lblDateBack")
    With lblDateBack
        .Height = 26
        .Width = 30
        .Picture = dForm.lblDateBack.Picture
        .Visible = False
        .ZOrder (1)

    End With

End Sub

Public Sub createMonth(snMonth As Integer)
    '@AssignedModule clsDate
    '@INCLUDE PROCEDURE getMonth
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    Dim lTop As Integer, i As Integer
    Dim fDate       As Integer

    Set dForm = uCalendar
    Set mFrame = dForm.frameMonth

    With mFrame
        .clear
    End With

    lTop = 4
    For i = 1 To 12
        Set lblMonth = mFrame.Controls.Add("Forms.Label.1", "lblMonth" & i)
        With lblMonth
            .Left = 0
            .Top = lTop
            .Width = 90
            .Height = 14
            .Font.Name = "MontSerrat Medium"
            .Font.Size = 9

            If i = snMonth Then
                .ForeColor = vbWhite
                PassiveMonth = lblMonth.Name
            Else
                .ForeColor = vbGrayText
            End If

            .TextAlign = fmTextAlignCenter
            .BackStyle = fmBackStyleTransparent
            .Caption = getMonth(i)
            .Tag = i
        End With
        lTop = lTop + 18

        Set aMonth = New clsDate
        Set aMonth.lblMonth = lblMonth
        Set aMonth.mFrame = mFrame
        Set aMonth.dForm = dForm
        colMonth.Add aMonth

    Next i
    Set lblMonth = Nothing

    Set lblMonthBack = mFrame.Controls.Add("Forms.Label.1", "lblMonthBack")
    With lblMonthBack
        .Left = 0
        .Top = (snMonth - 1) * 18
        .Width = 90
        .BackColor = &H2144FF
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleNone
        .ZOrder (1)
    End With

End Sub

Public Sub createYear(snYear As Integer)
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    Dim lTop As Integer, i As Integer

    Set dForm = uCalendar
    Set yFrame = dForm.frameYear

    With yFrame
        .clear
    End With

    lTop = 4
    For i = 1 To 12
        Set lblYear = yFrame.Controls.Add("Forms.Label.1", "lblYear" & i)
        With lblYear
            .Left = 0
            .Top = lTop
            .Width = 90
            .Height = 14
            .Font.Name = "MontSerrat Medium"
            .Font.Size = 9
            .ForeColor = vbGrayText
            .Caption = Format(Date, "yyyy") + i - 4

            If .Caption = snYear Then
                .ForeColor = vbWhite
                PassiveYear = lblYear.Name
            Else
                .ForeColor = vbGrayText
            End If

            .TextAlign = fmTextAlignCenter
            .BackStyle = fmBackStyleTransparent

            .Tag = i
        End With
        lTop = lTop + 18

        Set aYear = New clsDate
        Set aYear.lblYear = lblYear
        Set aYear.yFrame = yFrame
        Set aYear.dForm = dForm
        colYear.Add aYear

    Next i

    Set lblYearBack = yFrame.Controls.Add("Forms.Label.1", "lblYearBack")
    With lblYearBack
        .Left = 0
        .Top = yFrame.Controls(PassiveYear).Top - 4
        .Width = 90
        .BackColor = &H2144FF
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleNone
        .ZOrder (1)
    End With

End Sub

Private Sub dForm_Click()
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    framevisibleFalse
End Sub

Private Sub dFrame_Click()
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    framevisibleFalse
End Sub

Sub framevisibleFalse()
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    uCalendar.frameMonth.Visible = False
    uCalendar.frameYear.Visible = False
End Sub

Private Sub lblDate_Click()
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate

    If SelectedDay <> "" Then
        If Weekday(SelectedDayTag, 2) = 7 Or Weekday(SelectedDayTag, 2) = 6 Then
            dFrame.Controls(SelectedDay).ForeColor = &H2144FF
        Else
            dFrame.Controls(SelectedDay).ForeColor = &H48372C
        End If

    Else

    End If
    SelectDate lblDate
End Sub

Private Sub lblDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    SelectDate lblDate
    Unload uCalendar

End Sub

Sub SelectDate(ctrl As control)
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    '@INCLUDE DECLARATION Ctrl
    Dim curDate     As Date
    With uCalendar
        .lblSelectedDate = ctrl.Caption
        .lblSelectedMonth = .lblMonthName
        .lblSelectedYear = .txtYearName
        .lblSelectedDateName = Format(ctrl.Tag, "dddd")
        .txtSelectedDate = Format(ctrl.Tag, "dd.mm.yyyy")

        lblSelectedDateBackPosition ctrl
        ctrl.ForeColor = vbWhite
        SelectedDay = ctrl.Name
        SelectedDayTag = ctrl.Tag
    End With
End Sub

Private Sub lblDate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    framevisibleFalse
    lblDateBackPosition lblDate
    dayMouseOut
    lblDate.ForeColor = vbWhite
    PassiveDay = lblDate.Name
    PassiveDayTag = lblDate.Tag
End Sub

Private Sub lblDateBackPosition(ctrl As control)
    '@AssignedModule clsDate
    '@INCLUDE PROCEDURE MouseMoveIcon
    '@INCLUDE CLASS clsDate
    '@INCLUDE DECLARATION Ctrl
    MouseMoveIcon
    With dFrame.Controls("lblDateBack")
        .Visible = True
        .Top = ctrl.Top - 5
        .Left = ctrl.Left + 1
    End With
End Sub

Private Sub lblMonth_Click()
    '@AssignedModule clsDate
    '@INCLUDE PROCEDURE getMonth
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    With uCalendar
        .txtMonthNumber = lblMonth.Tag
        .frameMonth.Visible = False
        .lblMonthName = getMonth(.txtMonthNumber)
        createDates .txtYearName, .txtMonthNumber
    End With
End Sub

Private Sub lblYear_Click()
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    '@INCLUDE USERFORM uCalendar
    With uCalendar
        .txtYearName = lblYear
        .frameYear.Visible = False
        createDates .txtYearName, .txtMonthNumber
    End With

End Sub

Private Sub lblMonth_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '    MsgBox PassiveMonth
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    lblMonthBackPosition lblMonth
    monthMouseOut
    lblMonth.ForeColor = vbWhite
    PassiveMonth = lblMonth.Name

End Sub

Private Sub lblYear_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate

    lblYearBackPosition lblYear
    yearMouseOut
    lblYear.ForeColor = vbWhite
    PassiveYear = lblYear.Name

End Sub

Private Sub lblSelectedDateBackPosition(ctrl As control)
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    '@INCLUDE DECLARATION Ctrl
    With dFrame.Controls("lblSelectedDateBack")
        .Visible = True
        .Top = ctrl.Top - 5
        .Left = ctrl.Left + 1
    End With
End Sub

Private Sub lblYearBackPosition(ctrl As control)
    '@AssignedModule clsDate
    '@INCLUDE PROCEDURE MouseMoveIcon
    '@INCLUDE CLASS clsDate
    '@INCLUDE DECLARATION Ctrl
    MouseMoveIcon
    dForm.Controls("lblYearBack").Top = ctrl.Top - 4
End Sub

Private Sub lblMonthBackPosition(ctrl As control)
    '@AssignedModule clsDate
    '@INCLUDE PROCEDURE MouseMoveIcon
    '@INCLUDE CLASS clsDate
    '@INCLUDE DECLARATION Ctrl
    MouseMoveIcon
    dForm.Controls("lblMonthBack").Top = ctrl.Top - 4
End Sub

Public Sub dayMouseOut()
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate
    On Error Resume Next
    If PassiveDay <> "" Then

        If Weekday(PassiveDayTag, 2) = 7 Or Weekday(PassiveDayTag, 2) = 6 Then
            dFrame.Controls(PassiveDay).ForeColor = &H2144FF
        Else
            dFrame.Controls(PassiveDay).ForeColor = &H48372C
        End If

    End If
    If SelectedDay <> "" Then
        dFrame.Controls(SelectedDay).ForeColor = vbWhite
    End If

End Sub

Private Sub monthMouseOut()
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate

    If PassiveMonth <> "" Then
        mFrame.Controls(PassiveMonth).ForeColor = vbGrayText
    End If
End Sub

Private Sub yearMouseOut()
    '@AssignedModule clsDate
    '@INCLUDE CLASS clsDate

    If PassiveYear <> "" Then
        yFrame.Controls(PassiveYear).ForeColor = vbGrayText
    End If
End Sub

