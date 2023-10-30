VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aColorScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aColorScheme
'* Purpose    :
'* Copyright  :
'*
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 05-07-2023 12:04    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private myTarget    As Object
Public WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Public NormalBackColor, NormalForeColor, HoverBackColor, HoverForeColor, SelectedColor
Public DefaultNormalBackColor, DefaultNormalForeColor

Dim cWhite          As Long
Dim cWhiteHover     As Long
Dim cBlueLight      As Long
Dim cBlueLightHover As Long
Dim cGreenLight     As Long
Dim cGreenLightHover As Long
Dim cPurpleLight    As Long
Dim cPurpleLightHover As Long
Dim cYellowLight    As Long
Dim cYellowLightHover As Long
Dim cRedLight       As Long
Dim cRedLightHover  As Long
Dim cOrangeLight    As Long
Dim cOrangeLightHover As Long
Dim cGrayLight      As Long
Dim cGrayLightHover As Long
Dim cBrownLight     As Long
Dim cBrownLightHover As Long
Dim cPinkLight      As Long
Dim cPinkLightHover As Long

Dim cBlack          As Long
Dim cBlackHover     As Long
Dim cBlueDark       As Long
Dim cBlueDarkHover  As Long
Dim cGreenDark      As Long
Dim cGreenDarkHover As Long
Dim cPurpleDark     As Long
Dim cPurpleDarkHover As Long
Dim cYellowDark     As Long
Dim cYellowDarkHover As Long
Dim cRedDark        As Long
Dim cRedDarkHover   As Long
Dim cOrangeDark     As Long
Dim cOrangeDarkHover As Long
Dim cGrayDark       As Long
Dim cGrayDarkHover  As Long
Dim cBrownDark      As Long
Dim cBrownDarkHover As Long
Dim cPinkDark       As Long
Dim cPinkDarkHover  As Long

Sub AssignColors()
    ' Light Themes
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    '@INCLUDE DECLARATION red
    cWhite = RGB(255, 255, 255)
    cWhiteHover = RGB(230, 230, 230)
    cBlueLight = RGB(20, 61, 89)    ' Darker blue color
    cBlueLightHover = RGB(61, 129, 168)
    cGreenLight = RGB(33, 138, 33)    ' Darker green color
    cGreenLightHover = RGB(78, 160, 78)
    cPurpleLight = RGB(85, 47, 113)    ' Darker purple color
    cPurpleLightHover = RGB(130, 84, 151)
    cYellowLight = RGB(153, 123, 0)    ' Darker yellow color
    cYellowLightHover = RGB(193, 156, 0)
    cRedLight = RGB(169, 0, 0)    ' Darker red color
    cRedLightHover = RGB(204, 0, 0)
    cOrangeLight = RGB(201, 87, 0)    ' Darker orange color
    cOrangeLightHover = RGB(234, 113, 0)
    cGrayLight = RGB(51, 51, 51)    ' Darker gray color
    cGrayLightHover = RGB(85, 85, 85)
    cBrownLight = RGB(102, 34, 34)    ' Darker brown color
    cBrownLightHover = RGB(128, 42, 42)
    cPinkLight = RGB(163, 67, 111)    ' Darker pink color
    cPinkLightHover = RGB(207, 112, 145)

    ' Dark Themes
    cBlack = RGB(0, 0, 0)
    cBlackHover = RGB(64, 64, 64)
    cBlueDark = RGB(0, 0, 139)
    cBlueDarkHover = RGB(0, 0, 205)
    cGreenDark = RGB(0, 100, 0)
    cGreenDarkHover = RGB(0, 128, 0)
    cPurpleDark = RGB(128, 0, 128)
    cPurpleDarkHover = RGB(139, 0, 139)
    cYellowDark = RGB(128, 128, 0)
    cYellowDarkHover = RGB(184, 134, 11)
    cRedDark = RGB(139, 0, 0)
    cRedDarkHover = RGB(178, 34, 34)
    cOrangeDark = RGB(255, 140, 0)
    cOrangeDarkHover = RGB(255, 69, 0)
    cGrayDark = RGB(105, 105, 105)
    cGrayDarkHover = RGB(128, 128, 128)
    cBrownDark = RGB(139, 69, 19)
    cBrownDarkHover = RGB(160, 82, 45)
    cPinkDark = RGB(255, 105, 180)
    cPinkDarkHover = RGB(255, 20, 147)
End Sub

Sub ThemeWhiteAndBlueLight()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cWhite, cBlueLight, cWhiteHover, cBlueLightHover)
End Sub

Sub ThemeWhiteAndGreenLight()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cWhite, cGreenLight, cWhiteHover, cGreenLightHover)
End Sub

Sub ThemeWhiteAndPurpleLight()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cWhite, cPurpleLight, cWhiteHover, cPurpleLightHover)
End Sub

Sub ThemeWhiteAndYellowLight()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cWhite, cYellowLight, cWhiteHover, cYellowLightHover)
End Sub

Sub ThemeWhiteAndRedLight()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cWhite, cRedLight, cWhiteHover, cRedLightHover)
End Sub

Sub ThemeWhiteAndOrangeLight()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cWhite, cOrangeLight, cWhiteHover, cOrangeLightHover)
End Sub

Sub ThemeWhiteAndGrayLight()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cWhite, cGrayLight, cWhiteHover, cGrayLightHover)
End Sub

Sub ThemeWhiteAndBrownLight()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cWhite, cBrownLight, cWhiteHover, cBrownLightHover)
End Sub

Sub ThemeWhiteAndPinkLight()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cWhite, cPinkLight, cWhiteHover, cPinkLightHover)
End Sub

Sub ThemeBlueAndGreenLight()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cBlueLight, cGreenLight, cBlueLightHover, cGreenLightHover)
End Sub

Sub ThemeBlackAndBlueDark()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cBlack, cBlueDark, cBlackHover, cBlueDarkHover)
End Sub

Sub ThemeBlackAndGreenDark()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cBlack, cGreenDark, cBlackHover, cGreenDarkHover)
End Sub

Sub ThemeBlackAndPurpleDark()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cBlack, cPurpleDark, cBlackHover, cPurpleDarkHover)
End Sub

Sub ThemeBlackAndYellowDark()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cBlack, cYellowDark, cBlackHover, cYellowDarkHover)
End Sub

Sub ThemeBlackAndRedDark()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cBlack, cRedDark, cBlackHover, cRedDarkHover)
End Sub

Sub ThemeBlackAndOrangeDark()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cBlack, cOrangeDark, cBlackHover, cOrangeDarkHover)
End Sub

Sub ThemeBlackAndGrayDark()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cBlack, cGrayDark, cBlackHover, cGrayDarkHover)
End Sub

Sub ThemeBlackAndBrownDark()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cBlack, cBrownDark, cBlackHover, cBrownDarkHover)
End Sub

Sub ThemeBlackAndPinkDark()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    Call color(cBlack, cPinkDark, cBlackHover, cPinkDarkHover)
End Sub


Public Sub color( _
        NormalBackColorValue, _
        NormalForeColorValue, _
        HoverBackColorValue, _
        HoverForeColorValue, _
        Optional SelectedColorValue)
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme

    NormalBackColor = NormalBackColorValue
    NormalForeColor = NormalForeColorValue
    HoverBackColor = HoverBackColorValue
    HoverForeColor = HoverForeColorValue
    SelectedColor = SelectedColorValue

    ApplyColorToParent myTarget
End Sub

Private Sub ApplyColorToParent(obj As Object)
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aMultiPage
    '@INCLUDE CLASS aColorScheme
    '@INCLUDE DECLARATION Ctrl
    On Error Resume Next
    obj.BackColor = NormalBackColor
    obj.ForeColor = NormalForeColor
    On Error GoTo 0
    Dim Ctrl        As msforms.control
    For Each Ctrl In obj.Controls
        Select Case TypeName(Ctrl)
            Case "CommandButton", "Label", "CheckBox", "OptionButton"
                Ctrl.BackColor = NormalBackColor
                Ctrl.ForeColor = NormalForeColor
            Case "Frame"
                Ctrl.BackColor = NormalBackColor
                Ctrl.ForeColor = NormalForeColor
                ApplyColorToParent Ctrl
            Case "MultiPage"
                Dim i As Long
                With aMultiPage.Init(Ctrl)
                    For i = 0 To Ctrl.Pages.Count - 1
                        .SetBackColor i, CLng(NormalBackColor)
                    Next
                End With
            Case Else

        End Select
    Next
End Sub

Private Sub SetEmitter()
    '@AssignedModule aColorScheme
    '@INCLUDE PROCEDURE isUserform
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aColorScheme
    '@INCLUDE DECLARATION Emitter
    Set Emitter = Nothing
    Set Emitter = New EventListenerEmitter
    Select Case TypeName(myTarget)
        Case "CommandButton", "Label", "OptionButton", "CheckBox"    '@TODO add controls to emitter class, this isn't working for all
            Emitter.AddEventListener myTarget
        Case "MultiPage", "Frame"
            Emitter.AddEventListenerAll myTarget
        Case Else
            If isUserform(myTarget) Then
                Emitter.AddEventListenerAll myTarget
            End If
    End Select

End Sub

Private Sub resetColors()
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS aColorScheme
    NormalBackColor = 0
    NormalForeColor = 0
    HoverBackColor = 0
    HoverForeColor = 0
    SelectedColor = 0
End Sub

Public Function Init(FormFrameOrControl As Object) As aColorScheme
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aColorScheme
    resetColors
    Set myTarget = FormFrameOrControl
    AssignColors
    SetEmitter
    Set Init = Me
End Function

Private Sub Emitter_Click(control As Object)
    '@AssignedModule aColorScheme
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aColorScheme
    If control.Caption Like "Theme*" Then
        CallByName Me, control.Caption, VbMethod
    End If
End Sub

Private Sub Emitter_MouseOut(control As Object)
    '@AssignedModule aColorScheme
    '@INCLUDE PROCEDURE isUserform
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aColorScheme
    If NormalBackColor + NormalForeColor + HoverBackColor + HoverForeColor = 0 Then Exit Sub
    If isUserform(control) Then Exit Sub
    Select Case TypeName(control)
        Case "MultiPage", "Frame": Exit Sub
    End Select
    control.BackColor = NormalBackColor
    control.ForeColor = NormalForeColor
End Sub

Private Sub Emitter_MouseOver(control As Object)
    '@AssignedModule aColorScheme
    '@INCLUDE PROCEDURE isUserform
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aColorScheme
    If NormalBackColor + NormalForeColor + HoverBackColor + HoverForeColor = 0 Then Exit Sub
    If isUserform(control) Then Exit Sub
    Select Case TypeName(control)
        Case "MultiPage", "Frame": Exit Sub
    End Select
    control.BackColor = HoverBackColor
    control.ForeColor = HoverForeColor
End Sub

