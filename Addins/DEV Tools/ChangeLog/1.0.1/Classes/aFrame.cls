VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aFrame
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
'* Created    : 05-07-2023 07:26    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private oFrame      As MSForms.Frame

Public Function Init(TargetFrame As Object) As aFrame
    '@AssignedModule aFrame
    '@INCLUDE CLASS aFrame
    Set oFrame = TargetFrame
    Set Init = Me
End Function

Public Sub ResizeToFitControls(Optional marginRight = 0, Optional marginBottom = 0)
    '@AssignedModule aFrame
    '@INCLUDE CLASS aFrame
    oFrame.Width = 0
    oFrame.Height = 0
    Dim ctr         As MSForms.control
    Dim myWidth: myWidth = 0
    Dim myHeight: myHeight = 0
    For Each ctr In oFrame.Controls
        If ctr.Parent.Name = oFrame.Name Then
            If ctr.Visible = True Then
                If ctr.Left + ctr.Width > myWidth Then myWidth = ctr.Left + ctr.Width
                If ctr.Top + ctr.Height > myHeight Then myHeight = ctr.Top + ctr.Height
            End If
        End If
    Next
    oFrame.Width = myWidth + (oFrame.Width - oFrame.InsideWidth) + marginRight
    oFrame.Height = myHeight + (oFrame.Height - oFrame.InsideHeight) + marginBottom
End Sub

Public Sub AddThemeControls()
    '@AssignedModule aFrame
    '@INCLUDE PROCEDURE aSwitch
    '@INCLUDE CLASS aFrame
    '@INCLUDE CLASS aColorScheme
    '@INCLUDE DECLARATION Ctrl
    Dim ctrl        As MSForms.control
    Dim i           As Long
    For i = 1 To 19
        Set ctrl = oFrame.Controls.Add("Forms.Label.1")
        ctrl.Visible = False
        ctrl.Top = i * ctrl.Height
        ctrl.Left = 12
        ctrl.Height = 20
        ctrl.Width = 120
        ctrl.Font.Size = 9
        ctrl.Font.Name = "Consolas"
        ctrl.WordWrap = False
        ctrl.Caption = aSwitch(i, 1, "ThemeBlackAndBlueDark", 2, "ThemeBlackAndBrownDark", 3, "ThemeBlackAndGrayDark", 4, "ThemeBlackAndGreenDark", 5, "ThemeBlackAndOrangeDark", 6, "ThemeBlackAndPinkDark", 7, "ThemeBlackAndPurpleDark", 8, "ThemeBlackAndRedDark", 9, "ThemeBlackAndYellowDark", 10, "ThemeBlueAndGreenLight", 11, "ThemeWhiteAndBlueLight", 12, "ThemeWhiteAndBrownLight", 13, "ThemeWhiteAndGrayLight", 14, "ThemeWhiteAndGreenLight", 15, "ThemeWhiteAndOrangeLight", 16, "ThemeWhiteAndPinkLight", 17, "ThemeWhiteAndPurpleLight", 18, "ThemeWhiteAndRedLight", 19, "ThemeWhiteAndYellowLight")
        ctrl.Visible = True
    Next
    '
    '    On Error Resume Next
    '        For Each ctrl In oframe.parent.Controls
    '            ctrl.Font.Name = "Consolas"
    '            ctrl.Font.Size = 9
    '            ctrl.Font.Bold = True
    '        Next
    '    On Error GoTo 0

    aColorScheme.Init oFrame.Parent
End Sub

