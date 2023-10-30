VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uCodeOnTheFly 
   Caption         =   "Immediate Windows"
   ClientHeight    =   7332
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12900
   OleObjectBlob   =   "uCodeOnTheFly.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uCodeOnTheFly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm   : uCodeOnTheFly
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
'* Created    : 30-06-2023 11:53    Alex
'* Modified   : 18-07-2023 08:14    Alex                modified from frame method to aMultipage class
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private CodeOnTheFlyTextbox As MSForms.Textbox

Private Sub GetInfo_Click()
    '@AssignedModule uCodeOnTheFly
    '@INCLUDE PROCEDURE PlayTheSound
    '@INCLUDE CLASS aUserform
    '@INCLUDE USERFORM uCodeOnTheFly
    '@INCLUDE USERFORM uAuthor
    With aUserform.Init(Me)
        .Transition .Effect(GetInfo, "Top", GetInfo.Top - 10, 150)

        .Transition .Effect(GetInfo, "Top", GetInfo.Top + 10, 150)
    End With
    uAuthor.Show
End Sub

Private Sub labelRun_Click()
    '@AssignedModule uCodeOnTheFly
    '@INCLUDE PROCEDURE ShowUserformCodeOnTheFly
    '@INCLUDE PROCEDURE EvaluateQuestion
    '@INCLUDE PROCEDURE RunCodeOnTheFly
    '@INCLUDE PROCEDURE TextOfControl
    '@INCLUDE PROCEDURE appRunOnTime
    '@INCLUDE PROCEDURE PlayTheSound
    '@INCLUDE CLASS aUserform
    '@INCLUDE USERFORM uCodeOnTheFly
    '@INCLUDE DECLARATION CodeOnTheFlyTextbox
    aUserform.Init(Me).SaveOptions
    Dim Code        As String
    Code = TextOfControl(CodeOnTheFlyTextbox)
    If Left(Code, 1) <> "?" Then
        'if pointer width=0 then code=all text, else code=selection
        appRunOnTime Now, "ShowUserformCodeOnTheFly"
        RunCodeOnTheFly Code
    Else
        Dim qAsked  As Long
        qAsked = Len(Code) - Len(Replace(Code, "?", ""))
        If qAsked > 1 Then
            MsgBox qAsked & " questions detected. I can only answer one."
            Exit Sub
        End If
        
        appRunOnTime Now, "ShowUserformCodeOnTheFly"
        appRunOnTime Now, _
                "RunCodeOnTheFly", _
                Me.Name & "." & CodeOnTheFlyTextbox.Name & ".text=" & _
                Me.Name & "." & CodeOnTheFlyTextbox.Name & ".text & vbNewLine & "" & ""  " & Mid(Code, 2)
        '        EvaluateQuestion Mid(code, 2)
    End If
End Sub


Private Sub MultiPage1_Change()
    '@AssignedModule uCodeOnTheFly
    '@INCLUDE USERFORM uCodeOnTheFly
    '@INCLUDE DECLARATION CodeOnTheFlyTextbox
    Set CodeOnTheFlyTextbox = Me.Controls("Textbox" & MultiPage1.Value + 1)
End Sub


Private Sub UserForm_Activate()
    '@AssignedModule uCodeOnTheFly
    '@INCLUDE USERFORM uCodeOnTheFly
    With labelRun
        .BorderStyle = fmBorderStyleNone
        .Picture = LoadPicture(ThisWorkbook.path & "\Lib\img\MagicHat.bmp")
        .MouseIcon = LoadPicture(ThisWorkbook.path & "\Lib\img\wand.ico")
        .MousePointer = fmMousePointerCustom
    End With
End Sub

Private Sub UserForm_Initialize()
    '@AssignedModule uCodeOnTheFly
    '@INCLUDE CLASS aUserform
    '@INCLUDE CLASS aMultiPage
    '@INCLUDE USERFORM uCodeOnTheFly
    '@INCLUDE DECLARATION MyColors
    '@INCLUDE DECLARATION CodeOnTheFlyTextbox

    Me.Caption = "if pointer width=0 then code=all text, else code=selection"
    Dim i           As Long
    For i = 1 To 9
        With Me.Controls("Textbox" & i)
            .Width = 1500
            .Height = 1500
            .Left = 0
            .Top = 0
            .Font.Size = 10
            .Font.Name = "Consolas"
        End With
    Next

    aUserform.Init(Me).ShowMinimizeButton

    Me.Height = 200
    Me.Width = 430


    aUserform.Init(Me).LoadOptions
    Set CodeOnTheFlyTextbox = Me.Controls("Textbox" & MultiPage1.Value + 1)
    aMultiPage.Init(MultiPage1).BuildMenu createSidebarMinimizers:=False
    Dim ctl         As MSForms.control
    For Each ctl In Me.Controls("sidebarleft").Controls
        If ctl.Name = "sidebarLabel" & MultiPage1.SelectedItem.Name Then
            ctl.BackColor = MyColors.FormSelectedGreen
        Else
            ctl.BackColor = Me.Controls("sidebarleft").BackColor
        End If
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '@AssignedModule uCodeOnTheFly
    '@INCLUDE CLASS aUserform
    '@INCLUDE USERFORM uCodeOnTheFly
    aUserform.Init(Me).SaveOptions
End Sub

