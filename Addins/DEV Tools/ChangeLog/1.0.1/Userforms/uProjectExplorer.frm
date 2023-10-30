VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uProjectExplorer 
   Caption         =   "Project Explorer"
   ClientHeight    =   9732.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3636
   OleObjectBlob   =   "uProjectExplorer.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uProjectExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Userform   : uProjectExplorer
'* Created    : 06-10-2022 10:39
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Option Explicit


Private Sub GetInfo_Click()
    '@AssignedModule uProjectExplorer
    '@INCLUDE PROCEDURE PlayTheSound
    '@INCLUDE CLASS aUserform
    '@INCLUDE USERFORM uProjectExplorer
    '@INCLUDE USERFORM uAuthor
    With aUserform.Init(Me)
        .Transition .Effect(GetInfo, "Top", GetInfo.Top - 10, 150)

        .Transition .Effect(GetInfo, "Top", GetInfo.Top + 10, 150)
    End With
    uAuthor.Show
End Sub



Private Sub TreeView1_Click()
    '@AssignedModule uProjectExplorer
    '@INCLUDE CLASS aTreeView
    '@INCLUDE USERFORM uProjectExplorer
    aTreeView.Init(TreeView1).ActivateProjectElement
End Sub

Private Sub UserForm_Initialize()
    '@AssignedModule uProjectExplorer
    '@INCLUDE USERFORM uProjectExplorer
    InitializeProjectExplorer
End Sub

Sub InitializeProjectExplorer()
    '@AssignedModule uProjectExplorer
    '@INCLUDE PROCEDURE ActiveModule
    '@INCLUDE PROCEDURE ActiveCodepaneWorkbook
    '@INCLUDE CLASS aUserform
    '@INCLUDE CLASS aTreeView
    '@INCLUDE USERFORM uProjectExplorer
    Application.VBE.MainWindow.Visible = True
    aUserform.Init(Me).ParentIsVBE
    aTreeView.Init(TreeView1).LoadVBProjects
    With TreeView1
        .Sorted = True
        .Appearance = ccFlat
        .LineStyle = tvwRootLines
        .Style = tvwTreelinesPlusMinusPictureText
        .Font.Size = 10
        .indentation = 2
    End With
    aTreeView.Init(TreeView1).ImageListLoadProjectIcons ImageList1
    Dim TargetWorkbook As Workbook
    If Application.VBE.MainWindow.Visible = False Then
        Set TargetWorkbook = ActiveWorkbook
    Else
        Set TargetWorkbook = ActiveCodepaneWorkbook
    End If
    aTreeView.Init(TreeView1).SelectNodes True, TargetWorkbook.Name, Array(ActiveModule.Name)
End Sub

