VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uCodeFinder 
   Caption         =   "Code Finder"
   ClientHeight    =   8964.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4308
   OleObjectBlob   =   "uCodeFinder.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uCodeFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Userform   : uCodeFinder
'* Created    : 06-10-2022 10:34
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Dim AT As aTreeView

Dim CalledFromModule As VBComponent
Dim CalledFromProcedure As String

Private Sub CommandButton2_Click()
'@AssignedModule uCodeFinder
'@INCLUDE USERFORM uCodeFinder
    ReturnToCaller
End Sub

Private Sub GetInfo_Click()
'@AssignedModule uCodeFinder
'@INCLUDE PROCEDURE PlayTheSound
'@INCLUDE CLASS aUserform
'@INCLUDE USERFORM uCodeFinder
'@INCLUDE USERFORM uAuthor
    With aUserform.Init(Me)
        .Transition .Effect(GetInfo, "Top", GetInfo.Top - 10, 150)
        PlayTheSound ThisWorkbook.Path & "\sound\coin.wav"
        .Transition .Effect(GetInfo, "Top", GetInfo.Top + 10, 150)
    End With
    uAuthor.Show
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
'@AssignedModule uCodeFinder
'@INCLUDE USERFORM uCodeFinder
    Cancel = True
End Sub

Private Sub UserForm_Activate()
'@AssignedModule uCodeFinder
'@INCLUDE CLASS aUserform
'@INCLUDE USERFORM uCodeFinder
    With aUserform.Init(Me)
        .Resizable
        .ParentIsVBE
    End With

End Sub

Private Sub UserForm_Initialize()
'@AssignedModule uCodeFinder
'@INCLUDE PROCEDURE ActiveProcedure
'@INCLUDE PROCEDURE ActiveModule
'@INCLUDE CLASS aTreeView
'@INCLUDE USERFORM uCodeFinder
'@INCLUDE DECLARATION AT
'@INCLUDE DECLARATION CalledFromModule
'@INCLUDE DECLARATION CalledFromProcedure
    
    With TreeView1
        .Sorted = True
        .Appearance = ccFlat
        .LineStyle = tvwRootLines
        .Style = tvwTreelinesPlusMinusPictureText
        .Font.Size = 10
        .indentation = 2
    End With
    Set AT = aTreeView.Init(TreeView1)
    With AT
        .ImageListLoadProjectIcons ImageList1
        .CollapseAll
    End With
    Set CalledFromModule = ActiveModule
    CalledFromProcedure = ActiveProcedure

End Sub

Private Sub ReturnToCaller()
'@AssignedModule uCodeFinder
'@INCLUDE CLASS aModule
'@INCLUDE USERFORM uCodeFinder
'@INCLUDE DECLARATION CalledFromModule
'@INCLUDE DECLARATION CalledFromProcedure
    On Error GoTo HELL
    aModule.Init(CalledFromModule).Activate
    Dim i As Long
    For i = 1 To CalledFromModule.CodeModule.CountOfLines
        If InStr(1, CalledFromModule.CodeModule.Lines(i, 1), "Sub " & CalledFromProcedure) > 0 Or _
            InStr(1, CalledFromModule.CodeModule.Lines(i, 1), "Function " & CalledFromProcedure) > 0 Then
            CalledFromModule.CodeModule.CodePane.SetSelection i, 1, i, 1
            Exit Sub
        End If
    Next
HELL:
End Sub

Private Sub CommandButton1_Click()
'@AssignedModule uCodeFinder
'@INCLUDE CLASS aTreeView
'@INCLUDE USERFORM uCodeFinder
'@INCLUDE DECLARATION AT
    Dim tvtop As Long, tvleft As Long

    'TreeView1.Visible = False
    With AT
        .Clear
        .FindCodeEverywhere TextBox1.text
        .TreeviewAssignProjectImages
        .ExpandAll
    End With
    'TreeView1.Visible = True
    TreeView1.Nodes(1).Expanded = True
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'@AssignedModule uCodeFinder
'@INCLUDE USERFORM uCodeFinder
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub TreeView1_Click()
'@AssignedModule uCodeFinder
'@INCLUDE CLASS aTreeView
'@INCLUDE USERFORM uCodeFinder
'@INCLUDE DECLARATION AT
    If TreeView1.Nodes.count > 0 Then AT.ActivateProjectElement
End Sub

