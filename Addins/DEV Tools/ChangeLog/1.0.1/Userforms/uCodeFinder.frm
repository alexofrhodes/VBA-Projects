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

Dim AT              As aTreeView

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

        .Transition .Effect(GetInfo, "Top", GetInfo.Top + 10, 150)
    End With
    uAuthor.Show
End Sub

Private Sub Label2_Click()
    Me.TextBox1.Value = Replace(uCalendar.Datepicker, ".", "-")
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
    aUserform.Init(Me).LoadPosition
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
        '        .CollapseAll
    End With
    Set CalledFromModule = ActiveModule
    CalledFromProcedure = ActiveProcedure
    TextBox1.SetFocus

End Sub

Private Sub ReturnToCaller()
    '@AssignedModule uCodeFinder
    '@INCLUDE CLASS aModule
    '@INCLUDE USERFORM uCodeFinder
    '@INCLUDE DECLARATION CalledFromModule
    '@INCLUDE DECLARATION CalledFromProcedure
    On Error GoTo HELL
    aModule.Init(CalledFromModule).Activate
    Dim i           As Long
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
    DoSearch
End Sub

Public Sub DoSearch()
    '@LastModified 2308180711
    '@AssignedModule uCodeFinder
    '@INCLUDE CLASS aTreeView
    '@INCLUDE USERFORM uCodeFinder
    '@INCLUDE DECLARATION AT
    Dim tvtop As Long, tvleft As Long

    Me.Hide
    AT.clear
    AT.FindCodeEverywhere TextBox1.TEXT
    AT.ExpandAll
    If TreeView1.Nodes.Count > 0 Then
        TreeView1.Nodes(1).Expanded = True
        AT.TreeviewAssignProjectImages
    End If
    Me.Show
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '@AssignedModule uCodeFinder
    '@INCLUDE USERFORM uCodeFinder
    If KeyCode = 13 Then
        DoSearch
    End If
End Sub

Private Sub TreeView1_Click()
    '@LastModified 2308181007
    '@AssignedModule uCodeFinder
    '@INCLUDE CLASS aTreeView
    '@INCLUDE USERFORM uCodeFinder
    '@INCLUDE DECLARATION AT
    If TreeView1.Nodes.Count > 0 Then AT.ActivateProjectElement
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    aUserform.Init(Me).SavePosition
End Sub
