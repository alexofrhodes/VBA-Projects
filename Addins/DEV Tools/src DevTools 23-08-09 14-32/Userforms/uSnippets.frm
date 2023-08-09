VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uSnippets 
   Caption         =   "SnippetsManager"
   ClientHeight    =   8016
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7968
   OleObjectBlob   =   "uSnippets.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uSnippets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Userform   : uSnippets
'* Created    : 06-10-2022 10:41
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Option Explicit

Private SnippetsFolder As String

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'@AssignedModule uSnippets
'@INCLUDE USERFORM uSnippets
    Dim uSnipIndex As Long:     uSnipIndex = -1
    Dim uSnipFilter As String:  uSnipFilter = ""
    'TODO
'    ThisWorkbook.Sheets("uSnippets").Range("B1") = uSnipFilter
'    ThisWorkbook.Sheets("uSnippets").Range("B2") = uSnipIndex
    Unload Me
End Sub

Private Sub UserForm_Initialize()
'@AssignedModule uSnippets
'@INCLUDE PROCEDURE FoldersCreate
'@INCLUDE CLASS aListBox
'@INCLUDE CLASS aUserform
'@INCLUDE USERFORM uSnippets
'@INCLUDE DECLARATION ShowInVBE
'@INCLUDE DECLARATION SnippetsFolder
    If ShowInVBE = True Then
        Application.VBE.MainWindow.Visible = True
        aUserform.Init(Me).ParentIsVBE
    End If
    SnippetsFolder = LOCAL_LIBRARY_PROCEDURES

    FoldersCreate SnippetsFolder

    If Right(SnippetsFolder, 1) <> "\" Then SnippetsFolder = SnippetsFolder & "\"
    GetFilesUSnippets

'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets("uSnippets")
'    Dim uSnipFilter As String
'    uSnipFilter = ws.Range("B1")
'    tFilterSnippets.TEXT = uSnipFilter
'    Dim uSnipIndex As String
'    uSnipIndex = ws.Range("B2").TEXT
'    aListBox.Init(LSnippets).SelectItems uSnipIndex
    Dim myForm As New aUserform
    myForm.Init(Me).Resizable
End Sub

Sub SwitchParent()
    'Debug.Print Application.r
'@AssignedModule uSnippets
'@INCLUDE USERFORM uSnippets
    Stop
End Sub

Sub GetFilesUSnippets()
'@AssignedModule uSnippets
'@INCLUDE PROCEDURE LoopThroughFiles
'@INCLUDE USERFORM uSnippets
'@INCLUDE DECLARATION SnippetsFolder
    LSnippets.Clear
    Dim Files As Collection: Set Files = LoopThroughFiles(SnippetsFolder, "*.txt")
    Dim File
    For Each File In Files
        LSnippets.AddItem File
    Next
End Sub

Private Sub CommandButton1_Click()
'@AssignedModule uSnippets
'@INCLUDE USERFORM uSnippets
    tFilterSnippets.text = ""
    LSnippets.ListIndex = -1
End Sub

Private Sub cResize_Click()
'@AssignedModule uSnippets
'@INCLUDE USERFORM uSnippets
    If Me.Height < 429 Then
        Me.Height = 429
    Else
        Me.Height = 60
        Me.Width = 100
    End If

    Me.Show
End Sub

Private Sub cSnippetFolder_Click()
'@AssignedModule uSnippets
'@INCLUDE PROCEDURE FollowLink
'@INCLUDE USERFORM uSnippets
'@INCLUDE DECLARATION SnippetsFolder
    FollowLink SnippetsFolder
End Sub

Private Sub GetInfo_Click()
'@AssignedModule uSnippets
'@INCLUDE PROCEDURE PlayTheSound
'@INCLUDE CLASS aUserform
'@INCLUDE USERFORM uSnippets
'@INCLUDE USERFORM uAuthor
    With aUserform.Init(Me)
        .Transition .Effect(GetInfo, "Top", GetInfo.Top - 10, 150)
        PlayTheSound ThisWorkbook.Path & "\sound\coin.wav"
        .Transition .Effect(GetInfo, "Top", GetInfo.Top + 10, 150)
    End With
    uAuthor.Show
End Sub


Private Sub LSnippets_Click()
'@AssignedModule uSnippets
'@INCLUDE PROCEDURE TxtRead
'@INCLUDE USERFORM uSnippets
'@INCLUDE DECLARATION SnippetsFolder
    Dim sPath As String
    sPath = SnippetsFolder & LSnippets.List(LSnippets.ListIndex)
    LSnippetsPreview.text = TxtRead(sPath)
End Sub

Private Sub cCopySnippet_Click()
'@AssignedModule uSnippets
'@INCLUDE PROCEDURE CLIP
'@INCLUDE USERFORM uSnippets
    If Len(LSnippetsPreview.text) = 0 Then Exit Sub
    Dim s As String
    If LSnippetsPreview.SelLength = 0 Then
        s = LSnippetsPreview.text
    Else
        s = LSnippetsPreview.SelText
    End If
    CLIP s
    cResize_Click
    MsgBox "Snipet copied"
End Sub

Private Sub cOverwriteSnippet_Click()
'@AssignedModule uSnippets
'@INCLUDE PROCEDURE TxtOverwrite
'@INCLUDE PROCEDURE InputboxString
'@INCLUDE USERFORM uSnippets
'@INCLUDE DECLARATION SnippetsFolder
    Dim sPath As String
    Dim isNew As Boolean
    Dim wasResized As Boolean
    If LSnippets.ListIndex = -1 Then
        Dim ans As String
        cResize_Click
        ans = InputboxString(, "Select name for new file")
        If Len(ans) = 0 Then GoTo ExitHandler
        sPath = SnippetsFolder & ans & ".txt"
        isNew = True
        wasResized = True
    Else
        sPath = SnippetsFolder & LSnippets.List(LSnippets.ListIndex)
    End If
    TxtOverwrite sPath, LSnippetsPreview.text
    If isNew = True Then
        LSnippets.AddItem ans & ".txt"
        LSnippets.ListIndex = LSnippets.ListCount - 1
    End If
ExitHandler:
    If wasResized = True Then cResize_Click
End Sub

Private Sub cSnippetDelete_Click()
'@AssignedModule uSnippets
'@INCLUDE USERFORM uSnippets
'@INCLUDE DECLARATION SnippetsFolder
    cResize_Click
    Dim Proceed As Long
    Proceed = MsgBox("Delete " & LSnippets.List(LSnippets.ListIndex) & "?", vbYesNo)
    If Proceed = vbNo Then Exit Sub
    Dim sPath As String
    sPath = SnippetsFolder & LSnippets.List(LSnippets.ListIndex)
    Dim FSO As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject
    FSO.DeleteFile sPath
    LSnippets.RemoveItem LSnippets.ListIndex
    LSnippetsPreview.text = ""
    LSnippets.ListIndex = -1
    cResize_Click
End Sub

Private Sub cSnippetStartNew_Click()
'@AssignedModule uSnippets
'@INCLUDE PROCEDURE FileExists
'@INCLUDE PROCEDURE TxtOverwrite
'@INCLUDE USERFORM uSnippets
'@INCLUDE DECLARATION SnippetsFolder
    Dim NewName As String
    cResize_Click
    NewName = InputBox("New Snippet Name")
    If NewName = "" Then GoTo ExitHandler
    Dim sPath As String
    sPath = SnippetsFolder & NewName & ".txt"
    If FileExists(sPath) Then Exit Sub
    LSnippets.ListIndex = -1
    LSnippetsPreview.text = ""
    TxtOverwrite sPath, ""
    LSnippets.AddItem NewName & ".txt"
    LSnippets.ListIndex = LSnippets.ListCount - 1
    LSnippetsPreview.SetFocus
ExitHandler:
    cResize_Click
End Sub

Private Sub LSnippetsPreview_Enter()
'@AssignedModule uSnippets
'@INCLUDE USERFORM uSnippets
    LSnippetsPreview.SelStart = 0
End Sub

Private Sub tFilterSnippets_Change()
'@AssignedModule uSnippets
'@INCLUDE CLASS aListBox
'@INCLUDE USERFORM uSnippets
    GetFilesUSnippets
    aListBox.Init(LSnippets).FilterByColumn tFilterSnippets.text
End Sub

