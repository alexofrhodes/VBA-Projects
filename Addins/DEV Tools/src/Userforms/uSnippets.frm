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
Dim moResizer As New CFormResizer

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim uSnipIndex As Long:     uSnipIndex = -1
    Dim uSnipFilter As String:  uSnipFilter = ""
    'TODO
'    ThisWorkbook.Sheets("uSnippets").Range("B1") = uSnipFilter
'    ThisWorkbook.Sheets("uSnippets").Range("B2") = uSnipIndex
    Unload Me
End Sub

Private Sub UserForm_Resize()
    moResizer.FormResize
End Sub

Private Sub UserForm_Initialize()
    If ShowInVBE = True Then
        Application.VBE.MainWindow.Visible = True
        aUserform.Init(Me).ParentIsVBE
    End If
    Set moResizer.Form = Me
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

End Sub

Sub SwitchParent()
    'Debug.Print Application.r
    Stop
End Sub

Sub GetFilesUSnippets()
    LSnippets.Clear
    Dim Files As Collection: Set Files = LoopThroughFiles(SnippetsFolder, "*.txt")
    Dim File
    For Each File In Files
        LSnippets.AddItem File
    Next
End Sub

Private Sub CommandButton1_Click()
    tFilterSnippets.TEXT = ""
    LSnippets.ListIndex = -1
End Sub

Private Sub cResize_Click()
    If Me.Height < 429 Then
        Me.Height = 429
    Else
        Me.Height = 60
        Me.Width = 100
    End If

    Me.Show
End Sub

Private Sub cSnippetFolder_Click()
    FollowLink SnippetsFolder
End Sub

Private Sub GetInfo_Click()
    uAuthor.Show
End Sub

Private Sub LSnippets_Click()
    Dim sPath As String
    sPath = SnippetsFolder & LSnippets.list(LSnippets.ListIndex)
    LSnippetsPreview.TEXT = TxtRead(sPath)
End Sub

Private Sub cCopySnippet_Click()
    If Len(LSnippetsPreview.TEXT) = 0 Then Exit Sub
    Dim s As String
    If LSnippetsPreview.SelLength = 0 Then
        s = LSnippetsPreview.TEXT
    Else
        s = LSnippetsPreview.SelText
    End If
    CLIP s
    cResize_Click
    MsgBox "Snipet copied"
End Sub

Private Sub cOverwriteSnippet_Click()
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
        sPath = SnippetsFolder & LSnippets.list(LSnippets.ListIndex)
    End If
    TxtOverwrite sPath, LSnippetsPreview.TEXT
    If isNew = True Then
        LSnippets.AddItem ans & ".txt"
        LSnippets.ListIndex = LSnippets.ListCount - 1
    End If
ExitHandler:
    If wasResized = True Then cResize_Click
End Sub

Private Sub cSnippetDelete_Click()
    cResize_Click
    Dim Proceed As Long
    Proceed = MsgBox("Delete " & LSnippets.list(LSnippets.ListIndex) & "?", vbYesNo)
    If Proceed = vbNo Then Exit Sub
    Dim sPath As String
    sPath = SnippetsFolder & LSnippets.list(LSnippets.ListIndex)
    Dim fso As scripting.FileSystemObject
    Set fso = New scripting.FileSystemObject
    fso.DeleteFile sPath
    LSnippets.RemoveItem LSnippets.ListIndex
    LSnippetsPreview.TEXT = ""
    LSnippets.ListIndex = -1
    cResize_Click
End Sub

Private Sub cSnippetStartNew_Click()
    Dim NewName As String
    cResize_Click
    NewName = InputBox("New Snippet Name")
    If NewName = "" Then GoTo ExitHandler
    Dim sPath As String
    sPath = SnippetsFolder & NewName & ".txt"
    If FileExists(sPath) Then Exit Sub
    LSnippets.ListIndex = -1
    LSnippetsPreview.TEXT = ""
    TxtOverwrite sPath, ""
    LSnippets.AddItem NewName & ".txt"
    LSnippets.ListIndex = LSnippets.ListCount - 1
    LSnippetsPreview.SetFocus
ExitHandler:
    cResize_Click
End Sub

Private Sub LSnippetsPreview_Enter()
    LSnippetsPreview.SelStart = 0
End Sub

Private Sub tFilterSnippets_Change()
    GetFilesUSnippets
    aListBox.Init(LSnippets).FilterByColumn tFilterSnippets.TEXT
End Sub

