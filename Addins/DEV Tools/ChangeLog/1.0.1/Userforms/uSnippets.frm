VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uSnippets 
   Caption         =   "SnippetsManager"
   ClientHeight    =   8052
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7980
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
'* UserForm   : uSnippets
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
'* Created    : 17-08-2023 13:28    Alex
'* Modified   : 17-08-2023 13:28    Alex                added roundabout way to inject directly
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private SnippetsFolder As String


Private Sub SNIP_Click()
    If Len(LSnippetsPreview.TEXT) = 0 Then Exit Sub
    Dim S           As String
    If LSnippetsPreview.SelLength = 0 Then
        S = LSnippetsPreview.TEXT
    Else
        S = LSnippetsPreview.SelText
    End If

    SaveSettings

    If ShowInVBE Then
        Application.OnTime Now, "uShow_SnippetsVBE"
    Else
        Application.OnTime Now, "uShow_SnippetsWorkbook"
    End If

    aCodeModule.Active.Inject S
    '    cResize_Click
End Sub

Sub SaveSettings()
    IniWrite ThisWorkbook.path & "config\UserformSettings.ini", Me.Name, "SelStart", LSnippetsPreview.selStart
    IniWrite ThisWorkbook.path & "config\UserformSettings.ini", Me.Name, "SelLength", LSnippetsPreview.SelLength
    IniWrite ThisWorkbook.path & "config\UserformSettings.ini", Me.Name, "Filter", Me.tFilterSnippets.Value
    If LSnippets.ListIndex = -1 Then
        IniWrite ThisWorkbook.path & "config\UserformSettings.ini", Me.Name, "File", ""
    Else
        IniWrite ThisWorkbook.path & "config\UserformSettings.ini", Me.Name, "File", LSnippets.list(LSnippets.ListIndex)
    End If
End Sub

Private Sub UserForm_Activate()
    With aUserform.Init(Me)
        .Resizable True
        If ShowInVBE Then .ParentIsVBE
    End With
End Sub

Private Sub UserForm_Initialize()
    '@AssignedModule uSnippets
    '@INCLUDE PROCEDURE FoldersCreate
    '@INCLUDE CLASS aListBox
    '@INCLUDE CLASS aUserform
    '@INCLUDE USERFORM uSnippets
    '@INCLUDE DECLARATION ShowInVBE
    '@INCLUDE DECLARATION SnippetsFolder

    SnippetsFolder = LOCAL_LIBRARY_PROCEDURES

    FoldersCreate SnippetsFolder

    If Right(SnippetsFolder, 1) <> "\" Then SnippetsFolder = SnippetsFolder & "\"
    GetFilesUSnippets

    tFilterSnippets.TEXT = IniReadKey(ThisWorkbook.path & "config\UserformSettings.ini", Me.Name, "Filter")
    Dim fileName    As String
    fileName = IniReadKey(ThisWorkbook.path & "config\UserformSettings.ini", Me.Name, "File")
    Dim i           As Long
    For i = LBound(LSnippets.list) To UBound(LSnippets.list)
        If LSnippets.list(i) = fileName Then
            LSnippets.ListIndex = i
            Exit For
        End If
    Next

    LSnippetsPreview.selStart = IniReadKey(ThisWorkbook.path & "config\UserformSettings.ini", Me.Name, "SelStart", 0)
    LSnippetsPreview.SelLength = IniReadKey(ThisWorkbook.path & "config\UserformSettings.ini", Me.Name, "SelLength", 0)
    
End Sub

Sub SwitchParent()
    '@AssignedModule uSnippets
    '@INCLUDE USERFORM uSnippets
    '@todo

End Sub

Sub GetFilesUSnippets()
    '@AssignedModule uSnippets
    '@INCLUDE PROCEDURE LoopThroughFiles
    '@INCLUDE USERFORM uSnippets
    '@INCLUDE DECLARATION SnippetsFolder
    LSnippets.clear
    Dim Files       As Collection: Set Files = LoopThroughFiles(SnippetsFolder, "*.txt")
    Dim File
    For Each File In Files
        LSnippets.AddItem File
    Next
End Sub

Private Sub CommandButton1_Click()
    '@AssignedModule uSnippets
    '@INCLUDE USERFORM uSnippets
    tFilterSnippets.TEXT = ""
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

        .Transition .Effect(GetInfo, "Top", GetInfo.Top + 10, 150)
    End With
    uAuthor.Show
End Sub


Private Sub LSnippets_Click()
    '@AssignedModule uSnippets
    '@INCLUDE PROCEDURE TxtRead
    '@INCLUDE USERFORM uSnippets
    '@INCLUDE DECLARATION SnippetsFolder
    Dim sPath       As String
    sPath = SnippetsFolder & LSnippets.list(LSnippets.ListIndex)
    LSnippetsPreview.TEXT = TxtRead(sPath)
End Sub

Private Sub cCopySnippet_Click()
    '@AssignedModule uSnippets
    '@INCLUDE PROCEDURE CLIP
    '@INCLUDE USERFORM uSnippets
    If Len(LSnippetsPreview.TEXT) = 0 Then Exit Sub
    Dim S           As String
    If LSnippetsPreview.SelLength = 0 Then
        S = LSnippetsPreview.TEXT
    Else
        S = LSnippetsPreview.SelText
    End If
    CLIP S
    cResize_Click
    MsgBox "Snipet copied"
End Sub

Private Sub cOverwriteSnippet_Click()
    '@AssignedModule uSnippets
    '@INCLUDE PROCEDURE TxtOverwrite
    '@INCLUDE PROCEDURE InputboxString
    '@INCLUDE USERFORM uSnippets
    '@INCLUDE DECLARATION SnippetsFolder
    Dim sPath       As String
    Dim isNew       As Boolean
    Dim wasResized  As Boolean
    If LSnippets.ListIndex = -1 Then
        Dim ans     As String
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
    '@AssignedModule uSnippets
    '@INCLUDE USERFORM uSnippets
    '@INCLUDE DECLARATION SnippetsFolder
    cResize_Click
    Dim proceed     As Long
    proceed = MsgBox("Delete " & LSnippets.list(LSnippets.ListIndex) & "?", vbYesNo)
    If proceed = vbNo Then Exit Sub
    Dim sPath       As String
    sPath = SnippetsFolder & LSnippets.list(LSnippets.ListIndex)
    Dim FSO         As Scripting.FileSystemObject
    Set FSO = New Scripting.FileSystemObject
    FSO.DeleteFile sPath
    LSnippets.RemoveItem LSnippets.ListIndex
    LSnippetsPreview.TEXT = ""
    LSnippets.ListIndex = -1
    cResize_Click
End Sub

Private Sub cSnippetStartNew_Click()
    '@AssignedModule uSnippets
    '@INCLUDE PROCEDURE FileExists
    '@INCLUDE PROCEDURE TxtOverwrite
    '@INCLUDE USERFORM uSnippets
    '@INCLUDE DECLARATION SnippetsFolder
    Dim NewName     As String
    cResize_Click
    NewName = InputBox("New Snippet Name")
    If NewName = "" Then GoTo ExitHandler
    Dim sPath       As String
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
    '@AssignedModule uSnippets
    '@INCLUDE USERFORM uSnippets
    LSnippetsPreview.selStart = 0
End Sub

Private Sub tFilterSnippets_Change()
    '@AssignedModule uSnippets
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM uSnippets
    GetFilesUSnippets
    aListBox.Init(LSnippets).FilterByColumn tFilterSnippets.TEXT
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveSettings
End Sub
