VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uIniEditor 
   Caption         =   "                                                     ini Editor"
   ClientHeight    =   7944
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5988
   OleObjectBlob   =   "uIniEditor.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uIniEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Private Sub Label1_Click()
    FollowLink ThisProjectPath & "ExtraIniPaths.txt"
End Sub

Private Sub UserForm_Initialize()
    cbIniFiles.columnCount = 2
    cbIniFiles.ColumnWidths = "200;0"
    cbIniFiles.Font.Size = 12
    Dim out As New Collection
    FolderContents ThisProjectPath, False, True, True, out, "*.ini"
    
    Dim fileArray
    fileArray = Split(TxtRead("ExtraIniPaths.txt"), vbLf)
    Dim item
    For Each item In fileArray
        If Trim(item) <> "" And FileExists(item) Then out.Add item, item
    Next
    If out.count > 0 Then
        Dim FileList()
        ReDim FileList(1 To out.count, 1 To 2)
        Dim i As Long
        For i = 1 To out.count
            FileList(i, 1) = Mid(out(i), InStrRev(out(i), "\") + 1)
            FileList(i, 2) = out(i)
        Next
        cbIniFiles.List = FileList
    End If
    
    cbSections.Font.Size = 12
    
    If out.count = 1 Then cbIniFiles.ListIndex = 0
    
    TextBox1.MultiLine = True
    TextBox1.Font.Size = 12
    
    bSaveChanges.Caption = vbNewLine & bSaveChanges.Caption
    
    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll Me
End Sub

Private Sub bSaveChanges_Click()
    Dim ini As String: ini = SelectedIniFile
    Dim Changes, aChange, aKey, aValue
        Changes = Split(TextBox1.Text, vbNewLine)
    For Each aChange In Changes
        aKey = Trim(Split(aChange, "=")(0))
        aValue = Trim(Split(aChange, "=")(1))
        If aKey <> "" Then IniWriteKey ini, cbSections.Text, aKey, aValue
    Next
    If SelectedIniFile = RibbonIni Then Ribbon.InvalidateControl cbSections.Text
End Sub

Private Function SelectedIniFile()
    If cbIniFiles.ListIndex >= 0 Then SelectedIniFile = cbIniFiles.List(cbIniFiles.ListIndex, 1)
End Function

Private Sub cbIniFiles_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    Dim index As Long: index = cbIniFiles.ListIndex
    If index > -1 Then Exit Sub
    
    cbIniFiles.Text = Replace(cbIniFiles.Text, """", "")
    Dim ini As String: ini = cbIniFiles.Text: ini = Replace(ini, "/", "\")
    If Not ini Like "*.ini" Then Exit Sub
    If InStr(1, ini, ":") = 0 Then ini = Replace(ThisProjectPath & ini, "\\", "\")
    If Not FileExists(ini) Then
        Dim folder As String: folder = getFileFolder(ini)
        If MsgBox("Create this?" & vbNewLine & ini, vbYesNo) = vbNo Then GoTo ExitPoint
        If folder <> "" Then FoldersCreate folder
        TxtOverwrite ini, ""
    End If
    If Not ListContains(cbIniFiles, ini, 1, False) Then
        cbIniFiles.AddItem
        cbIniFiles.List(cbIniFiles.ListCount - 1, 0) = getFileName(ini) & getFileExtension(ini)
        cbIniFiles.List(cbIniFiles.ListCount - 1, 1) = ini

        If InStr(1, ini, ThisProjectPath) = 0 Then
            Dim ExtraIniPathsFile As String
                ExtraIniPathsFile = ThisProjectPath & "ExtraIniPaths.txt"
            Dim fileText As String
                fileText = TxtRead(ExtraIniPathsFile)
            If InStr(1, fileText, "ini") = 0 Then
                TxtOverwrite ExtraIniPathsFile, fileText & vbLf & ini
            End If
        End If
        cbIniFiles.ListIndex = cbIniFiles.ListCount - 1
    End If
    cbIniFiles_Change
    
ExitPoint:

End Sub

Private Sub cbIniFiles_Change()
    If cbIniFiles.ListIndex > -1 Then
        cbSections.Clear
        cbSections.Text = ""
        TextBox1.Text = ""
        Dim arr: arr = IniSections(SelectedIniFile)
        cbSections.List = arr
    End If
End Sub

Private Sub GetInfo_Click()
    uAuthor.Show
End Sub

Private Sub cbSections_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim ini As String: ini = SelectedIniFile: If ini = "" Then Exit Sub
    If Not IniSectionExists(ini, cbSections.Text) Then
        TextBox1.Value = ""
    ElseIf cbSections.ListIndex >= 0 Then
        TextBox1.Value = Replace(Join(IniReadSection(ini, cbSections.Text), vbLf), "=", " = ")
    End If
End Sub

Private Sub cbSections_Change()
    Dim ini As String: ini = SelectedIniFile: If ini = "" Then Exit Sub
    If cbSections.ListIndex >= 0 Then
        TextBox1.Value = Replace(Join(IniReadSection(ini, cbSections.Text), vbLf), "=", " = ")
    End If
End Sub

Private Sub Emitter_LabelMouseOut(Label As MSForms.Label)
    If Label.Picture = 0 Then If Label.BackColor <> MyColors.FormBackgroundDarkGray Then Label.BackColor = MyColors.FormBackgroundDarkGray '&H80B91E'&H534848
End Sub

Private Sub Emitter_LabelMouseOver(Label As MSForms.Label)
    If Label.Picture = 0 Then If Label.BackColor <> MyColors.FormSelectedGreen Then Label.BackColor = MyColors.FormSelectedGreen '&H80B91E '&H808080
End Sub
