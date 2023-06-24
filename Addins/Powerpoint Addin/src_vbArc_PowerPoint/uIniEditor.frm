VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uIniEditor 
   Caption         =   "UserForm1"
   ClientHeight    =   8040
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5988
   OleObjectBlob   =   "uIniEditor.frx":0000
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

Private Sub bSaveChanges_Click()
    Dim INI As String: INI = cbIniFiles.Text 'RibbonIni      ' <===== change
    Dim Changes, aChange, aKey, aValue
        Changes = Split(TextBox1.Text, vbNewLine)
    For Each aChange In Changes
        aKey = Trim(Split(aChange, "=")(0))
        aValue = Trim(Split(aChange, "=")(1))
        If aKey <> "" Then IniWriteKey INI, cbSections.Text, aKey, aValue
    Next
    myRibbon.InvalidateControl cbSections.Text
End Sub

Private Sub cbIniFiles_Change()
    If cbIniFiles.ListIndex >= 0 Then
        cbSections.List = IniSections(cbIniFiles.List(cbIniFiles.ListIndex, 1))
    End If
End Sub



Private Sub Image1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uAuthor.Show
End Sub

Private Sub UserForm_Initialize()
    cbIniFiles.Style = fmStyleDropDownList
    cbIniFiles.ColumnCount = 2
    cbIniFiles.ColumnWidths = "200;0"
    cbIniFiles.Font.Size = 12
    Dim out As New Collection
    FolderContents ThisProjectPath, False, True, True, out, "*.ini"
    
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
    
    cbSections.Style = fmStyleDropDownList
'    cbSections.List = IniSections(cbIniFiles.List(cbIniFiles.ListIndex))
    cbSections.Font.Size = 12
'    cbSections.ListIndex = 0
    
    If out.count = 1 Then cbIniFiles.ListIndex = 0
    
    TextBox1.MultiLine = True
    TextBox1.Font.Size = 12
    
    bSaveChanges.Caption = vbNewLine & bSaveChanges.Caption
    
    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll Me
End Sub

Private Sub cbSections_Change()
    If cbSections.ListIndex >= 0 Then
        TextBox1.Value = Join(IniReadSection(RibbonIni, cbSections.List(cbSections.ListIndex)), vbNewLine)
    End If
End Sub

Private Sub Emitter_LabelMouseOut(Label As MSForms.Label)
    If Label.BackColor <> MyColors.FormBackgroundDarkGray Then Label.BackColor = MyColors.FormBackgroundDarkGray '&H80B91E'&H534848
End Sub

Private Sub Emitter_LabelMouseOver(Label As MSForms.Label)
    If Label.BackColor <> MyColors.FormSelectedGreen Then Label.BackColor = MyColors.FormSelectedGreen  '&H80B91E '&H808080
End Sub

'Private Sub Emitter_LabelMouseOut(Label As MSForms.Label)
'    If Label.BackColor <> &H80B91E Then Label.BackColor = &H534848
'End Sub
'
'Private Sub Emitter_LabelMouseOver(Label As MSForms.Label)
'    If Label.BackColor <> &H80B91E Then Label.BackColor = &H808080
'End Sub

