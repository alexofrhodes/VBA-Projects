VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uVBProjects 
   Caption         =   "Export code - Edit Addins"
   ClientHeight    =   3708
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6684
   OleObjectBlob   =   "uVBProjects.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uVBProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Private Project As vbProject

Private Sub UserForm_Initialize()
    ListVBProjects ListBox1
    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll Me
End Sub

Private Sub Emitter_LabelMouseOut(Label As MSForms.Label)
    If Label.BackColor <> MyColors.FormBackgroundDarkGray Then Label.BackColor = MyColors.FormBackgroundDarkGray '&H80B91E'&H534848
End Sub

Private Sub Emitter_LabelMouseOver(Label As MSForms.Label)
    If Label.BackColor <> MyColors.FormSelectedGreen Then Label.BackColor = MyColors.FormSelectedGreen  '&H80B91E '&H808080
End Sub

Private Sub bEdit_Click()
    If ListBox1.ListIndex = -1 Then Exit Sub
    If Project.FileName = ThisProject.FileName Then
        MsgBox "Permission denied."
        Exit Sub
    End If
    EditAddin Project
    Dim s As String
    Dim i  As Long
    For i = 0 To 1
        s = ListBox1.List(ListBox1.ListIndex, i)
        ListBox1.List(ListBox1.ListIndex, i) = Replace(s, ".ppa", ".ppt")
    Next
End Sub

Private Sub bEndEdit_Click()
    If ListBox1.ListIndex = -1 Then Exit Sub
    Set Project = getProjectByPath(ListBox1.List(ListBox1.ListIndex, 1))
    FinishEditing Project
End Sub

Private Sub bExport_Click()
    If ListBox1.ListIndex = -1 Then Exit Sub
    ExportModules Project
End Sub

Private Sub bImport_Click()
    ImportModules Project, True
End Sub

Private Sub Image1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uAuthor.Show
End Sub

Private Sub ListBox1_Click()
    Set Project = getProjectByPath(ListBox1.List(ListBox1.ListIndex, 1))
End Sub

Private Sub bFolder_Click()
    If ListBox1.ListIndex = -1 Then Exit Sub
    Dim FolderPath As String
        FolderPath = ListBox1.List(ListBox1.ListIndex, 1)
    FollowLink getFileFolder(FolderPath)
End Sub

