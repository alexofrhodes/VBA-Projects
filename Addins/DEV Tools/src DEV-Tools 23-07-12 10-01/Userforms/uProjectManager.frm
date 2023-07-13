VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uProjectManager 
   Caption         =   "github.com/AlexOfRhodes"
   ClientHeight    =   3804
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5688
   OleObjectBlob   =   "uProjectManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uProjectManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Userform   : uProjectManager
'* Created    : 06-10-2022 10:39
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private Sub goToFolder_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    FollowLink Environ("USERprofile") & "\Documents\vbArc\"
End Sub

Private Sub GetInfo_Click()
    uAuthor.Show
End Sub

Private Sub ReloadListbox_Click()
    listOpenBooks.Clear
    aListBox.Init(listOpenBooks).LoadVBProjects
    Label1.Caption = ""
End Sub

Private Sub listOpenBooks_Click()
    AssignPathLabel
End Sub

Private Sub AssignPathLabel()
    If listOpenBooks.ListIndex = -1 Then Exit Sub
    Label1.Caption = IIf(UseWBpath, _
                        Workbooks(listOpenBooks.List(listOpenBooks.ListIndex)).Path & "\", _
                        Environ("USERprofile") & "\Documents\" & "vbArc\Backups\")
End Sub

Private Sub SelectFromList_Click()
    If listOpenBooks.ListIndex = -1 Then
        MsgBox "No book selected"
        Exit Sub
    End If
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(listOpenBooks.List(listOpenBooks.ListIndex))
    DoExport TargetWorkbook
End Sub

Private Sub UserForm_Initialize()
    aListBox.Init(listOpenBooks).LoadVBProjects
    aListBox.Init(listOpenBooks).SortOnColumn 0
    aUserform.Init(Me).LoadOptions
    AssignPathLabel
End Sub

Private Sub DoExport(TargetWorkbook As Workbook)


    If WorkbookProjectProtected(TargetWorkbook) Then
        Toast "Project of " & TargetWorkbook.Name & " is protected."
        Exit Sub
    End If

    Me.Hide

    aWorkbook.Init(TargetWorkbook).ExportProject _
                                                 bExportComponents:=chExportComponents.Value _
                                                , bSeparateProcedures:=chExportProcedures.Value _
                                                , bExportXML:=chExportXML.Value _
                                                , bExportReferences:=chExportReferences.Value _
                                                , bExportDeclarations:=chExportDeclarations.Value _
                                                , bExportUnified:=chExportUnified.Value _
                                                , bWorkbookBackup:=chWorkbookBackup.Value _
                                                , UseWorkbookFolder:=UseWBpath _
                                                , OpenFolderAfterExport:=OpenFolder.Value
        
    Me.Show
End Sub

Private Sub SelectFile_Click()
    Dim fPath As String: fPath = PickExcelFile
    If fPath = "" Then Exit Sub
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks.Open(FileName:=fPath, UpdateLinks:=0, ReadOnly:=False)
    DoExport TargetWorkbook
    TargetWorkbook.Close True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    aUserform.Init(Me).SaveOptions includeListbox:=False
End Sub

Private Sub UseWBpath_Click()
    AssignPathLabel
End Sub
