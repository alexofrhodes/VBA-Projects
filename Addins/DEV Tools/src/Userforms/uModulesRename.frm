VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uModulesRename 
   Caption         =   "Rename Components"
   ClientHeight    =   6564
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10236
   OleObjectBlob   =   "uModulesRename.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uModulesRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Userform   : RenameComps
'* Created    : 06-10-2022 10:33
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Private Sub UserForm_Initialize()
    aListBox.Init(listOpenBooks).LoadVBProjects
End Sub

Private Sub GetInfo_Click()
    uAuthor.Show
End Sub

Private Sub listOpenBooks_Click()
    addCompsList Workbooks(listOpenBooks.List(listOpenBooks.ListIndex))
End Sub

Private Sub RenameComponents_Click()
    If listOpenBooks.ListIndex = -1 Then
        MsgBox "No book selected"
        Exit Sub
    End If
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(listOpenBooks.List(listOpenBooks.ListIndex))
    Dim NewNames As Variant
    Dim i As Long
    NewNames = Split(textboxNewName, vbNewLine)
    For i = 0 To UBound(NewNames)
        If NewNames(i) = vbNullString Then
            NewNames(i) = LComponents.List(i)
        End If
    Next i
    For i = 0 To UBound(NewNames)
Retry:
        On Error GoTo EH
'        Select Case LComponents.list(i, 0)
'        Case Is = "Module", "Class", "UserForm"
            If LComponents.List(i, 1) <> NewNames(i) Then
                TargetWorkbook.VBProject.VBComponents(LComponents.List(i, 1)).Name = NewNames(i)
            End If
'        Case Is = "Document"
'            If LComponents.list(i, 1) <> NewNames(i) Then
'                GetSheetByCodeName(TargetWorkbook, LComponents.list(i, 1)).name = NewNames(i)
'            End If
'        End Select
    Next
    For i = 0 To LComponents.ListCount - 1
        LComponents.List(i, 1) = NewNames(i)
    Next i
    textboxNewName.TEXT = vbNullString
    Dim str As String
    str = Join(NewNames, vbNewLine)
    textboxNewName.TEXT = str
    MsgBox "Components renamed"
    Exit Sub
EH:
    NewNames(i) = NewNames(i) & "_R"
    Resume Retry:
End Sub


Private Sub addCompsList(TargetWorkbook As Workbook)
    LComponents.Clear
    Dim vbcomp As VBComponent
    For Each vbcomp In TargetWorkbook.VBProject.VBComponents
        If vbcomp.Name <> "ThisWorkbook" Then
            LComponents.AddItem
            LComponents.List(LComponents.ListCount - 1, 0) = aModule.Init(vbcomp).TypeToString
            LComponents.List(LComponents.ListCount - 1, 1) = vbcomp.Name
            If vbcomp.Type = vbext_ct_Document Then
                LComponents.List(LComponents.ListCount - 1, 2) = GetSheetByCodeName(TargetWorkbook, vbcomp.Name).Name
            End If
        End If
    Next
    Me.Caption = "Comps of " & TargetWorkbook.Name
    aListBox.Init(LComponents).SortOnColumn 0
    SyncNames TargetWorkbook
End Sub
Private Sub SyncNames(TargetWorkbook As Workbook)
    Dim str As String
    str = LComponents.List(0, 1)
    Dim i As Long
    For i = 1 To LComponents.ListCount - 1
        str = str & vbNewLine & LComponents.List(i, 1)
    Next
    textboxNewName.TEXT = str
End Sub
