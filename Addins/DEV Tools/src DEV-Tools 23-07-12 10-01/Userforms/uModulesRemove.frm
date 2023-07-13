VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uModulesRemove 
   Caption         =   "Remove Code or Components"
   ClientHeight    =   5736
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7572
   OleObjectBlob   =   "uModulesRemove.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uModulesRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Userform   : RemoveComps
'* Created    : 06-10-2022 10:33
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Option Explicit

Private Sub GetInfo_Click()
    uAuthor.Show
End Sub


Private Sub listOpenBooks_Click()
    addCompsList Workbooks(listOpenBooks.List(listOpenBooks.ListIndex))
End Sub

Private Sub Remover_Click()
    If listOpenBooks.ListIndex = -1 Then
        MsgBox "No book selected"
        Exit Sub
    End If
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(listOpenBooks.List(listOpenBooks.ListIndex))
    RemoveModules TargetWorkbook
End Sub

Private Sub RemoveModules(TargetWorkbook As Workbook)
    If LComponents.ListCount = 0 Then Exit Sub
    Dim Module As VBComponent
    Dim i As Long
    For i = 0 To LComponents.ListCount - 1
        If LComponents.Selected(i) Then
            If oCode.Value = True Then
                Set Module = TargetWorkbook.VBProject.VBComponents(LComponents.List(i, 1))
                aModule.Init(Module).CodeRemove
            ElseIf oComps.Value = True Then
                Set Module = TargetWorkbook.VBProject.VBComponents(LComponents.List(i, 1))
                aModule.Init(Module).Delete
            End If
        End If
    Next i
    addCompsList TargetWorkbook
End Sub

Private Sub addCompsList(TargetWorkbook As Workbook)
    '@INCLUDE SortListboxOnColumn
    '@INCLUDE ModuleTypeToString
    '@INCLUDE GetSheetByCodeName
    '@INCLUDE ResizeUserformToFitControls
    '@INCLUDE ResizeControlColumns
    LComponents.Clear
    Dim vbComp As VBComponent
    For Each vbComp In TargetWorkbook.VBProject.VBComponents
        If vbComp.Name <> "ThisWorkbook" Then
            LComponents.AddItem
            LComponents.List(LComponents.ListCount - 1, 0) = aModule.Init(vbComp).TypeToString
            LComponents.List(LComponents.ListCount - 1, 1) = vbComp.Name
            If vbComp.Type = vbext_ct_Document Then
                LComponents.List(LComponents.ListCount - 1, 2) = GetSheetByCodeName(TargetWorkbook, vbComp.Name).Name
            End If
        End If
    Next
    aListBox.Init(LComponents).SortOnColumn 0
    Me.Caption = "Comps of " & TargetWorkbook.Name
    '    aListBox.Init(LComponents).SortOnColumn 0

End Sub

Private Sub UserForm_Initialize()
    aListBox.Init(listOpenBooks).LoadVBProjects
End Sub
