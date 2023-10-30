VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uModules 
   Caption         =   "UserForm1"
   ClientHeight    =   9936.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   18828
   OleObjectBlob   =   "uModules.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm   : uModules
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
'* Created    : 17-08-2023 13:31    Alex                merged forms for modules Add / Rename / Remove
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private Sub listOpenBooks_Click()
    '@AssignedModule uModulesRemove
    '@INCLUDE USERFORM uModulesRemove
    addCompsList Workbooks(listOpenBooks.list(listOpenBooks.ListIndex))
End Sub

Private Sub UserForm_Initialize()
    SidebarRight.Visible = True
    SidebarBottom.Visible = False

    '// the class is predeclaredId = true but shouldn't this way still work?
    '    Dim am1 As aMultiPage
    '    Set am1 = New aMultiPage
    '    am1.Init(MultiPage1).BuildMenu createSidebarMinimizers:=True

    aMultiPage.Init(MultiPage1).BuildMenu createSidebarMinimizers:=True

    aListBox.Init(listOpenBooks).LoadVBProjects
End Sub

Private Sub SelectFromList_Click()
    '@AssignedModule uModulesAdd
    '@INCLUDE USERFORM uModulesAdd
    If listOpenBooks.ListIndex = -1 Then
        MsgBox "No book selected"
        Exit Sub
    End If
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(listOpenBooks.list(listOpenBooks.ListIndex))
    ModulesAdd TargetWorkbook
End Sub



Private Sub AddComponent(TargetWorkbook As Workbook, Module_Class_Form_Sheet As Long, componentArray As Variant)
    '@INCLUDE ModuleExists
    '@AssignedModule uModulesAdd
    '@INCLUDE PROCEDURE ModuleExists
    '@INCLUDE PROCEDURE CreateOrSetSheet
    '@INCLUDE USERFORM uModulesAdd
    Dim CompType    As Long
    CompType = Module_Class_Form_Sheet
    Dim vbProj      As VBProject
    Set vbProj = TargetWorkbook.VBProject
    Dim vbComp      As VBComponent
    Dim NewSheet    As Worksheet
    Dim i           As Long
    Dim counter     As Long
    On Error GoTo ErrorHandler
    For i = LBound(componentArray) To UBound(componentArray)
        If componentArray(i) <> vbNullString Then
            Select Case CompType
                Case Is = 1, 2, 3
                    If ModuleExists(CStr(componentArray(i)), TargetWorkbook) = False Then
                        If CompType = 1 Then Set vbComp = vbProj.VBComponents.Add(vbext_ct_StdModule)
                        If CompType = 2 Then Set vbComp = vbProj.VBComponents.Add(vbext_ct_ClassModule)
                        If CompType = 3 Then Set vbComp = vbProj.VBComponents.Add(vbext_ct_MSForm)
                    End If
                    vbComp.Name = componentArray(i)
                Case Is = 4
                    If CompType = 4 Then
                        Set NewSheet = CreateOrSetSheet(CStr(componentArray(i)), TargetWorkbook)
                        NewSheet.Name = componentArray(i)
                    End If
            End Select
        End If
loop1:
    Next i
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    counter = counter + 1
    componentArray(i) = componentArray(i) & counter
    GoTo loop1
End Sub


Private Sub ModulesAdd(TargetWorkbook As Workbook)
    '@AssignedModule uModulesAdd
    '@INCLUDE USERFORM uModulesAdd
    Dim coll        As Collection
    Set coll = New Collection
    Dim element     As Variant
    coll.Add Split(Me.tModule.TEXT, vbNewLine)
    coll.Add Split(Me.tClass.TEXT, vbNewLine)
    coll.Add Split(Me.tUserform.TEXT, vbNewLine)
    coll.Add Split(Me.tDocument.TEXT, vbNewLine)
    Dim typeCounter As Long
    For Each element In coll
        If UBound(element) <> -1 Then
            typeCounter = typeCounter + 1
            AddComponent TargetWorkbook, typeCounter, element
        End If
    Next element
    MsgBox typeCounter & " components added to " & TargetWorkbook.Name
End Sub


Private Sub GetInfo_Click()
    '@AssignedModule uModulesAdd
    '@INCLUDE USERFORM uModulesAdd
    '@INCLUDE USERFORM uAuthor
    uAuthor.Show
End Sub


Private Sub Remover_Click()
    '@AssignedModule uModulesRemove
    '@INCLUDE USERFORM uModulesRemove
    If listOpenBooks.ListIndex = -1 Then
        MsgBox "No book selected"
        Exit Sub
    End If
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(listOpenBooks.list(listOpenBooks.ListIndex))
    RemoveModules TargetWorkbook
End Sub

Private Sub RemoveModules(TargetWorkbook As Workbook)
    '@AssignedModule uModulesRemove
    '@INCLUDE CLASS aModule
    '@INCLUDE USERFORM uModulesRemove
    If LComponents.ListCount = 0 Then Exit Sub
    Dim module      As VBComponent
    Dim i           As Long
    For i = 0 To LComponents.ListCount - 1
        If LComponents.Selected(i) Then
            If oCode.Value = True Then
                Set module = TargetWorkbook.VBProject.VBComponents(LComponents.list(i, 1))
                aModule.Init(module).CodeRemove
            ElseIf oComps.Value = True Then
                Set module = TargetWorkbook.VBProject.VBComponents(LComponents.list(i, 1))
                aModule.Init(module).Delete
            End If
        End If
    Next i
    addCompsList TargetWorkbook
End Sub


Private Sub addCompsList(TargetWorkbook As Workbook)
    '@AssignedModule uModulesRename
    '@INCLUDE PROCEDURE GetSheetByCodeName
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM uModulesRename
    LComponents.clear
    Dim vbComp      As VBComponent
    For Each vbComp In TargetWorkbook.VBProject.VBComponents
        If vbComp.Name <> "ThisWorkbook" Then
            LComponents.AddItem
            LComponents.list(LComponents.ListCount - 1, 0) = aModule.Init(vbComp).TypeToString
            LComponents.list(LComponents.ListCount - 1, 1) = vbComp.Name
            If vbComp.Type = vbext_ct_Document Then
                LComponents.list(LComponents.ListCount - 1, 2) = GetSheetByCodeName(TargetWorkbook, vbComp.Name).Name
            End If
        End If
    Next
    Me.Caption = "Comps of " & TargetWorkbook.Name
    aListBox.Init(LComponents).SortOnColumn 1

    SyncNames TargetWorkbook
End Sub

Private Sub RenameComponents_Click()
    '@AssignedModule uModulesRename
    '@INCLUDE PROCEDURE GetSheetByCodeName
    '@INCLUDE USERFORM uModulesRename
    If listOpenBooks.ListIndex = -1 Then
        MsgBox "No book selected"
        Exit Sub
    End If
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(listOpenBooks.list(listOpenBooks.ListIndex))
    Dim NewNames    As Variant
    Dim i           As Long
    NewNames = Split(textboxNewName, vbNewLine)
    For i = 0 To UBound(NewNames)
        If NewNames(i) = vbNullString Then
            NewNames(i) = LComponents.list(i)
        End If
    Next i
    For i = 0 To UBound(NewNames)
retry:
        On Error GoTo EH
        '        Select Case LComponents.list(i, 0)
        '        Case Is = "Module", "Class", "UserForm"
        If LComponents.list(i, 1) <> NewNames(i) Then
            TargetWorkbook.VBProject.VBComponents(LComponents.list(i, 1)).Name = NewNames(i)
        End If
        '        Case Is = "Document"
        '            If LComponents.list(i, 1) <> NewNames(i) Then
        '                GetSheetByCodeName(TargetWorkbook, LComponents.list(i, 1)).name = NewNames(i)
        '            End If
        '        End Select
    Next
    For i = 0 To LComponents.ListCount - 1
        LComponents.list(i, 1) = NewNames(i)
    Next i
    textboxNewName.TEXT = vbNullString
    Dim str         As String
    str = Join(NewNames, vbNewLine)
    textboxNewName.TEXT = str
    MsgBox "Components renamed"
    Exit Sub
EH:
    NewNames(i) = NewNames(i) & "_R"
    Resume retry:
End Sub



Private Sub SyncNames(TargetWorkbook As Workbook)
    '@AssignedModule uModulesRename
    '@INCLUDE USERFORM uModulesRename
    Dim str         As String
    Dim i           As Long
    For i = 0 To LComponents.ListCount - 1
        str = str & IIf(i > 0, vbNewLine, "") & LComponents.list(i, 1)
    Next
    textboxNewName.TEXT = str
    textboxNewName.ScrollBars = fmScrollBarsVertical
    textboxNewName.SetFocus
    textboxNewName.selStart = 0
End Sub
