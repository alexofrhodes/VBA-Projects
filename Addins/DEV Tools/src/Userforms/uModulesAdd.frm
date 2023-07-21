VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uModulesAdd 
   Caption         =   "Add Components"
   ClientHeight    =   4704
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11520
   OleObjectBlob   =   "uModulesAdd.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uModulesAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Userform   : AddComps
'* Created    : 06-10-2022 10:33
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Option Explicit





Private Sub ModulesAdd(TargetWorkbook As Workbook)
    Dim coll As Collection
    Set coll = New Collection
    Dim element As Variant
    coll.Add Split(Me.tModule.text, vbNewLine)
    coll.Add Split(Me.tClass.text, vbNewLine)
    coll.Add Split(Me.tUserform.text, vbNewLine)
    coll.Add Split(Me.tDocument.text, vbNewLine)
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
    uAuthor.Show
End Sub



Private Sub SelectFromList_Click()
    If listOpenBooks.ListIndex = -1 Then
        MsgBox "No book selected"
        Exit Sub
    End If
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(listOpenBooks.List(listOpenBooks.ListIndex))
    ModulesAdd TargetWorkbook
End Sub



Private Sub AddComponent(TargetWorkbook As Workbook, Module_Class_Form_Sheet As Long, componentArray As Variant)
    '@INCLUDE ModuleExists
    Dim CompType As Long
    CompType = Module_Class_Form_Sheet
    Dim vbProj As VBProject
    Set vbProj = TargetWorkbook.VBProject
    Dim vbComp As VBComponent
    Dim NewSheet As Worksheet
    Dim i As Long
    Dim counter As Long
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

Private Sub UserForm_Initialize()
    aListBox.Init(listOpenBooks).LoadVBProjects
End Sub
