VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uReferences 
   Caption         =   "References"
   ClientHeight    =   8664.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11484
   OleObjectBlob   =   "uReferences.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uReferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Userform   : uReferences
'* Created    : 06-10-2022 10:40
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private Sub cADD_Click()
    '@AssignedModule uReferences
    '@INCLUDE USERFORM uReferences
    If ListBox1.ListIndex = -1 Then
        MsgBox "Select target workbook first"
        Exit Sub
    End If
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(ListBox1.list(ListBox1.ListIndex))
    On Error Resume Next
    TargetWorkbook.VBProject.REFERENCES.AddFromGuid LReferences.list(LReferences.ListIndex, 1), 0, 0

    Call PopulateLRefActive

End Sub

Private Sub cClearFilter_Click()
    '    LReferences.Clear
    '    Call PopulateLReferences
    '@AssignedModule uReferences
    '@INCLUDE USERFORM uReferences
    tFilterReferences.TEXT = ""
End Sub

Private Sub cRemove_Click()
    '@AssignedModule uReferences
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE USERFORM uReferences
    If ListBox1.ListIndex = -1 Then
        MsgBox "Select target workbook first"
        Exit Sub
    End If
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(ListBox1.list(ListBox1.ListIndex))
    On Error Resume Next
    aWorkbook.Init(TargetWorkbook).RemoveReferenceByGUID LRefActive.list(LReferences.ListIndex, 2)
    PopulateLRefActive
End Sub

Private Sub GetInfo_Click()
    '@AssignedModule uReferences
    '@INCLUDE PROCEDURE PlayTheSound
    '@INCLUDE CLASS aUserform
    '@INCLUDE USERFORM uReferences
    '@INCLUDE USERFORM uAuthor
    With aUserform.Init(Me)
        .Transition .Effect(GetInfo, "Top", GetInfo.Top - 10, 150)

        .Transition .Effect(GetInfo, "Top", GetInfo.Top + 10, 150)
    End With
    uAuthor.Show
End Sub


Private Sub listbox1_change()
    '@AssignedModule uReferences
    '@INCLUDE USERFORM uReferences
    Call PopulateLRefActive
End Sub

Private Sub tFilterReferences_Change()
    'Reload list so if you type and delete you'll get the items back
    '@AssignedModule uReferences
    '@INCLUDE USERFORM uReferences
    LReferences.clear
    Call PopulateLReferences

    Dim i           As Long
    Dim N           As Long
    Dim str         As String
    Dim sTemp       As String

    'Equals is always case sensitive
    'Remove LCase if you want it to be case sensitive
    str = LCase(tFilterReferences.TEXT)

    N = LReferences.ListCount

    'Work backwards when deleting items
    For i = N - 1 To 0 Step -1
        'Equals is always case sensitive
        'Remove LCase if you want it to be case sensitive
        sTemp = LCase(LReferences.list(i, 0))

        If InStr(sTemp, str) = 0 Then
            LReferences.RemoveItem (i)
            'Exit Sub   'Uncomment to Exit if value found
        End If
    Next i
End Sub

Private Sub UserForm_Activate()
    'MakeUserFormChildOfVBEditor Me.Caption
    '@AssignedModule uReferences
    '@INCLUDE USERFORM uReferences
End Sub

Private Sub UserForm_Initialize()
    '@AssignedModule uReferences
    '@INCLUDE PROCEDURE WorkbookProjectProtected
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM uReferences
    aListBox.Init(ListBox1).LoadVBProjects
    Call PopulateLReferences

    '    Dim x, y As Variant
    '    On Error Resume Next
    '    For Each x In Array(Workbooks, AddIns)
    '        For Each y In x
    '            If Not WorkbookProjectProtected(Workbooks(y.Name)) Then
    '                If Err.Number = 0 Then
    '                    ListBox1.AddItem y.Name
    '                End If
    '            End If
    '            Err.Clear
    '        Next
    '    Next
End Sub

Function PopulateLRefActive()
    '@INCLUDE SortListboxOnColumn
    '@INCLUDE DP
    '@AssignedModule uReferences
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM uReferences
    '@INCLUDE DECLARATION GUID
    Dim FromWorkbook As Workbook
    Set FromWorkbook = Workbooks(ListBox1.list(ListBox1.ListIndex))
    Dim i           As Long
    i = 0
    LRefActive.clear
    Dim myRef       As Reference
    For Each myRef In FromWorkbook.VBProject.REFERENCES
        LRefActive.AddItem
        LRefActive.list(i, 0) = myRef.IsBroken
        LRefActive.list(i, 1) = IIf(myRef.Description <> "", myRef.Description, myRef.Name)
        LRefActive.list(i, 2) = myRef.GUID
        i = i + 1
    Next myRef
    aListBox.Init(LRefActive).SortOnColumn 1
End Function

Sub PopulateLReferences()
    '@AssignedModule uReferences
    '@INCLUDE USERFORM uReferences
    Dim i           As Long
    i = 0
    Dim rng         As Range
    Set rng = ThisWorkbook.Sheets("REFERENCES").Range("A1").CurrentRegion
    Set rng = rng.offset(1).Resize(rng.rows.Count - 1)
    Dim cell        As Range
    For Each cell In rng.Columns(1).Cells
        LReferences.AddItem
        LReferences.list(i, 0) = cell.TEXT
        LReferences.list(i, 1) = cell.offset(0, 1).TEXT
        LReferences.list(i, 2) = cell.offset(0, 2).TEXT
        LReferences.list(i, 3) = cell.offset(0, 3).TEXT
        i = i + 1
    Next cell

End Sub

