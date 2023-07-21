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
    If ListBox1.ListIndex = -1 Then
        MsgBox "Select target workbook first"
        Exit Sub
    End If
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(ListBox1.List(ListBox1.ListIndex))
    On Error Resume Next
    TargetWorkbook.VBProject.REFERENCES.AddFromGuid LReferences.List(LReferences.ListIndex, 1), 0, 0

    Call PopulateLRefActive

End Sub

Private Sub cClearFilter_Click()
    '    LReferences.Clear
    '    Call PopulateLReferences
    tFilterReferences.text = ""
End Sub

Private Sub cRemove_Click()
    If ListBox1.ListIndex = -1 Then
        MsgBox "Select target workbook first"
        Exit Sub
    End If
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(ListBox1.List(ListBox1.ListIndex))
    On Error Resume Next
    aWorkbook.Init(TargetWorkbook).RemoveReferenceByGUID LRefActive.List(LReferences.ListIndex, 2)
    PopulateLRefActive
End Sub

Private Sub GetInfo_Click()
    With aUserform.Init(Me)
        .Transition .Effect(GetInfo, "Top", GetInfo.Top - 10, 150)
        PlayTheSound ThisWorkbook.Path & "\sound\coin.wav"
        .Transition .Effect(GetInfo, "Top", GetInfo.Top + 10, 150)
    End With
    uAuthor.Show
End Sub


Private Sub listbox1_change()
    Call PopulateLRefActive
End Sub

Private Sub tFilterReferences_Change()
    'Reload list so if you type and delete you'll get the items back
    LReferences.Clear
    Call PopulateLReferences

    Dim i               As Long
    Dim N               As Long
    Dim str             As String
    Dim sTemp           As String

    'Equals is always case sensitive
    'Remove LCase if you want it to be case sensitive
    str = LCase(tFilterReferences.text)

    N = LReferences.ListCount

    'Work backwards when deleting items
    For i = N - 1 To 0 Step -1
        'Equals is always case sensitive
        'Remove LCase if you want it to be case sensitive
        sTemp = LCase(LReferences.List(i, 0))

        If InStr(sTemp, str) = 0 Then
            LReferences.RemoveItem (i)
            'Exit Sub   'Uncomment to Exit if value found
        End If
    Next i
End Sub

Private Sub UserForm_Activate()
  'MakeUserFormChildOfVBEditor Me.Caption
End Sub

Private Sub UserForm_Initialize()
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
    Dim FromWorkbook As Workbook
    Set FromWorkbook = Workbooks(ListBox1.List(ListBox1.ListIndex))
    Dim i As Long
    i = 0
    LRefActive.Clear
    Dim myRef As Reference
    For Each myRef In FromWorkbook.VBProject.REFERENCES
        LRefActive.AddItem
        LRefActive.List(i, 0) = myRef.IsBroken
        LRefActive.List(i, 1) = IIf(myRef.Description <> "", myRef.Description, myRef.Name)
        LRefActive.List(i, 2) = myRef.GUID
        i = i + 1
    Next myRef
    aListBox.Init(LRefActive).SortOnColumn 1
End Function

Sub PopulateLReferences()
    Dim i As Long
    i = 0
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("REFERENCES").Range("A1").CurrentRegion
    Set rng = rng.OFFSET(1).Resize(rng.rows.count - 1)
    Dim cell As Range
    For Each cell In rng.Columns(1).Cells
        LReferences.AddItem
        LReferences.List(i, 0) = cell.text
        LReferences.List(i, 1) = cell.OFFSET(0, 1).text
        LReferences.List(i, 2) = cell.OFFSET(0, 2).text
        LReferences.List(i, 3) = cell.OFFSET(0, 3).text
        i = i + 1
    Next cell

End Sub

