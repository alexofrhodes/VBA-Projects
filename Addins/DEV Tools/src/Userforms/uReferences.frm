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
    Set TargetWorkbook = Workbooks(ListBox1.list(ListBox1.ListIndex))
    On Error Resume Next
    TargetWorkbook.VBProject.REFERENCES.AddFromGuid LReferences.list(LReferences.ListIndex, 1), 0, 0

    Call PopulateLRefActive

End Sub

Private Sub cClearFilter_Click()
    '    LReferences.Clear
    '    Call PopulateLReferences
    tFilterReferences.TEXT = ""
End Sub

Private Sub cRemove_Click()
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
    uAuthor.Show
End Sub

Private Sub ListBox1_Change()
    Call PopulateLRefActive
End Sub

Private Sub tFilterReferences_Change()
    'Reload list so if you type and delete you'll get the items back
    LReferences.Clear
    Call PopulateLReferences

    Dim i               As Long
    Dim n               As Long
    Dim str             As String
    Dim sTemp           As String

    'Equals is always case sensitive
    'Remove LCase if you want it to be case sensitive
    str = LCase(tFilterReferences.TEXT)

    n = LReferences.ListCount

    'Work backwards when deleting items
    For i = n - 1 To 0 Step -1
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
End Sub

Private Sub UserForm_Initialize()
    Call PopulateLReferences

    Dim X, Y As Variant
    On Error Resume Next
    For Each X In Array(Workbooks, AddIns)
        For Each Y In X
            If Not WorkbookProjectProtected(Workbooks(Y.Name)) Then
                If Err.Number = 0 Then
                    ListBox1.AddItem Y.Name
                End If
            End If
            Err.Clear
        Next
    Next

End Sub

Function PopulateLRefActive()
    '@INCLUDE SortListboxOnColumn
    '@INCLUDE DP
    Dim FromWorkbook As Workbook
    Set FromWorkbook = Workbooks(ListBox1.list(ListBox1.ListIndex))
    Dim i As Long
    i = 0
    LRefActive.Clear
    Dim myRef As Reference
    For Each myRef In FromWorkbook.VBProject.REFERENCES
        uReferences.LRefActive.AddItem
        uReferences.LRefActive.list(i, 0) = myRef.IsBroken
        uReferences.LRefActive.list(i, 1) = IIf(myRef.Description <> "", myRef.Description, myRef.Name)
        uReferences.LRefActive.list(i, 2) = myRef.GUID
        dp myRef.Name
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
        uReferences.LReferences.AddItem
        uReferences.LReferences.list(i, 0) = cell.TEXT
        uReferences.LReferences.list(i, 1) = cell.OFFSET(0, 1).TEXT
        uReferences.LReferences.list(i, 2) = cell.OFFSET(0, 2).TEXT
        uReferences.LReferences.list(i, 3) = cell.OFFSET(0, 3).TEXT
        i = i + 1
    Next cell

End Sub

