VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} z_ListBox 
   Caption         =   "UserForm1"
   ClientHeight    =   5640
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8376.001
   OleObjectBlob   =   "z_ListBox.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "z_ListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private al          As aListBox

Sub LoadSample()
    '@AssignedModule z_ListBox
    '@INCLUDE USERFORM z_ListBox
    With ListBox1
        .columnCount = 10
        Dim x As Long, y As Long
        Dim var(1 To 10, 1 To 10)
        For x = 1 To 10
            For y = 1 To 10
                var(x, y) = x * y
            Next
        Next
        .list = var
        .multiSelect = fmMultiSelectSingle

    End With
End Sub
Private Sub Label10_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE PROCEDURE RandomStringArray
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    al.removeHeaders
    With ListBox1
        .columnCount = 10
        .clear
        .list = RandomStringArray(10, 10, 5)
    End With
End Sub


Private Sub CheckBox2_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    If CheckBox2.value = True Then
        ListBox1.multiSelect = fmMultiSelectExtended
    Else
        ListBox1.multiSelect = fmMultiSelectSingle
    End If
End Sub



Private Sub Label11_Click()
    '    dp al.SelectedRowsText
    '@AssignedModule z_ListBox
    '@INCLUDE PROCEDURE dp
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    dp al.SelectedRowsArray
End Sub



Private Sub Label12_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    al.ListenToDoubleClick
End Sub

Private Sub Label13_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    al.ListenToExtendedSelection
End Sub

Private Sub Label14_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    al.ListenToDragDrop ListBox2
End Sub

Private Sub Label15_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    Dim arr
    arr = Split(TextBox1.value, ",")
    al.ShowTheseColumns arr
End Sub

Private Sub Label2_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE CLASS aUserform
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    al.AutofitColumns True
    aUserform.Init(Me).ResizeToFitControls 10, 10
End Sub

Private Sub Label3_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    Me.Caption = "Add files to Listbox by Drag and Drop"
    al.AcceptFiles sExpansion:="*", iDeepSubPath:=999
End Sub

Private Sub Label4_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    al.AddFilter
End Sub

Private Sub Label5_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE USERFORM z_ListBox
    LoadSample
End Sub

Private Sub Label6_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    al.RememberList
End Sub

Private Sub Label7_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    Dim i           As Long: i = InputBox("How many rows to display?")
    If i >= 1 And i <= ListBox1.ListCount Then al.HeightToEntries i
End Sub

Private Sub Label8_Click()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    al.AddHeader
    Me.Caption = "Click on header to sort A to Z"
End Sub

Private Sub UserForm_Initialize()
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    Set al = New aListBox
    al.Init ListBox1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '@AssignedModule z_ListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM z_ListBox
    '@INCLUDE DECLARATION al
    Set al = Nothing
End Sub

