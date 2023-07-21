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

Private al As aListBox

Sub LoadSample()
    With ListBox1
        .columnCount = 10
        Dim X As Long, Y As Long
        Dim var(1 To 10, 1 To 10)
        For X = 1 To 10
            For Y = 1 To 10
                var(X, Y) = X * Y
            Next
        Next
        .List = var
        .multiSelect = fmMultiSelectSingle
        
    End With
End Sub
Private Sub Label10_Click()
    al.removeHeaders
    With ListBox1
        .columnCount = 10
        .Clear
        .List = RandomStringArray(10, 10, 5)
    End With
End Sub

Private Sub CheckBox1_Click()
    If CheckBox1.Value = False Then
        ListBox1.Width = 185
    Else
        al.AutofitColumns CheckBox1.Value
    End If
    aUserform.Init(Me).ResizeToFitControls 10, 10
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2.Value = True Then
        ListBox1.multiSelect = fmMultiSelectExtended
    Else
        ListBox1.multiSelect = fmMultiSelectSingle
    End If
End Sub



Private Sub Label11_Click()
'    dp al.SelectedRowsText
    dp al.SelectedRowsArray
End Sub



Private Sub Label12_Click()
    al.ListenToDoubleClick
End Sub

Private Sub Label13_Click()
    al.ListenToExtendedSelection
End Sub

Private Sub Label14_Click()
    al.ListenToDragDrop ListBox2
End Sub

Private Sub Label2_Click()
    al.AutofitColumns CheckBox1.Value
End Sub

Private Sub Label3_Click()
    al.AcceptFiles sExpansion:="*", iDeepSubPath:=999
    Me.Caption = "Add files to Listbox by Drag and Drop"
End Sub

Private Sub Label4_Click()
    al.AddFilter
End Sub

Private Sub Label5_Click()
    LoadSample
End Sub

Private Sub Label6_Click()
    al.RememberList
End Sub

Private Sub Label7_Click()
    Dim i As Long: i = InputBox("How many rows to display?")
    If i >= 1 And i <= ListBox1.ListCount Then al.HeightToEntries i
End Sub

Private Sub Label8_Click()
    al.AddHeader
    Me.Caption = "Click on header to sort A to Z"
End Sub

Private Sub UserForm_Initialize()
    Set al = New aListBox
    al.Init ListBox1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set al = Nothing
End Sub

