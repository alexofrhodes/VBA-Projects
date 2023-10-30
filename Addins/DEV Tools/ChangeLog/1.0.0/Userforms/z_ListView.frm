VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} z_ListView 
   Caption         =   "UserForm1"
   ClientHeight    =   6456
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9732.001
   OleObjectBlob   =   "z_ListView.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "z_ListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Label2_Click()
    '@AssignedModule z_ListView
    '@INCLUDE USERFORM z_ListView

End Sub

Private Sub UserForm_Initialize()
    '@AssignedModule z_ListView
    '@INCLUDE CLASS aListView
    '@INCLUDE USERFORM z_ListView
    Dim i           As Long
    For i = 1 To 4
        ListView1.ListItems.Add , , "Test" & i
    Next
    With ListView1
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .View = lvwReport
        .ColumnHeaders.Add , , "Filepath"
        .multiSelect = False    '@TODO multi drag drop
    End With

    With aListView.Init(ListView1)
        .AutofitColumns
        .EnableDropFilesFolders True, False, False, "*"
        .EnableDragSort
    End With


End Sub

