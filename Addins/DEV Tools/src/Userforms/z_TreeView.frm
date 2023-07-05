VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} z_TreeView 
   Caption         =   "UserForm1"
   ClientHeight    =   6600
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7056
   OleObjectBlob   =   "z_TreeView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "z_TreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With aTreeView.Init(TreeView1)
        .ApplyStandardStyle
        .LoadRange ThisWorkbook.Sheets("TV_Data").Range("A1"), True, True
    End With
End Sub
