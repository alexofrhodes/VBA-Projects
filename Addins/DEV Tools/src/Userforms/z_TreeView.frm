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

Private AT As aTreeView
Private Sub Label1_Click()
    dp AT.ToArray
End Sub

Private Sub Label2_Click()
    AT.LoadRange ThisWorkbook.Sheets("TV_Data").Range("A1"), True, True
End Sub

Private Sub Label3_Click()
    dp AT.TreeviewArrayPaths
End Sub

Private Sub Label4_Click()
    AT.LoadVBProjects
End Sub

Private Sub Label5_Click()
    AT.Clear
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    AT.ChildrenCheck Node, Node.Checked
End Sub

Private Sub UserForm_Initialize()
    Set AT = New aTreeView
    AT.Init TreeView1
    TreeView1.CheckBoxes = True
End Sub

