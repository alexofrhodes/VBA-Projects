VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} zTest 
   Caption         =   "UserForm1"
   ClientHeight    =   6456
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9732.001
   OleObjectBlob   =   "zTest.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "zTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
'    ListView1.ListItems.Add , , "Test"

'    aUserform.Init(Me).LoadPosition
    aMultiPage.Init(MultiPage1).BuildMenu
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    aUserform.Init(Me).SavePosition
End Sub
