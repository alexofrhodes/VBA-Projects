VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uMultiPageMenu 
   Caption         =   "UserForm1"
   ClientHeight    =   5736
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9036.001
   OleObjectBlob   =   "uMultiPageMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uMultiPageMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    aMultiPage.Init(MultiPage1).BuildMenu
End Sub
