VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} z_MutiPage 
   Caption         =   "UserForm1"
   ClientHeight    =   9312.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13512
   OleObjectBlob   =   "z_MutiPage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "z_MutiPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    SidebarRight.Visible = True
    SidebarBottom.Visible = True
    aMultiPage.Init(MultiPage1).BuildMenu
    
'    aColorScheme.Init(me).ThemeBlackAndGreenDark
End Sub
