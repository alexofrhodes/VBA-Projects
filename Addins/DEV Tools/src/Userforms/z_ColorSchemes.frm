VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} z_ColorSchemes 
   Caption         =   "UserForm1"
   ClientHeight    =   10008
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9612.001
   OleObjectBlob   =   "z_ColorSchemes.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "z_ColorSchemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
'    aColorScheme.Init(Me.Frame1).ThemeBlackAndGreenDark
'    aColorScheme.Init(Me).ThemeBlackAndGreenDark
    aFrame.Init(fColorSchemes).AddThemeControls
    
    
    Dim Ctrl As MSForms.control
    On Error Resume Next
        For Each Ctrl In Me.Controls
            Ctrl.Font.Name = "Consolas"
            Ctrl.Font.Size = 9
            Ctrl.Font.Bold = True
        Next
    On Error GoTo 0
    
End Sub
