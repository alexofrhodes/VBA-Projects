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
    '@AssignedModule z_ColorSchemes
    '@INCLUDE CLASS aFrame
    '@INCLUDE CLASS aColorScheme
    '@INCLUDE USERFORM z_ColorSchemes
    '@INCLUDE DECLARATION Ctrl
    aFrame.Init(fColorSchemes).AddThemeControls


    Dim ctrl        As MSForms.control
    On Error Resume Next
    For Each ctrl In Me.Controls
        ctrl.Font.Name = "Consolas"
        ctrl.Font.Size = 9
        ctrl.Font.Bold = True
    Next
    On Error GoTo 0

End Sub
