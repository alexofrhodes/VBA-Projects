VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} z_MultiPage 
   Caption         =   "UserForm1"
   ClientHeight    =   9312.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13512
   OleObjectBlob   =   "z_MultiPage.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "z_MultiPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub UserForm_Initialize()
    '@AssignedModule z_MutiPage
    '@INCLUDE CLASS aMultiPage
    '@INCLUDE CLASS aColorScheme
    '@INCLUDE USERFORM z_MutiPage
    SidebarRight.Visible = False
    SidebarBottom.Visible = False

    '// the class is predeclaredId = true but shouldn't this way still work?
    '    Dim am1 As aMultiPage
    '    Set am1 = New aMultiPage
    '    am1.Init(MultiPage1).BuildMenu createSidebarMinimizers:=True

    aMultiPage.Init(MultiPage1).BuildMenu createSidebarMinimizers:=True

End Sub

