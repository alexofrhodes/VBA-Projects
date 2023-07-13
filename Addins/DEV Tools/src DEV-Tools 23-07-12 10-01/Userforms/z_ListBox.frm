VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} z_ListBox 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "z_ListBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "z_ListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private al As aListBox

Private Sub UserForm_Initialize()
    With ListBox1
        .ColumnCount = 10
        Dim x As Long, y As Long
        Dim var(1 To 10, 1 To 10)
        For x = 1 To 10
            For y = 1 To 10
                var(x, y) = x * y
            Next
        Next
        .List = var
        .multiSelect = fmMultiSelectExtended
    End With
    aListBox.Init ListBox1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set al = Nothing
End Sub
