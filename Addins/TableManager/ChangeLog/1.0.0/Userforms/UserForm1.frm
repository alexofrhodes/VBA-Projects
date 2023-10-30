VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4176
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoadedList As Variant
Private previousInputLength As Long
Public TargetControl As MSForms.control

Private Sub ListBox1_Enter()
    If ListBox1.ListCount > 0 Then
        If ListBox1.ListIndex = -1 Then ListBox1.ListIndex = 0
    End If
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    TargetControl.Value = ListBox1.List(ListBox1.ListIndex)
    uTableManager.Show
    Unload Me
End Sub

Private Sub ListBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            If ListBox1.ListIndex <> -1 Then TargetControl.Value = ListBox1.List(ListBox1.ListIndex)
            uTableManager.Show
            Unload Me
        Case Is = 27
            ListBox1.ListIndex = -1
            TextBox1.SetFocus
    End Select
End Sub

Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 27
            If TextBox1.Value <> "" Then
                TextBox1.Value = ""
            Else
                uTableManager.Show
                Unload Me
            End If
    End Select
End Sub

Private Sub TextBox1_Change()
    Select Case Len(TextBox1.Value)
    Case Is = 0 'show whole range.value
        ListBox1.List = LoadedList
    Case Is = 1 'filter the whole array
        ListBox1.List = ArrayFilterLike(LoadedList, TextBox1.Value, False)
    Case Is > 1
        If Len(TextBox1.Value) > previousInputLength Then
            ListBox1.List = ArrayFilterLike(ListBox1.List, TextBox1.Value, False)
        Else
            ListBox1.List = ArrayFilterLike(LoadedList, TextBox1.Value, False)
        End If
    End Select
    previousInputLength = Len(TextBox1.Value)
End Sub
