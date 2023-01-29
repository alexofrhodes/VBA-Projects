VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uAddinManager 
   Caption         =   "ADDINS MANAGER"
   ClientHeight    =   7584
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4896
   OleObjectBlob   =   "uAddinManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uAddinManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    For i = 0 To body.ListCount - 1
        If body.Selected(i) = True Then AddIns(body.List(i, 0)).Installed = Not AddIns(body.List(i, 0)).Installed
    Next
    LoadAddins
End Sub

Private Sub CommandButton2_Click()
    Dim ans As Long
    ans = MsgBox("Irreversible. Proceed?", vbYesNo)
    If ans = vbNo Then Exit Sub
    AddIns(body.List(body.ListIndex, 0)).Installed = False
    Kill AddIns(body.List(body.ListIndex, 0)).FullName
End Sub

Private Sub CommandButton3_Click()
    SortListboxOnColumn body, 0
End Sub

Private Sub CommandButton4_Click()
    SortListboxOnColumn body, 1
End Sub

Private Sub CommandButton5_Click()
    vbArcAddinsForm.Show
    Unload Me
End Sub

Private Sub info_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Admin_AddinsModified
End Sub

Private Sub UserForm_Initialize()
    LoadAddins
End Sub

Private Sub LoadAddins()
'#IMPORTS SortListboxOnColumn
    body.Clear
    Dim ad As AddIn
    On Error Resume Next
    For Each ad In AddIns
        body.AddItem
        body.List(body.ListCount - 1, 0) = Left(ad.Name, InStr(1, ad.Name, ".") - 1)
        body.List(body.ListCount - 1, 1) = IIf(ad.Installed, " ENABLED", "-")
    Next
    SortListboxOnColumn body, 1
End Sub

Sub SortListboxOnColumn(lBox As MSForms.ListBox, Optional OnColumn As Long = 0)
    Dim vntData As Variant
    Dim vntTempItem As Variant
    Dim lngOuterIndex As Long
    Dim lngInnerIndex As Long
    Dim lngSubItemIndex As Long
    'Store the list in an array for sorting
    vntData = lBox.List
    'Bubble sort the array on the first value
    For lngOuterIndex = LBound(vntData, 1) To UBound(vntData, 1) - 1
        For lngInnerIndex = lngOuterIndex + 1 To UBound(vntData, 1)
            If vntData(lngOuterIndex, OnColumn) > vntData(lngInnerIndex, OnColumn) Then
                'Swap values
                For lngSubItemIndex = 0 To lBox.ColumnCount - 1
                    vntTempItem = vntData(lngOuterIndex, lngSubItemIndex)
                    vntData(lngOuterIndex, lngSubItemIndex) = vntData(lngInnerIndex, lngSubItemIndex)
                    vntData(lngInnerIndex, lngSubItemIndex) = vntTempItem
                Next
            End If
        Next lngInnerIndex
    Next lngOuterIndex
    'Remove the contents of the listbox
    lBox.Clear
    'Repopulate with the sorted list
    lBox.List = vntData
End Sub

