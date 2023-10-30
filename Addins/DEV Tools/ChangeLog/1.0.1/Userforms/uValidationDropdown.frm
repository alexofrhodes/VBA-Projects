VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uValidationDropdown 
   Caption         =   "DropDown Validation"
   ClientHeight    =   6432
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4188
   OleObjectBlob   =   "uValidationDropdown.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uValidationDropdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'Private Sub CommandButton1_Click()
'    If Me.Height > 224 Then
'        Me.Height = 175
'        CommandButton1.Caption = "V"
'    Else
'        Me.Height = 225
'        CommandButton1.Caption = "Ë"
'    End If
'End Sub

Private Sub GetInfo_Click()
    '@AssignedModule uValidationDropdown
    '@INCLUDE PROCEDURE PlayTheSound
    '@INCLUDE CLASS aUserform
    '@INCLUDE USERFORM uAuthor
    '@INCLUDE USERFORM uValidationDropdown
    With aUserform.Init(Me)
        .Transition .Effect(GetInfo, "Top", GetInfo.Top - 10, 150)
        .Transition .Effect(GetInfo, "Top", GetInfo.Top + 10, 150)
    End With
    uAuthor.Show
End Sub


Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '@AssignedModule uValidationDropdown
    '@INCLUDE USERFORM uValidationDropdown
    Dim lngValType  As Long
    On Error Resume Next
    lngValType = ActiveCell.Validation.Type
    On Error GoTo 0
    If lngValType = 3 Then updateData
End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '    If KeyCode = 13 Then
    '        updateData
    '    End If
    '    UVDRefreshList
    '@AssignedModule uValidationDropdown
    '@INCLUDE USERFORM uValidationDropdown
End Sub

Private Sub TextBox1_Change()
    '@AssignedModule uValidationDropdown
    '@INCLUDE USERFORM uValidationDropdown
    UVDRefreshList
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '@AssignedModule uValidationDropdown
    '@INCLUDE USERFORM uValidationDropdown
    If KeyCode = 13 Then
        updateData
    End If
End Sub

Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    '@AssignedModule uValidationDropdown
    '@INCLUDE USERFORM uValidationDropdown
    Select Case KeyCode
        Case 27    'ESC
            Unload Me
    End Select
End Sub

Private Sub UserForm_Activate()
    '@AssignedModule uValidationDropdown
    '@INCLUDE USERFORM uValidationDropdown
    UVDRefreshList
End Sub

Sub updateData()
    '@INCLUDE UVDRefreshList
    '@AssignedModule uValidationDropdown
    '@INCLUDE USERFORM uValidationDropdown
    ActiveCell.Value = Me.ListBox1.Value
    If Me.oClose.Value = True Then
        Me.TextBox1.Value = ""
        Unload Me
    ElseIf Me.oDown.Value = True Then
        ActiveCell.offset(1, 0).Select
        If Me.TextBox1.Value <> "" Then
            Me.TextBox1.Value = ""
        Else
            UVDRefreshList
        End If
        If Me.ListBox1.ListCount = 0 Then Unload Me
    ElseIf Me.oRight.Value = True Then
        ActiveCell.offset(0, 1).Select
        If Me.TextBox1.Value <> "" Then
            Me.TextBox1.Value = ""
        Else
            UVDRefreshList
        End If
        If Me.ListBox1.ListCount = 0 Then Unload Me
    ElseIf Me.oNone.Value = True Then
        Me.TextBox1.Value = ""
    End If
End Sub

Sub UVDRefreshList()
    '@INCLUDE ValidRange
    '@INCLUDE isValidation
    '@AssignedModule uValidationDropdown
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM uValidationDropdown
    Me.ListBox1.clear
    Dim rng         As Range
    Dim cell        As Range
    If isValidation(ActiveCell) = False Then Exit Sub
    If ValidRange(ActiveCell.Validation.Formula1) = True Then
        Set rng = Range(Replace(ActiveCell.Validation.Formula1, "=", ""))
        For Each cell In rng
            If Me.TextBox1.Value = "" Then
                Me.ListBox1.AddItem cell.Value
            Else
                If InStr(1, cell.Value, Me.TextBox1.Value, vbTextCompare) > 0 Then
                    Me.ListBox1.AddItem cell.Value
                End If
            End If
        Next
    ElseIf InStr(1, ActiveCell.Validation.Formula1, ",") > 0 Then
        Dim var
        var = Split(ActiveCell.Validation.Formula1, ",")
        Dim element
        For Each element In var
            Me.ListBox1.AddItem element
        Next
    Else
        Unload Me
    End If

    aListBox.Init(ListBox1).FilterByColumn TextBox1.TEXT

    On Error Resume Next
    Me.ListBox1.ListIndex = 0
    On Error GoTo 0

End Sub

Function ValidRange(str As String) As Boolean
    '@AssignedModule uValidationDropdown
    '@INCLUDE USERFORM uValidationDropdown
    Dim rng         As Range
    On Error Resume Next
    Set rng = Range(Replace(str, "=", ""))
    On Error GoTo 0
    If rng Is Nothing Then
        ValidRange = False
    Else
        ValidRange = True
    End If
End Function

Function isValidation(rng As Range) As Boolean
    '@AssignedModule uValidationDropdown
    '@INCLUDE USERFORM uValidationDropdown
    Dim dvtype      As Integer
    On Error Resume Next
    dvtype = rng.Validation.Type
    On Error GoTo 0
    If dvtype = 3 Then
        isValidation = True
    Else
        isValidation = False
    End If
End Function

Private Sub UserForm_Initialize()
    '@INCLUDE LoadUserformOptions
    '@AssignedModule uValidationDropdown
    '@INCLUDE CLASS aUserform
    '@INCLUDE USERFORM uValidationDropdown
    aUserform.Init(Me).LoadOptions
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '@INCLUDE SaveUserformOptions
    '@AssignedModule uValidationDropdown
    '@INCLUDE CLASS aUserform
    '@INCLUDE USERFORM uValidationDropdown
    aUserform.Init(Me).SaveOptions , , False, False
End Sub

