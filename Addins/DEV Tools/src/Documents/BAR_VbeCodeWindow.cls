VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
'@AssignedModule Sheet34
'@INCLUDE PROCEDURE BarExists
    If Selection.Cells.count > 1 Then Exit Sub
    If BarExists("testtag") Then
        CommandBars("testtag").ShowPopup
        Cancel = True
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'@AssignedModule Sheet34
'@INCLUDE DECLARATION rBAR_LOCATION
'@INCLUDE DECLARATION rMENU_TYPE
Application.EnableEvents = False

    If Replace(Target.Address, "$", "") = rMENU_TYPE Then
        Range(rBAR_LOCATION).ClearContents
    End If
    Application.EnableEvents = True
End Sub

Public Sub Worksheet_SelectionChange(ByVal Target As Range)
'@AssignedModule Sheet34
'@INCLUDE PROCEDURE openUValiationDropdown
On Error GoTo EH
    If Selection.Cells.count > 1 And Cells(Selection.row, 1) <> "" Then Exit Sub
    Calculate
    openUValiationDropdown
EH:
End Sub

