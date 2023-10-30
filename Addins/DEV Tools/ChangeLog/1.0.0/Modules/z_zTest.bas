Attribute VB_Name = "z_zTest"

Option Explicit


Function tm2(x As Long, y As Long) As Long

End Function

Function tm3() As Collection

End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 18-08-2023 13:12    Alex                (z_zTest.bas > GotoSelectedShapeOnAction)

Sub GotoSelectedShapeOnAction()
'@LastModified 2308181312
'@Description Something
'@INCLUDE PROCEDURE GotoShapeOnaction
    Dim shapeCount  As Long
    On Error Resume Next
    shapeCount = Selection.ShapeRange.Count
    On Error GoTo 0
    If shapeCount <> 1 Then Exit Sub
    Dim shp         As Shape
    Set shp = ActiveSheet.Shapes(Selection.Name)
    GotoShapeOnaction ActiveSheet.Shapes(Selection.Name)
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 18-08-2023 13:12    Alex                (z_zTest.bas > GotoShapeOnaction)

Sub GotoShapeOnaction(shp As Shape)
'@LastModified 2308181312
'@INCLUDE CLASS aProcedure
    Dim Procedure   As String
    Procedure = shp.OnAction
    If Procedure = "" Then Exit Sub
    On Error GoTo ErrorHandler
    aProcedure.Init(ActiveWorkbook, , Procedure).Activate
    Exit Sub
ErrorHandler:
    MsgBox Procedure & " not found"
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 21-08-2023 13:59    Alex                (z_zTest.bas > ShowInNotepad)

Sub ShowInNotepad(txt As String)
'@LastModified 2308211359
'@INCLUDE PROCEDURE FollowLink
'@INCLUDE PROCEDURE TxtOverwrite
    Dim TargetFile  As String
    TargetFile = ThisWorkbook.path & "\tmp.txt"
    TxtOverwrite TargetFile, txt
    FollowLink TargetFile
End Sub
