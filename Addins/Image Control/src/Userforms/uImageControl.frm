VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uImageControl 
   ClientHeight    =   5976
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6036
   OleObjectBlob   =   "uImageControl.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uImageControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private mainPath As String
Private Shrink As Double
Private eTime As Variant

Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1


Private Sub UserForm_Activate()
    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll Me
End Sub

Private Sub UserForm_Initialize()
Dim anc As MSForms.control

    For Each c In Me.Controls
        If TypeName(c) = "Frame" Then
            'c.Caption = ""
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                c.Visible = False
                If InStr(1, c.Tag, "anchor") > 0 Then
                    On Error Resume Next
                    Set anc = Me.Controls("Anchor" & Mid(c.Tag, InStr(1, c.Tag, "Anchor", vbTextCompare) + Len("Anchor"), 2))
                    If anc Is Nothing Then Stop
                    On Error GoTo 0
                    c.Top = anc.Top 'Anchor01.Top
                    c.left = anc.left ' Anchor01.Left
                    Set anc = Nothing
                End If
            End If
        End If
    Next
    
    Dim mainPath As String
    mainPath = Environ$("USERPROFILE") & "\My Documents\vbArc\ExportedImages\"
    FoldersCreate mainPath
    ComboBox1.list = Array("GIF", "JPG", "ICO", "BMP", "CUR", "WMF")
    ComboBox1.ListIndex = 0
    Shrink = CInt(lbShrink.Caption) * 0.1
    
    Me.Height = 259
    Me.Width = 166
End Sub

Private Sub bExportRange_Click()
    ExportAsImage
End Sub

Private Sub bFitToText_Click()
    TextBoxResizeTB
End Sub

Private Sub bExportShapes_Click()
    ExportShapeAsPicture
End Sub

Private Sub bSelectByRange_Click()
    SelectShapesWithinSelectedRange
End Sub

Private Sub bOB_Click()
    ShapesOutsideVisibleRange
End Sub

Private Sub bSelectByName_Click()
    SelectShapesByName
End Sub

Private Sub bInsertToRange_Click()
    InsertPictures
End Sub

Private Sub bInsertToComment_Click()
    InsertImageInActivecellComment
End Sub



Private Sub bPastePicture_Click()
    PasteAsPicture
End Sub

Private Sub bPasteLinked_Click()
    PasteAsLinkedPicture
End Sub

Private Sub bAlignHorizontal_Click()
    GridHorizontal
End Sub

Private Sub bAlignVertical_Click()
    GridVertical
End Sub



Private Sub bFitCell_Click()
    PicturesFitCenter
End Sub

Private Sub bSelectByText_Click()
    SelectShapesByText
End Sub

Private Sub Emitter_LabelMouseOut(Label As MSForms.Label)
    If InStr(1, Label.Tag, "reframe", vbTextCompare) > 0 Then
        If Label.BackColor <> &H80B91E Then Label.BackColor = &H534848
    End If
End Sub

Private Sub Emitter_LabelMouseOver(Label As MSForms.Label)
    If InStr(1, Label.Tag, "reframe", vbTextCompare) > 0 Then
        If Label.BackColor <> &H80B91E Then Label.BackColor = &H808080
    End If
End Sub

Sub Emitter_LabelClick(ByRef Label As MSForms.Label)
    If InStr(1, Label.Tag, "reframe", vbTextCompare) > 0 Then Reframe Label
End Sub


Private Sub Reframe(control As MSForms.control)
    For Each c In Me.Controls
        If TypeName(c) = "Frame" Then
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                If c.Name <> control.Parent.Parent.Name Then c.Visible = False
            End If
        End If
    Next
    Me.Controls(control.Caption).Visible = True
    For Each c In Me.Controls
        If TypeName(c) = "Label" Then
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                c.BackColor = &H534848
                'c.SpecialEffect = fmSpecialEffectFlat
            End If
        End If
    Next
    control.BackColor = &H80B91E
    'Control.SpecialEffect = fmSpecialEffectRaised

End Sub

Private Sub iFolder_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    FollowLink mainPath
End Sub

Private Sub Image2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    uDEV.Show
End Sub
Sub FollowLink(FolderPath As String)
        Dim oShell As Object
        Dim Wnd As Object
        Set oShell = CreateObject("Shell.Application")
        For Each Wnd In oShell.Windows
            If Wnd.Name = "File Explorer" Then
               If Wnd.Document.Folder.Self.Path = FolderPath Then Exit Sub
            End If
        Next Wnd
        Application.ThisWorkbook.FollowHyperlink Address:=FolderPath, NewWindow:=True
End Sub


Private Sub SpinButton1_Change()
    lbShrink.Caption = SpinButton1.Value
    Shrink = CInt(lbShrink.Caption) * 0.1
End Sub

Sub FoldersCreate(FolderPath As String)
    'Create all the folders in a folder path
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim arrayElement As Variant
    'Split the folder path into individual folder names
    individualFolders = Split(FolderPath, "\")
    'Loop though each individual folder name
    On Error Resume Next
    For Each arrayElement In individualFolders
        'Build string of folder path
        tempFolderPath = tempFolderPath & arrayElement & "\"
        'If folder does not exist, then create it
        If Dir(tempFolderPath, vbDirectory) = "" Then
            MkDir tempFolderPath
        End If
    Next arrayElement
End Sub
Sub GridVertical()
    Dim shp As Shape
    Dim lCnt As Long
    Dim dTop As Double
    Dim dLeft As Double
    Dim dHeight As Double
    Dim dWidth As Double
    Dim dSPACE As Variant
    Dim lRowCnt As Variant
    Dim dStart As Double
    Dim dMaxHeight As Double
    If TypeName(Selection) = "Range" Then
        MsgBox "Please select shapes before running the macro."
        Exit Sub
    End If
    lRowCnt = Application.InputBox("Enter the number of columns for the vertical shape grid.", "Vertical Shape Grid", Type:=1)
    If lRowCnt <= 0 Or lRowCnt = False Then
        Exit Sub
    End If
    dSPACE = Application.InputBox("Enter the space between shapes in points.", "Vertical Shape Grid", Type:=1)
    If TypeName(dSPACE) = "Boolean" Then
        Exit Sub
    End If
    lCnt = 1
    For Each shp In Selection.ShapeRange
        With shp
            If lCnt = 1 Then
                dStart = .left
            Else
                If lCnt Mod lRowCnt = 1 Or lRowCnt = 1 Then
                    .Top = dTop + dMaxHeight + dSPACE
                    .left = dStart
                    dMaxHeight = .Height
                Else
                    .Top = dTop
                    .left = dLeft + dWidth + dSPACE
                End If
            End If
            dTop = .Top
            dLeft = .left
            dHeight = .Height
            dWidth = .Width
            dMaxHeight = WorksheetFunction.Max(dMaxHeight, .Height)
        End With
        lCnt = lCnt + 1
    Next shp
End Sub

Sub GridHorizontal()
    Dim shp As Shape
    Dim lCnt As Long
    Dim dTop As Double
    Dim dLeft As Double
    Dim dHeight As Double
    Dim dWidth As Double
    Dim dSPACE As Variant
    Dim lColCnt As Variant
    Dim lCol As Long
    Dim dStart As Double
    Dim lRow As Double
    Dim dMaxWidth As Double
    If TypeName(Selection) = "Range" Then
        MsgBox "Please select shapes before running the macro."
        Exit Sub
    End If
    lColCnt = Application.InputBox("Enter the number of rows for the horizontal shape grid.", "Horizontal Shape Grid", Type:=1)
    If lColCnt <= 0 Or lColCnt = False Then
        Exit Sub
    End If
    dSPACE = Application.InputBox("Enter the space between shapes in points.", "Horizontal Shape Grid", Type:=1)
    If TypeName(dSPACE) = "Boolean" Then
        Exit Sub
    End If
    lCnt = 1
    For Each shp In Selection.ShapeRange
        With shp
            If lCnt = 1 Then
                dStart = .Top
            Else
                If lCnt Mod lColCnt = 1 Or lColCnt = 1 Then
                    .Top = dStart
                    .left = dLeft + dMaxWidth + dSPACE
                    dMaxWidth = .Width
                Else
                    .Top = dTop + dHeight + dSPACE
                    .left = dLeft
                End If
            End If
            dTop = .Top
            dLeft = .left
            dHeight = .Height
            dWidth = .Width
            dMaxWidth = WorksheetFunction.Max(dMaxWidth, .Width)
        End With
        lCnt = lCnt + 1
    Next shp
End Sub


Sub ExportShapeAsPicture()
    Dim cht As ChartObject
    Dim ActiveShape As Shape
    Dim EXT As String
    EXT = uImageControl.ComboBox1.Text
    If TypeName(Selection) = "Range" Then GoTo NoShapeSelected
    For Each ActiveShape In Selection.ShapeRange
        Set cht = ActiveSheet.ChartObjects.Add( _
                  left:=ActiveCell.left, _
                  Width:=ActiveShape.Width, _
                  Top:=ActiveCell.Top, _
                  Height:=ActiveShape.Height)
        cht.ShapeRange.Fill.Visible = msoFalse
        cht.ShapeRange.line.Visible = msoFalse
        ActiveShape.Copy
        cht.Activate
        ActiveChart.Paste
        cht.Chart.Export mainPath & ActiveShape.Name & "." & EXT
        cht.Delete
        ActiveShape.Select
    Next ActiveShape
    Exit Sub
NoShapeSelected:
    MsgBox "Please select shapes before running the macro."
    Exit Sub
End Sub

Sub ExportAsImage()
    If Not TypeName(Selection) = "Range" Then
        MsgBox "Please select shapes before running the macro."
        Exit Sub
    End If
    Dim cell As Range
    Dim EXT As String
    EXT = uImageControl.ComboBox1.Text
    Dim Action As Long
    Action = MsgBox("(YES) = for each area in selection" & Chr(10) & _
                    "(NO) = for each cell in selection", vbYesNoCancel)
    If Action = vbCancel Then Exit Sub
On Error Resume Next
    Application.DisplayAlerts = False
    Select Case Action
        Case Is = vbNo
            For Each cell In Selection
                Call ExportRangeAsImage(ActiveSheet, cell, mainPath, cell.Value, EXT)
                Application.Wait (Now + TimeValue("0:00:01"))
            Next cell
        Case Is = vbYes
            Dim result As String
            For i = 1 To Selection.Areas.Count
                result = ""
                result = InputBox("name for image of area: " & Selection.Areas(i).Address)
                If CStr(result) = "" Then result = Format(Now, "hhmmss")
                Call ExportRangeAsImage(ActiveSheet, Selection.Areas(i), mainPath, result, EXT)
                Application.Wait (Now + TimeValue("0:00:01"))
            Next i
    End Select
    Application.DisplayAlerts = True
    Shell "explorer.exe" & " " & mainPath, vbNormalFocus
End Sub
' Procedure : ExportRangeAsImage
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Capture a picture of a worksheet range and save it to disk
'               Returns True if the operation is successful
' Note      : *** Overwrites files, if already exists, without any warning! ***
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: Uses Late Binding, so none required
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' ws            : Worksheet to capture the image of the range from
' rng           : Range to capture an image of
' sPath         : Fully qualified path where to export the image to
' sFileName     : filename to save the image to WITHOUT the extension, just the name
' sImgExtension : The image file extension, commonly: JPG, GIF, PNG, BMP
'                   If omitted will be JPG format
'
' Usage:
' ~~~~~~
' ? ExportRangeAsImage(Sheets("Sheet1"), Range("A1"), "C:\Temp\Charts\", "test01". "JPG")
' ? ExportRangeAsImage(Sheets("Products"), Range("D5:F23"), "C:\Temp\Charts", "test02")
' ? ExportRangeAsImage(Sheets("Sheet1"), Range("A1"), "C:\Temp\Charts\", "test01", "PNG")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2020-04-06              Initial Release
'---------------------------------------------------------------------------------------
Function ExportRangeAsImage(ws As Worksheet, _
                            rng As Range, _
                            sPath As String, _
                            sFileName As String, _
                            Optional sImgExtension As String = "JPG") As Boolean
    Dim oChart                As ChartObject
    On Error GoTo Error_Handler
    If right(sPath, 1) <> "\" Then sPath = sPath & "\"
    Application.ScreenUpdating = False
    ws.Activate
    rng.CopyPicture xlScreen, xlPicture          'Copy Range Content
    Set oChart = ws.ChartObjects.Add(0, 0, rng.Width, rng.Height) 'Add chart
    oChart.Activate
    With oChart.Chart
        .Paste                                   'Paste our Range
        .Export sPath & sFileName & "." & LCase(sImgExtension), sImgExtension 'Export the chart as an image
    End With
    oChart.Delete                                'Delete the chart
    ExportRangeAsImage = True
Error_Handler_Exit:
    On Error Resume Next
    Application.ScreenUpdating = True
    If Not oChart Is Nothing Then Set oChart = Nothing
    Exit Function
Error_Handler:
    '76 - Path not found
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: ExportRangeAsImage" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Sub InsertPictures()
    Dim PicList() As Variant
    Dim PicFormat As String
    Dim rng As Range
    Dim sShape As Shape
    On Error Resume Next
    PicList = Application.GetOpenFilename(PicFormat, multiSelect:=True)
    xColIndex = Application.ActiveCell.Column
    If IsArray(PicList) Then
        xRowIndex = Application.ActiveCell.Row
        For lLoop = LBound(PicList) To UBound(PicList)
            Set rng = Cells(xRowIndex, xColIndex)
            Set sShape = ActiveSheet.Shapes.AddPicture(PicList(lLoop), msoFalse, msoCTrue, _
                                                       rng.left, rng.Top, -1, -1)
            With sShape
.LockAspectRatio = msoTrue
                If .Height > .Width Then
                    .Height = rng.Height - (rng.Height * Shrink)
                Else
                    .Width = rng.Width - (rng.Width * Shrink)
                End If
                .Top = rng.MergeArea.Top + (rng.MergeArea.Height - .Height) / 2
                .left = rng.MergeArea.left + (rng.MergeArea.Width - .Width) / 2
            End With
            xRowIndex = xRowIndex + 1
        Next
    End If
End Sub

Sub InsertImageInActivecellComment()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell before running the macro."
        Exit Sub
    End If
    Dim cell As Range
    Dim cmt As Comment
    Dim PicPath As String
    Dim str As String
    Dim myObj As Object
    Dim myDirString As String
    Set myObj = Application.FileDialog(msoFileDialogFilePicker)
    With myObj
        .InitialFileName = "C:\Users\" & Environ$("Username") & "\Pictures"
        .Filters.Add "Images", "*.png, *jpeg, *.jpg, *.gif, *.ico, *.cur, *.wmf"
        .FilterIndex = 2
        If .Show = False Then MsgBox "No picture selected", vbExclamation: Exit Sub
        PicPath = .SelectedItems(1)
    End With
    On Error Resume Next
    Set cell = Selection.MergeArea
    With cell
        If .Comment Is Nothing Then
            Set cmt = .AddComment
            str = cmt.Text
        Else
            Set cmt = .Comment
            str = cmt.Text
        End If
    End With
    With cmt
        .Text ((Replace(str, Application.UserName & ":", "")))
        .Shape.Fill.UserPicture PicPath
        .Visible = False
    End With
End Sub

Sub ScreenRefresh()
    Dim s As Shape
    For Each s In Workbooks("").Sheets("Sheet1")
s.Top = ThisWorkbook.Windows(1).VisibleRange.Top
    Next s
End Sub

Sub StartTimedRefresh()
    Call ScreenRefresh
    eTime = Now + TimeValue("00:00:01")
    Application.OnTime eTime, "StartTimedRefresh"
End Sub

Sub StopTimer()
    Application.OnTime eTime, "StartTimedRefresh", , False
End Sub

Sub TextBoxResizeTB()
    Dim xShape As Shape
    Dim xSht As Worksheet
    On Error Resume Next
    For Each xSht In ActiveWorkbook.Worksheets
        For Each xShape In xSht.Shapes
            xShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
            xShape.TextFrame2.WordWrap = False
        Next
    Next
End Sub

Sub PicturesFitCenter()
    If TypeName(Selection) = "Range" Then
        MsgBox "Please select shapes before running the macro."
        Exit Sub
    End If
    Dim ans As Long
    ans = MsgBox("Lock Aspect Ratio?", vbYesNoCancel)
    If ans = vbCancel Then Exit Sub
    Dim p As Shape
For Each p In Selection.ShapeRange
        Dim cell As Range: Set cell = Cells(p.TopLeftCell.Row, p.TopLeftCell.Column)
        With p
            If ans = vbYes Then
                .LockAspectRatio = True
                If .Height > .Width Then
                    .Height = cell.Height - (cell.Height * Shrink)
                Else
                    .Width = cell.Width - (cell.Width * Shrink)
                End If
            Else
                .LockAspectRatio = False
                If .Height > .Width Then
                    .Width = cell.Width - (cell.Width * Shrink)
                    .Height = cell.Height - (cell.Height * Shrink)
                Else
                    .Height = cell.Height - (cell.Height * Shrink)
                    .Width = cell.Width - (cell.Width * Shrink)
                End If
            End If
            .Top = cell.MergeArea.Top + (cell.MergeArea.Height - .Height) / 2
            .left = cell.MergeArea.left + (cell.MergeArea.Width - .Width) / 2
        End With
    Next
End Sub

Sub ShapesOutsideVisibleRange()
    If ActiveSheet.Shapes.Count = 0 Then
        MsgBox "No shapes in active sheet"
        Exit Sub
    End If
    Dim s As Shape
    Dim rngholder As String
    For Each s In ActiveSheet.Shapes
        If Range(s.BottomRightCell.Address).Row > ActiveWindow.VisibleRange.rows.Count Then
            rngholder = _
                      rngholder & Chr(10) & s.BottomRightCell.Address
        End If
    Next s
    If rngholder = "" Then
        MsgBox "No shape out of range"
        Exit Sub
    End If
    Dim arr
    arr = Split(rngholder, Chr(10))
    Dim lastSposition As String
    lastSposition = arr(UBound(arr))
    If Range(lastSposition).Row > ActiveWindow.VisibleRange.rows.Count Then
        MsgBox "There are shapes after the last visible row." _
             & Chr(10) & "Their BottomRight cells span the following ranges: " _
             & rngholder
    Else
        MsgBox "All shapes positioned inside visible range"
    End If
End Sub

Sub PasteAsPicture()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range before running the macro."
        Exit Sub
    End If
    For i = 1 To Selection.Areas.Count
        Application.CutCopyMode = False
        Selection.Areas(i).Copy
        ActiveSheet.Pictures.Paste
    Next
    Application.CutCopyMode = False
End Sub

Sub PasteAsLinkedPicture()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range before running the macro."
        Exit Sub
    End If
    Dim coll As New Collection
    For i = 1 To Selection.Areas.Count
        coll.Add Selection.Areas(i).Address
    Next
    Dim element As Variant
    Range(coll(1)).Select
    For Each element In coll
        Application.CutCopyMode = False
        Range(element).Copy
        ActiveSheet.Pictures.Paste Link:=True
    Next
    Application.CutCopyMode = False
End Sub

Sub SelectShapesWithinSelectedRange()
    On Error Resume Next
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range first"
        Exit Sub
    End If
    Dim shp As Shape
    Dim R As Range
    Set R = Selection
    For Each shp In ActiveSheet.Shapes
        If Not Intersect(Range(shp.TopLeftCell, shp.BottomRightCell), R) Is Nothing Then
            shp.Select Replace:=False
        End If
    Next shp
End Sub

Sub SelectShapesByName()
    On Error Resume Next
    Dim shp As Shape
    ActiveSheet.Range("A1").Select
    Dim str As String
    str = InputBox("contains in NAME?")
    For Each shp In ActiveSheet.Shapes
        If InStr(shp.Name, str) Then
            shp.Select Replace:=False
        End If
    Next shp
End Sub

Sub SelectShapesByText()
    Dim shp As Shape
    Dim str As String
    str = InputBox("contains in TEXT?")
    ActiveSheet.Range("A1").Select
    On Error GoTo nxt
For Each shp In ActiveWorkbook.ActiveSheet.Shapes
        If shp.Type <> 13 Then
            With shp.TextFrame.Characters
                If InStr(1, .Text, str) Then
                    shp.Select Replace:=False
                End If
            End With
        End If
nxt:
    Next shp
End Sub


Private Sub ResizeUserformToFitControls(FORM As Object)
    FORM.Width = 0
    FORM.Height = 0
    Dim ctr As MSForms.control
    Dim myWidth
    myWidth = FORM.InsideWidth
    For Each ctr In FORM.Controls
        If ctr.Visible = True Then
            If ctr.left + ctr.Width > myWidth Then myWidth = ctr.left + ctr.Width
        End If
    Next
    FORM.Width = myWidth + FORM.Width - FORM.InsideWidth + 10
    Dim myHeight
    myHeight = FORM.InsideHeight
    For Each ctr In FORM.Controls
        If ctr.Visible = True Then
            If ctr.Top + ctr.Height > myHeight Then myHeight = ctr.Top + ctr.Height
        End If
    Next
    FORM.Height = myHeight + (FORM.Height - FORM.InsideHeight) + 10
End Sub

