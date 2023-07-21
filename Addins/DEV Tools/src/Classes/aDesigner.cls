VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aDesigner
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 04-05-2023 12:18    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Option Explicit

Private oModule As VBComponent

Function Active() As aDesigner
    Set Active = Me
    Class_Initialize
End Function

Private Sub Class_Initialize()
    Set oModule = ActiveModule
    If oModule.Type <> vbext_ct_MSForm Then End
End Sub

Private Sub Class_Terminate()
'exit
End Sub

Function SelectedControl() As MSForms.control
    Dim col As New Collection
    Set col = SelectedControls
    If col.count = 1 Then
        Set SelectedControl = col(1)
    End If
End Function

Function SelectedControls() As Collection
    Dim ctl    As control
    Dim out As New Collection
    For Each ctl In oModule.Designer.Selected
        out.Add ctl
    Next ctl
    
    Dim ac As New aCollection
    Set out = ac.Init(out).Reverse.Items
    Set SelectedControls = out
End Function

Function SelectedFrameOrMultipageControl() As MSForms.control
    Dim col As New Collection
    Set col = SelectedFrameOrMultipageControls
    If col.count = 1 Then
        Set SelectedFrameOrMultipageControl = col(1)
    End If
End Function


Function SelectedFrameOrMultipageControls() As Collection
    Dim ctl       As MSForms.control
    Dim subctl    As MSForms.control
    Dim out As New Collection
    For Each ctl In SelectedControls
        If InStr("FrameMultiPage", TypeName(ctl)) > 0 Then
            For Each subctl In ctl.Controls
                If InStr("FrameMultiPage", TypeName(subctl)) = 0 And subctl.InSelection Then
                    out.Add subctl
                End If
            Next
        End If
    Next ctl
    Dim ac As New aCollection
    Set out = ac.Init(out).Reverse.Items
    Set SelectedFrameOrMultipageControls = out
End Function

'Function RealActiveControl(oForm As Variant) As MSForms.control
''Jaafar Tribak
'
'    Dim oControl As MSForms.control
'
'    On Error Resume Next
'    Set oControl = oForm.ActiveControl
'    Do
'        Set oControl = CallByName(oControl, "ActiveControl", VbGet)
'    Loop Until TypeName(oControl) <> "Frame"
'    Set RealActiveControl = oControl
'End Function

Sub RemoveCaption()
    Dim c As MSForms.control
    For Each c In SelectedControls
        c.Caption = ""
    Next
End Sub

Sub SetHandCursor()
    Dim ctr As MSForms.control, ctr2 As MSForms.control
    For Each ctr In SelectedControls
        AssignHandCursor ctr
    Next
End Sub

Sub SetHandCursorToSubControls()
    Dim ctr As MSForms.control, ctr2 As MSForms.control
    For Each ctr In SelectedFrameOrMultipageControls
        AssignHandCursor ctr
    Next
End Sub

Private Sub AssignHandCursor(control As MSForms.control)
    On Error GoTo Catch
    With control
        .MouseIcon = LoadPicture(ThisWorkbook.Path & "\img\Hand Cursor Pointer.ico")
        .MousePointer = fmMousePointerCustom
    End With
Catch:
End Sub

Sub SwitchNames()
    If SelectedControls.count <> 2 Then Exit Sub

    Dim ctrls As Collection
    Set ctrls = SelectedControls
    If ctrls.count <> 2 Then Exit Sub
    Dim tmp1 As String
    tmp1 = ctrls(1).Name
    Dim tmp2 As String
    tmp2 = ctrls(2).Name
    ctrls(1).Name = "tmp1"
    ctrls(2).Name = "tmp2"
    ctrls(1).Name = tmp2
    ctrls(2).Name = tmp1
End Sub

Sub SwitchPositions()
    If SelectedControls.count <> 2 Then
        MsgBox "Select exactly 2 controls to swirtch their position"
        Exit Sub
    End If
    Dim ctrls As Collection
    Set ctrls = SelectedControls
    If ctrls.count <> 2 Then Exit Sub
    Dim left1 As Long, left2 As Long
    Dim top1 As Long, top2 As Long
    left1 = ctrls(1).Left
    top1 = ctrls(1).Top
    left2 = ctrls(2).Left
    top2 = ctrls(2).Top
    ctrls(1).Left = left2
    ctrls(1).Top = top2
    ctrls(2).Left = left1
    ctrls(2).Top = top1
End Sub


Sub ReplaceCommandButtonWithLabel()
    Dim ctr As MSForms.control, ctr2 As MSForms.control
    Dim lbl As MSForms.Label
    Dim ctrName As String
    Dim col As New Collection: Set col = SelectedControls
    For Each ctr In col
        Set lbl = oModule.Designer.Controls.Add("Forms.label.1")
        ctrName = ctr.Name
        ctr.Name = "tmp" & ctr.Name
        With lbl
            .Name = ctrName
            .Caption = ctr.Caption
            .Tag = ctr.Tag
            .Top = ctr.Top
            .Left = ctr.Left
            .Width = ctr.Width
            .Height = ctr.Height
            .BackColor = ctr.BackColor
            .ForeColor = ctr.ForeColor
            .Font.Size = ctr.Font.Size
            .Font.Name = ctr.Font.Name
        End With
        AssignHandCursor lbl
        oModule.Designer.Controls.Remove ctr.Name
    Next
End Sub

Sub CenterLabelCaption()
    Dim lbl As MSForms.Label
    Set lbl = SelectedControl
    lbl.Font.Size = IIf(lbl.SpecialEffect = fmSpecialEffectFlat, 18 / 22, 16 / 22) * lbl.Height
    lbl.TextAlign = fmTextAlignCenter
End Sub

Sub EditObjectProperties(ParamArray Args() As Variant)
    Rem EditObjectProperties "left",0,"top",0
    Dim obj As control: Set obj = SelectedControl
    If obj Is Nothing Or SelectedControls.count <> 1 Then Exit Sub
    If Not WorksheetFunction.IsEven(UBound(Args)) Then Exit Sub
    Dim i As Long: i = 0
    Do While i < UBound(Args)
        CallByName obj, Args(i), VbLet, _
                    IIf(IsNumeric(Args(i + 1)), _
                    CLng(Args(i + 1)), _
                    Args(i + 1))
        i = i + 2
    Loop
End Sub

Sub EditObjectsProperty(objProperty As String, ParamArray Args() As Variant)
    Dim obj As New Collection: Set obj = SelectedControls
    If obj.count <> UBound(Args) + 1 Then
        Toast "selected controls count <> arguements count"
        Exit Sub
    End If
    Dim ArgItem
    Dim i As Long
    i = obj.count
    Dim element As Variant
    For Each element In obj
        CallByName element, objProperty, VbLet, _
                    IIf(IsNumeric(Args(i - 1)), _
                    CLng(Args(i - 1)), _
                    Args(i - 1))
        i = i - 1
    Next
End Sub

Sub RenameControlAndCode()
    Dim ctr As MSForms.control: Set ctr = SelectedControl
    If ctr Is Nothing Then Exit Sub
    Dim OldName As String: OldName = ctr.Name
    Dim NewName As String: NewName = InputboxString
    If NewName = "" Then Exit Sub
    ctr.Name = NewName
    Dim CountOfLines As Long: CountOfLines = oModule.CodeModule.CountOfLines
    If CountOfLines = 0 Then Exit Sub
    Dim strline As String
    Dim i As Long
    For i = 1 To CountOfLines
        strline = oModule.CodeModule.Lines(i, 1)
        If InStr(1, strline, " " & OldName & "_") > 0 Then
            oModule.CodeModule.ReplaceLine (i), Replace(strline, OldName, NewName)
        End If
    Next
End Sub

Sub SortControlsHorizontally()
    SortControls False
End Sub

Sub SortControlsVertivally()
    SortControls True
End Sub

Private Sub SortControls(Optional SortVertically As Boolean = True)
    Rem call from immediate window while looking at userform
    Dim ctr As MSForms.control
    Dim coll As New Collection
    Dim lastTop As Long
    Dim lastLeft As Long
    Dim element As Variant
    For Each element In SelectedControls
        coll.Add element.Name
    Next
    Set coll = aCollection.Init(coll).Sort.Items
    lastTop = 2000
    For Each element In coll
        If oModule.Designer.Controls(element).Top < lastTop Then lastTop = oModule.Designer.Controls(element).Top
        If oModule.Designer.Controls(element).Left < lastLeft Then lastLeft = oModule.Designer.Controls(element).Left
    Next
    For Each element In coll
        If SortVertically = True Then
            lastTop = lastTop + oModule.Designer.Controls(element).Height + 6
        Else
            lastLeft = lastLeft + oModule.Designer.Controls(element).Width + 6
        End If
        oModule.Designer.Controls(element).Top = lastTop
        oModule.Designer.Controls(element).Left = lastLeft
    Next
End Sub

Public Sub CopyControlProperties()
    If SelectedControls.count <> 1 Then Exit Sub
    Dim ctr As control: Set ctr = SelectedControl
    Dim ws As Worksheet: Set ws = CreateOrSetSheet("CopyControlProperties", ThisWorkbook)
    Dim PropertiesArray As Variant
    PropertiesArray = Array("Accelerator", "Alignment", "AutoSize", "AutoTab", "BackColor", "BackStyle", "BorderColor", "BorderStyle", "BoundColumn", _
    "Caption", "Children", "columnCount", "ColumnHeads", "ColumnWidths", "ControlSource", "ControlTipText", "Cycle", "DrawBuffer", "Enabled", "EnterKeyBehavior", "Expanded", _
    "FirstSibling", "FontBold", "FontSize", "ForeColor", "FullPath", "GroupName", "Height", "HelpContextID", "KeepScrollBarsVisible", "LargeChange", "LastSibling", "LineStyle", "ListRows", "Locked", _
    "Max", "MaxLength", "Min", "MouseIcon", "MousePointer", "MultiLine", "MultiSelect", "Next", "Nodes", "Orientation", _
    "Parent", "PasswordChar", "PathSeparator", "Picture", "PictureAlignment", "PictureSizeMode", "PictureTiling", "Previous", "RightToLeft", "Root", "RowSource", _
    "ScrollBars", "ScrollHeight", "ScrollLeft", "ScrollTop", "ScrollWidth", "Selected", "SelectedItem", "ShowModal", "SmallChange", "Sorted", "SpecialEffect", "StartUpPosition", _
    "Style", "Tag", "Text", "TextColumn", "TripleState", "WhatsThisHelp", "Width", "Zoom")
    If ws.Range("A1") = "" Then ws.Range("A1").Resize(UBound(PropertiesArray) + 1) = WorksheetFunction.Transpose(PropertiesArray)
    Dim PropertiesRange As Range: Set PropertiesRange = ws.Range("A1").CurrentRegion.Resize(, 1)
    Dim Property As Range
    On Error Resume Next
    For Each Property In PropertiesRange
        Property.OFFSET(0, 1) = CallByName(ctr, Property.Value, VbGet)
    Next
End Sub

Public Sub PasteControlProperties()
    Dim control As MSForms.control
    Dim Controls As Collection: Set Controls = SelectedControls
    If Controls.count = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("CopyControlProperties")
    If ws.Columns(2).SpecialCells(xlCellTypeConstants).count = 0 Then
        Toast "You haven't saved properties before"
        Exit Sub
    End If
    Dim PropertiesRange As Range: Set PropertiesRange = ws.Range("A1").CurrentRegion.Resize(, 1)
    Dim Property As Range
    On Error Resume Next
    For Each control In Controls
        For Each Property In PropertiesRange
            CallByName control, Property.Value, VbLet, Property.OFFSET(0, 1).Value
        Next
    Next
End Sub



Public Sub CopySubControlProperties()
    If SelectedFrameOrMultipageControls.count <> 1 Then Exit Sub
    Dim ctr As control: Set ctr = SelectedFrameOrMultipageControl
    Dim ws As Worksheet: Set ws = CreateOrSetSheet("CopyControlProperties", ThisWorkbook)
    Dim PropertiesArray As Variant
    PropertiesArray = Array("Accelerator", "Alignment", "AutoSize", "AutoTab", "BackColor", "BackStyle", "BorderColor", "BorderStyle", "BoundColumn", _
    "Caption", "Children", "columnCount", "ColumnHeads", "ColumnWidths", "ControlSource", "ControlTipText", "Cycle", "DrawBuffer", "Enabled", "EnterKeyBehavior", "Expanded", _
    "FirstSibling", "FontBold", "FontSize", "ForeColor", "FullPath", "GroupName", "Height", "HelpContextID", "KeepScrollBarsVisible", "LargeChange", "LastSibling", "LineStyle", "ListRows", "Locked", _
    "Max", "MaxLength", "Min", "MouseIcon", "MousePointer", "MultiLine", "MultiSelect", "Next", "Nodes", "Orientation", _
    "Parent", "PasswordChar", "PathSeparator", "Picture", "PictureAlignment", "PictureSizeMode", "PictureTiling", "Previous", "RightToLeft", "Root", "RowSource", _
    "ScrollBars", "ScrollHeight", "ScrollLeft", "ScrollTop", "ScrollWidth", "Selected", "SelectedItem", "ShowModal", "SmallChange", "Sorted", "SpecialEffect", "StartUpPosition", _
    "Style", "Tag", "Text", "TextColumn", "TripleState", "WhatsThisHelp", "Width", "Zoom")
    If ws.Range("A1") = "" Then ws.Range("A1").Resize(UBound(PropertiesArray) + 1) = WorksheetFunction.Transpose(PropertiesArray)
    Dim PropertiesRange As Range: Set PropertiesRange = ws.Range("A1").CurrentRegion.Resize(, 1)
    Dim Property As Range
    On Error Resume Next
    For Each Property In PropertiesRange
        Property.OFFSET(0, 1) = CallByName(ctr, Property.Value, VbGet)
    Next
End Sub

Public Sub PasteSubControlProperties()
    Dim control As MSForms.control
    Dim Controls As Collection: Set Controls = SelectedControls
    If Controls.count = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("CopyControlProperties")
    If ws.Columns(2).SpecialCells(xlCellTypeConstants).count = 0 Then
        Toast "You haven't saved properties before"
        Exit Sub
    End If
    Dim PropertiesRange As Range: Set PropertiesRange = ws.Range("A1").CurrentRegion.Resize(, 1)
    Dim Property As Range
    On Error Resume Next
    For Each control In Controls
        For Each Property In PropertiesRange
            CallByName control, Property.Value, VbLet, Property.OFFSET(0, 1).Value
        Next
    Next
End Sub


