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

Private oModule     As VBComponent

Function Active() As aDesigner
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    Set Active = Me
    Class_Initialize
End Function

Private Sub Class_Initialize()
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE ActiveModule
    '@INCLUDE CLASS aDesigner
    Set oModule = ActiveModule
    If oModule.Type <> vbext_ct_MSForm Then End
End Sub

Private Sub Class_Terminate()
    'exit
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
End Sub

Function SelectedControl() As msforms.control
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    Dim col         As New Collection
    Set col = SelectedControls
    If col.Count = 1 Then
        Set SelectedControl = col(1)
    End If
End Function

Function SelectedControls() As Collection
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aCollection
    '@INCLUDE CLASS aDesigner
    Dim ctl         As control
    Dim out         As New Collection
    For Each ctl In oModule.Designer.Selected
        out.Add ctl
    Next ctl

    Dim ac          As New aCollection
    Set out = ac.Init(out).Reverse.items
    Set SelectedControls = out
End Function

Function SelectedFrameOrMultipageControl() As msforms.control
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    Dim col         As New Collection
    Set col = SelectedFrameOrMultipageControls
    If col.Count = 1 Then
        Set SelectedFrameOrMultipageControl = col(1)
    End If
End Function


Function SelectedFrameOrMultipageControls() As Collection
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aCollection
    '@INCLUDE CLASS aDesigner
    Dim ctl         As msforms.control
    Dim subctl      As msforms.control
    Dim out         As New Collection
    For Each ctl In SelectedControls
        If InStr("FrameMultiPage", TypeName(ctl)) > 0 Then
            For Each subctl In ctl.Controls
                If InStr("FrameMultiPage", TypeName(subctl)) = 0 And subctl.InSelection Then
                    out.Add subctl
                End If
            Next
        End If
    Next ctl
    Dim ac          As New aCollection
    Set out = ac.Init(out).Reverse.items
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
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    Dim c           As msforms.control
    For Each c In SelectedControls
        c.Caption = ""
    Next
End Sub

Sub SetHandCursor()
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    Dim ctr As msforms.control, ctr2 As msforms.control
    For Each ctr In SelectedControls
        AssignHandCursor ctr
    Next
End Sub

Sub SetHandCursorToSubControls()
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    Dim ctr As msforms.control, ctr2 As msforms.control
    For Each ctr In SelectedFrameOrMultipageControls
        AssignHandCursor ctr
    Next
End Sub

Private Sub AssignHandCursor(control As msforms.control)
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    On Error GoTo catch
    With control
        .MouseIcon = LoadPicture(ThisWorkbook.path & "\Lib\img\Hand Cursor Pointer.ico")
        .MousePointer = fmMousePointerCustom
    End With
catch:
End Sub

Sub SwitchNames()
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    If SelectedControls.Count <> 2 Then Exit Sub

    Dim ctrls       As Collection
    Set ctrls = SelectedControls
    If ctrls.Count <> 2 Then Exit Sub
    Dim tmp1        As String
    tmp1 = ctrls(1).Name
    Dim tmp2        As String
    tmp2 = ctrls(2).Name
    ctrls(1).Name = "tmp1"
    ctrls(2).Name = "tmp2"
    ctrls(1).Name = tmp2
    ctrls(2).Name = tmp1
End Sub

Sub SwitchPositions()
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    If SelectedControls.Count <> 2 Then
        MsgBox "Select exactly 2 controls to swirtch their position"
        Exit Sub
    End If
    Dim ctrls       As Collection
    Set ctrls = SelectedControls
    If ctrls.Count <> 2 Then Exit Sub
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
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE tmp
    '@INCLUDE CLASS aDesigner
    Dim ctr As msforms.control, ctr2 As msforms.control
    Dim lbl         As msforms.Label
    Dim ctrName     As String
    Dim col         As New Collection: Set col = SelectedControls
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

'* Modified   : Date and Time       Author              Description
'* Updated    : 24-10-2023 19:12    Alex                (aDesigner.cls > CenterLabelCaption)

Sub CenterLabelCaption()
'@LastModified 2310241912
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    Dim lbl         As msforms.Label
    Set lbl = SelectedControl
    For Each lbl In SelectedControls
        If TypeName(lbl) = "Label" Then
            lbl.Font.Size = IIf(lbl.SpecialEffect = fmSpecialEffectFlat, 18 / 22, 16 / 22) * lbl.Height
            lbl.TextAlign = fmTextAlignCenter
        End If
    Next
End Sub

Sub EditObjectProperties(ParamArray Args() As Variant)
    Rem EditObjectProperties "left",0,"top",0
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    Dim obj         As control: Set obj = SelectedControl
    If obj Is Nothing Or SelectedControls.Count <> 1 Then Exit Sub
    If Not WorksheetFunction.IsEven(UBound(Args)) Then Exit Sub
    Dim i           As Long: i = 0
    Do While i < UBound(Args)
        CallByName obj, Args(i), VbLet, _
                IIf(IsNumeric(Args(i + 1)), _
                CLng(Args(i + 1)), _
                Args(i + 1))
        i = i + 2
    Loop
End Sub

Sub EditObjectsProperty(objProperty As String, ParamArray Args() As Variant)
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE Toast
    '@INCLUDE CLASS aDesigner
    Dim obj         As New Collection: Set obj = SelectedControls
    If obj.Count <> UBound(Args) + 1 Then
        Toast "selected controls count <> arguements count"
        Exit Sub
    End If
    Dim ArgItem
    Dim i           As Long
    i = obj.Count
    Dim element     As Variant
    For Each element In obj
        CallByName element, objProperty, VbLet, _
                IIf(IsNumeric(Args(i - 1)), _
                CLng(Args(i - 1)), _
                Args(i - 1))
        i = i - 1
    Next
End Sub

Sub RenameControlAndCode()
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE InputboxString
    '@INCLUDE CLASS aDesigner
    Dim ctr         As msforms.control: Set ctr = SelectedControl
    If ctr Is Nothing Then Exit Sub
    Dim OldName     As String: OldName = ctr.Name
    Dim NewName     As String: NewName = InputboxString
    If NewName = "" Then Exit Sub
    ctr.Name = NewName
    Dim CountOfLines As Long: CountOfLines = oModule.CodeModule.CountOfLines
    If CountOfLines = 0 Then Exit Sub
    Dim strline     As String
    Dim i           As Long
    For i = 1 To CountOfLines
        strline = oModule.CodeModule.Lines(i, 1)
        If InStr(1, strline, " " & OldName & "_") > 0 Then
            oModule.CodeModule.ReplaceLine (i), Replace(strline, OldName, NewName)
        End If
    Next
End Sub

Sub SortControlsHorizontally()
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    SortControls False
End Sub

Sub SortControlsVertically()
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    SortControls True
End Sub

Private Sub SortControls(Optional SortVertically As Boolean = True)
    Rem call from immediate window while looking at userform
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aCollection
    '@INCLUDE CLASS aDesigner
    Dim ctr         As msforms.control
    Dim coll        As New Collection
    Dim lastTop     As Long
    Dim lastLeft    As Long
    Dim element     As Variant
    For Each element In SelectedControls
        coll.Add element.Name
    Next
    Set coll = aCollection.Init(coll).Sort.items
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
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE Min
    '@INCLUDE PROCEDURE CreateOrSetSheet
    '@INCLUDE CLASS aDesigner
    If SelectedControls.Count <> 1 Then Exit Sub
    Dim ctr         As control: Set ctr = SelectedControl
    Dim ws          As Worksheet: Set ws = CreateOrSetSheet("CopyControlProperties", ThisWorkbook)
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
    Dim Property    As Range
    On Error Resume Next
    For Each Property In PropertiesRange
        Property.offset(0, 1) = CallByName(ctr, Property.value, VbGet)
    Next
End Sub

Public Sub PasteControlProperties()
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE Toast
    '@INCLUDE CLASS aDesigner
    Dim control     As msforms.control
    Dim Controls    As Collection: Set Controls = SelectedControls
    If Controls.Count = 0 Then Exit Sub
    Dim ws          As Worksheet: Set ws = ThisWorkbook.Sheets("CopyControlProperties")
    If ws.Columns(2).SpecialCells(xlCellTypeConstants).Count = 0 Then
        Toast "You haven't saved properties before"
        Exit Sub
    End If
    Dim PropertiesRange As Range: Set PropertiesRange = ws.Range("A1").CurrentRegion.Resize(, 1)
    Dim Property    As Range
    On Error Resume Next
    For Each control In Controls
        For Each Property In PropertiesRange
            CallByName control, Property.value, VbLet, Property.offset(0, 1).value
        Next
    Next
End Sub



Public Sub CopySubControlProperties()
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE Min
    '@INCLUDE PROCEDURE CreateOrSetSheet
    '@INCLUDE CLASS aDesigner
    If SelectedFrameOrMultipageControls.Count <> 1 Then Exit Sub
    Dim ctr         As control: Set ctr = SelectedFrameOrMultipageControl
    Dim ws          As Worksheet: Set ws = CreateOrSetSheet("CopyControlProperties", ThisWorkbook)
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
    Dim Property    As Range
    On Error Resume Next
    For Each Property In PropertiesRange
        Property.offset(0, 1) = CallByName(ctr, Property.value, VbGet)
    Next
End Sub

Public Sub PasteSubControlProperties()
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE Toast
    '@INCLUDE CLASS aDesigner
    Dim control     As msforms.control
    Dim Controls    As Collection: Set Controls = SelectedControls
    If Controls.Count = 0 Then Exit Sub
    Dim ws          As Worksheet: Set ws = ThisWorkbook.Sheets("CopyControlProperties")
    If ws.Columns(2).SpecialCells(xlCellTypeConstants).Count = 0 Then
        Toast "You haven't saved properties before"
        Exit Sub
    End If
    Dim PropertiesRange As Range: Set PropertiesRange = ws.Range("A1").CurrentRegion.Resize(, 1)
    Dim Property    As Range
    On Error Resume Next
    For Each control In Controls
        For Each Property In PropertiesRange
            CallByName control, Property.value, VbLet, Property.offset(0, 1).value
        Next
    Next
End Sub


Public Sub addFrameFormCode()
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE CLIP
    '@INCLUDE PROCEDURE Toast
    '@INCLUDE PROCEDURE Reframe
    '@INCLUDE PROCEDURE dp
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aDesigner
    '@INCLUDE DECLARATION Emitter
    Dim S           As String
    Dim q: q = """"
    S = Join( _
            Array(, _
                "Private WithEvents Emitter As EventListeneRemitter", _
                "", _
                "Private Sub Emitter_LabelMouseOut(Label As MSForms.Label)", _
                "    If InStr(1, Label.Tag, " & q & "reframe" & q & ", vbTextCompare) > 0 Then", _
                "        If Label.BackColor <> &H80B91E Then Label.BackColor = &H534848", _
                "    End If", _
                "End Sub", _
                "", _
                "Private Sub Emitter_LabelMouseOver(Label As MSForms.Label)", _
                "    If InStr(1, Label.Tag, " & q & "reframe" & q & ", vbTextCompare) > 0 Then", _
                "        If Label.BackColor <> &H80B91E Then Label.BackColor = &H808080", _
                "    End If", _
                "End Sub", _
                "", _
                "Sub Emitter_LabelClick(ByRef Label As MSForms.Label)", _
                "    If InStr(1, Label.Tag, " & q & "reframe" & q & ", vbTextCompare) > 0 Then Reframe Me, Label", _
                "End Sub", _
                "", _
                "Private Sub UserForm_Initialize()", _
                "    startFrameForm Me", _
                "End Sub"), _
            vbLf)
    If InStr(1, aModule.Init(oModule).Code, oModule.Name & "_Initialize") Then
        Toast "Threre is already _Initialize_ code in this form. Code will be put in cilpboard and immediate window."
        dp S
        CLIP S
    Else
        oModule.CodeModule.AddFromString S
    End If
End Sub


Public Sub CreateFrameMenu()
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE ActiveModule
    '@INCLUDE CLASS aDesigner
    '@INCLUDE DECLARATION MyColors
    Select Case SelectedControls.Count
        Case Is = 0
            ActiveModule.Designer.BackColor = MyColors.FormBackgroundDarkGray
            addFrameSidebar ActiveModule
            addFrameFormCode
        Case Is = 1
            addFrameSidebar SelectedControl
    End Select
End Sub

Private Sub addFrameSidebar(Form As Object, Optional dockRight As Boolean)
    '@INCLUDE askFormMenuElements
    '@INCLUDE UnderlineFrameName
    '@INCLUDE CreateOrSetFrame
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE CreateOrSetFrame
    '@INCLUDE CLASS aDesigner
    '@INCLUDE DECLARATION ControlIDLabel
    Dim F           As msforms.control
    Dim l           As msforms.control
    Set F = CreateOrSetFrame(Form, "SideBar" & Form.Name)
    F.Tag = "skip"
    F.BackColor = 5457992
    F.ForeColor = vbWhite
    F.BorderStyle = 1
    F.BorderStyle = 0
    F.Width = 80
    If TypeName(Form) = "VBComponent" Then
        F.Height = 800
    Else
        F.Height = Form.Height
    End If
    dockRight = IIf(TypeName(Form) = "VBComponent", False, True)
    If dockRight = True Then
        F.Left = Form.Width - F.Width
    Else
        F.Left = 0
    End If
    UnderlineFrameName Form, F
    If TypeName(Form) = "VBComponent" Then
        Set l = Form.Designer.Controls.Add(ControlIDLabel, "Anchor" & Form.Name)
    Else
        Set l = Form.Controls.Add(ControlIDLabel, "Anchor" & Form.Name)
    End If
    l.Visible = False
    l.Left = IIf(TypeName(Form) = "VBComponent", F.Left + F.Width + 9, 1)
    l.Top = 12
    l.Width = 1
    l.BackColor = vbWhite
    l.Visible = False
    askFormMenuElements Form
End Sub

Private Sub askFormMenuElements(Form As Object)
    '@INCLUDE InputboxString
    '@INCLUDE addFrameMenu
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE InputboxString
    '@INCLUDE CLASS aDesigner
    Dim FormElements As String
    FormElements = InputboxString("Form Menus", "Type comma delimited menu names")
    If FormElements = "" Then Exit Sub
    Dim var
    var = Split(FormElements, ",")
    Dim i           As Long
    For i = LBound(var) To UBound(var)
        var(i) = Trim(var(i))
    Next
    Dim coll        As New Collection
    Dim element
    On Error Resume Next
    For Each element In var
        If Not IsNumeric(Left(element, 1)) _
                And InStr(1, element, " ") = 0 Then
            coll.Add CStr(element), CStr(element)
        End If
    Next
    On Error GoTo 0
    For Each element In coll
        addFrameMenu Form, CStr(element)
    Next
End Sub

Private Sub addFrameMenu(Form As Object, FrameCaptionNoSpace As String)
    '@INCLUDE Reframe
    '@INCLUDE UnderlineFrameName
    '@INCLUDE CreateOrSetFrame
    '@INCLUDE AvailableFormOrFrameRow
    '@INCLUDE AvailableFormOrFrameColumn
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE CreateOrSetFrame
    '@INCLUDE PROCEDURE AvailableFormOrFrameRow
    '@INCLUDE PROCEDURE AvailableFormOrFrameColumn
    '@INCLUDE PROCEDURE Reframe
    '@INCLUDE CLASS aDesigner
    '@INCLUDE DECLARATION ControlIDFrame
    '@INCLUDE DECLARATION ControlIDLabel
    Dim F           As msforms.control
    Dim l           As msforms.control
    Dim module      As VBComponent
    If TypeName(Form) = "VBComponent" Then
        Set module = Form
        Set F = module.Designer.Controls.Add(ControlIDFrame, FrameCaptionNoSpace)
    Else
        Set module = ThisWorkbook.VBProject.VBComponents(Form.Parent.Name)
        Set F = CreateOrSetFrame(module.Designer.Controls(Form.Name), FrameCaptionNoSpace)
    End If
    F.Tag = "anchor" & Form.Name
    F.Caption = FrameCaptionNoSpace
    F.ForeColor = vbWhite
    F.Visible = False
    If TypeName(Form) = "VBComponent" Then
        F.Left = AvailableFormOrFrameColumn(Form.Designer)
    Else
        F.Left = 0
    End If
    F.Visible = True
    F.BorderStyle = 1
    F.BorderStyle = 0
    F.Top = 12
    F.Width = 100
    UnderlineFrameName Form, F
    If TypeName(Form) = "VBComponent" Then
        Set l = module.Designer.Controls("SideBar" & Form.Name).Controls.Add(ControlIDLabel)
    Else
        Set l = module.Designer.Controls("SideBar" & Form.Name).Add(ControlIDLabel)
    End If
    l.Caption = FrameCaptionNoSpace
    l.ForeColor = vbWhite
    l.Visible = False
    l.Top = AvailableFormOrFrameRow(module.Designer.Controls("SideBar" & Form.Name))
    l.Left = l.Left + 3
    l.Visible = True
    l.Tag = "reframe"
    l.Width = F.Width
End Sub

Private Sub AddControlsToFrame(isSubFrame As Boolean)
    '@INCLUDE SelectedControl
    '@INCLUDE SelectedControls
    '@INCLUDE SelectedFrameControl
    '@INCLUDE ActiveModule
    '@INCLUDE InputboxString
    '@AssignedModule aDesigner
    '@INCLUDE PROCEDURE ActiveModule
    '@INCLUDE PROCEDURE InputboxString
    '@INCLUDE CLASS aDesigner
    '@INCLUDE DECLARATION ControlIDCommandButton
    If ActiveModule.Type <> vbext_ct_MSForm Then Exit Sub
    Dim ad          As New aDesigner
    If ad.SelectedControls.Count <> 1 Then Exit Sub
    If TypeName(ad.SelectedControl) <> "Frame" Then Exit Sub
    Dim module      As VBComponent
    Dim TargetFrame As msforms.control
    If isSubFrame = False Then
        Set TargetFrame = ad.SelectedControl
        Set module = ActiveModule
    Else
        Set TargetFrame = ad.SelectedFrameOrMultipageControl
        Set module = ThisWorkbook.VBProject.VBComponents(TargetFrame.Parent.Parent.Name)
    End If
    Dim ControlNames As String
    ControlNames = InputboxString("Form Menus", "Type comma delimited menu names")
    If ControlNames = "" Then Exit Sub
    Dim var
    var = Split(ControlNames, ",")
    Dim i           As Long
    For i = LBound(var) To UBound(var)
        var(i) = Trim(var(i))
    Next
    Dim coll        As New Collection
    Dim element
    On Error Resume Next
    For Each element In var
        If Not IsNumeric(Left(element, 1)) _
                And InStr(1, element, " ") = 0 Then
            coll.Add CStr(element), CStr(element)
        End If
    Next
    On Error GoTo 0
    Dim l           As msforms.control
    For Each element In coll
        Set l = module.Designer.Controls(TargetFrame.Name).Controls.Add(ControlIDCommandButton, element)
        l.Top = 7 + ((TargetFrame.Controls.Count - 1) * l.Height)
        l.BackColor = vbWhite
    Next
End Sub

Private Sub UnderlineFrameName(Form As Object, F As msforms.control)
    '@AssignedModule aDesigner
    '@INCLUDE CLASS aDesigner
    '@INCLUDE DECLARATION ControlIDLabel
    Dim module      As VBComponent
    If TypeName(Form) = "VBComponent" Then
        Set module = Form
    Else
        Set module = ThisWorkbook.VBProject.VBComponents(Form.Parent.Name)
    End If
    Dim l           As msforms.Label
    Set l = module.Designer.Controls(F.Name).Controls.Add(ControlIDLabel)
    l.Top = 6
    l.Height = 1
    l.Width = 100
    l.BackColor = vbWhite
    l.Tag = "skip"
End Sub

