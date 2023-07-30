VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uTableManager 
   Caption         =   "Table Manager"
   ClientHeight    =   10104
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   20364
   OleObjectBlob   =   "uTableManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uTableManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Private TargetWorkbook As Workbook
Private TargetWorksheet As Worksheet
Public TargetTable As ListObject
Private aLV As aListView

Private isDoubleClick As Boolean
Private isAddingNewRow As Boolean

Private WithEvents myForm As UserForm
Attribute myForm.VB_VarHelpID = -1
Private currentSelection As String

Private previousInputLength As Long
Private editorIndex As Long
Private previousEditorIndex As Long

Private Sub Emitter_Focus(control As Object)
    If Not control.Name Like "Editor-*" Then Exit Sub
    
    Me.Controls("lbl-" & Split(control.Name, "-")(1)).BackColor = RGB(80, 200, 120)
    
    If InStr(1, control.Text, vbNewLine) > 0 Then
        control.ZOrder (fmTop)
        Dim dif As Long
        dif = CountOfCharacters(control.Text, vbNewLine) + 1
        control.Height = WorksheetFunction.Min( _
                                                Frame2.Height - control.Top - 12, _
                                                control.Height * dif)
        control.SelStart = 0
        control.ScrollBars = fmScrollBarsBoth
        control.BackColor = RGB(255, 255, 204)
    End If
    If TypeName(control) = "TextBox" Then control.Value = Replace(control.Value, vbTab, "")
    If TypeName(control) = "ComboBox" Then control.Text = Replace(control.Text, vbTab, "")
End Sub

Private Sub Emitter_Blur(control As Object)
    If Not control.Name Like "Editor-*" Then Exit Sub
    Me.Controls("lbl-" & Split(control.Name, "-")(1)).BackColor = Me.BackColor
    
    If InStr(1, control.Text, vbNewLine) > 0 Then
        control.Height = 18
        control.ScrollBars = fmScrollBarsNone
        control.BackColor = vbWhite
    End If
End Sub

Private Sub Emitter_Keydown(control As Object, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If Not control.Name Like "Editor-*" Then Exit Sub
    editorIndex = -1
    If KeyCode = vbKeyTab Then
        editorIndex = Split(control.Name, "-")(1)
        If Shift = 1 Then
            editorIndex = IIf(editorIndex = 1, Frame2.Controls.Count / 2, editorIndex - 1)
        Else
            editorIndex = IIf(editorIndex = Frame2.Controls.Count / 2, 1, editorIndex + 1)
        End If
    Else
        Exit Sub
    End If
    If editorIndex = -1 Then Exit Sub
    Emitter_Blur control
    Dim Editor As MSForms.control
    Set Editor = Frame2.Controls("Editor-" & editorIndex)
    If editorIndex < previousEditorIndex _
    Or (previousEditorIndex = 1 And editorIndex = Frame2.Controls.Count / 2) Then
        Frame2.ScrollTop = IIf(editorIndex = 1, 0, Controls("lbl-" & editorIndex).Top)
    End If
    Editor.SetFocus
    previousEditorIndex = editorIndex
End Sub

Private Sub GetInfo_Click()
    uAuthor.Show
End Sub

Private Sub Label1_Click()
    Select Case Label1.Tag
    Dim i As Long
    Case "Left"
        Label1.Picture = LoadPicture(ThisWorkbook.Path & "\img\right.jpg")
        Label1.Tag = "Right"
        Frame1.Width = Label1.Width + 6
    Case "Right"
        Label1.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg")
        Label1.Tag = "Left"
        Frame1.Width = ListBox1.Left + ListBox1.Width + 6
    End Select
    Frame4.Left = Frame1.Left + Frame1.Width
    Frame2.Left = Frame4.Left + Frame4.Width
    Frame3.Left = Frame2.Left
    
    aUserform.Init(Me).ResizeToFitControls (6)
End Sub

Private Sub Label2_Click()
    Select Case Label2.Tag
    Dim i As Long
    Case "Left"
        Label2.Picture = LoadPicture(ThisWorkbook.Path & "\img\right.jpg")
        Label2.Tag = "Right"
        Frame2.Width = Label2.Width + 6
    Case "Right"
        Label2.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg")
        Label2.Tag = "Left"
        Frame2.Width = 354
    End Select
    Frame3.Width = Frame2.Width
    aUserform.Init(Me).ResizeToFitControls (6)
End Sub

Private Sub Label3_Click()
    If isAddingNewRow Then
        SaveNewRow
    Else
        SaveChanges
    End If
End Sub

Private Sub Label4_Click()
    aLV.Clear
    RemoveEditorControls
    ListBox3.Clear
    Set TargetTable = Nothing
    ListBox2.Clear
    aListBox.Init(ListBox1).LoadVBProjects
    SelectActivePath
End Sub

Sub SelectActivePath()
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ActiveCell.ListObject
    If lo Is Nothing And ActiveSheet.ListObjects.Count = 1 Then Set lo = ActiveSheet.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    
    Set TargetTable = lo
    Dim i As Long
    If ListBox1.ListCount > 0 Then
        For i = 0 To ListBox1.ListCount - 1
            If ListBox1.List(i) = ActiveWorkbook.Name Then
                ListBox1.Selected(i) = True
                Exit For
            End If
        Next
    End If
    If ListBox2.ListCount > 0 Then
        For i = 0 To ListBox2.ListCount - 1
            If ListBox2.List(i) = ActiveSheet.Name Then
                ListBox2.Selected(i) = True
                Exit For
            End If
        Next
    End If
    If ListBox3.ListCount > 0 Then
        For i = 0 To ListBox3.ListCount - 1
            If ListBox3.List(i) = lo.Name Then
                ListBox3.Selected(i) = True
                Exit For
            End If
        Next
    End If
End Sub

Private Sub Label5_Click()
    isAddingNewRow = True
    CreateEditorControls
End Sub

Private Sub Label6_Click()
    If isAddingNewRow Then Exit Sub
    If MsgBox("Permanently delete this row?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    TargetTable.ListRows(ListView1.selectedItem.index).Delete
    ListView1.ListItems.Remove ListView1.selectedItem.index
End Sub

Private Sub Label7_Click()
    PrintOut
End Sub

Private Sub Label8_Click()
    aLV.AutofitColumns
End Sub

Private Sub Label9_Click()
    Select Case Label9.Tag
    Dim i As Long
    Case "Left"
        Label9.Picture = LoadPicture(ThisWorkbook.Path & "\img\right.jpg")
        Label9.Tag = "Right"
        Frame4.Width = Label9.Width + 6
    Case "Right"
        Label9.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg")
        Label9.Tag = "Left"
        Frame4.Width = 540
    End Select
    Frame2.Left = Frame4.Left + Frame4.Width
    Frame3.Left = Frame2.Left
    
    aUserform.Init(Me).ResizeToFitControls (6)
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    
    isAddingNewRow = False
    
    Dim selectedItem As MSComctlLib.ListItem
    
    On Error Resume Next
    Set selectedItem = ListView1.selectedItem
    On Error GoTo 0
    
    If selectedItem Is Nothing Then Exit Sub
    
    Dim newSelection As String
    newSelection = ListBox1.List(ListBox1.ListIndex) & _
                    ListBox2.List(ListBox2.ListIndex) & _
                    ListBox3.List(ListBox3.ListIndex) & _
                    ListView1.selectedItem.index
    
    If Not newSelection = currentSelection Then
        currentSelection = newSelection
        CreateEditorControls
    End If
    
End Sub

Private Sub ListView1_DblClick()
    isDoubleClick = True
End Sub

Private Sub ListView1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)
    isAddingNewRow = False
    CreateEditorControls
    
    If Not isDoubleClick Then Exit Sub
    
    Dim targetColumn As Long:   targetColumn = aLV.ClickedColumn(X, Y)
    Dim targetRow As Long:      targetRow = ListView1.selectedItem.index
    
'    If targetColumn = -1 Then
'
'    ElseIf targetColumn = 0 Then
'        MsgBox "Header: " & vbTab & targetColumn + 1 & vbTab & ListView1.ColumnHeaders(targetColumn) & vbNewLine & _
'               "Item: " & vbTab & targetRow & vbTab & ListView1.SelectedItem.TEXT
'    ElseIf targetColumn > 1 Then
'        MsgBox "Header: " & vbTab & targetColumn + 1 & vbTab & ListView1.ColumnHeaders(targetColumn) & vbNewLine & _
'               "Item: " & vbTab & targetRow & vbTab & ListView1.SelectedItem.TEXT & vbNewLine & _
'               "Subitem: " & vbTab & targetColumn + 1 & vbTab & ListView1.ListItems(targetRow).ListSubItems(targetColumn)
'    End If

    Dim Editor As MSForms.control
    Set Editor = Frame2.Controls("Editor-" & IIf(targetColumn = -1, 1, targetColumn + 1))
    Editor.SetFocus
    Editor.SelStart = 0
    Emitter_Focus Editor
    isDoubleClick = False
End Sub

Private Sub TextBox1_Change()
    Dim op As operator
    op = Choose(ComboBox2.ListIndex + 1, _
                operator.IS_LIKE, _
                operator.IS_EQUAL, _
                operator.NOT_EQUAL, _
                operator.IS_CONTAINS, _
                operator.NOT_CONTAINS, _
                operator.STARTS_WITH, _
                operator.ENDS_WITH, _
                operator.GREATER_THAN, _
                operator.GREATER_OR_EQUAL, _
                operator.LESS_THAN, _
                operator.LESS_OR_EQUAL)
    
    Select Case Len(TextBox1.Value)
    Case Is = 0
        aLV.InitializeFromArray TargetTable.Range.Value
    Case Is = 1
        previousInputLength = 1
        aLV.InitializeFromArray FilterArray2d(TargetTable.Range.Value, True, TextBox1.Value, operator.IS_LIKE, ComboBox1.Value)
    Case Is > 1
        If Len(TextBox1.Value) > previousInputLength Then
            aLV.InitializeFromArray FilterArray2d(aLV.Value, True, TextBox1.Value, operator.IS_LIKE, ComboBox1.Value)
        Else
            previousInputLength = Len(TextBox1.Value)
            aLV.InitializeFromArray FilterArray2d(TargetTable.Range.Value, True, TextBox1.Value, operator.IS_LIKE, ComboBox1.Value)
        End If
    End Select
    CreateEditorControls
End Sub

Private Sub UserForm_Initialize()
    aListBox.Init(ListBox1).LoadVBProjects
    Dim i As Long, ws As Worksheet, counter As Long
    For i = ListBox1.ListCount - 1 To 0 Step -1
        counter = 0
        With Workbooks(ListBox1.List(i))
            For Each ws In .Worksheets
                If ws.ListObjects.Count > 0 Then GoTo SKIP
            Next
            ListBox1.RemoveItem i
        End With
SKIP:
    Next
    
    With ListView1
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .Font.Name = "Segoe UI"
    End With
    
    Set aLV = aListView.Init(ListView1)
'----------------------------------------------------------------------------
    Frame1.BorderStyle = fmBorderStyleNone: Frame1.Caption = ""
    Frame2.BorderStyle = fmBorderStyleNone: Frame2.Caption = ""
    With Me.Frame2
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = .InsideHeight * 8.5
        .ScrollWidth = .InsideWidth * 9
    End With
    
    Frame3.BorderStyle = fmBorderStyleNone
    Frame3.Caption = ""
    Frame4.BorderStyle = fmBorderStyleNone
    Frame4.Caption = ""
    
'    Application.WindowState = xlMaximized
'    Me.Top = Application.Left
'    Me.Left = Application.Top
'    Me.Height = Application.Height
'    Me.Width = Application.Width

    Frame4.Left = Frame1.Left + Frame1.Width
'    Frame4.Width = Frame2.Left - Frame4.Left
    Frame2.Left = Frame4.Left + Frame4.Width
    Frame3.Left = Frame2.Left
'--------------------------------------------------------------------------
    Label1.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg")
    Label1.Tag = "Left"
    Label2.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg")
    Label2.Tag = "Left"
    
    Label3.Picture = LoadPicture(ThisWorkbook.Path & "\img\save.jpg")
    Label4.Picture = LoadPicture(ThisWorkbook.Path & "\img\refresh.jpg")
    Label5.Picture = LoadPicture(ThisWorkbook.Path & "\img\add.jpg")
    Label6.Picture = LoadPicture(ThisWorkbook.Path & "\img\delete.jpg")
    Label7.Picture = LoadPicture(ThisWorkbook.Path & "\img\printer.jpg")
    Label8.Picture = LoadPicture(ThisWorkbook.Path & "\img\autofit.jpg")
    
    Label9.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg")
    Label9.Tag = "Left"
    
'----------------------------------------------------------------------------
    setControlStyle
'----------------------------------------------------------------------------
    ComboBox2.List = Array("like", "contains", "starts", "ends", "=", "<>", ">", ">=", "<", "<=")
    ComboBox2.ListIndex = 0
'----------------------------------------------------------------------------
    
    SelectActivePath
    
'    aUserform.Init(Me).ResizeToFitControls (6)
'    aUserform.Init(Me).MaximizeButton
End Sub
Sub setControlStyle()
    Dim Ctrl As MSForms.control
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "Label" Then
            Ctrl.MouseIcon = LoadPicture(ThisWorkbook.Path & "\img\Hand Cursor Pointer.ico")
            
            Ctrl.MousePointer = fmMousePointerCustom
            Ctrl.SpecialEffect = fmSpecialEffectRaised
            
            Ctrl.Width = 28 'ctrl.Width + 6
            Ctrl.Height = 28 'ctrl.Height + 6
        End If
    Next
End Sub

Private Sub SaveChanges()
    TargetTable.ListRows(ListView1.selectedItem.index).Range.Value = EditorsValueArray
    UpdateListviewRow ListView1.selectedItem.index, EditorsValueArray
End Sub
Private Sub SaveNewRow()
    Dim val
    val = EditorsValueArray
    TargetTable.ListRows.Add
    TargetTable.ListRows(TargetTable.ListRows.Count).Range.Value = val
    Dim out()
    ReDim out(1 To 1, 1 To UBound(val))
    Dim i As Long
    For i = LBound(val) To UBound(val)
        out(1, i) = val(i)
    Next
    aLV.AppendArray out
End Sub
Sub UpdateListviewRow(index As Long, newValues)
    Dim i As Long
    For i = 1 To ListView1.ColumnHeaders.Count
        If i = 1 Then
            ListView1.ListItems(index).Text = newValues(i)
        Else
            ListView1.ListItems(index).ListSubItems(i - 1).Text = newValues(i)
        End If
    Next
End Sub
Function EditorsValueArray()
    Dim out()
    ReDim out(1 To Frame2.Controls.Count / 2)
    Dim val
    Dim i As Long
    For i = 1 To Frame2.Controls.Count / 2
        val = Frame2.Controls("Editor-" & i).Value
        out(i) = val
    Next
    EditorsValueArray = out
End Function
Private Sub listbox1_change()
    If ListBox1.ListIndex = -1 Then Exit Sub
    Dim ListboxValue As String
    ListboxValue = ListBox1.List(ListBox1.ListIndex)

    Set TargetWorkbook = Workbooks(ListboxValue)
    ListBox2.Clear
    ListBox3.Clear
    Set TargetTable = Nothing
    aLV.Clear
    Dim ws As Worksheet
    For Each ws In TargetWorkbook.Worksheets
        If ws.ListObjects.Count > 0 Then
            ListBox2.AddItem ws.Name
        End If
    Next
    If ListBox2.ListCount = 1 Then ListBox2.ListIndex = 0
End Sub

Private Sub listbox2_change()
    If ListBox2.ListIndex = -1 Then Exit Sub
    Dim ListboxValue As String
    ListboxValue = ListBox2.List(ListBox2.ListIndex)
    Set TargetWorksheet = TargetWorkbook.Worksheets(ListboxValue)
    ListBox3.Clear
    Set TargetTable = Nothing
    Dim lo As ListObject
    For Each lo In TargetWorksheet.ListObjects
        ListBox3.AddItem lo.Name
    Next
    aLV.Clear
    RemoveEditorControls
    If ListBox3.ListCount = 1 Then ListBox3.ListIndex = 0
End Sub

Private Sub listbox3_change()
    If ListBox3.ListIndex = -1 Then Exit Sub
    Dim ListboxValue As String
    ListboxValue = ListBox3.List(ListBox3.ListIndex)
    Set TargetTable = TargetWorksheet.ListObjects(ListboxValue)
    aLV.InitializeFromArray TargetTable.Range.Value

    RemoveEditorControls
    
    ComboBox1.Clear
    ComboBox1.AddItem "-1"
    Dim i As Long
    For i = 1 To ListView1.ColumnHeaders.Count
        ComboBox1.AddItem CStr(i)
    Next
    ComboBox1.ListIndex = 0
    ComboBox2.ListIndex = 0
    
    ListView1.SetFocus
    CreateEditorControls
End Sub

Sub RemoveEditorControls()
    Set Emitter = Nothing
    Set Emitter = New EventListenerEmitter
    Dim Ctrl As MSForms.control
    If Frame2.Controls.Count > 0 Then
'        frame2.Visible = False
        For Each Ctrl In Frame2.Controls
            Frame2.Controls.Remove Ctrl.Name
        Next
'        frame2.Visible = True
    End If
End Sub

Function isValidation(rng As Range) As Boolean
    Dim dvtype As Integer
    On Error Resume Next
    dvtype = rng.Validation.Type
    On Error GoTo 0
    If dvtype = 3 Then
        isValidation = True
    Else
        isValidation = False
    End If
End Function

Sub CreateEditorControls()

    RemoveEditorControls
    
    Dim i As Long, lbl As MSForms.Label, txt As MSForms.Textbox, cbx As MSForms.ComboBox
    Dim tableCell As Range
    Dim validationArray
    For i = 1 To ListView1.ColumnHeaders.Count
        On Error Resume Next
        If isAddingNewRow Then
            Set tableCell = TargetTable.DataBodyRange(1, i)
        Else
            Set tableCell = TargetTable.DataBodyRange(ListView1.selectedItem.index, i)
        End If
        On Error GoTo 0
            If tableCell Is Nothing Then Exit Sub
        Set lbl = Frame2.Controls.Add("Forms.Label.1")
        lbl.Width = Frame2.Width - 32
        lbl.Height = 10
        lbl.Left = 6
        lbl.Top = AvailableFormOrFrameRow(Frame2, , , 3)
        lbl.Caption = i & "/" & ListView1.ColumnHeaders.Count & " - " & ListView1.ColumnHeaders(i).Text
        lbl.Font.Size = 6
        lbl.Font.Bold = True
        lbl.Name = "lbl-" & i
        
        If Not isValidation(tableCell) Then
            Set txt = Frame2.Controls.Add("Forms.Textbox.1")
            txt.Top = lbl.Top + lbl.Height
            txt.Left = lbl.Left
            txt.Width = Frame2.Width - 32
            txt.Height = 18
            txt.Name = "Editor-" & i
            txt.Font.Name = "Segoe UI"
            txt.EnterKeyBehavior = True
            txt.MultiLine = True
            txt.WordWrap = False
        Else
            Set cbx = Frame2.Controls.Add("Forms.ComboBox.1")
            If InStr(1, tableCell.Validation.Formula1, "=") > 0 Then
                Dim rng As Range
                Set rng = Nothing
                On Error Resume Next
                Set rng = TargetWorksheet.Range(Replace(tableCell.Validation.Formula1, "=", ""))
                On Error GoTo 0
                If Not rng Is Nothing Then
                    validationArray = rng.Value
                End If
            Else
                validationArray = Split(tableCell.Validation.Formula1, ",")
            End If
            
            Dim item
            For Each item In validationArray
                cbx.AddItem item
            Next
            cbx.Top = lbl.Top + lbl.Height
            cbx.Left = lbl.Left
            cbx.Width = Frame2.Width - 32
            cbx.Height = 18
            cbx.Name = "Editor-" & i
            cbx.Font.Name = "Segoe UI"
        End If
    Next
    
    If Not isAddingNewRow Then
        For i = 1 To ListView1.ColumnHeaders.Count
            Set tableCell = TargetTable.DataBodyRange(ListView1.selectedItem.index, i)
            Me.Controls("Editor-" & i).Value = tableCell.Value
        Next
    End If
    
    Emitter.AddEventListenerAll Frame2
    Frame2.ScrollTop = 0
End Sub

Sub PrintOut()
    TargetTable.Range.PrintOut , , , True
End Sub
