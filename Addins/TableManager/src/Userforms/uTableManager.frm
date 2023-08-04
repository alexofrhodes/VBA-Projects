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
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* BLOG       : https://alexofrhodes.github.io
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* DONATE     : http://paypal.me/alexofrhodes
'*
'* Project    : Table Manager
'* Purpose    : View and Edit Tables
'* Version    : 1.3.0
'* Date       : 2023-08-04
'*
'* Modifications:
'*
'*  + Added optional date picker for date fields
'*  + Cell validation applies to Editor Controls
'*  + Added Comments
'*  ~ Refactored code
'*  ~ Renamed Controls
'* Fixed too long scrollbar on Editor Frame
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Private TargetWorkbook      As Workbook
Private TargetWorksheet     As Worksheet
Private TargetTable         As ListObject
Private aLV                 As aListView
Private isDoubleClick       As Boolean
Private isAddingNewRow      As Boolean
Private currentSelection    As String
Private previousInputLength As Long
Private editorIndex         As Long
Private previousEditorIndex As Long
Private targetBanner        As Object

'/////////////////////////////////////////////
'///            INITIALIZE                 ///
'/////////////////////////////////////////////

Private Sub UserForm_Initialize()
    aUserform.Init(Me).MinimizeButton
    Set aLV = aListView.Init(ListViewTable)
    LoadWorkbooks
    LoadOptions
    setControlStyle
    SelectActivePath
End Sub

Sub LoadWorkbooks()
    aListBox.Init(ListboxWorkbook).LoadVBProjects
    Dim i As Long, ws As Worksheet, counter As Long
    'remove workbooks without any table
    For i = ListboxWorkbook.ListCount - 1 To 0 Step -1
        counter = 0
        With Workbooks(ListboxWorkbook.List(i))
            For Each ws In .Worksheets
                If ws.ListObjects.Count > 0 Then GoTo SKIP
            Next
            ListboxWorkbook.RemoveItem i
        End With
SKIP:
    Next
End Sub

Sub LoadOptions()
    'at this moment only if user checked Show Calendar Checkbox
    Dim boo: boo = IniReadKey(ThisWorkbook.Path & "\config.ini", Me.Name, "ShowCalendar")
    If boo <> vbNullString Then CheckBoxDatePicker.Value = boo
    ComboBoxFilterOperator.List = Array("like", "contains", "starts", "ends", "=", "<>", ">", ">=", "<", "<=")
    ComboBoxFilterOperator.ListIndex = 0
End Sub

Private Sub CheckBoxDatePicker_Click()
    'Save checked status
    IniWrite ThisWorkbook.Path & "\config.ini", Me.Name, "ShowCalendar", CheckBoxDatePicker.Value
End Sub

Sub setControlStyle()
    FormatListview
    FormatFrames
    SetImages
    FormatLabels
End Sub

Sub FormatListview()
    With ListViewTable
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .Font.Name = "Segoe UI"
    End With
End Sub

Sub FormatFrames()
    FrameTablePath.BorderStyle = fmBorderStyleNone: FrameTablePath.Caption = ""
    With FrameTableView
        .BorderStyle = fmBorderStyleNone
        .Caption = ""
        .Left = FrameTablePath.Left + FrameTablePath.Width
    End With
    With FrameTableEditor
        .BorderStyle = fmBorderStyleNone: FrameTableEditor.Caption = ""
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = .InsideHeight * 8.5
        .ScrollWidth = .InsideWidth * 9
        .Left = FrameTableView.Left + FrameTableView.Width
    End With
    With FrameTableEditorBottom
        .BorderStyle = fmBorderStyleNone
        .Caption = ""
        .Left = FrameTableEditor.Left
    End With
End Sub

Sub SetImages()
    LabelResize1.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg"): LabelResize1.Tag = "Left"
    LabelResize2.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg"): LabelResize2.Tag = "Left"
    LabelResize3.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg"): LabelResize3.Tag = "Left"
    LabelNew.Picture = LoadPicture(ThisWorkbook.Path & "\img\add.jpg")
    LabelSave.Picture = LoadPicture(ThisWorkbook.Path & "\img\save.jpg")
    LabelRefresh.Picture = LoadPicture(ThisWorkbook.Path & "\img\refresh.jpg")
    LabelDelete.Picture = LoadPicture(ThisWorkbook.Path & "\img\delete.jpg")
    LabelPrint.Picture = LoadPicture(ThisWorkbook.Path & "\img\printer.jpg")
    LabelAutofit.Picture = LoadPicture(ThisWorkbook.Path & "\img\autofit.jpg")
    LabelGoFirst.Picture = LoadPicture(ThisWorkbook.Path & "\img\top.jpg")
    LabelGoLast.Picture = LoadPicture(ThisWorkbook.Path & "\img\bottom.jpg")
    CheckBoxDatePicker.Picture = LoadPicture(ThisWorkbook.Path & "\img\calendar.jpg")
    CheckBoxDatePicker.PicturePosition = fmPicturePositionLeftCenter
End Sub

Sub FormatLabels()
    Dim Ctrl As MSForms.control
    For Each Ctrl In Me.Controls
        Select Case TypeName(Ctrl)
        Case "Label"
            With Ctrl
                .MouseIcon = LoadPicture(ThisWorkbook.Path & "\img\Hand Cursor Pointer.ico")
                .MousePointer = fmMousePointerCustom
                .SpecialEffect = fmSpecialEffectRaised
                .Width = 28
                .Height = 28
            End With
        'not labels but let's put it here
        Case "ComboBox", "TextBox", "ListBox"
            Ctrl.BorderStyle = fmBorderStyleSingle
        End Select
    Next
End Sub

Sub SelectActivePath()
    'if active cell inside table OR only 1 table in active sheet then
    'choose active workbook > active sheet > that table
    'and create the editors for the first row
    
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ActiveCell.ListObject
    On Error GoTo 0
    If lo Is Nothing And ActiveSheet.ListObjects.Count = 1 Then Set lo = ActiveSheet.ListObjects(1)
    If lo Is Nothing Then Exit Sub
    
    Set TargetTable = lo
    Dim i As Long
    If ListboxWorkbook.ListCount > 0 Then
        For i = 0 To ListboxWorkbook.ListCount - 1
            If ListboxWorkbook.List(i) = ActiveWorkbook.Name Then
                ListboxWorkbook.Selected(i) = True
                Exit For
            End If
        Next
    End If
    If ListboxWorksheet.ListCount > 0 Then
        For i = 0 To ListboxWorksheet.ListCount - 1
            If ListboxWorksheet.List(i) = ActiveSheet.Name Then
                ListboxWorksheet.Selected(i) = True
                Exit For
            End If
        Next
    End If
    If ListboxTable.ListCount > 0 Then
        For i = 0 To ListboxTable.ListCount - 1
            If ListboxTable.List(i) = lo.Name Then
                ListboxTable.Selected(i) = True
                Exit For
            End If
        Next
    End If
End Sub


'/////////////////////////////////////////////
'///     BOOK - SHEET - TABLE CHANGE       ///
'/////////////////////////////////////////////

Private Sub ListboxWorkbook_change()
    'LIST workbook's SHEETS that have tables
    If ListboxWorkbook.ListIndex = -1 Then Exit Sub
    Dim ListboxValue As String
    ListboxValue = ListboxWorkbook.List(ListboxWorkbook.ListIndex)

    Set TargetWorkbook = Workbooks(ListboxValue)
    ListboxWorksheet.Clear
    ListboxTable.Clear
    Set TargetTable = Nothing
    aLV.Clear
    Dim ws As Worksheet
    For Each ws In TargetWorkbook.Worksheets
        If ws.ListObjects.Count > 0 Then
            ListboxWorksheet.AddItem ws.Name
        End If
    Next
    If ListboxWorksheet.ListCount = 1 Then ListboxWorksheet.ListIndex = 0
End Sub

Private Sub ListboxWorksheet_change()
    'LIST worksheet's TABLES
    If ListboxWorksheet.ListIndex = -1 Then Exit Sub
    Dim ListboxValue As String
    ListboxValue = ListboxWorksheet.List(ListboxWorksheet.ListIndex)
    Set TargetWorksheet = TargetWorkbook.Worksheets(ListboxValue)
    ListboxTable.Clear
    Set TargetTable = Nothing
    Dim lo As ListObject
    For Each lo In TargetWorksheet.ListObjects
        ListboxTable.AddItem lo.Name
    Next
    aLV.Clear
    RemoveEditorControls
    'if there's only one table, load it
    If ListboxTable.ListCount = 1 Then ListboxTable.ListIndex = 0
End Sub

Sub RemoveEditorControls()
    Set Emitter = Nothing
    Set Emitter = New EventListenerEmitter
    Dim Ctrl As MSForms.control
    If FrameTableEditor.Controls.Count > 0 Then
        For Each Ctrl In FrameTableEditor.Controls
            FrameTableEditor.Controls.Remove Ctrl.Name
        Next
    End If
End Sub

Private Sub ListboxTable_change()
    'LOAD TABLE
    If ListboxTable.ListIndex = -1 Then Exit Sub
    Set TargetTable = TargetWorksheet.ListObjects(ListboxTable.List(ListboxTable.ListIndex))
    aLV.InitializeFromArray TargetTable.Range.Value
    SetListviewNumberFormat
    ResetFilters
    ListViewTable.SetFocus
    CreateEditorControls
End Sub

Private Sub SetListviewNumberFormat()
    'otherwise the value will be loaded as a weird number
    'this may be slow on large tables, @TODO confirm and modify aListView.InitializeFromArray if needed
    Dim x As ListItem, y As ListSubItem, val
    For Each x In ListViewTable.ListItems
        val = x.Text
        If IsDate(val) Or IsTime(val) Then
            applyFormat val, TargetTable.DataBodyRange(x.index, 1)
            x.Text = val
        End If
        For Each y In x.ListSubItems
            val = y.Text
            If IsDate(val) Or IsTime(val) Then
                applyFormat val, TargetTable.DataBodyRange(x.index, y.index + 1)
                y.Text = val
            End If
        Next
    Next
End Sub

Function IsTime(ByVal inputValue As Variant) As Boolean
    If IsNumeric(inputValue) Or IsDate(inputValue) Then
        ' Check if the numeric value is within the valid time range (0 to 1)
        If inputValue >= 0 And inputValue <= 1 Then
            ' Convert to total seconds in a day
            Dim totalSeconds As Double
            totalSeconds = inputValue * 24 * 60 * 60
            ' Validate the individual components of the time
            Dim hours As Long, minutes As Long, seconds As Long
            hours = Int(totalSeconds / 3600)
            totalSeconds = totalSeconds Mod 3600
            minutes = Int(totalSeconds / 60)
            seconds = totalSeconds Mod 60
            If hours >= 0 And hours < 24 And minutes >= 0 And minutes < 60 And seconds >= 0 And seconds < 60 Then
                IsTime = True
                Exit Function
            End If
        End If
    End If
    IsTime = False
End Function

Sub applyFormat(ByRef inputValue As Variant, ByVal inputRange As Range)
    If IsEmpty(inputValue) Then Exit Sub
    On Error Resume Next
    Dim cellFormat As String
    cellFormat = inputRange.NumberFormat
    On Error GoTo 0
    If cellFormat = "General" Then Exit Sub
    On Error Resume Next
    Dim formattedValue As Variant
    formattedValue = Format(inputValue, cellFormat)
    On Error GoTo 0
    If Not IsError(formattedValue) Then inputValue = formattedValue
End Sub

Sub ResetFilters()
    ComboBoxFilterColumn.Clear
    ComboBoxFilterColumn.AddItem "-1"
    Dim i As Long
    For i = 1 To ListViewTable.ColumnHeaders.Count
        ComboBoxFilterColumn.AddItem CStr(i)
    Next
    If ComboBoxFilterColumn.ListCount > 0 Then ComboBoxFilterColumn.ListIndex = 0
    If ComboBoxFilterOperator.ListCount > 0 Then ComboBoxFilterOperator.ListIndex = 0
End Sub

Private Sub CreateEditorControls()
    RemoveEditorControls
    Dim i As Long, lbl As MSForms.Label, txt As MSForms.Textbox, cbx As MSForms.ComboBox
    Dim cell As Range
    Dim validationArray
    For i = 1 To ListViewTable.ColumnHeaders.Count
        On Error Resume Next
        If isAddingNewRow Then
            Set cell = TargetTable.DataBodyRange(1, i)
        Else
            Set cell = TargetTable.DataBodyRange(ListViewTable.selectedItem.index, i)
        End If
        On Error GoTo 0
        If cell Is Nothing Then Exit Sub
        Set lbl = FrameTableEditor.Controls.Add("Forms.Label.1")
        lbl.Width = FrameTableEditor.Width - 32
        lbl.Height = 10
        lbl.Left = 6
        lbl.Top = AvailableFormOrFrameRow(FrameTableEditor, , , 3)
        lbl.Caption = i & "/" & ListViewTable.ColumnHeaders.Count & " - " & ListViewTable.ColumnHeaders(i).Text
        lbl.Font.Size = 6
        lbl.Font.Bold = True
        lbl.Name = "lbl-" & i
        Dim dvType As Integer
        On Error Resume Next
        dvType = cell.Validation.Type
        On Error GoTo 0
        'create a combobox if it has datavalidation list, for example = Cat, Dog, Horse
        '@TODO check if list is from range
        If isValidationList(cell) Then
            Set cbx = FrameTableEditor.Controls.Add("Forms.ComboBox.1")
            If InStr(1, cell.Validation.Formula1, "=") > 0 Then
                Dim rng As Range
                Set rng = Nothing
                On Error Resume Next
                Set rng = TargetWorksheet.Range(Replace(cell.Validation.Formula1, "=", ""))
                On Error GoTo 0
                If Not rng Is Nothing Then
                    validationArray = rng.Value
                End If
            Else
                validationArray = Split(cell.Validation.Formula1, ",")
            End If
            
            Dim item
            For Each item In validationArray
                cbx.AddItem item
            Next
            cbx.Top = lbl.Top + lbl.Height
            cbx.Left = lbl.Left
            cbx.Width = FrameTableEditor.Width - 32
            cbx.Height = 16
            cbx.Name = "Editor-" & i
            cbx.Font.Name = "Segoe UI"
            cbx.BorderStyle = fmBorderStyleSingle
'            cbx.Style=fmStyleDropDownList
        Else 'create a textbox
            Set txt = FrameTableEditor.Controls.Add("Forms.Textbox.1")
            txt.Top = lbl.Top + lbl.Height
            txt.Left = lbl.Left
            txt.Width = FrameTableEditor.Width - 32
            txt.Height = 16
            txt.Name = "Editor-" & i
            txt.Font.Name = "Segoe UI"
            txt.EnterKeyBehavior = True
            txt.MultiLine = True
            txt.WordWrap = False
            txt.BorderStyle = fmBorderStyleSingle
        End If
    Next
    If isAddingNewRow Then
        'if adding new row then editors would be empty, do nothing
    Else 'load the values properly formated
        Dim val As Variant
        For i = 1 To ListViewTable.ColumnHeaders.Count
            Set cell = TargetTable.DataBodyRange(ListViewTable.selectedItem.index, i)
            val = cell.Value
            applyFormat val, cell
            Me.Controls("Editor-" & i).Value = val
        Next
    End If
    'add event handling to the dynamically created Editor Controls
    Emitter.AddEventListenerAll FrameTableEditor
    FrameTableEditor.ScrollTop = 0
    SetScrollHeight
End Sub

Function isValidationList(cell As Range) As Boolean
    Dim dvType As Integer
    On Error Resume Next
    dvType = cell.Validation.Type
    On Error GoTo 0
    If dvType = 3 Then isValidationList = True
End Function

Sub SetScrollHeight()
    Dim ctr As MSForms.control
    Set ctr = FrameTableEditor.Controls("Editor-" & FrameTableEditor.Controls.Count / 2)
    'making it able to go a bit further down below last control
    'because if it is a multiline textbox it will be able to expand its height (see Emitter_Focus)
    FrameTableEditor.ScrollHeight = ctr.Top + ctr.Height + 100
End Sub


'/////////////////////////////////////////////
'///            TABLE EVENTS               ///
'/////////////////////////////////////////////

Private Sub ListViewTable_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub ListViewTable_DblClick()
    'we actually need to info from the _MouseUp parameters which are not available in _DblClick
    'so we put our DblClick code in the _MouseUp event, but allow it to run only if we Double Clicked
    isDoubleClick = True
End Sub

Private Sub ListViewTable_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
                                  ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    'we clicked a table row, so recreate the editors
    isAddingNewRow = False
    CreateEditorControls
    If Not isDoubleClick Then Exit Sub
    'detect in which column the cell we clicked belongs
    Dim targetColumn As Long:   targetColumn = aLV.ClickedColumn(x, y)
    Dim targetRow As Long:      targetRow = ListViewTable.selectedItem.index
    'set focus to the corresponding Editor
    Dim Editor As MSForms.control
    Set Editor = FrameTableEditor.Controls("Editor-" & IIf(targetColumn = -1, 1, targetColumn + 1))
    Editor.SetFocus
    Editor.SelStart = 0
    Emitter_Focus Editor
    isDoubleClick = False
End Sub

Private Sub ListViewTable_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    'if we're moving with up/down keys so changed row, recreate the editors
    isAddingNewRow = False
    Dim selectedItem As MSComctlLib.ListItem
    On Error Resume Next
    Set selectedItem = ListViewTable.selectedItem
    On Error GoTo 0
    If selectedItem Is Nothing Then Exit Sub
    Dim newSelection As String
    newSelection = ListboxWorkbook.List(ListboxWorkbook.ListIndex) & _
                    ListboxWorksheet.List(ListboxWorksheet.ListIndex) & _
                    ListboxTable.List(ListboxTable.ListIndex) & _
                    ListViewTable.selectedItem.index
    If Not newSelection = currentSelection Then
        currentSelection = newSelection
        CreateEditorControls
    End If
End Sub

Private Sub LabelAutoFit_Click()
    aLV.AutofitColumns
End Sub

Private Sub LabelRefresh_Click()
    'refresh if for example we have the form loaded and created a new table or opened a new workbook
    aLV.Clear
    RemoveEditorControls
    ListboxTable.Clear
    Set TargetTable = Nothing
    ListboxWorksheet.Clear
    aListBox.Init(ListboxWorkbook).LoadVBProjects
    SelectActivePath
End Sub

Private Sub LabelPrint_Click()
    TargetTable.Range.PrintOut , , , True
End Sub

Private Sub TextBoxFilterValue_Change()
    '@TODO consider actually filtering the table on the worksheet and loading its filtered data
    '      but the current method allows filtering if value matches any cell in the row
    Dim op As operator
    op = Choose(ComboBoxFilterOperator.ListIndex + 1, _
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
    
    Select Case Len(TextBoxFilterValue.Value)
    Case Is = 0 'show whole range.value
        aLV.InitializeFromArray TargetTable.Range.Value
    Case Is = 1 'filter the whole array
        previousInputLength = 1
        aLV.InitializeFromArray FilterArray2d(TargetTable.Range.Value, True, _
                                              TextBoxFilterValue.Value, _
                                              operator.IS_LIKE, _
                                              ComboBoxFilterColumn.Value)
    Case Is > 1
        If Len(TextBoxFilterValue.Value) > previousInputLength Then 'filter the current table's array for better speed
            aLV.InitializeFromArray FilterArray2d(aLV.Value, True, _
                                                  TextBoxFilterValue.Value, _
                                                  operator.IS_LIKE, _
                                                  ComboBoxFilterColumn.Value)
        Else
            previousInputLength = Len(TextBoxFilterValue.Value)     'filter the original table array
            aLV.InitializeFromArray FilterArray2d(TargetTable.Range.Value, True, _
                                                  TextBoxFilterValue.Value, _
                                                  operator.IS_LIKE, _
                                                  ComboBoxFilterColumn.Value)
        End If
    End Select
    CreateEditorControls
End Sub

'/////////////////////////////////////////////
'///            EDITORS EVENTS             ///
'/////////////////////////////////////////////

Private Sub Emitter_Focus(control As Object)
    If Not control.Name Like "Editor-*" Then Exit Sub
    editorIndex = Split(control.Name, "-")(1)
    'format editor's label to indicate focus
    Dim Label As MSForms.Label
    Set Label = Me.Controls("lbl-" & editorIndex)
    Label.BackColor = RGB(80, 200, 120)
    'if multiline textbox then make it taller
    ResizeIfNeeded control
    'if we focused by tabbing, remove the tab (we're using _KeyDown event, read that)
    If TypeName(control) = "TextBox" Then control.Value = Replace(control.Value, vbTab, "")
    If TypeName(control) = "ComboBox" Then control.Text = Replace(control.Text, vbTab, "")
    'if it is a date field then optionally show date picker
    If CheckBoxDatePicker.Value = True And UCase(Label.Caption) Like "*DATE*" Then
        Dim retVal As String: retVal = uCalendar.Datepicker
        If retVal <> "" Then control.Value = Format(Replace(retVal, ".", "/"), TableCell(control).NumberFormat)
    End If
    'display datavalidation information for respective cell
    ShowDatavalidationBanner control, Label
End Sub

Sub ResizeIfNeeded(control As Object)
    If Not TypeName(control) = "TextBox" Then Exit Sub
    control.ZOrder (fmTop)
    Dim dif As Long
    dif = CountOfCharacters(control.Text, vbNewLine) + 1
    Dim targetHeight
    targetHeight = WorksheetFunction.Min(FrameTableEditor.Height - control.Top - 12, 12 * dif)
    targetHeight = WorksheetFunction.Max(targetHeight, 18)
    If control.Height = targetHeight Then Exit Sub
    control.Height = targetHeight
    control.ScrollBars = fmScrollBarsBoth
    'change backcolor to make it more distinguishable as it is over other controls
    control.BackColor = RGB(255, 255, 204)
End Sub

Function TableCell(control As Object) As Range
    Set TableCell = TargetTable.DataBodyRange(ListViewTable.selectedItem.index, CInt(Split(control.Name, "-")(1)))
End Function

Sub ShowDatavalidationBanner(control As Object, Banner As Object)
    On Error Resume Next
    Dim dataValidation As Validation:   Set dataValidation = TableCell(control).Validation
    Dim ValidationType As XlDVType:     ValidationType = dataValidation.Type
    On Error GoTo 0
    If ValidationType = 0 Then Exit Sub
    Dim validationFormula1  As String:                      validationFormula1 = dataValidation.Formula1
    Dim validationFormula2  As String:                      validationFormula2 = dataValidation.Formula2
    Dim operator            As XlFormatConditionOperator:   operator = dataValidation.operator
    Dim msg As String
    If DatavalidationTypeToString(ValidationType) = "List" Then
        msg = validationFormula1
    Else
        Select Case dataValidation.operator
            Case xlBetween
                msg = validationFormula1 & " < VALUE < " & validationFormula2
            Case xlNotBetween
                msg = validationFormula1 & " > VALUE < " & validationFormula2
            Case xlEqual, xlNotEqual, xlGreater, xlLess, xlGreaterEqual, xlLessEqual
                msg = "VALUE " & OperatorToString(operator) & " " & validationFormula1
         End Select
     End If
     msg = "(" & msg & ")"
     Set targetBanner = Banner
     Banner.Tag = Banner.Caption
     Banner.Caption = Banner.Caption & Space(4) & "-" & Space(4) & _
                      "[" & DatavalidationTypeToString(ValidationType) & "]" & Space(4) & "-" & Space(4) & msg
End Sub

Public Function OperatorToString(operator As XlFormatConditionOperator) As String
    OperatorToString = CStr(aSwitch(operator, _
                            xlBetween, "<<", _
                            xlNotBetween, "><", _
                            xlEqual, "=", _
                            xlNotEqual, "<>", _
                            xlGreater, ">", _
                            xlLess, "<", _
                            xlGreaterEqual, ">=", _
                            xlLessEqual, "<="))
End Function

Public Function DatavalidationTypeToString(dvType As XlDVType)
    Select Case dvType
    Case xlValidateInputOnly:   DatavalidationTypeToString = "Value Change"
    Case xlValidateWholeNumber: DatavalidationTypeToString = "Whole Number"
    Case xlValidateDecimal:     DatavalidationTypeToString = "Decimal"
    Case xlValidateList:        DatavalidationTypeToString = "List"
    Case xlValidateDate:        DatavalidationTypeToString = "Date"
    Case xlValidateTime:        DatavalidationTypeToString = "Time"
    Case xlValidateCustom:      DatavalidationTypeToString = "Custom"
    Case xlValidateTextLength:  DatavalidationTypeToString = "Text Length"
    End Select
End Function

Private Sub Emitter_Blur(control As Object)
    If Not control.Name Like "Editor-*" Then Exit Sub
    'restore original view for the Editor and its Label
    Me.Controls("lbl-" & Split(control.Name, "-")(1)).BackColor = Me.BackColor
    If Not TypeName(control) = "TextBox" Then Exit Sub
'    If InStr(1, control.Text, vbNewLine) > 0 Then
        control.Height = 16
        control.ScrollBars = fmScrollBarsNone
        control.BackColor = vbWhite
'    End If
    If Not targetBanner Is Nothing Then targetBanner.Caption = targetBanner.Tag
End Sub

Private Sub Emitter_Keydown(control As Object, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If Not control.Name Like "Editor-*" Then Exit Sub
    ' we're using _KeyDown event to
    ' 1. capture the use tab or shift + tab and allow cycling from last to first or first to last ...
    editorIndex = -1
    If KeyCode = vbKeyTab Then
        editorIndex = Split(control.Name, "-")(1)
        If Shift = 1 Then
            editorIndex = IIf(editorIndex = 1, FrameTableEditor.Controls.Count / 2, editorIndex - 1)
        Else
            editorIndex = IIf(editorIndex = FrameTableEditor.Controls.Count / 2, 1, editorIndex + 1)
        End If
    Else
        Exit Sub
    End If
    If editorIndex = -1 Then Exit Sub
    ' 2. restore the view of the editor we're moving out from and it's label
    Emitter_Blur control
    ' ... continuing (1)
    Dim Editor As MSForms.control
    Set Editor = FrameTableEditor.Controls("Editor-" & editorIndex)
    ' if going backwards then scroll if needed to have the correct view
    If editorIndex < previousEditorIndex _
    Or (previousEditorIndex = 1 And editorIndex = FrameTableEditor.Controls.Count / 2) Then
        FrameTableEditor.ScrollTop = IIf(editorIndex = 1, 0, Controls("lbl-" & editorIndex).Top)
    End If
    Editor.SetFocus
    previousEditorIndex = editorIndex
End Sub

Private Sub Emitter_Change(control As Object)
    ' format for data validation pass/fail
    If InStr(1, control.Value, vbTab) > 0 Then Exit Sub
    ResizeIfNeeded control
    Dim cell As Range: Set cell = TableCell(control)
    If IsValueValidForCell(cell, control.Value) Then
        control.ForeColor = RGB(0, 100, 0)
        control.BorderColor = RGB(0, 100, 0)
    Else
        control.ForeColor = RGB(178, 34, 34)
        control.BorderColor = RGB(178, 34, 34)
    End If
End Sub

Public Function IsValueValidForCell(ByVal inputRange As Range, ByVal inputValue As Variant) As Boolean
    On Error Resume Next
    Dim dataValidation As Validation:   Set dataValidation = inputRange.Validation
    Dim ValidationType As XlDVType:     ValidationType = dataValidation.Type
    On Error GoTo 0
    If ValidationType = 0 Then
        IsValueValidForCell = True
    Else
        ' Check if the cell has data validation
        Dim validationFormula1 As String, validationFormula2 As String
        validationFormula1 = dataValidation.Formula1
        validationFormula2 = dataValidation.Formula2
        Dim operator As XlFormatConditionOperator:  operator = dataValidation.operator
        
        If inputValue = "" And Not dataValidation.IgnoreBlank Then
            IsValueValidForCell = False
        ElseIf ValidationType = xlValidateInputOnly Then
            ' Allow input only or custom validation types
            IsValueValidForCell = True
        ElseIf ValidationType = xlValidateCustom Then
            '@TODO consider whate this custom may be and try to evaluate it
            IsValueValidForCell = True
        Else
            ' Evaluate the validation formula with the input value
            Dim validationResult As Boolean
            Select Case ValidationType
                Case xlValidateWholeNumber
                    If IsNumeric(inputValue) Then
                        If CLng(inputValue) - Round(inputValue, 0) = 0 Then
                            validationResult = EvaluateComparison(CInt(inputValue), operator, _
                                                                  CInt(validationFormula1), _
                                                                  CInt(validationFormula2))
                        End If
                    End If
                Case xlValidateDecimal
                    If IsNumeric(inputValue) Then
                        validationResult = EvaluateComparison(CLng(inputValue), operator, _
                                                              CLng(validationFormula1), _
                                                              CLng(validationFormula2))
                    End If
                Case xlValidateTextLength
                    If validationFormula2 = "" Then
                        validationResult = EvaluateComparison(CInt(Len(inputValue)), operator, _
                                                              CInt(validationFormula1), _
                                                              0)
                    Else
                        validationResult = EvaluateComparison(CInt(Len(inputValue)), operator, _
                                                              CInt(validationFormula1), _
                                                              CInt(validationFormula2))
                    End If
                Case xlValidateDate
                    If IsDate(inputValue) Then
                        If validationFormula2 = "" Then
                            validationResult = EvaluateComparison(CDate(inputValue), operator, _
                                                                  CDate(validationFormula1), _
                                                                  "")
                        Else
                            validationResult = EvaluateComparison(CDate(inputValue), operator, _
                                                                  CDate(validationFormula1), _
                                                                  CDate(validationFormula2))
                        End If
                    End If
                Case xlValidateTime
                    If IsDate(inputValue) Then
                        If IsTime(CDate(inputValue)) Then
                            If validationFormula2 = "" Then
                                validationResult = EvaluateComparison(CDate(inputValue), operator, _
                                                                      CDate(validationFormula1), _
                                                                      "")
                            Else
                                validationResult = EvaluateComparison(CDate(inputValue), operator, _
                                                                      CDate(validationFormula1), _
                                                                      CDate(validationFormula2))
                            End If
                        End If
                    End If
                Case xlValidateList
                    Dim listValues As Variant
                    listValues = Split(Replace(validationFormula1, "=", ""), ",")
                    Dim i As Long
                    For i = LBound(listValues) To UBound(listValues)
                        listValues(i) = Trim(listValues(i))
                    Next i

                    validationResult = Not IsError(Application.match(inputValue, listValues, 0))
                Case xlValidateCustom
                    validationResult = True
                Case Else
                    validationResult = Application.Evaluate(inputValue & validationFormula1)
            End Select
            IsValueValidForCell = CBool(validationResult)
        End If
    End If
End Function

Public Function EvaluateComparison(ByVal inputValue As Variant, ByVal operator As XlFormatConditionOperator, _
                                   ByVal validationFormula1 As Variant, ByVal validationFormula2 As Variant) As Boolean
    If TypeName(inputValue) <> TypeName(validationFormula1) Then Exit Function
    Select Case operator
        Case xlBetween
            If TypeName(inputValue) <> TypeName(validationFormula2) Then Exit Function
            EvaluateComparison = (validationFormula1 <= inputValue And inputValue <= validationFormula2)
        Case xlNotBetween
            If TypeName(inputValue) <> TypeName(validationFormula2) Then Exit Function
            EvaluateComparison = (validationFormula1 > inputValue And inputValue < validationFormula2)
        Case xlEqual
            EvaluateComparison = (inputValue = validationFormula1)
        Case xlNotEqual
            EvaluateComparison = (inputValue <> validationFormula1)
        Case xlGreater
            EvaluateComparison = (inputValue > validationFormula1)
        Case xlLess
            EvaluateComparison = (inputValue < validationFormula1)
        Case xlGreaterEqual
            EvaluateComparison = (inputValue >= validationFormula1)
        Case xlLessEqual
            EvaluateComparison = (inputValue <= validationFormula1)
    End Select
End Function


'/////////////////////////////////////////////
'///            EDITORS COMMANDS           ///
'/////////////////////////////////////////////

Private Sub LabelSave_Click()
    ' check for data validation
    If Not PassValidation Then
        MsgBox "Failed to pass Data Validation"
        Exit Sub
    End If
    ' save depending on if we are new row or editing existing
    If isAddingNewRow Then
        SaveNewRow
    Else
        SaveChanges
    End If
End Sub

Function PassValidation() As Boolean
    PassValidation = True
    Dim val
    Dim cell As Range
    Dim i As Long
    For i = 1 To FrameTableEditor.Controls.Count / 2
        Set cell = TargetTable.DataBodyRange(ListViewTable.selectedItem.index, i)
        val = Controls("Editor-" & i).Value
        If Not IsValueValidForCell(cell, val) Then
            PassValidation = False
            Exit Function
        End If
    Next
End Function

Private Sub SaveChanges()
    TargetTable.ListRows(ListViewTable.selectedItem.index).Range.Value = EditorsValueArray
    UpdateListviewRow ListViewTable.selectedItem.index, EditorsValueArray
End Sub

Function EditorsValueArray()
    Dim out()
    ReDim out(1 To FrameTableEditor.Controls.Count / 2)
    Dim val
    Dim i As Long
    For i = 1 To FrameTableEditor.Controls.Count / 2
        val = FrameTableEditor.Controls("Editor-" & i).Value
        out(i) = val
    Next
    EditorsValueArray = out
End Function

Sub UpdateListviewRow(index As Long, newValues)
    Dim i As Long
    For i = 1 To ListViewTable.ColumnHeaders.Count
        If i = 1 Then
            ListViewTable.ListItems(index).Text = newValues(i)
        Else
            ListViewTable.ListItems(index).ListSubItems(i - 1).Text = newValues(i)
        End If
    Next
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

Private Sub LabelNew_Click()
    isAddingNewRow = True
    CreateEditorControls
End Sub

Private Sub LabelDelete_Click()
    'delete selected row
    If isAddingNewRow Then Exit Sub
    If MsgBox("Permanently delete this row?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    TargetTable.ListRows(ListViewTable.selectedItem.index).Delete
    ListViewTable.ListItems.Remove ListViewTable.selectedItem.index
End Sub

Private Sub LabelGoFirst_Click()
    FrameTableEditor.ScrollTop = 0
    FrameTableEditor.Controls("Editor-1").SetFocus
    If editorIndex > 0 Then Emitter_Blur Controls("Editor-" & editorIndex)
    Emitter_Focus Controls("Editor-1")
End Sub

Private Sub LabelGoLast_Click()
    FrameTableEditor.ScrollTop = Controls("lbl-" & (FrameTableEditor.Controls.Count / 2)).Top
    FrameTableEditor.Controls("Editor-" & (FrameTableEditor.Controls.Count / 2)).SetFocus
    If editorIndex > 0 Then Emitter_Blur Controls("Editor-" & editorIndex)
    Emitter_Focus Controls("Editor-" & (FrameTableEditor.Controls.Count / 2))
End Sub


'/////////////////////////////////////////////
'///            RESIZING                   ///
'/////////////////////////////////////////////

' @TODO consider resizing by dragging

Private Sub LabelResize1_Click()
    Select Case LabelResize1.Tag
    Dim i As Long
    Case "Left"
        LabelResize1.Picture = LoadPicture(ThisWorkbook.Path & "\img\right.jpg")
        LabelResize1.Tag = "Right"
        FrameTablePath.Width = LabelResize1.Width + 6
    Case "Right"
        LabelResize1.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg")
        LabelResize1.Tag = "Left"
        FrameTablePath.Width = ListboxWorkbook.Left + ListboxWorkbook.Width + 6
    End Select
    FrameTableView.Left = FrameTablePath.Left + FrameTablePath.Width
    FrameTableEditor.Left = FrameTableView.Left + FrameTableView.Width
    FrameTableEditorBottom.Left = FrameTableEditor.Left
    
    aUserform.Init(Me).ResizeToFitControls (6)
End Sub

Private Sub LabelResize2_Click()
    Select Case LabelResize2.Tag
    Dim i As Long
    Case "Left"
        LabelResize2.Picture = LoadPicture(ThisWorkbook.Path & "\img\right.jpg")
        LabelResize2.Tag = "Right"
        FrameTableView.Width = LabelResize2.Width + 6
    Case "Right"
        LabelResize2.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg")
        LabelResize2.Tag = "Left"
        FrameTableView.Width = 540
    End Select
    FrameTableEditor.Left = FrameTableView.Left + FrameTableView.Width
    FrameTableEditorBottom.Left = FrameTableEditor.Left
    
    aUserform.Init(Me).ResizeToFitControls (6)
End Sub

Private Sub LabelResize3_Click()
    Select Case LabelResize3.Tag
    Dim i As Long
    Case "Left"
        LabelResize3.Picture = LoadPicture(ThisWorkbook.Path & "\img\right.jpg")
        LabelResize3.Tag = "Right"
        FrameTableEditor.Width = LabelResize3.Width + 6
    Case "Right"
        LabelResize3.Picture = LoadPicture(ThisWorkbook.Path & "\img\left.jpg")
        LabelResize3.Tag = "Left"
        FrameTableEditor.Width = 354
    End Select
    FrameTableEditorBottom.Width = FrameTableEditor.Width
    aUserform.Init(Me).ResizeToFitControls (6)
End Sub


'/////////////////////////////////////////////
'///            OTHER                      ///
'/////////////////////////////////////////////

Private Sub GetInfo_Click()
    uAuthor.Show
End Sub





























