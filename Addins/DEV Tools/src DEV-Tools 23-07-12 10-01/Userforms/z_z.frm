VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} z_z 
   Caption         =   "UserForm1"
   ClientHeight    =   6756
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14028
   OleObjectBlob   =   "z_z.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "z_z"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TargetWorkbook As Workbook
Private TargetWorksheet As Worksheet
Private TargetTable As ListObject
Private aLV As aListView

Private byPass As Boolean

Private WithEvents myForm As UserForm
Attribute myForm.VB_VarHelpID = -1
Private currentSelection As String

' Custom event for ListView selection change
Public Event ListViewSelectionChanged(ByVal selectedItem As MSComctlLib.ListItem)

Private Sub ListView1_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    Dim selectedItem As MSComctlLib.ListItem
    
    On Error Resume Next
    Set selectedItem = ListView1.selectedItem
    On Error GoTo 0
    
    If selectedItem Is Nothing Then Exit Sub
    
    Dim newSelection As String
    newSelection = "|" & Join(aLV.SelectionArray, "|")
    
    If Not newSelection = currentSelection Then
        currentSelection = newSelection
        RemoveFrameControls
        AddFrameControls
    End If
End Sub


Private Sub ListView1_DblClick()
    byPass = False
End Sub

Private Sub ListView1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    RemoveFrameControls
    AddFrameControls
    
    If byPass Then Exit Sub
    
    Dim targetColumn As Long:   targetColumn = aLV.ClickedColumn(x, y)
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
    Set Editor = Frame1.Controls("FrameControl" & IIf(targetColumn = -1, 1, targetColumn + 1))
    Editor.SetFocus
    Editor.SelStart = 0
    Editor.SelLength = 500
    
    byPass = True
End Sub

Private Sub UserForm_Initialize()
    Set aLV = aListView.Init(ListView1)
    ListView1.Gridlines = True
    ListView1.FullRowSelect = True
    ListView1.HideSelection = False
    aListBox.Init(ListBox1).LoadVBProjects
    Set myForm = Me
End Sub

Public Sub SaveChanges()
    Dim i As Long, tableCell As Range, newVal
    For i = 1 To Frame1.Controls.count
        Set tableCell = TargetTable.DataBodyRange(ListView1.selectedItem.index, i)
        newVal = Frame1.Controls("FrameControl" & i).Value
        tableCell.Value = newVal
    Next
End Sub

Private Sub listbox1_change()
    If ListBox1.ListIndex = -1 Then Exit Sub
    Dim ListboxValue As String
    ListboxValue = ListBox1.List(ListBox1.ListIndex)

    Set TargetWorkbook = Workbooks(ListboxValue)
    ListBox2.Clear
    ListBox3.Clear
    aLV.Clear
    Dim ws As Worksheet
    For Each ws In TargetWorkbook.Worksheets
        If ws.ListObjects.count > 0 Then
            ListBox2.AddItem ws.Name
        End If
    Next

End Sub

Private Sub listbox2_change()
    If ListBox2.ListIndex = -1 Then Exit Sub
    Dim ListboxValue As String
    ListboxValue = ListBox2.List(ListBox2.ListIndex)
    Set TargetWorksheet = TargetWorkbook.Worksheets(ListboxValue)
    ListBox3.Clear
    Dim lo As ListObject
    For Each lo In TargetWorksheet.ListObjects
        ListBox3.AddItem lo.Name
    Next
End Sub

Private Sub listbox3_change()
    If ListBox3.ListIndex = -1 Then Exit Sub
    Dim ListboxValue As String
    ListboxValue = ListBox3.List(ListBox3.ListIndex)
    Set TargetTable = TargetWorksheet.ListObjects(ListboxValue)
    aLV.InitializeFromArray TargetTable.Range.Value
    RemoveFrameControls
    
End Sub

Sub RemoveFrameControls()
Dim ctrl As MSForms.control
    If Frame1.Controls.count > 0 Then
'        Frame1.Visible = False
        For Each ctrl In Frame1.Controls
            Frame1.Controls.Remove ctrl.Name
        Next
'        Frame1.Visible = True
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

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub


Sub AddFrameControls()
    Dim i As Long, lbl As MSForms.Label, txt As MSForms.Textbox, cbx As MSForms.ComboBox
    Dim tableCell As Range
    Dim validationArray
    For i = 1 To ListView1.ColumnHeaders.count
        Set tableCell = TargetTable.DataBodyRange(ListView1.selectedItem.index, i)
        Set lbl = Frame1.Controls.Add("Forms.Label.1")
        lbl.Width = 60
        lbl.Height = 18
        lbl.Left = 6
        lbl.Top = AvailableFormOrFrameRow(Frame1, , , 3)
        lbl.Caption = ListView1.ColumnHeaders(i).TEXT
        If Not isValidation(tableCell) Then
            Set txt = Frame1.Controls.Add("Forms.Textbox.1")
            txt.Top = lbl.Top
            txt.Left = lbl.Left + lbl.Width + 6
            txt.Width = 120
            txt.Height = lbl.Height
            txt.Name = "FrameControl" & i
        Else
            Set cbx = Frame1.Controls.Add("Forms.ComboBox.1")
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
            cbx.Top = lbl.Top
            cbx.Left = lbl.Left + lbl.Width + 6
            cbx.Width = 120
            cbx.Height = lbl.Height
            cbx.Name = "FrameControl" & i
        End If
    Next
    For i = 1 To ListView1.ColumnHeaders.count
        Set tableCell = TargetTable.DataBodyRange(ListView1.selectedItem.index, i)
        Frame1.Controls("FrameControl" & i).Value = tableCell.Value
    Next
End Sub
