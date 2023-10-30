VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uFinder 
   Caption         =   "Finder"
   ClientHeight    =   9456.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16296
   OleObjectBlob   =   "uFinder.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim moResizer As New CFormResizer

Private Sub clear_Click()
    IS_LIKE.Value = True
    FirstComp.Text = "*"
    SecondComp.Text = ""
    FindThis.Text = ""
    offR.Text = 0
    offC.Text = 0
    ListBox1.clear
    TextBox3.Text = ""
    FindThis.SetFocus
End Sub

Private Sub cLookAtRow_Click()
    TrySearch
End Sub

Private Sub GetInfo_Click()
uAuthor.Show
End Sub

Private Sub HideEmptyColumns_Click()
    TrySearch
End Sub




Private Sub ListBox1_Change()

DoEvents
If FindThis = "" Then Exit Sub
    Dim s As String
    s = ListBox1.list(Listbox_Selected(ListBox1, 2), 3)
    TextBox3.Value = s
    
'
'    Dim sStart As Long
'    sStart = InStr(1, UCase(s), UCase(Replace(FindThis.Value, "ò", "Ó"))) - 1
'    Dim sLen As Long
'    sLen = Len(FindThis.Value)
'    On Error GoTo EH
'    If TextBox3 <> "" Then
'        With TextBox3
'            .SelStart = sStart
'            .SelLength = sLen
'            .SetFocus
'        End With
'    End If
'EH:
End Sub


Private Sub FirstComp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
        DoEvents
        ListBox1.ListIndex = 0
    End If
End Sub

Private Sub oCurrentRegion_Click()
    TrySearch
End Sub

Private Sub oUsedRange_Click()
    TrySearch
End Sub

Private Sub oVBLF_Click()
    TrySearch
End Sub

Private Sub SecondComp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
        DoEvents
        ListBox1.ListIndex = 0
    End If
End Sub

Private Sub SpinButton1_Change()
    offC.Value = SpinButton1.Value
End Sub

Private Sub SpinButton2_Change()
    offR.Value = SpinButton2.Value
End Sub

Private Sub UserForm_Activate()
    Set moResizer.FORM = Me
    CreateListboxHeader ListBox1, ListBox2, Array("BOOK", "SHEEET", "RANGE", "ROW", "FORMULA")
    FindThis.SetFocus
'    on eror resume next
    FindThis.SelStart = 0
    FindThis.SelLength = Len(FindThis.Text)
End Sub

Private Sub UserForm_Initialize()
    LoadUserformOptions Me, Array("listbox1", "listbox2", "textbox3")
    AddMinimizeButtonToUserform Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveUserformOptions Me
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
    moResizer.FormResize
End Sub
Sub CreateListboxHeader(body As MSForms.ListBox, header As MSForms.ListBox, arrHeaders)

    header.Width = body.Width
    Dim i As Long
    'must have a listbox to use as headers
    header.ColumnCount = body.ColumnCount
    header.ColumnWidths = body.ColumnWidths
    'add headerelements
    header.clear
    header.AddItem
    
    If ArrayDimensions(arrHeaders) = 1 Then
        For i = 0 To UBound(arrHeaders)
            'make it prety
            header.list(0, i) = arrHeaders(i)
        Next i
    Else
        For i = 1 To UBound(arrHeaders, 2)
            header.list(0, i - 1) = arrHeaders(1, i)
        Next i
    End If
    body.ZOrder (1)
    header.ZOrder (0)
    header.SpecialEffect = fmSpecialEffectFlat
    header.BackColor = &H403636   'RGB(200, 200, 200)
    'align header to body
    header.Height = 15
    header.Width = body.Width
    header.Left = body.Left
    header.Top = body.Top - header.Height - 1
    header.Font.Bold = True
    header.Font.Name = "Comic Sans MS"
    header.Font.Size = 9
    header.ForeColor = vbWhite
'    header.BackColor =
End Sub
Private Sub FindThis_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
        DoEvents
        ListBox1.ListIndex = 0
    End If
End Sub

Function ContainsFilter(ws As Worksheet) As Boolean
    Filtered.Caption = ""
    If ws.FilterMode = True Then Filtered.Caption = "Some data is filtered"
End Function

Private Sub CommandButton1_Click()
    TrySearch
End Sub

Function GetSearchRange(ByRef RangeOrSheet) As Range
    Dim rng As Range
    If TypeName(RangeOrSheet) = "Range" Then
        If IsDate(FirstComp.Value) Then
            Set rng = FindDateRange(RangeOrSheet.Parent.Range(RangeOrSheet.Address))
        ElseIf IsNumeric(FirstComp.Value) Then
            Set rng = FindNumericRange(RangeOrSheet.Parent.Range(RangeOrSheet.Address))
        Else
            Set rng = RangeFindAll(RangeOrSheet.Parent.Range(RangeOrSheet.Address), FindThis)
        End If
    Else
        If IsDate(FirstComp.Value) Then
'            Set rng = FindDateRange(RangeOrSheet.UsedRange)
            Set rng = FindDateRange(Union(RangeOrSheet.Columns(FindThis), RangeOrSheet.UsedRange))
        ElseIf IsNumeric(FirstComp.Value) Then
'            Set rng = FindNumericRange(RangeOrSheet.UsedRange)
            Set rng = FindDateRange(Union(RangeOrSheet.Columns(FindThis), RangeOrSheet.UsedRange))
        Else
            Set rng = RangeFindAll(RangeOrSheet.UsedRange, FindThis)
        End If
    End If
    Set GetSearchRange = rng
End Function
Sub TrySearch()
    If IsNumeric(FindThis.Value) Then
'        Select Case FirstComp.Text
'        Case "", "*"
'            FirstComp.Text = FindThis.Text
'        End Select
    ElseIf FirstComp.Text = "" Then
        FirstComp.Text = "*"
    End If
    
    Dim op As OPERATOR
    op = WhichOperator(Frame2)

    Dim TargetWorkbook As Workbook
    Dim TargetWorkSheet As Worksheet
    Dim element As Variant
    Dim str As String
    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    Dim varCompare
    Dim varRange()
    ReDim Preserve varRange(1 To ListBox1.ColumnCount, 1 To 1)
    
    Dim FirstCell As Range
    
    Select Case whichOption(Frame1, "OptionButton").Caption
    
        Case Is = "Selection"
            ContainsFilter Selection.Parent
            If TypeName(Selection) <> "Range" Then
                MsgBox "Under current options, please select a Range."
                Exit Sub
            End If

            Set rng = GetSearchRange(Selection)
            
            Dim s As String
            If Not rng Is Nothing Then
                
                For Each cell In rng
                    testMatch cell, i, varRange, op
                Next
            End If
    
        Case Is = "Active Sheet"
            ContainsFilter ActiveSheet
            Set TargetWorkSheet = ActiveSheet
            
            Set rng = GetSearchRange(TargetWorkSheet)
            
            If Not rng Is Nothing Then
                For Each cell In rng
                    If IsEmpty(varRange(1, UBound(varRange, 2))) Then
                        testMatch cell, i, varRange, op
                    Else
                        If cell.Row <> Split(varRange(3, UBound(varRange, 2)), "$")(2) Then
                            testMatch cell, i, varRange, op
                        End If
                    End If
                Next
            End If
        Case Is = "Active Book"
            Set TargetWorkbook = ActiveWorkbook
            For Each TargetWorkSheet In TargetWorkbook.Worksheets
                ContainsFilter TargetWorkSheet
'                Set rng = RangeFindAll(TargetWorkSheet.UsedRange, FindThis)
                Set rng = GetSearchRange(TargetWorkSheet)
                If Not rng Is Nothing Then
                    For Each cell In rng
                        If IsEmpty(varRange(1, UBound(varRange, 2))) Then
                            testMatch cell, i, varRange, op
                        Else
                            If cell.Row <> Split(varRange(3, UBound(varRange, 2)), "$")(2) Then
                                testMatch cell, i, varRange, op
                            End If
                        End If
                    Next
                End If
            Next
        Case Is = "All Books"
            For Each TargetWorkbook In Workbooks
                For Each TargetWorkSheet In TargetWorkbook.Worksheets
                ContainsFilter TargetWorkSheet
'                    Set rng = RangeFindAll(TargetWorkSheet.UsedRange, FindThis)
                    Set rng = GetSearchRange(TargetWorkSheet)
                    If Not rng Is Nothing Then
                        For Each cell In rng
                            If IsEmpty(varRange(1, UBound(varRange, 2))) Then
                                testMatch cell, i, varRange, op
                            Else
                                If cell.Row <> Split(varRange(3, UBound(varRange, 2)), "$")(2) Then
                                    testMatch cell, i, varRange, op
                                End If
                            End If
                        Next
                    End If
                Next
            Next
    End Select
    
    ListBox1.list = Transpose2DArray(varRange)
    Rem ResizeControlColumns ListBox1, False
    '        ResizeUserformToFitControls Me
    '    var = Split(ListBox1.ColumnWidths, ";")
    '    ListBox1.ColumnWidths = Join(Array(var(0), var(1), var(2)), ";") & ";1500"

End Sub
Function testMatch(ByRef cell As Range, ByRef i As Long, ByRef varRange As Variant, op As OPERATOR) As Boolean
'On Error GoTo hell
    Dim del As String
    Dim dupdel As String
    del = IIf(oVBLF.Value = True, vbNewLine, "|")
    dupdel = IIf(oVBLF.Value = True, vbNewLine & vbNewLine, "||")
    
    s = RangeToString(CellRow(cell), del)
    If HideEmptyColumns = True Then
        Do While InStr(1, s, dupdel) > 0
            s = Replace(s, dupdel, del)
        Loop
    End If
    
    If cLookAtRow.Value = True Then
        If InStr(1, UCase(s), UCase(FirstComp.Value)) > 0 Then Exit Function
    Else
        If compare(cell.Offset(CInt(offR), CInt(offC)).Value, op, _
                    IIf(FirstComp = "", FindThis.Value, FirstComp.Value), _
                    IIf(SecondComp = "", FirstComp.Value, SecondComp.Value)) = False _
                    Then Exit Function
    End If
    
    i = i + 1
    ReDim Preserve varRange(1 To ListBox1.ColumnCount, 1 To i)
    varRange(1, i) = cell.Parent.Parent.Name
    varRange(2, i) = cell.Parent.Name
    varRange(3, i) = cell.Address
    varRange(4, i) = s
    varRange(5, i) = IIf(cell.HasFormula, cell.Formula, "")
    testMatch = True
    If IsEmpty(varRange(1, i)) Then Exit Function
'hell:
End Function
Function WhichOperator(Frame As MSForms.Frame) As OPERATOR
    Dim op As OPERATOR

    Select Case whichOption(Frame, "OptionButton").Name
        Case Is = "IS_LIKE"
            op = OPERATOR.IS_LIKE
        Case Is = "IS_EQUAL"
            op = OPERATOR.IS_EQUAL
        Case Is = "NOT_EQUAL"
            op = OPERATOR.NOT_EQUAL
        Case Is = "IS_CONTAINS"
            op = OPERATOR.IS_CONTAINS
        Case Is = "NOT_CONTAINS"
            op = OPERATOR.NOT_CONTAINS
        Case Is = "STARTS_WITH"
            op = OPERATOR.STARTS_WITH
        Case Is = ".ENDS_WITH"
            op = OPERATOR.ENDS_WITH
        Case Is = "GREATER_THAN"
            op = OPERATOR.GREATER_THAN
        Case Is = "GREATER_OR_EQUAL"
            op = OPERATOR.GREATER_OR_EQUAL
        Case Is = "LESS_THAN"
            op = OPERATOR.LESS_THAN
        Case Is = "LESS_OR_EQUAL"
            op = OPERATOR.LESS_OR_EQUAL
        Case Is = "IS_BETWEEN"
            op = OPERATOR.IS_BETWEEN
        Case Is = "NOT_BETWEEN"
            op = OPERATOR.NOT_BETWEEN
    End Select
    WhichOperator = op
End Function

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Long
    i = ListBox1.ListIndex
    Dim wb As Workbook
    Set wb = Workbooks(ListBox1.list(i, 0))
    Dim ws As Worksheet
    Set ws = wb.Sheets(ListBox1.list(i, 1))
    Dim rng As Range
    Set rng = ws.Range(ListBox1.list(i, 2))
    ws.Activate
    rng.Select
End Sub

'
'Private Sub CommandButton1_Click()
'    If IsNumeric(FindThis.Value) Then
'        Select Case FirstComp.Text
'        Case "", "*"
'            FirstComp.Text = FindThis.Text
'        End Select
'    ElseIf FirstComp.Text = "" Then
'        FirstComp.Text = "*"
'    End If
'
'    Dim op As OPERATOR
'    op = WhichOperator(Frame2)
'
'    Dim TargetWorkbook As Workbook
'    Dim TargetWorkSheet As Worksheet
'    Dim element As Variant
'    Dim str As String
'    Dim rng As Range
'    Dim cell As Range
'    Dim i As Long
'    Dim varCompare
'    Dim varRange()
'    ReDim Preserve varRange(1 To ListBox1.ColumnCount, 1 To 1)
'
'    Select Case whichOption(Frame1, "OptionButton").Caption
'
'        Case Is = "Selection"
'            ContainsFilter Selection.Parent
'            If TypeName(Selection) <> "Range" Then
'                MsgBox "Under current options, please select a Range."
'                Exit Sub
'            End If
'
'
'            If IsDate(FirstComp.Value) Then
'                Set rng = FindDateRange(Selection) '(Selection.SpecialCells(xlCellTypeConstants))
'            ElseIf IsNumeric(FirstComp.Value) Then
'                Set rng = FindNumericRange(Selection) '(Selection.SpecialCells(xlCellTypeConstants))
'            Else
'                Set rng = RangeFindAll(Selection, FindThis) '(Selection.SpecialCells(xlCellTypeConstants), FindThis)
'            End If
'
'            Dim s As String
'            If Not rng Is Nothing Then
'                For Each cell In rng
'                    If compare(cell.Offset(CInt(offR), CInt(offC)).Value, op, IIf(FirstComp = "", FindThis.Value, FirstComp.Value), IIf(SecondComp = "", FirstComp.Value, SecondComp.Value)) = True Then
'                        i = i + 1
'                        ReDim Preserve varRange(1 To ListBox1.ColumnCount, 1 To i)
'                        varRange(1, i) = cell.Parent.Parent.Name
'                        varRange(2, i) = cell.Parent.Name
'                        varRange(3, i) = cell.Address
'                        s = RangeToString(CellRow(cell), "|")
'                        If HideEmptyColumns = True Then
'                            Do While InStr(1, s, "||") > 0
'                                s = Replace(s, "||", "|")
'                            Loop
'                        End If
'                        varRange(4, i) = s
'                        varRange(5, i) = IIf(cell.HasFormula, cell.Formula, "")
'                    End If
'                Next
'            End If
'
'        Case Is = "Active Sheet"
'            ContainsFilter ActiveSheet
'            Set TargetWorkSheet = ActiveSheet
'            If IsDate(FindThis.Value) Then
'                Set rng = FindDateRange(TargetWorkSheet.UsedRange)
'            ElseIf IsNumeric(FindThis.Value) Then
'                Set rng = FindNumericRange(TargetWorkSheet.UsedRange)
'            Else
'                Set rng = RangeFindAll(TargetWorkSheet.UsedRange, FindThis)
'            End If
'
'            If Not rng Is Nothing Then
'                For Each cell In rng
'                    If compare(cell.Offset(offR, offC).Value, op, IIf(FirstComp = "", FindThis.Value, FirstComp.Value), IIf(SecondComp = "", FirstComp.Value, SecondComp.Value)) = True Then
'                        i = i + 1
'                        ReDim Preserve varRange(1 To ListBox1.ColumnCount, 1 To i)
'                        varRange(1, i) = cell.Parent.Parent.Name
'                        varRange(2, i) = cell.Parent.Name
'                        varRange(3, i) = cell.Address
'                        s = RangeToString(CellRow(cell), "|")
'                        If HideEmptyColumns = True Then
'                            Do While InStr(1, s, "||") > 0
'                                s = Replace(s, "||", "|")
'                            Loop
'                        End If
'                        varRange(4, i) = s
'                        varRange(5, i) = IIf(cell.HasFormula, cell.Formula, "")
'                    End If
'                Next
'            End If
'        Case Is = "Active Book"
'            Set TargetWorkbook = ActiveWorkbook
'            For Each TargetWorkSheet In TargetWorkbook.Worksheets
'                ContainsFilter TargetWorkSheet
'                Set rng = RangeFindAll(TargetWorkSheet.UsedRange, FindThis)
'                If Not rng Is Nothing Then
'                    For Each cell In rng
'                        If compare(cell.Offset(offR, offC).Value, op, IIf(FirstComp = "", FindThis.Value, FirstComp.Value), IIf(SecondComp = "", FirstComp.Value, SecondComp.Value)) = True Then
'                            i = i + 1
'                            ReDim Preserve varRange(1 To ListBox1.ColumnCount, 1 To i)
'                            varRange(1, i) = cell.Parent.Parent.Name
'                            varRange(2, i) = cell.Parent.Name
'                            varRange(3, i) = cell.Address
'                            s = RangeToString(CellRow(cell), "|")
'                            If HideEmptyColumns = True Then
'                                Do While InStr(1, s, "||") > 0
'                                    s = Replace(s, "||", "|")
'                                Loop
'                            End If
'                            varRange(4, i) = s
'                            varRange(5, i) = IIf(cell.HasFormula, cell.Formula, "")
'                        End If
'                    Next
'                End If
'            Next
'        Case Is = "All Books"
'            For Each TargetWorkbook In Workbooks
'                For Each TargetWorkSheet In TargetWorkbook.Worksheets
'                ContainsFilter TargetWorkSheet
'                    Set rng = RangeFindAll(TargetWorkSheet.UsedRange, FindThis)
'                    If Not rng Is Nothing Then
'                        For Each cell In rng
'                            If compare(cell.Offset(offR, offC).Value, op, IIf(FirstComp = "", FindThis.Value, FirstComp.Value), IIf(SecondComp = "", FirstComp.Value, SecondComp.Value)) = True Then
'                                i = i + 1
'                                ReDim Preserve varRange(1 To ListBox1.ColumnCount, 1 To i)
'                                varRange(1, i) = cell.Parent.Parent.Name
'                                varRange(2, i) = cell.Parent.Name
'                                varRange(3, i) = cell.Address
'                                s = RangeToString(CellRow(cell), "|")
'                                If HideEmptyColumns = True Then
'                                    Do While InStr(1, s, "||") > 0
'                                        s = Replace(s, "||", "|")
'                                    Loop
'                                End If
'                                varRange(4, i) = s
'                                varRange(5, i) = IIf(cell.HasFormula, cell.Formula, "")
'                            End If
'                        Next
'                    End If
'                Next
'            Next
'    End Select
'
'    ListBox1.list = Transpose2DArray(varRange)
'    Rem ResizeControlColumns ListBox1, False
''        ResizeUserformToFitControls Me
''    var = Split(ListBox1.ColumnWidths, ";")
''    ListBox1.ColumnWidths = Join(Array(var(0), var(1), var(2)), ";") & ";1500"
'
'End Sub
