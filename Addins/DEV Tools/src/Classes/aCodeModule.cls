VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aCodeModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aCodeModule
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 04-05-2023 12:17    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Dim oWorkbook As Workbook
Dim oModule   As VBComponent
Private cm    As CodeModule

Public ColumnFirst As Long
Public ColumnLast  As Long
Public RowCount    As Long
Public RowFirst    As Long
Public RowLast     As Long

Public Selection          As String
Public SelectionAfter     As String
Public SelectionBefore    As String
Public SelectionLength    As Long
Public SelectionLinesCode As String

Sub ImportProcedure()
    aProcedure.Active.ImportProcedure Me.Selection, False
End Sub

Private Sub Class_Initialize()
    Init ActiveModule
End Sub

Public Function Active() As aCodeModule
    Init ActiveModule
    Set Active = Me
End Function

Public Function Init(Module As VBComponent) As aCodeModule
    Set oModule = Module
    Set cm = oModule.CodeModule
    Set oWorkbook = WorkbookOfModule(Module)
    AssignVariables
    Set Init = Me
End Function

Private Sub Class_Terminate()
    'do something

End Sub

Private Sub AssignVariables()

    Dim startLine As Long, StartColumn As Long, endLine As Long, EndColumn As Long
    cm.CodePane.GetSelection startLine, StartColumn, endLine, EndColumn
    
    ColumnFirst = StartColumn
    ColumnLast = EndColumn
    
    RowFirst = startLine
    RowLast = endLine
    If RowLast > startLine And EndColumn = 1 Then
        RowLast = endLine - 1
    End If
    RowCount = RowLast - RowFirst + 1
    
    SelectionLinesCode = cm.Lines(RowFirst, RowCount)
    

    If ColumnLast - ColumnFirst = 0 Then
        If endLine > RowLast Then
            ColumnLast = Len(cm.Lines(RowLast, 1)) + 1
        End If
    End If
    
    SelectionAfter = Mid(cm.Lines(RowLast, 1), ColumnLast) ' + IIf(Len(Me.Selection) > 0, 1, 0))
    SelectionBefore = Mid(cm.Lines(RowFirst, 1), 1, ColumnFirst - 1)
    
    If RowCount = 1 And (ColumnLast - ColumnFirst > 0) Then
        SelectionLength = ColumnLast - ColumnFirst
    Else
        SelectionLength = Len(SelectionLinesCode) - Len(SelectionBefore) - Len(SelectionAfter) - 1 - Len(CountOfCharacters(Selection, vbNewLine)) * 2
    End If
    
    If RowLast - RowFirst = 0 Then
        Me.Selection = Mid(cm.Lines(RowFirst, 1), ColumnFirst, ColumnLast - ColumnFirst)
        Exit Sub
    End If
    Dim str As String
    Dim i As Long
    For i = RowFirst To RowLast
        If str = "" Then
            str = Mid(cm.Lines(i, 1), ColumnFirst)
        ElseIf i < RowLast Then
            str = str & vbNewLine & cm.Lines(i, 1)
        Else
            str = str & vbNewLine & Left(cm.Lines(i, 1), ColumnLast - 1)
        End If
    Next
    Me.Selection = str

End Sub

Sub Todo()
    Const tFormat = "dd-mm-yyyy hh:nn"
    Dim Module As VBComponent
    Set Module = oModule
    Dim txtName As String
        txtName = AUTHOR_NAME
    If txtName = vbNullString Then txtName = Environ("UserName")
    Dim nLine As Long
        nLine = aCodeModule.Init(Module).RowFirst
        Do While nLine > 1 And Right(Module.CodeModule.Lines(nLine, 1), 1) = "_"
            nLine = nLine - 1
        Loop
    Dim sLine As String
        sLine = Module.CodeModule.Lines(nLine, 1)
    Dim sSpec As String
        sSpec = Space(Len(sLine) - Len(Trim(sLine)))
    Dim sTodo As String
        sTodo = sSpec & "'* @TODO Created: " & _
                VBA.Format$(Now, tFormat) & _
                " Author: " & txtName & vbCrLf & _
                sSpec & "'* @TODO "
        Module.CodeModule.InsertLines nLine, vbNewLine & sTodo & vbNewLine
End Sub

Sub ToggleComments()
    Dim sLine As String
    Dim i As Long
    For i = RowFirst To RowLast
        sLine = cm.Lines(i, 1)
        If IsCommentLine(sLine) Then
            If Trim(sLine) Like "Rem*" Then
                UnremTargetLine i
            ElseIf Trim(sLine) Like "'*" Then
                UncommentTargetLine i
            End If
        Else
            CommentTargetLine i
        End If
    Next
    SetSelection RowFirst, 1, RowLast, 10000
End Sub





Public Sub DimSeparate()
    Dim coll As New Collection
    Dim Code As String:         Code = SelectionLinesCode
    Dim FirstLine As Long:      FirstLine = RowFirst
    Dim iSpace As Long
    Dim iCount As Long:         iCount = RowCount
    
    If iCount = 1 Then
        iSpace = IndentationCount(cm.Lines(FirstLine, 1))
        Code = cm.Lines(FirstLine, iCount)
        Code = Join(Split(Code, ", "), vbNewLine & Space(iSpace) & "Dim ")
    Else
        Dim element
        For Each element In Split(Code, vbNewLine)
            iSpace = IndentationCount(CStr(element))
            coll.Add Join(Split(element, ", "), vbNewLine & Space(iSpace) & "Dim ")
        Next
        Code = aCollection.Init(coll).ToString(vbNewLine)
    End If
    Code = StringFormatAlignRowsElements(Code, " As ", True)

    cm.DeleteLines FirstLine, iCount
    cm.InsertLines FirstLine, Code
    SetSelection FirstLine, 1, FirstLine + UBound(Split(Code, vbNewLine)), 10000
End Sub

Sub DimMerge()
    Dim Code      As String:   Code = SelectionLinesCode
    Dim FirstLine As Long:     FirstLine = RowFirst
    Dim iCount    As Long:     iCount = RowCount
    Dim iSpace    As Long:     iSpace = IndentationCount(cm.Lines(FirstLine, 1))
    Code = cm.Lines(FirstLine, iCount)
    If Code Like "*" & vbNewLine Then Code = Left(Code, Len(Code) - Len(vbNewLine))
    Code = Replace(Join(Split(Code, vbNewLine)), "Dim ", ", ")
    Code = Replace(Code, ",", "Dim", , 1)
    Code = Space(iSpace) & RegExpReplace(Code, "\s{2,}", " ")
    Code = Replace(Code, " ,", ",")
    Code = Mid(Code, 2)
    cm.DeleteLines FirstLine, iCount
    cm.InsertLines FirstLine, Code
    SetSelection FirstLine, 1, FirstLine, 10000
End Sub

Public Sub SetSelection(startLine As Long, StartColumn As Long, endLine As Long, EndColumn As Long)
    cm.CodePane.SetSelection startLine, StartColumn, endLine, EndColumn
    AssignVariables
End Sub

Public Sub Move_Up()
    Dim blockStart        As Long:      blockStart = RowFirst
    Dim blockEnd          As Long:      blockEnd = RowLast
    Dim blockCountOfLines As Long:      blockCountOfLines = blockEnd - blockStart + 1
    Dim blockString       As String:    blockString = cm.Lines(blockStart, blockCountOfLines)
    Dim insertBlockAtLine As Long:      insertBlockAtLine = blockStart - 1
    Dim AP                As String:    AP = ActiveProcedure
    If AP = "" Then Exit Sub
    Select Case insertBlockAtLine
        Case 1, aProcedure.Init(, , AP).LineIndex(Declaration_First)
            Exit Sub
    End Select
    cm.DeleteLines blockStart, blockCountOfLines
    cm.InsertLines insertBlockAtLine, blockString
    SetSelection insertBlockAtLine, 1, insertBlockAtLine + blockCountOfLines - 1, 10000
End Sub

Public Sub Move_Down()
    Dim blockStart        As Long:      blockStart = RowFirst
    Dim blockEnd          As Long:      blockEnd = RowLast
    Dim blockCountOfLines As Long:      blockCountOfLines = blockEnd - blockStart + 1
    Dim blockString       As String:    blockString = cm.Lines(blockStart, blockCountOfLines)
    Dim insertBlockAtLine As Long:      insertBlockAtLine = blockStart + 1
    Dim AP                As String:    AP = ActiveProcedure
    If AP = "" Then Exit Sub
    Select Case insertBlockAtLine + blockCountOfLines - 1
        Case cm.CountOfLines, aProcedure.Init(, , AP).LineIndex(Procedure_Last)
            Exit Sub
    End Select
    cm.DeleteLines blockStart, blockCountOfLines
    cm.InsertLines insertBlockAtLine, blockString
    SetSelection insertBlockAtLine, 1, insertBlockAtLine + blockCountOfLines - 1, 10000
End Sub

Public Sub Inject(str As String)
    Dim Code    As String: Code = SelectionBefore & str & SelectionAfter
    Dim endLine As Long
    'catch error when selected all lines
    If RowCount > cm.CountOfLines Then endLine = endLine - 1
    cm.DeleteLines RowFirst, RowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, _
                 ColumnFirst, _
                 RowFirst + UBound(Split(str, vbLf)), _
                 IIf((UBound(Split(str, vbLf))) = 0, ColumnLast, 1) + Len(Split(str, vbLf)(UBound(Split(str, vbLf))))
End Sub

Sub InjectTemplateModule()
    Inject TemplateModule
End Sub

Private Function TemplateModule(Optional Module As VBComponent) As String
    Dim ComponentName As String: ComponentName = oModule.Name
    Dim s As String
    s = s & DevInfo & vbNewLine & vbNewLine
    s = s & "'   Version:    <Last Update Date goes here>" & vbNewLine
    s = s & "'   Description: General purpose library included in all projects" & vbNewLine & vbNewLine
    s = s & "'   Changelog" & vbNewLine
    s = s & "'   Date" & vbTab & vbTab & "Modification" & vbNewLine
    s = s & "'   " & Format(Date, "dd/mm/yy") & vbTab & "Initial Development" & vbNewLine
    s = s & vbNewLine
    s = s & "'Options" & vbNewLine
    s = s & "    Option Explicit" & vbNewLine
    s = s & "    Option Private Module" & vbNewLine
    s = s & "    Option Compare Text" & vbNewLine & vbNewLine
    s = s & "'Private Constants" & vbNewLine
    s = s & "    Private Const cModule    As String = " & ComponentName & vbNewLine
    TemplateModule = s
End Function

Function CopyTemplateFromSheet(Template As String)
    Dim cell As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Templates")
    On Error Resume Next
    Set cell = ws.Columns(1).SpecialCells(xlCellTypeConstants).Find(Template, LookAt:=xlWhole)
    On Error GoTo 0

    CopyTemplateFromSheet = cell.OFFSET(0, 1)
End Function

Sub InjectTemplateFromSheet(Template As String)
    If Len(Selection) > 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Templates")
    Dim cell As Range
    On Error Resume Next
    Set cell = ws.Columns(1).SpecialCells(xlCellTypeConstants).Find(Template, LookAt:=xlWhole)
    On Error GoTo 0
    If Not cell Is Nothing Then Inject cell.OFFSET(0, 1)
End Sub

Sub InjectTemplateProcedure(functionName As String)
    Inject TemplateProcedure(functionName)
End Sub

Function ErrHandlerTemplate(ProcedureName As String) As String
    ErrHandlerTemplate = CopyTemplateFromSheet("ErrHandler")
End Function

Private Function TemplateProcedure(functionName As String) As String
    Dim Q As String: Q = """"
    Dim s As String
    s = s & "Function " & functionName & "(ByVal MyParameter as String) As Variant" & vbNewLine
    s = s & DevInfo & vbNewLine
    s = s & vbNewLine
    s = s & "'   Description:Procedure Description" & vbNewLine
    s = s & "'   Inputs:     MyParameter  Describe its purpose" & vbNewLine
    s = s & "'   Outputs:    Success: <return this>" & vbNewLine
    s = s & "'               Failure: <return this>" & vbNewLine
    s = s & "'   Requisites  Routines    ModuleName.ProcedureName" & vbNewLine
    s = s & "'               Classes      Class Module Name" & vbNewLine
    s = s & "'               Forms        User Form Name" & vbNewLine
    s = s & "'               Tables       Table Name" & vbNewLine
    s = s & "'               References   Reference" & vbNewLine
    s = s & "'   Notes <add if needed>" & vbNewLine
    s = s & "'   Example: ?" & functionName & "(MyParameter)" & vbNewLine & vbNewLine
    s = s & "'   Changelog" & vbNewLine
    s = s & "'   Date        Modification" & vbNewLine
    s = s & "'   " & Format(Date, "DD/MM/YY") & "    Initial Release" & vbNewLine
    s = s & vbNewLine
    s = s & "'   Check Inputs and Requisites" & vbNewLine
    s = s & "    If sParameter = vbNullString then Err.Raise DspError, , ""Parameter missing""" & vbNewLine
    s = s & vbNewLine
    s = s & "'   Declarations" & vbNewLine
    s = s & "    Const cRoutine as String = " & Q & functionName & Q
    s = s & "    LogProcedureUsage cRoutine"
    s = s & "'   Error Handling Initialization" & vbNewLine
    s = s & "    On Error GoTo ErrHandler" & vbNewLine
    s = s & "    " & functionName & " = Failure    'Assume failure" & vbNewLine
    s = s & vbNewLine
    s = s & "'   Initialize Variables" & vbNewLine & vbNewLine
    s = s & "'   Procedure" & vbNewLine
    s = s & "    Application.screenupdating=false" & vbNewLine
    s = s & "    StartTimer " & functionName & vbNewLine & "    " & vbNewLine
    s = s & "    " & functionName & " = Success    'Successful finish" & vbNewLine
    s = s & "    EndTimer" & vbNewLine & vbNewLine
    s = s & "NormalExit:" & vbNewLine
    s = s & "    Application.screenupdating=false" & vbNewLine
    s = s & "    Exit Sub" & vbNewLine
    s = s & vbNewLine
    s = s & ErrHandlerTemplate(functionName) & vbNewLine
    s = s & "End Function"
    TemplateProcedure = s
End Function


Sub Format_VBA7()
    SelectRows
    Inject FormatVBA7(Selection)
End Sub

Private Function StringToVBA7() As String
  
End Function

Public Sub Substitute(oldValue As String, NewValue As String, Optional count As Long = -1)
    Dim str           As String: str = Selection
    Dim lastRowLength As Long
        lastRowLength = 1 + Len(Replace(Split(str, vbNewLine)(UBound(Split(str, vbNewLine))), oldValue, NewValue))
    Dim Code          As String
    Code = SelectionBefore & _
           Replace(str, oldValue, NewValue, , count, vbTextCompare) & _
           SelectionAfter
    cm.DeleteLines RowFirst, RowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, _
                 ColumnFirst, _
                 RowFirst + CountOfCharacters(Code, vbNewLine), _
                 IIf(CountOfCharacters(Code, vbNewLine) > 0, lastRowLength, ColumnFirst + Len(Code))
End Sub

Public Sub CaseLower()
    Inject LCase(Selection)
End Sub

Public Sub CaseProper()
    Inject WorksheetFunction.Proper(Selection)
End Sub

Public Sub CaseUpper()
    Inject UCase(Selection)
End Sub

Sub Sort_Comma()
    Sort ","
End Sub

Private Sub SelectRows()
    SetSelection RowFirst, 1, RowLast, 10000
End Sub

Sub Sort_Lines()
    SelectRows
    Sort vbNewLine
End Sub

Public Sub Sort(delimeter As String)
    Dim Code As String:     Code = Me.Selection
    Dim arr
        arr = Split(Code, delimeter)
        arr = SortSelectionArray(arr)
    Dim str As String:  str = Join(arr, delimeter)
    Code = SelectionBefore & _
           str & _
           SelectionAfter

           
    cm.DeleteLines RowFirst, RowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, _
                  ColumnFirst, _
                  RowLast, _
                  IIf(InStr(1, Code, vbNewLine) = 0, ColumnLast, Len(arr(UBound(arr)))) + IIf(InStr(1, Code, vbNewLine) = 0, 0, 1)
End Sub

Sub InjectArgumentStyle()
    Dim Result As String: Result = aProcedure.Init(ActiveCodepaneWorkbook, , Selection).ArgumentStyle
    If Result <> "" Then Inject Result
End Sub

Sub Encapsulate_Quotes()
    Encapsulate Chr(34), Chr(34)
End Sub

Sub Encapsulate_Parenthesis()
    Encapsulate "(", ")"
End Sub

Public Sub Encapsulate(Before As String, After As String)
    Dim str  As String:  str = Selection
    Dim Code As String
    Code = SelectionBefore & _
           Before & str & After & _
           SelectionAfter
    cm.DeleteLines RowFirst, RowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, ColumnFirst, RowLast, ColumnLast + 2
End Sub

'Sub Encapsulate_CommaSeparatedValues_Quotes()
'    cpsFormatEncapsulateMultiple Chr(34), Chr(34), ","
'End Sub
'
'Sub Encapsulate_CommaSeparatedValues_Parenthesis()
'    cpsFormatEncapsulateMultiple "(", ")", ","
'End Sub
'
'Sub EncapsulateMultipleLinesWithParenthesis()
'    cpsFormatEncapsulateMultiple "(", ")", ","
'End Sub
'
'Sub EncapsulateMultipleLinesWithQuotes()
'    cpsFormatEncapsulateMultiple Chr(34), Chr(34), vbNewLine
'End Sub
'
Public Sub EncapsulateMultiple(Before As String, After As String, Splitter As String)
    Dim Code    As String: Code = Selection
    Dim arr:               arr = Split(Code, Splitter)
    Dim counter As Long
    For counter = LBound(arr) To UBound(arr) - IIf(Right(UBound(arr), Len(Splitter)) = Splitter, Len(Splitter), 0)
        arr(counter) = Before & arr(counter) & After
    Next
    Code = Join(arr, Splitter)
    Code = SelectionBefore & Code & SelectionAfter
    cm.DeleteLines RowFirst, RowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, ColumnFirst, RowLast, ColumnLast + ((Len(Before) + (Len(After))) * (UBound(arr) + 1))
End Sub

Public Sub UnFoldLine()
    Dim str  As String:  str = Selection
    Dim Code As String
    Code = SelectionBefore & _
           Join(Split(str, "_" & vbNewLine), " ") & _
           SelectionAfter

    cm.DeleteLines RowFirst, RowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, ColumnFirst, RowFirst, 10000
End Sub

Public Sub FoldLine()
    If RowCount > 1 Then UnFoldLine
    If InStr(1, Me.Selection, ",") = 0 And InStr(1, Me.Selection, "(") = 0 Then
        SetSelection RowFirst, ColumnFirst, RowFirst, 10000
'        Debug.Print "Select a whole line or part of a line with comma separated arguments."
'        Exit Sub
    End If
    Dim str As String
    str = Me.Selection
    
    str = Join(Split(str, ", "), ", _" & vbNewLine & Space(ColumnFirst - 1))
    str = Replace(str, "(", "( _" & vbNewLine & Space(ColumnFirst - 1), , 1)

    Dim Code As String
    Code = SelectionBefore & _
           str & _
           SelectionAfter
    cm.DeleteLines RowFirst, RowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, ColumnFirst, RowFirst + UBound(Split(Code, " _" & vbNewLine)), 10000
End Sub

Sub RotateCommas()
    Rotate ","
End Sub

Sub RotateLines()
    Rotate vbNewLine
End Sub

Public Sub Rotate(delim As String)
    Rem Rotate multiple  eg. a,b,c,d -> b,c,d,a
    Dim s As String
        s = Selection
        s = Join(ArrayRotate(Split(s, delim)), delim)
        s = SelectionBefore & s & SelectionAfter
    Dim idx As Long:    idx = RowFirst
    cm.DeleteLines idx, RowCount
    cm.InsertLines idx, s
    SetSelection RowFirst, ColumnFirst, RowLast, ColumnLast
End Sub

Sub rotateEqualInLines()
    RotateMultiple "=", vbNewLine
End Sub

Public Sub RotateMultiple(flipper As String, Splitter As String)
    Dim startLine As Long, StartColumn As Long, endLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, endLine, EndColumn
    Dim Code As String:    Code = Selection
    Dim arr As Variant
        arr = Split(Code, Splitter)
        arr = ArrayRemoveEmptyElements(arr)
    Dim counter As Long
    For counter = LBound(arr) To UBound(arr) - IIf(Right(UBound(arr), Len(Splitter)) = Splitter, Len(Splitter), 0)
        arr(counter) = Split(arr(counter), flipper)(1) & flipper & Split(arr(counter), flipper)(0)
    Next
    Code = Join(arr, Splitter)
    Code = SelectionBefore & _
            Code & _
            SelectionAfter
    cm.DeleteLines startLine, endLine - startLine + 1
    cm.InsertLines startLine, Code
    SetSelection startLine, StartColumn, endLine, EndColumn
End Sub

Sub Duplicate()
    Dim s As String
    If RowCount > 1 Then
        SelectRows
        s = Selection & vbNewLine & Selection
    Else
        Dim RowNumber As Long:      RowNumber = RowFirst
        Dim AP As String:           AP = ActiveProcedure:       If AP = "" Then Exit Sub
        Dim Limit As Long:          Limit = aProcedure.Init(, , AP).LineIndex(Body_First_AfterComments)
        Do While Len(Trim(cm.Lines(RowNumber, 1))) = 0 _
        And RowNumber - 1 >= Limit
            RowNumber = RowNumber - 1
        Loop
        If Len(Trim(cm.Lines(RowNumber, 1))) > 0 Then
            s = cm.Lines(RowNumber, 1)
        Else
            Exit Sub
        End If
    End If
    cm.InsertLines RowFirst, s
End Sub

Sub Increment()
    If RowCount > 1 Then Exit Sub
    Dim RowNumber As Long:      RowNumber = RowFirst
    Dim AP        As String:    AP = ActiveProcedure:       If AP = "" Then Exit Sub
    Dim Limit     As Long:      Limit = aProcedure.Init(, , AP).LineIndex(Body_First_AfterComments)
    Dim s         As String
    Do While Len(Trim(cm.Lines(RowNumber, 1))) = 0 _
    And RowNumber - 1 >= Limit
        RowNumber = RowNumber - 1
    Loop
    If Len(Trim(cm.Lines(RowNumber, 1))) > 0 Then
        s = cm.Lines(RowNumber, 1)
    Else
        Exit Sub
    End If
    s = IncreaseAllNumbersInString(s)
    cm.InsertLines RowFirst + 1, s
    SetSelection RowFirst + 1, 1, RowFirst + 1, 10000
End Sub

Sub Cut()
    CLIP SelectionLinesCode
    Debug.Print Selection
    cm.DeleteLines RowFirst, RowCount
    SetSelection RowFirst, 1, RowFirst, 1
    AssignVariables
End Sub

Sub Copy()
    SelectRows
    CLIP Selection
    AssignVariables
End Sub

Public Sub BeautifyFunction()
    Dim txt As String: txt = Selection
    txt = StringFormatFunctionNested(txt)
    IndentForInject txt
    Inject txt
End Sub

Private Sub IndentForInject(ByRef txt As String)
    Dim var:            var = Split(txt, vbNewLine): If UBound(var) < 1 Then Exit Sub
    Dim Indent As Long: Indent = ColumnFirst:        If Indent < 2 Then Exit Sub
    Dim i As Long
    For i = LBound(var) + 1 To UBound(var)
        var(i) = Space(Indent) & var(i)
        txt = Join(var, vbNewLine)
    Next
End Sub
'
'Function cpsCurrentRegionLineBlockFirst() As String
'    Dim Module As VBComponent
'    Set Module = ActiveModule
'    If Len(Module.codeModule.Lines(cpsLineFirst, 1)) = 0 Then Exit Function
'
'    Dim Procedure As String
'        Procedure = ActiveProcedure
'    Dim StartLine As Long
'        StartLine = cpsLineFirst
'    Dim nLine As Long
'        nLine = StartLine
'    If nLine - 1 > 1 _
'    And nLine - 1 > ProcedureTitleLineLast(Module, Procedure) Then
'        Do Until Len(Module.codeModule.Lines(nLine, 1)) _
'               - Len(Trim(Module.codeModule.Lines(nLine, 1))) = 4
'            nLine = nLine - 1
'        Loop
'    End If
'    StartLine = nLine
'    cpsCurrentRegionLineBlockFirst = StartLine
'End Function
'
'Function cpsCurrentRegionLineBlockLast() As String
'    Dim Module As VBComponent
'    Set Module = ActiveModule
'    If Len(Module.codeModule.Lines(cpsLineFirst, 1)) = 0 Then Exit Function
'    Dim Procedure As String
'        Procedure = ActiveProcedure
'    Dim EndLine As Long
'        EndLine = cpsLineLast
'    Dim nLine As Long
'        nLine = EndLine
'    If nLine + 1 < ProcedureLinesLast(Module, Procedure) Then
'        Do Until Len(Module.codeModule.Lines(nLine, 1)) _
'               - Len(Trim(Module.codeModule.Lines(nLine, 1))) = 4
'            nLine = nLine + 1
'        Loop
'    End If
'    EndLine = nLine
'    cpsCurrentRegionLineBlockLast = EndLine
'End Function
'
'Function cpsCurrentRegionLineFirst() As String
'    Dim Module As VBComponent
'    Set Module = ActiveModule
'    If Len(Module.codeModule.Lines(cpsLineFirst, 1)) = 0 Then Exit Function
'
'    Dim Procedure As String
'        Procedure = ActiveProcedure
'    Dim StartLine As Long
'        StartLine = cpsLineFirst
'    Dim nLine As Long
'        nLine = StartLine
'    Do While nLine - 1 > 1 _
'    And nLine - 1 > ProcedureTitleLineLast(Module, Procedure) _
'    And Len(Trim(Module.codeModule.Lines(nLine - 1, 1))) > 0
'        nLine = nLine - 1
'    Loop
'    StartLine = nLine
'    cpsCurrentRegionLineFirst = StartLine
'End Function
'
'Function cpsCurrentRegionLineLast() As String
'    Dim Module As VBComponent
'    Set Module = ActiveModule
'    If Len(Module.codeModule.Lines(cpsLineFirst, 1)) = 0 Then Exit Function
'    Dim Procedure As String
'        Procedure = ActiveProcedure
'    Dim EndLine As Long
'        EndLine = cpsLineLast
'    Dim nLine As Long
'        nLine = EndLine
'    Do While nLine + 1 < ProcedureLinesLast(Module, Procedure) _
'    And Len(Trim(Module.codeModule.Lines(nLine + 1, 1))) > 0
'        nLine = nLine + 1
'    Loop
'    EndLine = nLine
'    cpsCurrentRegionLineLast = EndLine
'End Function
'
'Function cpsCurrentRegionLinesCode() As String
'    Dim Module As VBComponent
'    Set Module = ActiveModule
'    If Len(Module.codeModule.Lines(cpsLineFirst, 1)) = 0 Then Exit Function
'
'    Dim Procedure As String
'        Procedure = ActiveProcedure
'    Dim StartLine As Long
'        StartLine = cpsLineFirst
'    Dim nLine As Long
'        nLine = StartLine
'    Do While nLine - 1 > 1 _
'    And nLine - 1 > ProcedureTitleLineLast(Module, Procedure) _
'    And Len(Trim(Module.codeModule.Lines(nLine - 1, 1))) > 0
'        nLine = nLine - 1
'    Loop
'    StartLine = nLine
'
'    Dim EndLine As Long
'        EndLine = cpsLineLast
'    nLine = EndLine
'    Do While nLine + 1 < ProcedureLinesLast(Module, Procedure) _
'    And Len(Trim(Module.codeModule.Lines(nLine + 1, 1))) > 0
'        nLine = nLine + 1
'    Loop
'    EndLine = nLine
'    cpsCurrentRegionLinesCode = Module.codeModule.Lines(StartLine, EndLine - StartLine + 1)
'End Function

Sub Comment()
    Dim rowLine As String
    Dim i As Long
    For i = RowFirst To RowLast
        CommentTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection RowFirst, 1, RowLast, 1000
End Sub

Public Sub CommentTargetLine(targetLine As Long)
    Dim sLine As String:  sLine = oModule.CodeModule.Lines(targetLine, 1)
    If IsCommentLine(sLine) Then Exit Sub
    Dim dif As Long:        dif = Len(sLine) - Len(LTrim(sLine))
    oModule.CodeModule.ReplaceLine targetLine, Space(dif) & "'" & Trim(sLine)
End Sub

Sub UnComment()
    Dim i As Long
    For i = RowFirst To RowLast
        UncommentTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection RowFirst, 1, RowLast, 1000
End Sub

Public Sub UncommentTargetLine(targetLine As Long)
    Dim sLine As String
    With oModule.CodeModule
        sLine = .Lines(targetLine, 1)
        If Left(Trim(sLine), 1) = "'" Then
            .ReplaceLine targetLine, Replace(sLine, "'", "", , 1)
        End If
    End With
End Sub

Sub RemAdd()
    Dim i As Long
    For i = RowFirst To RowLast
        RemTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection RowFirst, 1, RowLast, 1000
End Sub

Sub RemTargetLine(targetLine As Long)
    Dim rowLine As String: rowLine = oModule.CodeModule.Lines(targetLine, 1)
    If IsCommentLine(rowLine) Then Exit Sub
    oModule.CodeModule.ReplaceLine targetLine, Space(Len(rowLine) - Len(LTrim(rowLine))) & "Rem " & Trim(rowLine)
End Sub

Sub RemRemove()
    Dim i As Long
    For i = RowFirst To RowLast
        UnremTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection RowFirst, 1, RowLast, 1000
End Sub

Sub UnremTargetLine(targetLine As Long)
    Dim sLine As String
    With oModule.CodeModule
        sLine = .Lines(targetLine, 1)
        If Left(Trim(sLine), 4) = "Rem " Then
            .ReplaceLine targetLine, Replace(sLine, "Rem ", "", , 1)
        End If
    End With
End Sub

Sub BringProcedureHere()
    aProcedure.Init(ActiveCodepaneWorkbook, ActiveModule, ActiveProcedure).BringProcedureHere Selection
End Sub

Sub AlignAs()
    Align " As ", SearchFromLeft:=True
End Sub

Sub AlignColumn()
    Align ":", SearchFromLeft:=True
    Align ":", SearchFromLeft:=False
End Sub

Sub AlignComments()
    Align "'", SearchFromLeft:=False
End Sub

Sub Align(AlignString As String, SearchFromLeft As Boolean, Optional AlignAtColumn As Long)

    Dim s As String:        s = StringFormatAlignRowsElements(SelectionLinesCode, AlignString, SearchFromLeft, AlignAtColumn)
    Dim LineFirst As Long:  LineFirst = RowFirst
    oModule.CodeModule.DeleteLines LineFirst, RowCount
    oModule.CodeModule.InsertLines LineFirst, s
    SelectRows
End Sub

Public Property Get ProceduresByName() As Collection
    Dim out           As New Collection
    Dim s             As String: s = Selection
    Dim ProcedureName As String
    Dim Procedures    As Collection
    Set Procedures = aWorkbook.Init(oWorkbook).Procedures(True, False, False, False)
    Dim Procedure     As aProcedure
    For Each Procedure In Procedures
        ProcedureName = Procedure.Name
        If RegexTest(s, ProcedureName) Then out.Add ProcedureName, ProcedureName
    Next
    Set ProceduresByName = out
End Property

Function ProceduresByDeclaration() As Collection
    Dim out           As New Collection
    Dim s             As String: s = Selection
    Dim ProcedureName As String
    Dim Procedures    As Collection
    Set Procedures = ProceduresOfTXT(s)
    Dim Procedure     As aProcedure
    For Each Procedure In Procedures
        ProcedureName = Procedure.Name
        If RegexTest(s, ProcedureName) Then out.Add ProcedureName, ProcedureName
    Next
    Set ProceduresByDeclaration = out
End Function




Public Sub AssignEnumValues(Optional ToThePower As Boolean = True)
    Dim row1 As Long, row2 As Long
    row1 = ActiveEnumStartLine
    row2 = ActiveEnumEndLine
    Dim Code As String: Code = cm.Lines(row1 + 1, row2 - row1 - 1)
    Dim arr: arr = Split(Code, vbNewLine)
    Code = ""
    Dim out As String
    Dim i As Long
    For i = 0 To UBound(arr)
        If InStr(1, arr(i), "=") > 0 Then arr(i) = Split(arr(i), "=")(0)
        arr(i) = Space(4) & Trim(arr(i))
    Next
    If ToThePower = True Then
        For i = 0 To UBound(arr)
            out = arr(i) & "= 2 ^ " & i
            Code = IIf(Code = "", out, Code & vbNewLine & out)
        Next
    Else
        For i = 0 To UBound(arr)
            out = arr(i) & "= " & i + 1
            Code = IIf(Code = "", out, Code & vbNewLine & out)
        Next
    End If
    cm.DeleteLines row1 + 1, row2 - row1 - 1
    cm.InsertLines row1 + 1, Code
End Sub

Public Sub EnumToCase()
    Rem point inside enum before calling this from immediate window or vbe menu button
    Dim enumName As String: enumName = ActiveEnumName
    Dim arr:                arr = Split(oModule.CodeModule.Lines(ActiveEnumStartLine + 1, ActiveEnumEndLine - ActiveEnumStartLine - 1), vbNewLine)
    Dim out As String:      out = "Select case Variable "
    Dim Code As String:     Code = out
    Dim i As Long
    For i = 0 To UBound(arr)
        If InStr(1, arr(i), "=") > 0 Then arr(i) = Split(arr(i), "=")(0)
        arr(i) = Trim(arr(i))
    Next
    For i = 0 To UBound(arr)
        If arr(i) <> "" Then
            out = "    Case is = " & enumName & "." & arr(i) & vbNewLine
            Code = IIf(Code = "", out, Code & vbNewLine & out)
        End If
    Next
    Code = Code & vbNewLine & "End Select"
    Debug.Print "The following was copied to clipboard:"
    Debug.Print String(20, "-")
    Debug.Print Code
    CLIP Code
End Sub

Private Function ActiveEnumName() As String
    Dim Line As String
    Dim i As Long
    For i = aCodeModule.Init(oModule).RowFirst To 1 Step -1
        Line = oModule.CodeModule.Lines(i, 1)
        If InStr(1, Line, "Enum ") > 0 Then
            ActiveEnumName = Trim(Split(Line, "Enum ")(1))
            Exit Function
        End If
    Next
End Function

Private Function ActiveEnumStartLine() As Long
    Dim i    As Long
    Dim Line As String
    For i = aCodeModule.Init(oModule).RowFirst To 1 Step -1
        Line = oModule.CodeModule.Lines(i, 1)
        If InStr(1, Line, "Enum ") > 0 Then
            ActiveEnumStartLine = i
            Exit Function
        End If
    Next
End Function

Function ActiveEnumEndLine() As Long
    Dim i        As Long
    Dim enumName As String
    Dim Line     As String
    For i = aCodeModule.Init(oModule).RowFirst To oModule.CodeModule.CountOfLines
        Line = oModule.CodeModule.Lines(i, 1)
        If InStr(1, Line, "End Enum") > 0 Then
            ActiveEnumEndLine = i
            Exit Function
        End If
    Next
End Function

Function Procedure() As aProcedure
'@INCLUDE CLASS aProcedure
    If ActiveProcedure = "" Then
        Toast "The carret is not inside a procedure"
        Exit Function
    End If
    Set Procedure = aProcedure.Active
End Function

Sub injectDivider(Optional Character As String = "~")
    If Len(Trim(cm.Lines(RowFirst, 1))) = 0 Then
        cm.ReplaceLine RowFirst, "'" & String(Procedure.LargestLineLength - 1, "character")
    Else
        cm.InsertLines RowFirst, "'" & String(Procedure.LargestLineLength - 1, "character")
    End If
End Sub