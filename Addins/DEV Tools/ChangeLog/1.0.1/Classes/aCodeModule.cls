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

Dim oWorkbook       As Workbook
Dim oModule         As VBComponent
Private cm          As CodeModule

Public ColumnFirst  As Long
Public ColumnLast    As Long
Public rowCount      As Long
Public RowFirst      As Long
Public RowLast       As Long

Public CodemoduleSelection     As String
Public SelectionAfter          As String
Public SelectionBefore         As String
Public SelectionLength         As Long
Public SelectionLinesCode      As String

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 11:01    Alex                (aCodeModule.cls > ActivateProcedure)

Public Sub ActivateProcedure()
'@LastModified 2308221101
    aProcedure.Init(ActiveCodepaneWorkbook, ActiveModule, CodemoduleSelection).Activate
End Sub

Public Sub ImportProcedure()
    '@INCLUDE CLASS aCodeModule
    '@INCLUDE CLASS aProcedure
    '@AssignedModule aCodeModule
    aProcedure.Active.ImportProcedure CodemoduleSelection, False
End Sub

Public Function Parent() As VBComponent
    Set Parent = oModule
End Function
Public Function Active() As aCodeModule
    '@INCLUDE PROCEDURE ActiveModule
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Init ActiveModule
    Set Active = Me
End Function

Public Function Init(module As VBComponent) As aCodeModule
    '@INCLUDE PROCEDURE WorkbookOfModule
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Set oModule = module
    Set cm = oModule.CodeModule
    Set oWorkbook = WorkbookOfModule(module)
    AssignVariables
    Set Init = Me
End Function

Private Sub Class_Terminate()
    'do something
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule

End Sub

Private Sub AssignVariables()
    '@INCLUDE PROCEDURE CountOfCharacters
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule

    Dim startLine As Long, StartColumn As Long, endLine As Long, EndColumn As Long
    cm.CodePane.GetSelection startLine, StartColumn, endLine, EndColumn

    ColumnFirst = StartColumn
    ColumnLast = EndColumn

    RowFirst = startLine
    RowLast = endLine
    If RowLast > startLine And EndColumn = 1 Then
        RowLast = endLine - 1
    End If
    rowCount = RowLast - RowFirst + 1

    SelectionLinesCode = cm.Lines(RowFirst, rowCount)


    If ColumnLast - ColumnFirst = 0 Then
        If endLine > RowLast Then
            ColumnLast = Len(cm.Lines(RowLast, 1)) + 1
        End If
    End If

    SelectionAfter = Mid(cm.Lines(RowLast, 1), ColumnLast)    ' + IIf(Len(CodemoduleSelection) > 0, 1, 0))
    SelectionBefore = Mid(cm.Lines(RowFirst, 1), 1, ColumnFirst - 1)

    If rowCount = 1 And (ColumnLast - ColumnFirst > 0) Then
        SelectionLength = ColumnLast - ColumnFirst
    Else
        SelectionLength = Len(SelectionLinesCode) - Len(SelectionBefore) - Len(SelectionAfter) - 1 - Len(CountOfCharacters(SelectionLinesCode, vbNewLine)) * 2
    End If

    If RowLast - RowFirst = 0 Then
        CodemoduleSelection = Mid(cm.Lines(RowFirst, 1), ColumnFirst, ColumnLast - ColumnFirst)
        Exit Sub
    End If
    Dim str         As String
    Dim i           As Long
    For i = RowFirst To RowLast
        If str = "" Then
            str = Mid(cm.Lines(i, 1), ColumnFirst)
        ElseIf i < RowLast Then
            str = str & vbNewLine & cm.Lines(i, 1)
        Else
            str = str & vbNewLine & Left(cm.Lines(i, 1), ColumnLast - 1)
        End If
    Next
    CodemoduleSelection = str

End Sub

Sub Todo()
    '@INCLUDE CLASS aCodeModule
    '@INCLUDE DECLARATION AUTHOR_NAME
    '@AssignedModule aCodeModule
    Const tFormat = "dd-mm-yyyy hh:nn"
    Dim module      As VBComponent
    Set module = oModule
    Dim txtName     As String
    txtName = AUTHOR_NAME
    If txtName = vbNullString Then txtName = Environ("UserName")
    Dim nLine       As Long
    nLine = aCodeModule.Init(module).RowFirst
    Do While nLine > 1 And Right(module.CodeModule.Lines(nLine, 1), 1) = "_"
        nLine = nLine - 1
    Loop
    Dim sLine       As String
    sLine = module.CodeModule.Lines(nLine, 1)
    Dim sSpec       As String
    sSpec = Space(Len(sLine) - Len(Trim(sLine)))
    Dim sTodo       As String
    sTodo = sSpec & "'* @TODO Created: " & _
            VBA.Format$(Now, tFormat) & _
            " Author: " & txtName & vbCrLf & _
            sSpec & "'* @TODO "
    module.CodeModule.InsertLines nLine, vbNewLine & sTodo & vbNewLine
End Sub

Sub ToggleComments()
    '@INCLUDE PROCEDURE IsCommentLine
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim sLine       As String
    Dim i           As Long
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
    '@INCLUDE PROCEDURE StringFormatAlignRowsElements
    '@INCLUDE PROCEDURE IndentationCount
    '@INCLUDE CLASS aCodeModule
    '@INCLUDE CLASS aCollection
    '@AssignedModule aCodeModule
    Dim coll        As New Collection
    Dim Code        As String: Code = SelectionLinesCode
    Dim FirstLine   As Long: FirstLine = RowFirst
    Dim iSpace      As Long
    Dim iCount      As Long: iCount = rowCount

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

Public Sub DimMerge()
    '@INCLUDE PROCEDURE RegExpReplace
    '@INCLUDE PROCEDURE IndentationCount
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim Code        As String: Code = SelectionLinesCode
    Dim FirstLine   As Long: FirstLine = RowFirst
    Dim iCount      As Long: iCount = rowCount
    Dim iSpace      As Long: iSpace = IndentationCount(cm.Lines(FirstLine, 1))
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
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    cm.CodePane.SetSelection startLine, StartColumn, endLine, EndColumn
    AssignVariables
End Sub

Public Sub Move_Up()
    '@INCLUDE PROCEDURE ActiveProcedure
    '@INCLUDE CLASS aCodeModule
    '@INCLUDE CLASS aProcedure
    '@AssignedModule aCodeModule
    Dim blockStart  As Long: blockStart = RowFirst
    Dim blockEnd    As Long: blockEnd = RowLast
    Dim blockCountOfLines As Long: blockCountOfLines = blockEnd - blockStart + 1
    Dim blockString As String: blockString = cm.Lines(blockStart, blockCountOfLines)
    Dim insertBlockAtLine As Long: insertBlockAtLine = blockStart - 1
    Dim ap          As String: ap = ActiveProcedure
    If ap = "" Then Exit Sub
    Select Case insertBlockAtLine
        Case 1, aProcedure.Init(, , ap).Line_Declaration_First
            Exit Sub
    End Select
    cm.DeleteLines blockStart, blockCountOfLines
    cm.InsertLines insertBlockAtLine, blockString
    SetSelection insertBlockAtLine, 1, insertBlockAtLine + blockCountOfLines - 1, 10000
End Sub

Public Sub Move_Down()
    '@INCLUDE PROCEDURE ActiveProcedure
    '@INCLUDE CLASS aCodeModule
    '@INCLUDE CLASS aProcedure
    '@AssignedModule aCodeModule
    Dim blockStart  As Long: blockStart = RowFirst
    Dim blockEnd    As Long: blockEnd = RowLast
    Dim blockCountOfLines As Long: blockCountOfLines = blockEnd - blockStart + 1
    Dim blockString As String: blockString = cm.Lines(blockStart, blockCountOfLines)
    Dim insertBlockAtLine As Long: insertBlockAtLine = blockStart + 1
    Dim ap          As String: ap = ActiveProcedure
    If ap = "" Then Exit Sub
    Select Case insertBlockAtLine + blockCountOfLines - 1
        Case cm.CountOfLines, aProcedure.Init(, , ap).Line_Last
            Exit Sub
    End Select
    cm.DeleteLines blockStart, blockCountOfLines
    cm.InsertLines insertBlockAtLine, blockString
    SetSelection insertBlockAtLine, 1, insertBlockAtLine + blockCountOfLines - 1, 10000
End Sub

Public Sub Inject(str As String)
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim Code        As String: Code = SelectionBefore & str & SelectionAfter
    Dim endLine     As Long
    'catch error when selected all lines
    If rowCount > cm.CountOfLines Then endLine = endLine - 1
    cm.DeleteLines RowFirst, rowCount
    cm.InsertLines RowFirst, Code
'    Debug.Print cm.Name & vbTab & RowFirst
    SetSelection RowFirst, _
            ColumnFirst, _
            RowFirst + UBound(Split(str, vbLf)), _
            IIf((UBound(Split(str, vbLf))) = 0, ColumnLast, 1) + Len(Split(str, vbLf)(UBound(Split(str, vbLf))))
End Sub

Public Sub Insert(str As String)
    cm.InsertLines RowFirst, str
End Sub


Private Function CopyTemplateFromSheet(Template As String)
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim cell        As Range
    Dim ws          As Worksheet
    Set ws = ThisWorkbook.Sheets("Templates")
    On Error Resume Next
    Set cell = ws.Columns(1).SpecialCells(xlCellTypeConstants).Find(Template, LookAt:=xlWhole)
    On Error GoTo 0

    CopyTemplateFromSheet = cell.offset(0, 1)
End Function

Private Sub InjectTemplateFromSheet(Template As String)
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    If Len(Selection) > 0 Then Exit Sub
    Dim ws          As Worksheet: Set ws = ThisWorkbook.Sheets("Templates")
    Dim cell        As Range
    On Error Resume Next
    Set cell = ws.Columns(1).SpecialCells(xlCellTypeConstants).Find(Template, LookAt:=xlWhole)
    On Error GoTo 0
    If Not cell Is Nothing Then Inject cell.offset(0, 1)
End Sub

Public Sub Format_VBA7()
    '@INCLUDE PROCEDURE FormatVBA7
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    SelectRows
    Inject FormatVBA7(CodemoduleSelection)
End Sub

Public Sub Substitute(oldValue As String, NewValue As String, Optional Count As Long = -1)
    '@INCLUDE PROCEDURE CountOfCharacters
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim str         As String: str = CodemoduleSelection
    Dim lastRowLength As Long
    lastRowLength = 1 + Len(Replace(Split(str, vbNewLine)(UBound(Split(str, vbNewLine))), oldValue, NewValue))
    Dim Code        As String
    Code = SelectionBefore & _
            Replace(str, oldValue, NewValue, , Count, vbTextCompare) & _
            SelectionAfter
    cm.DeleteLines RowFirst, rowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, _
            ColumnFirst, _
            RowFirst + CountOfCharacters(Code, vbNewLine), _
            IIf(CountOfCharacters(Code, vbNewLine) > 0, lastRowLength, ColumnFirst + Len(Code))
End Sub

Public Sub CaseLower()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Inject LCase(CodemoduleSelection)
End Sub

Public Sub CaseProper()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Inject WorksheetFunction.Proper(CodemoduleSelection)
End Sub

Public Sub CaseUpper()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Inject UCase(CodemoduleSelection)
End Sub

Sub Sort_Comma()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Sort ","
End Sub

Private Sub SelectRows()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    SetSelection RowFirst, 1, RowLast, 10000
End Sub

Public Sub Sort_Lines()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    SelectRows
    Sort vbNewLine
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 12-09-2023 14:24    Alex                (aCodeModule.cls > SortLinesByLength) initial release

Public Sub SortLinesByLength()
'@LastModified 2309121424
    SelectRows
    Dim arr
    arr = Split(CodemoduleSelection, vbNewLine)
    ArraySortByLength arr, LBound(arr), UBound(arr)
    Inject Join(arr, vbNewLine)
End Sub

Public Sub Sort(delimeter As String)
    '@INCLUDE PROCEDURE SortSelectionArray
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim Code        As String: Code = CodemoduleSelection
    Dim arr
    arr = Split(Code, delimeter)
    arr = SortSelectionArray(arr)
    Dim str         As String: str = Join(arr, delimeter)
    Code = SelectionBefore & _
            str & _
            SelectionAfter


    cm.DeleteLines RowFirst, rowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, _
            ColumnFirst, _
            RowLast, _
            IIf(InStr(1, Code, vbNewLine) = 0, ColumnLast, Len(arr(UBound(arr)))) + IIf(InStr(1, Code, vbNewLine) = 0, 0, 1)
End Sub

Public Sub InjectArgumentStyleFolded()
    '@INCLUDE PROCEDURE ActiveCodepaneWorkbook
    '@INCLUDE CLASS aCodeModule
    '@INCLUDE CLASS aProcedure
    '@AssignedModule aCodeModule
    Dim result      As String: result = aProcedure.Init(ActiveCodepaneWorkbook, , CodemoduleSelection).ArgumentStyleFolded
    If result <> "" Then Inject result
End Sub

Public Sub Encapsulate_Quotes()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Encapsulate Chr(34), Chr(34)
End Sub

Public Sub Encapsulate_Parenthesis()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Encapsulate "(", ")"
End Sub

Public Sub Encapsulate(Before As String, After As String)
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim str         As String: str = CodemoduleSelection
    Dim Code        As String
    Code = SelectionBefore & _
            Before & str & After & _
            SelectionAfter
    cm.DeleteLines RowFirst, rowCount
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
Public Sub EncapsulateMultiple( _
                              Before As String, _
                              After As String, _
                              Splitter As String)
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim Code        As String: Code = CodemoduleSelection
    Dim arr: arr = Split(Code, Splitter)
    Dim counter     As Long
    For counter = LBound(arr) To UBound(arr) - IIf(Right(UBound(arr), Len(Splitter)) = Splitter, Len(Splitter), 0)
        arr(counter) = Before & arr(counter) & After
    Next
    Code = Join(arr, Splitter)
    Code = SelectionBefore & Code & SelectionAfter
    cm.DeleteLines RowFirst, rowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, _
                 ColumnFirst, _
                 RowLast, _
                 ColumnLast + ( _
                 (Len(Before) + (Len(After))) * (UBound(arr) + 1))
End Sub

Public Sub UnFoldLine()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim str         As String: str = CodemoduleSelection
    If Trim(str) = "" Then Exit Sub
    Dim Code        As String
    Code = SelectionBefore & _
            Join(Split(str, "_" & vbNewLine), " ") & _
            SelectionAfter

    cm.DeleteLines RowFirst, rowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, ColumnFirst, RowFirst, 10000
End Sub

Public Sub FoldLine()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    If rowCount > 1 Then UnFoldLine
    If InStr(1, CodemoduleSelection, ",") = 0 And InStr(1, CodemoduleSelection, "(") = 0 Then
        SetSelection RowFirst, ColumnFirst, RowFirst, 10000
        '        Debug.Print "Select a whole line or part of a line with comma separated arguments."
        '        Exit Sub
    End If
    Dim str         As String
    str = CodemoduleSelection

    str = Join(Split(str, ", "), ", _" & vbNewLine & Space(ColumnFirst - 1))
    str = Replace(str, "(", "( _" & vbNewLine & Space(ColumnFirst - 1), , 1)

    Dim Code        As String
    Code = SelectionBefore & _
            str & _
            SelectionAfter
    cm.DeleteLines RowFirst, rowCount
    cm.InsertLines RowFirst, Code
    SetSelection RowFirst, ColumnFirst, RowFirst + UBound(Split(Code, " _" & vbNewLine)), 10000
End Sub

Public Sub RotateCommas()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Rotate ","
End Sub

Public Sub RotateLines()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Rotate vbNewLine
End Sub

Public Sub Rotate(delim As String)
    Rem Rotate multiple  eg. a,b,c,d -> b,c,d,a
    '@INCLUDE PROCEDURE ArrayRotate
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim S           As String
    S = CodemoduleSelection
    S = Join(ArrayRotate(Split(S, delim)), delim)
    S = SelectionBefore & S & SelectionAfter
    Dim idx         As Long: idx = RowFirst
    cm.DeleteLines idx, rowCount
    cm.InsertLines idx, S
    SetSelection RowFirst, ColumnFirst, RowLast, ColumnLast
End Sub

Public Sub rotateEqualInLines()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    RotateMultiple "=", vbNewLine
End Sub

Public Sub RotateMultiple(flipper As String, Splitter As String)
    '@INCLUDE PROCEDURE ArrayRemoveEmptyElements
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim startLine   As Long
    Dim StartColumn As Long
    Dim endLine     As Long
    Dim EndColumn   As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, _
                                                StartColumn, _
                                                endLine, _
                                                EndColumn
    Dim Code        As String: Code = CodemoduleSelection
    Dim arr         As Variant
    arr = Split(Code, Splitter)
    arr = ArrayRemoveEmptyElements(arr)
    Dim counter     As Long
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

Public Sub Duplicate()
    '@INCLUDE PROCEDURE ActiveProcedure
    '@INCLUDE CLASS aCodeModule
    '@INCLUDE CLASS aProcedure
    '@AssignedModule aCodeModule
    Dim S           As String
    If rowCount > 1 Then
        SelectRows
        S = CodemoduleSelection & vbNewLine & CodemoduleSelection
    Else
        Dim RowNumber As Long: RowNumber = RowFirst
        Dim ap      As String: ap = ActiveProcedure: If ap = "" Then Exit Sub
        Dim Limit   As Long: Limit = aProcedure.Init(, , ap).Line_Body_FirstAfterComments
        Do While Len(Trim(cm.Lines(RowNumber, 1))) = 0 _
                And RowNumber - 1 >= Limit
            RowNumber = RowNumber - 1
        Loop
        If Len(Trim(cm.Lines(RowNumber, 1))) > 0 Then
            S = cm.Lines(RowNumber, 1)
        Else
            Exit Sub
        End If
    End If
    cm.InsertLines RowFirst, S
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 17-08-2023 08:24    Alex                (Increment)

Public Sub Increment()
    '@LastModified 2308170824
    '@INCLUDE PROCEDURE IncreaseAllNumbersInString
    '@INCLUDE CLASS aCodeModule
    '@INCLUDE CLASS aProcedure
    '@AssignedModule aCodeModule
    If rowCount > 1 Then Exit Sub
    Dim RowNumber   As Long: RowNumber = RowFirst

    Do While Len(Trim(cm.Lines(RowNumber, 1))) = 0    'And RowNumber - 1 >= Limit
        RowNumber = RowNumber - 1
    Loop

    Dim S           As String
    If Len(Trim(cm.Lines(RowNumber, 1))) > 0 Then
        S = cm.Lines(RowNumber, 1)
    Else
        Exit Sub
    End If

    Dim ap          As String
    Dim i           As Long
    For i = 0 To 3
        On Error Resume Next
        ap = cm.ProcOfLine(RowNumber, i)
        On Error GoTo 0
        If ap <> "" Then Exit For
    Next
    If ap = "" Then Exit Sub

    S = IncreaseAllNumbersInString(S)
    cm.InsertLines RowNumber + 1, S
    SetSelection RowNumber + 1, 1, RowFirst + 1, 10000
End Sub

Public Sub Cut()
    '@INCLUDE PROCEDURE CLIP
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    CLIP SelectionLinesCode
    Debug.Print CodemoduleSelection
    cm.DeleteLines RowFirst, rowCount
    SetSelection RowFirst, 1, RowFirst, 1
    AssignVariables
End Sub

Public Sub Copy()
    '@INCLUDE PROCEDURE CLIP
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    SelectRows
    CLIP SelectionLinesCode
    AssignVariables
End Sub

Public Sub BeautifyFunction()
    '@INCLUDE PROCEDURE StringFormatFunctionNested
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim txt         As String: txt = CodemoduleSelection
    If txt = "" Then Exit Sub
    txt = StringFormatFunctionNested(txt)
    IndentForInject txt
    Inject txt
End Sub

Private Sub IndentForInject(ByRef txt As String)
    '@INCLUDE PROCEDURE ActiveProcedure
    '@INCLUDE PROCEDURE ActiveModule
    '@INCLUDE PROCEDURE ProcedureLinesLast
    '@INCLUDE PROCEDURE ProcedureTitleLineLast
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim var: var = Split(txt, vbNewLine): If UBound(var) < 1 Then Exit Sub
    Dim Indent      As Long: Indent = ColumnFirst: If Indent < 2 Then Exit Sub
    Dim i           As Long
    For i = LBound(var) + 1 To UBound(var)
        var(i) = Space(Indent) & var(i)
        txt = Join(var, vbNewLine)
    Next
End Sub

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

Public Sub Comment()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim rowLine     As String
    Dim i           As Long
    For i = RowFirst To RowLast
        CommentTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection RowFirst, 1, RowLast, 1000
End Sub

Public Sub CommentTargetLine(targetLine As Long)
    '@INCLUDE PROCEDURE IsCommentLine
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim sLine       As String: sLine = oModule.CodeModule.Lines(targetLine, 1)
    If IsCommentLine(sLine) Then Exit Sub
    Dim dif         As Long: dif = Len(sLine) - Len(LTrim(sLine))
    oModule.CodeModule.ReplaceLine targetLine, Space(dif) & "'" & Trim(sLine)
End Sub

Public Sub UnComment()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim i           As Long
    For i = RowFirst To RowLast
        UncommentTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection RowFirst, 1, RowLast, 1000
End Sub

Public Sub UncommentTargetLine(targetLine As Long)
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim sLine       As String
    With oModule.CodeModule
        sLine = .Lines(targetLine, 1)
        If Left(Trim(sLine), 1) = "'" Then
            .ReplaceLine targetLine, Replace(sLine, "'", "", , 1)
        End If
    End With
End Sub

Public Sub RemAdd()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim i           As Long
    For i = RowFirst To RowLast
        RemTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection RowFirst, 1, RowLast, 1000
End Sub

Private Sub RemTargetLine(targetLine As Long)
    '@INCLUDE PROCEDURE IsCommentLine
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim rowLine     As String: rowLine = oModule.CodeModule.Lines(targetLine, 1)
    If IsCommentLine(rowLine) Then Exit Sub
    oModule.CodeModule.ReplaceLine targetLine, _
                                   Space( _
                                   Len(rowLine) - Len(LTrim(rowLine))) & "Rem " & Trim(rowLine)
End Sub

Public Sub RemRemove()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim i           As Long
    For i = RowFirst To RowLast
        UnremTargetLine i
    Next
    oModule.CodeModule.CodePane.SetSelection RowFirst, 1, RowLast, 1000
End Sub

Public Sub UnremTargetLine(targetLine As Long)
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Dim sLine       As String
    With oModule.CodeModule
        sLine = .Lines(targetLine, 1)
        If Left(Trim(sLine), 4) = "Rem " Then
            .ReplaceLine targetLine, Replace(sLine, "Rem ", "", , 1)
        End If
    End With
End Sub

Public Sub BringProcedureHere()
    '@INCLUDE PROCEDURE ActiveProcedure
    '@INCLUDE PROCEDURE ActiveModule
    '@INCLUDE PROCEDURE ActiveCodepaneWorkbook
    '@INCLUDE CLASS aCodeModule
    '@INCLUDE CLASS aProcedure
    '@AssignedModule aCodeModule
    aProcedure.Init(ActiveCodepaneWorkbook, ActiveModule, ActiveProcedure).BringProcedureHere CodemoduleSelection
End Sub

Public Sub AlignAs()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Align " As ", SearchFromLeft:=True
End Sub

Public Sub AlignColumn()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Align ":", SearchFromLeft:=True
    Align ":", SearchFromLeft:=False
End Sub

Public Sub AlignComments()
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule
    Align "'", SearchFromLeft:=False
End Sub

Public Sub Align( _
                AlignString As String, _
                SearchFromLeft As Boolean, _
                Optional AlignAtColumn As Long)
    '@INCLUDE PROCEDURE StringFormatAlignRowsElements
    '@INCLUDE CLASS aCodeModule
    '@AssignedModule aCodeModule

    Dim S           As String: S = StringFormatAlignRowsElements( _
                                   SelectionLinesCode, _
                                   AlignString, _
                                   SearchFromLeft, _
                                   AlignAtColumn)
    Dim LineFirst   As Long: LineFirst = RowFirst
    oModule.CodeModule.DeleteLines LineFirst, rowCount
    oModule.CodeModule.InsertLines LineFirst, S
    SelectRows
End Sub

Public Property Get ProceduresByName() As Collection
    Dim out           As New Collection
    Dim S             As String: S = CodemoduleSelection
    Dim ProcedureName As String
    Dim Procedures    As Collection
    Set Procedures = aWorkbook.Init(oWorkbook).Procedures(True, False, False, False)
    Dim Procedure   As aProcedure
    For Each Procedure In Procedures
        ProcedureName = Procedure.Name
        If RegexTest(S, ProcedureName) Then out.Add ProcedureName, ProcedureName
    Next
    Set ProceduresByName = out
End Property

Public Function ProceduresByDeclaration() As Collection
    '@AssignedModule aCodeModule
    '@INCLUDE PROCEDURE ProceduresOfTXT
    '@INCLUDE PROCEDURE RegexTest
    '@INCLUDE CLASS aCodeModule
    '@INCLUDE CLASS aProcedure
    Dim out         As New Collection
    Dim S           As String: S = CodemoduleSelection
    Dim ProcedureName As String
    Dim Procedures  As Collection
    Set Procedures = ProceduresOfTXT(S)
    Dim Procedure   As aProcedure
    For Each Procedure In Procedures
        ProcedureName = Procedure.Name
        If RegexTest(S, ProcedureName) Then out.Add ProcedureName, ProcedureName
    Next
    Set ProceduresByDeclaration = out
End Function




Public Sub AssignEnumValues(Optional ToThePower As Boolean = True)
    '@AssignedModule aCodeModule
    '@INCLUDE CLASS aCodeModule
    Dim row1 As Long, row2 As Long
    row1 = ActiveEnumStartLine
    row2 = ActiveEnumEndLine
    Dim Code        As String: Code = cm.Lines(row1 + 1, row2 - row1 - 1)
    Dim arr: arr = Split(Code, vbNewLine)
    Code = ""
    Dim out         As String
    Dim i           As Long
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
    '@AssignedModule aCodeModule
    '@INCLUDE PROCEDURE CLIP
    '@INCLUDE CLASS aCodeModule
    Dim enumName    As String: enumName = ActiveEnumName
    Dim arr: arr = Split(oModule.CodeModule.Lines(ActiveEnumStartLine + 1, ActiveEnumEndLine - ActiveEnumStartLine - 1), vbNewLine)
    Dim out         As String: out = "Select case Variable "
    Dim Code        As String: Code = out
    Dim i           As Long
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
    '@AssignedModule aCodeModule
    '@INCLUDE CLASS aCodeModule
    Dim line        As String
    Dim i           As Long
    For i = aCodeModule.Init(oModule).RowFirst To 1 Step -1
        line = oModule.CodeModule.Lines(i, 1)
        If InStr(1, line, "Enum ") > 0 Then
            ActiveEnumName = Trim(Split(line, "Enum ")(1))
            Exit Function
        End If
    Next
End Function

Private Function ActiveEnumStartLine() As Long
    '@AssignedModule aCodeModule
    '@INCLUDE CLASS aCodeModule
    Dim i           As Long
    Dim line        As String
    For i = aCodeModule.Init(oModule).RowFirst To 1 Step -1
        line = oModule.CodeModule.Lines(i, 1)
        If InStr(1, line, "Enum ") > 0 Then
            ActiveEnumStartLine = i
            Exit Function
        End If
    Next
End Function

Private Function ActiveEnumEndLine() As Long
    '@AssignedModule aCodeModule
    '@INCLUDE CLASS aCodeModule
    Dim i           As Long
    Dim enumName    As String
    Dim line        As String
    For i = aCodeModule.Init(oModule).RowFirst To oModule.CodeModule.CountOfLines
        line = oModule.CodeModule.Lines(i, 1)
        If InStr(1, line, "End Enum") > 0 Then
            ActiveEnumEndLine = i
            Exit Function
        End If
    Next
End Function

Public Function Procedure() As aProcedure
    '@INCLUDE CLASS aProcedure
    '@AssignedModule aCodeModule
    '@INCLUDE PROCEDURE ActiveProcedure
    '@INCLUDE PROCEDURE Toast
    '@INCLUDE CLASS aCodeModule
    If ActiveProcedure = "" Then
        Toast "The carret is not inside a procedure"
        Exit Function
    End If
    Set Procedure = aProcedure.Active
End Function

Public Sub injectDivider(Optional Character As String = "~")
    '@AssignedModule aCodeModule
    '@INCLUDE CLASS aCodeModule
    If Len(Trim(cm.Lines(RowFirst, 1))) = 0 Then
        cm.ReplaceLine RowFirst, "'" & String(Procedure.LargestLineLength - 1, "character")
    Else
        cm.InsertLines RowFirst, "'" & String(Procedure.LargestLineLength - 1, "character")
    End If
End Sub
