VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aProcedure
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

Private oProcedure  As String
Private oWorkbook   As Workbook
Private oModule     As VBComponent
Private oProcKind   As vbext_ProcKind

Private coll_LinkedProcedures As Collection
Private coll_LinkedProceduresDeep As Collection

Public Enum Code_Part
    All_Code
    Declaration_Code
    Declaration_Clean_Code
    Body_Code
    Header_Code
End Enum

Public Enum Property_Type
    Modified
    Ignore
    Parent_Assigned
End Enum

Public Sub InjectTemplate()
    InjectToBodyBottom _
        NewProcedureTemplate( _
            ReturnsObject:=False, _
            addErrorHandling:=True), _
        SkipIfExists:=True
    InjectModification
'    Debug.Print "You may want to use the following after writing the procedure's code:"
'    Debug.Print String(50, "-")
'    Debug.Print "aprocedure.active.InjectObjectsRelease"
'    Debug.Print "aProcedure.Active.InjectLinkedLists"
End Sub

Public Sub InjectTemplateObject()
    InjectToBodyBottom _
        NewProcedureTemplate( _
            ReturnsObject:=True, _
            addErrorHandling:=True), _
        SkipIfExists:=True
    InjectModification
'    Debug.Print "You may want to use the following after writing the procedure's code:"
'    Debug.Print String(50, "-")
'    Debug.Print "aprocedure.active.InjectObjectsRelease"
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 29-08-2023 12:07    Alex                (z_zTest.bas > NewProcedure) initial release

Private Function NewProcedureTemplate(ReturnsObject As Boolean, addErrorHandling As Boolean) As String
'@LastModified 2308291207
'@INCLUDE CLASS aProcedure
'@INCLUDE CLASS aModule
'@INCLUDE CLASS LogFile
'@INCLUDE CLASS CallStack

'
    If ActiveProcedure = "" Then Exit Function
    Dim am As New aModule:      Set am = am.Active
    Dim ap As New aProcedure:   Set ap = aProcedure.Active
    If Trim(Join(Split(ap.Code_Body))) <> vbNullString Then Exit Function
    Const C_Tab = vbTab
    Dim isFunction As Boolean:  isFunction = (ap.KindAsString = "Function" Or ap.KindAsString = "Get")
    
    Dim q As String: q = """"
    If Not am.HeaderContains("C_MODULE_NAME") Then oModule.CodeModule.AddFromString "Private Const C_MODULE_NAME = " & q & am.Name & q
    Dim S As String, msg As String

    S = Join( _
            Array( _
                "'=== Constants ===", _
                C_Tab & "Const C_PROCEDURE_NAME = " & q & ap.Name & q, _
                vbLf, _
                "'=== Variables ===", _
                ""), _
            vbLf)
    If isFunction Then S = Join(Array(S, C_Tab & "Dim Output As " & ap.returnType), vbLf)
    
    S = Join( _
            Array( _
                S, _
                "'=== Initialize ==="), _
            vbLf)
    
    msg = q & Now & vbTab & "Start:" & vbTab & q & " & " & "C_PROCEDURE_NAME"
    S = Join( _
            Array( _
                S, _
                C_Tab & "debug.print " & msg, _
                C_Tab & "LogFile.log " & msg), _
            vbLf)
    
    Dim Args As String: Args = ap.Arguments
    
    S = Join( _
            Array( _
                S, _
                C_Tab & "CallStack.Push C_MODULE_NAME, C_PROCEDURE_NAME" & IIf(Args <> "", ", " & ap.Arguments, ""), _
                IIf(addErrorHandling, C_Tab & "On error goto ErrorHandler", ""), _
                vbLf, _
                "'=== Code ===", _
                vbLf, _
                "'=== Terminate ===", _
                "NormalExit:"), _
            vbLf)

    If isFunction Then S = Join(Array(S, C_Tab & "" & IIf(ReturnsObject, "Set ", "") & ap.Name & " = output", vbLf), vbLf)
    
    msg = VBA.Replace(msg, "Start", "End", , 1, vbTextCompare)
    S = Join( _
            Array( _
                S, _
                C_Tab & "debug.print " & msg, _
                C_Tab & "LogFile.log " & msg, _
                C_Tab & "CallStack.Pop", _
                "'Bellow you may want to call aprocedure.active.InjectObjectsReleaseHere", _
                "'=== Clean Up ===", _
                "", _
                C_Tab & "Exit " & ap.KindAsString, _
                ""), _
            vbLf)
    If addErrorHandling Then
        S = Join( _
                Array( _
                    S, _
                    "ErrorHandler:", _
                    C_Tab & "CallStack.PrintErrorInfo", _
                    C_Tab & "CallStack.LogErrorInfo", _
                    C_Tab & "ERR.CLEAR", _
                    C_Tab & "ERR.Raise 999,C_PROCEDURE_NAME,CALLSTACK.ErrorInfo"), _
                vbLf)
    End If
    NewProcedureTemplate = S
    

'note'
''''''
'    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
'            "Error Number: " & Err.Number & vbCrLf & _
'            "Error Source: TxtOverwrite" & vbCrLf & _
'            "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
End Function

Public Sub Activate()
    oModule.Activate
    Dim startPos    As Long: startPos = InStr(1, oModule.CodeModule.Lines(Me.Line_Declaration_First, 1), Me.Name, vbTextCompare)
    oModule.CodeModule.CodePane.SetSelection Line_Declaration_First, _
                                             startPos, _
                                             Line_Declaration_First, _
                                             startPos + Len(Me.Name)
End Sub

Public Function Init( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Procedure As String, _
        Optional ProcKind As vbext_ProcKind) As aProcedure
    '@AssignedModule aProcedure
    '@INCLUDE PROCEDURE AssignCPSvariables
    '@INCLUDE CLASS aProcedure
    If Not AssignCPSvariables(TargetWorkbook, module, Procedure) Then End
    oProcedure = Procedure
    Set oWorkbook = TargetWorkbook
    Set oModule = module
    oProcKind = ProcKind
    Set Init = Me
    Set coll_LinkedProceduresDeep = Nothing
End Function

Public Function Active() As aProcedure
    '@AssignedModule aProcedure
    '@INCLUDE PROCEDURE ActiveModule
    '@INCLUDE PROCEDURE ActiveCodepaneWorkbook
    '@INCLUDE CLASS aCodeModule
    '@INCLUDE CLASS aProcedure
    Dim module      As VBComponent: Set module = ActiveModule
    Dim Procedure   As String
    Dim i           As Long
    Dim ProcKind    As vbext_ProcKind
    Dim lineNum     As Long: lineNum = aCodeModule.Init(module).RowFirst
    For i = 0 To 4
        ProcKind = i
        Procedure = module.CodeModule.ProcOfLine(lineNum, ProcKind)
        If Procedure <> vbNullString Then
            Init ActiveCodepaneWorkbook, module, Procedure, ProcKind
            Set Active = Me
            Exit For
        End If
    Next
End Function

Public Sub Copy( _
        TargetWorkbook As Workbook, _
        Overwrite As Boolean)
    '@AssignedModule aProcedure
    '@INCLUDE PROCEDURE LinkedProceduresDeep
    '@INCLUDE PROCEDURE ModuleAddOrSet
    '@INCLUDE PROCEDURE ModuleOfProcedure
    '@INCLUDE PROCEDURE ProcedureExists
    '@INCLUDE PROCEDURE ProcedureReplace
    '@INCLUDE CLASS aProcedure
    Dim module      As VBComponent
    Dim S           As String
    Dim Procedure   As Variant
    Dim ap          As aProcedure
    For Each Procedure In LinkedProceduresDeep
        Set ap = New aProcedure
        ap.Init oWorkbook, , CStr(Procedure)
        S = ap.Code_All
        If Not ProcedureExists(TargetWorkbook, CStr(Procedure)) Then
            Set module = ModuleAddOrSet(TargetWorkbook, "vbArcImports", vbext_ct_StdModule)
            module.CodeModule.AddFromString S
        Else
            Set module = ModuleOfProcedure(TargetWorkbook, CStr(Procedure))
            If Overwrite = True Then ProcedureReplace module, CStr(Procedure), S
        End If
    Next
End Sub

Public Function Contains( _
        this As String, _
        WholeWord As Boolean, _
        MatchCase As Boolean, _
        PatternSearch As Boolean) As String
    '@AssignedModule aProcedure
    '@INCLUDE CLASS aProcedure

    Contains = oModule.CodeModule.Find( _
            this, _
            1, _
            Line_First, _
            Line_Last, _
            10000, _
            WholeWord, _
            MatchCase, _
            PatternSearch)
End Function

Public Function Contains_InHeader( _
        this As String, _
        WholeWord As Boolean, _
        MatchCase As Boolean, _
        PatternSearch As Boolean) As String
    '@AssignedModule aProcedure
    '@INCLUDE CLASS aProcedure

    Contains_InHeader = oModule.CodeModule.Find( _
            this, _
            1, _
            Line_Header_First, _
            Line_Header_Last, _
            10000, _
            WholeWord, _
            MatchCase, _
            PatternSearch)
End Function

Public Function Contains_InBody( _
        this As String, _
        WholeWord As Boolean, _
        MatchCase As Boolean, _
        PatternSearch As Boolean) As String
    '@AssignedModule aProcedure
    '@INCLUDE CLASS aProcedure

    Contains_InBody = oModule.CodeModule.Find( _
            this, _
            1, _
            Line_Body_First, _
            Line_Body_Last, _
            10000, _
            WholeWord, _
            MatchCase, _
            PatternSearch)
End Function

Public Function Scope() As String
    '@AssignedModule aProcedure
    '@INCLUDE CLASS aProcedure
    Dim sTitle      As String
    sTitle = Code_Declaration_Clean
    If UCase(sTitle) Like UCase("*" & "Public " & "*" & oProcedure & "*") Then
        Scope = "Public"
    ElseIf UCase(sTitle) Like UCase("*" & "Private " & "*" & oProcedure & "*") Then
        Scope = "Private"
    Else
        Select Case Parent.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                Scope = "Public"
            Case Else
                Scope = "Private"
        End Select
    End If
End Function

Private Sub AssignLineStartEnd(CodePart As Code_Part, ByRef LineStart As Long, ByRef LineEnd As Long)
    '@AssignedModule aProcedure
    '@INCLUDE CLASS aProcedure
    Select Case CodePart
        Case Code_Part.All_Code
            LineStart = Line_First
            LineEnd = Line_Last
        Case Code_Part.Body_Code
            LineStart = Line_Body_First
            LineEnd = Line_Body_Last
        Case Code_Part.Header_Code
            LineStart = Line_Header_First
            LineEnd = Line_Header_Last
    End Select
End Sub

Public Sub CommentsToOwnLine()
    '@AssignedModule aProcedure
    '@INCLUDE PROCEDURE CommentsMoveToOwnLine
    '@INCLUDE CLASS aProcedure
    Dim S           As String: S = CommentsMoveToOwnLine(Code_All)
    Dim i           As Long: i = Line_First
    Delete
    oModule.CodeModule.InsertLines i, S
End Sub

Public Function Property(PropertyType As Property_Type) As Variant
    '@AssignedModule aProcedure
    '@INCLUDE CLASS aProcedure
    Select Case PropertyType

        Case Property_Type.Modified: Property = PropertyGet(Modified)
        Case Property_Type.Parent_Assigned: Property = PropertyGet(Parent_Assigned)

    End Select
End Function

Public Property Get ArgumentCount() As Long
    Dim S           As String: S = Code_Declaration_Clean
    If S Like "*" & oProcedure & "()*" Then Exit Property
    ArgumentCount = UBound(Split(S, ",")) + 1
End Property

Public Property Get Name() As String
    Name = oProcedure
End Property

Public Property Get returnType() As String
    returnType = "Invalid"
    If Not (KindAsString = "Function" Or KindAsString = "Get") Then Exit Property
    Dim Title       As String: Title = Code_Declaration_Clean
    If Title Like "*) As *" Then
        returnType = Mid(Title, InStrRev(Title, ") As ") + Len(") As ")) ' Split(Title, ") As ")(1)
    Else
        returnType = "Variant"
    End If
End Property

Public Property Get Parent() As VBComponent
    Set Parent = oModule
End Property

Public Property Get ParentAssigned() As VBComponent
    Dim ModuleName  As String: ModuleName = Property(Parent_Assigned)
    If ModuleName = "" Then Exit Property
    Set ParentAssigned = ModuleAddOrSet(oWorkbook, ModuleName, vbext_ct_StdModule)
    'End Property

    'Public Property Set ParentAssigned(Module As VBComponent)
    '
    '    Set oModule = Module
End Property

Public Property Get KindAsLong() As Long
    Dim S           As String: S = Code_Declaration_Clean
    Select Case True
        Case InStr(1, S, "Get " & oProcedure) > 0: KindAsLong = vbext_pk_Get    '3
        Case InStr(1, S, "Let " & oProcedure) > 0: KindAsLong = vbext_pk_Let    '1
        Case InStr(1, S, "Set " & oProcedure) > 0: KindAsLong = vbext_pk_Set    '2
        Case Else: KindAsLong = vbext_pk_Proc    '0
    End Select
End Property

Public Property Get KindAsString() As String
    Dim S           As String: S = Code_Declaration_Clean
    Select Case True
        Case InStr(1, S, "Get " & oProcedure) > 0: KindAsString = "Get"
        Case InStr(1, S, "Let " & oProcedure) > 0: KindAsString = "Let"
        Case InStr(1, S, "Set " & oProcedure) > 0: KindAsString = "Set"
        Case Else
            If InStr(1, S, "Function " & oProcedure) > 0 Then
                KindAsString = "Function"
            ElseIf InStr(1, S, "Sub " & oProcedure) > 0 Then
                KindAsString = "Sub"
            End If
    End Select
End Property


Private Sub Class_Terminate()
    ReSetVariables
End Sub

Private Sub ReSetVariables()
    Set oWorkbook = Nothing
    Set oModule = Nothing
    oProcedure = ""
End Sub

'Public Function lineIndex(idx As Line_Index) As Long
'    Select Case idx
'        Case Line_Index.Procedure_First: lineIndex = Line_First
'        Case Line_Index.Procedure_Last: lineIndex = Line_Last
'        Case Line_Index.Procedure_Count: lineIndex = Line_Count
'        Case Line_Index.Declaration_First: lineIndex = Line_Declaration_First
'        Case Line_Index.Declaration_Last: lineIndex = Line_Declaration_Last
'        Case Line_Index.Declaration_Count: lineIndex = Line_Declaration_Count
'        Case Line_Index.Body_First: lineIndex = Line_Body_First
'        Case Line_Index.Body_First_AfterComments: lineIndex = Line_Body_FirstAfterComments
'        Case Line_Index.Body_Last: lineIndex = Line_Body_Last
'        Case Line_Index.Body_Count: lineIndex = Line_Body_Count
'        Case Line_Index.Header_First: lineIndex = Line_Header_First
'        Case Line_Index.Header_Last: lineIndex = Line_Header_Last
'        Case Line_Index.Header_Count: lineIndex = Line_Header_Count
'
'    End Select
'End Function

Public Function LineLike(this As String, CodePart As Code_Part) As Long
    Dim LineStart As Long, LineEnd As Long
    AssignLineStartEnd CodePart, LineStart, LineEnd
    Dim targetLine  As String
    Dim i           As Long
    For i = LineStart To LineEnd
        targetLine = Trim(oModule.CodeModule.Lines(i, 1))
        If UCase(targetLine) Like UCase(this) Then
            LineLike = i
            Exit Function
        End If
    Next
End Function

Public Sub EnsureBlankLineBeforeProcedure()
    Dim line        As Long: line = Line_Declaration_First
    If line = 1 Then oModule.CodeModule.InsertLines 1, "": Exit Sub
    Dim check       As String
    On Error Resume Next
    check = oModule.CodeModule.ProcOfLine(line - 1, oProcKind)
    On Error GoTo 0
    If check <> Name Then oModule.CodeModule.InsertLines line, ""
End Sub

Public Function Line_Header_First() As Long
    EnsureBlankLineBeforeProcedure
    Line_Header_First = Line_First
End Function

Public Function Line_Header_Last() As Long
    EnsureBlankLineBeforeProcedure
    Line_Header_Last = Line_Declaration_First - 1
End Function

Public Function Line_Header_Count()
    Line_Header_Count = Line_Header_Last - Line_Header_First + 1
End Function

Public Function Line_First() As Long
    '@todo
    Line_First = oModule.CodeModule.procStartLine(oProcedure, oProcKind)    'findProcedureKind(omodule,oprocedure)
    If Line_First = 0 Then Debug.Print "error in firstline (wrong procKind)": End
End Function

Public Function Line_Count() As Long
    Line_Count = oModule.CodeModule.ProcCountLines(oProcedure, oProcKind)
End Function

Public Function Line_Last() As Long
    Line_Last = Line_First + Line_Count - 1
End Function

Public Function Line_Body_First() As Long
    Line_Body_First = Line_Declaration_First + Line_Declaration_Count
End Function

Public Function Line_Body_FirstAfterComments() As Long
    Dim i           As Long
    Dim S           As String
    For i = Line_Body_First To Line_Body_Last
        S = Trim(oModule.CodeModule.Lines(i, 1))
        If S = vbNullString Then
            Exit For
        ElseIf Left(S, 1) = "'" Then
        ElseIf Left(S, 3) = "Rem" Then
        ElseIf Right(Trim(oModule.CodeModule.Lines(i - 1, 1)), 1) = "_" Then
        ElseIf Right(S, 1) = "_" Then
        Else
            Exit For
        End If
    Next
    Line_Body_FirstAfterComments = i
End Function

Public Function Line_Body_Last() As Long
    Line_Body_Last = Line_Last - 1
End Function

Public Function Line_Body_Count() As Long
    Line_Body_Count = Line_Body_Last - Line_Body_First + 1
End Function

Public Function Line_Declaration_First() As Long
    Line_Declaration_First = oModule.CodeModule.ProcBodyLine(oProcedure, oProcKind)
End Function

Public Function Line_Declaration_Last() As Long
    Line_Declaration_Last = Line_Declaration_First + Line_Declaration_Count - 1
End Function

Public Function Line_Declaration_Count() As Long
    Dim targetLine  As Long: targetLine = Line_Declaration_First
    Dim counter     As Long: counter = 1
    Dim S           As String: S = oModule.CodeModule.Lines(targetLine, counter)
    Do While Right(S, 1) = "_"
        counter = counter + 1
        S = oModule.CodeModule.Lines(targetLine, counter)
    Loop
    Line_Declaration_Count = counter
End Function

Public Function Code_All()
    Code_All = oModule.CodeModule.Lines(Line_First, Line_Count)
End Function
Public Function Code_Body()
    Code_Body = oModule.CodeModule.Lines(Line_Body_First, Line_Body_Count)
End Function
Public Function Code_Header()
    Code_Header = oModule.CodeModule.Lines(Line_Header_First, Line_Header_Count)
End Function
Public Function Code_Declaration()
    Code_Declaration = oModule.CodeModule.Lines(Line_Declaration_First, Line_Declaration_Count)
End Function
Public Function Code_Declaration_Clean()
    If InStr(1, Code_Declaration, "_") = 0 Then
        Code_Declaration_Clean = Trim(Code_Declaration)
    Else
        Code_Declaration_Clean = VBA.Replace(Join(ArrayRemoveEmptyElements(ArrayTrim(Split(Code_Declaration, "_" & vbNewLine))), " "), "( ", "(")
    End If
End Function

Public Function collLinkedProceduresDeep() As Collection
    If coll_LinkedProceduresDeep Is Nothing Then
        Set coll_LinkedProceduresDeep = LinkedProceduresDeep
    End If
    Set collLinkedProceduresDeep = coll_LinkedProceduresDeep
End Function

Public Function collLinkedProcedures() As Collection
    If coll_LinkedProcedures Is Nothing Then
        Set coll_LinkedProcedures = LinkedProcedures
    End If
    Set collLinkedProcedures = coll_LinkedProcedures
End Function

Public Function LinkedProcedures() As Collection
    '@TODO befor Procedues was a collection of strings, now objects. Use aWorkbook...ProceduresArray???
    Dim Procedures  As Collection
    Set Procedures = aWorkbook.Init(oWorkbook).ProceduresNames(True, False, False, False)
    Dim S           As String: S = Code_All
    Dim coll        As New Collection
    Dim Procedure   As Variant
    For Each Procedure In Procedures
        If UCase(CStr(Procedure)) <> UCase(CStr(oProcedure)) Then
            If RegexTest(S, "\W" & CStr(Procedure) & "[.(\W]") = True Then
                coll.Add Procedure, CStr(Procedure)
            End If
        End If
    Next
    Set LinkedProcedures = coll
End Function

Public Function LinkedProceduresDeep() As Collection
    '@LastModified 2307201045
    Dim AllProcedures As Collection: Set AllProcedures = aWorkbook.Init(oWorkbook).ProceduresNames(True, False, False, False)
    Dim Processed   As Collection: Set Processed = New Collection
    Dim CalledProcedures As Collection: Set CalledProcedures = New Collection

    Dim Procedure   As Variant
    Dim module      As VBComponent

    Processed.Add CStr(oProcedure), CStr(oProcedure)
    CalledProcedures.Add CStr(oProcedure), CStr(oProcedure)
    On Error Resume Next
    For Each Procedure In LinkedProcedures
        CalledProcedures.Add CStr(Procedure), CStr(Procedure)
    Next
    On Error GoTo 0

    Dim ap          As aProcedure
    Dim CalledProceduresCount As Long
    CalledProceduresCount = CalledProcedures.Count
    Dim element
repeat:
    For Each element In CalledProcedures
        If Not aCollection.Init(Processed).Contains(, CStr(element)) Then
            On Error Resume Next
            Set ap = New aProcedure
            For Each Procedure In ap.Init(oWorkbook, , CStr(element)).LinkedProcedures
                CalledProcedures.Add CStr(Procedure), CStr(Procedure)
            Next
            On Error GoTo 0
            Processed.Add CStr(element), CStr(element)
        End If
    Next
    If CalledProcedures.Count > CalledProceduresCount Then
        CalledProceduresCount = CalledProcedures.Count
        GoTo repeat
    End If

    Set LinkedProceduresDeep = aCollection.Init(CalledProcedures).Sort.items
End Function

Sub UnfoldDeclaration()
    Dim result      As String: result = Code_Declaration_Clean & vbLf & Code_Body & vbLf & oModule.CodeModule.Lines(Line_Last, 1)
    Replace result
End Sub
Sub FoldDeclaration()
    Dim startLine   As Long
    Dim LastLine    As Long
    Dim FirstColumn As Long
    Dim LastColumn  As Long

    If ArgumentCount = 0 Then Exit Sub
    Dim cp          As aCodeModule
    Set cp = aCodeModule.Init(oModule)

    startLine = Line_Declaration_First
    FirstColumn = InStr(1, oModule.CodeModule.Lines(startLine, 1), "(")
    LastLine = Line_Declaration_Last
    '    Do While InStr(1, oModule.CodeModule.Lines(LastLine, 1), ")") = 0
    '        LastLine = LastLine - 1
    '    Loop
    LastColumn = InStr(1, oModule.CodeModule.Lines(LastLine, 1), ")") + 1
    If LastColumn = 0 Then Exit Sub

    cp.SetSelection startLine, FirstColumn, LastLine, LastColumn
    cp.FoldLine
    cp.SetSelection startLine, FirstColumn, Line_Declaration_Last, 10000
End Sub

Public Sub Export()
    Dim Procedure
    For Each Procedure In collLinkedProceduresDeep
        LinkedExport CStr(Procedure)
    Next
End Sub

Private Sub LinkedExport(Procedure As String)
    Dim module      As VBComponent
    Set module = ModuleOfProcedure(oWorkbook, Procedure)
    Dim c           As New aProcedure
    c.Init oWorkbook, module, Procedure
    Dim proclastmod As String: proclastmod = c.Property(Modified)
    If proclastmod = vbNullString Then
        c.InjectLinkedLists
        c.PropertySet Modified, Format(Now, "yymmddhhnn")
        proclastmod = c.Property(Modified)
    End If
    Dim S           As String: S = c.Code_All
    Dim FileFullName As String: FileFullName = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"

    If FileExists(FileFullName) Then
        Dim filelastmod
        filelastmod = StringLastModified(TxtRead(FileFullName))
        If proclastmod > filelastmod Then
            Debug.Print "OVERWROTE " & Procedure
            TxtOverwrite FileFullName, S
        Else
            Debug.Print "Skipping " & Procedure & " because:"
            Debug.Print vbTab & "StringLastModified(TxtRead(FileFullName)) >= Property(Modified)"
        End If
    Else
        Debug.Print "NEW " & Procedure
        TxtOverwrite FileFullName, S
    End If
    Dim element
    For Each element In c.LinkedUserforms
        oWorkbook.VBProject.VBComponents(element).Export LOCAL_LIBRARY_USERFORMS & element & ".frm"
    Next
    For Each element In c.LinkedClasses
        oWorkbook.VBProject.VBComponents(element).Export LOCAL_LIBRARY_CLASSES & element & ".cls"
    Next
    For Each element In c.LinkedDeclarations
        ExportDeclaration CStr(element)
    Next
End Sub

Sub ExportDeclaration(DeclarationName As String)
    aWorkbook.Init(oWorkbook).DeclarationsTableCreate
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")

    Dim codeName    As String
    Dim codeText    As String
    Dim cell        As Range
    On Error Resume Next
    Set cell = TargetWorksheet.Columns(3).Find(DeclarationName, LookAt:=xlWhole)
    On Error GoTo 0
    If cell Is Nothing Then Exit Sub

    codeName = DeclarationName
    codeText = cell.offset(0, 1).TEXT
    TxtOverwrite LOCAL_LIBRARY_DECLARATIONS & DeclarationName & ".txt", codeText

End Sub
Public Sub InjectLinkedLists()
    RemoveIncludeLines
    InjectLinkedDeclarations
    InjectLinkedClasses
    InjectLinkedUserforms
    InjectLinkedProcedures
End Sub

Public Sub Indent()
    Dim S           As String
    Dim nIndent     As Long
    Dim i           As Long
    For i = Line_First To Line_Last
        S = LTrim$(oModule.CodeModule.Lines(i, 1))
        If Trim(S) <> "" Then
            If IsBlockEnd(S) Then nIndent = nIndent - 1
            If nIndent < 0 Then nIndent = 0
            S = Space$(nIndent * 4) & S
            oModule.CodeModule.ReplaceLine i, S
            If IsBlockStart(LTrim$(S)) Then nIndent = nIndent + 1
        End If
    Next
End Sub

Public Sub NumbersAdd()
    Dim counter     As Long: counter = 1
    Dim LineStart As Long, LineEnd As Long
    AssignLineStartEnd Body_Code, LineStart, LineEnd
    Dim S           As String
    Dim i           As Long
    For i = LineStart To LineEnd
        S = oModule.CodeModule.Lines(i, 1)
        If IsLineNumberAble(S) _
                And Right(Trim(oModule.CodeModule.Lines(i - 1, 1)), 1) <> "_" Then
            oModule.CodeModule.ReplaceLine i, counter & ":" & S
            counter = counter + 1
        End If
    Next i
End Sub

Public Sub NumbersRemove()
    Dim counter     As Long: counter = 1
    Dim LineStart As Long, LineEnd As Long
    AssignLineStartEnd Body_Code, LineStart, LineEnd
    Dim S           As String
    Dim i           As Long
    For i = LineStart To LineEnd
        S = oModule.CodeModule.Lines(i, 1)
        If IsNumeric(Left(Trim(S), 1)) Then
            oModule.CodeModule.ReplaceLine i, _
                    Mid(S, InStr(S, ":") + 1)
            'Space(InStr(s, ":"))
        End If
    Next i
End Sub

Public Sub RemoveEmptyLines()
    Dim LineStart As Long, LineEnd As Long
    AssignLineStartEnd Body_Code, LineStart, LineEnd
    Dim i           As Long
    For i = LineEnd To LineStart Step -1
        If Len(Trim(oModule.CodeModule.Lines(i, 1))) = 0 Then
            oModule.CodeModule.DeleteLines i, 1
        End If
    Next
End Sub

Public Sub CommentsRemove(RemoveRem As Boolean, CodePart As Code_Part)
    Replace StringCommentsRemove(CodePart, RemoveRem)

    '    Dim N               As Long
    '    Dim lineText        As String
    '    Dim QUOTES          As Long
    '    Dim Q               As Long
    '    Dim StartPos        As Long
    '    Dim L As Long
    '    Dim LineStart As Long, LineEnd As Long
    '    AssignLineStartEnd CodePart, LineStart, LineEnd
    '    Dim j As Long
    '    For j = LineEnd To LineStart Step -1
    '        lineText = LTrim(oModule.CodeModule.Lines(j, 1))
    '        If Not RemoveRem Then If lineText Like "Rem *" Then GoTo SKIP
    '        StartPos = 1
    'Retry:
    '        N = InStr(StartPos, lineText, "'")
    '        Q = InStr(StartPos, lineText, """")
    '        QUOTES = 0
    '        If Q < N Then
    '            For L = 1 To N
    '                If Mid(lineText, L, 1) = """" Then
    '                    QUOTES = QUOTES + 1
    '                End If
    '            Next L
    '        End If
    '        If QUOTES = Application.WorksheetFunction.Odd(QUOTES) Then
    '            StartPos = N + 1
    '            GoTo Retry:
    '        Else
    '            Select Case N
    '                Case Is = 0
    '
    '                Case Is = 1
    '                    oModule.CodeModule.DeleteLines j, 1
    '                Case Is > 1
    '                    oModule.CodeModule.ReplaceLine j, Left(oModule.CodeModule.Lines(j, 1), N - 1)
    '            End Select
    '        End If
    'SKIP:
    '    Next j
End Sub

Public Sub RemoveIncludeLines()
    RemoveLinesLike "'@INCLUDE *"
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 13:56    Alex                (aProcedure.cls > RemoveLinesLike)

Public Sub RemoveLinesLike(this As String)
'@LastModified 2308221356
    this = UCase(this)
    Dim S           As String
    Dim i           As Long
    For i = Line_Body_Last To Line_Header_First Step -1
        S = Trim(oModule.CodeModule.Lines(i, 1))
        S = UCase(S)
        If S Like this Then
            oModule.CodeModule.DeleteLines i
        End If
    Next
End Sub

Public Sub InjectLinkedClasses()
    Dim ListOfImports As String
    Dim S           As String: S = Code_All
    Dim element     As Variant
    For Each element In LinkedClasses
        If InStr(1, S, "@INCLUDE CLASS " & element) = 0 _
                And InStr(1, ListOfImports, "@INCLUDE CLASS " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE CLASS " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE CLASS " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        oModule.CodeModule.InsertLines _
                Line_Body_FirstAfterComments, ListOfImports
    End If
End Sub

Public Sub InjectLinkedProcedures()
    Dim Procedures  As Collection
    Set Procedures = collLinkedProcedures
    Dim ListOfImports As String
    Dim S           As String: S = Code_All
    Dim Procedure   As Variant
    For Each Procedure In Procedures
        If InStr(1, S, "@INCLUDE PROCEDURE " & Procedure) = 0 And InStr(1, ListOfImports, "@INCLUDE PROCEDURE " & Procedure) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE PROCEDURE " & Procedure
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE PROCEDURE " & Procedure
            End If
        End If
    Next
    If ListOfImports <> "" Then
        oModule.CodeModule.InsertLines Line_Body_FirstAfterComments, ListOfImports
    End If
End Sub

Public Sub InjectLinkedUserforms()
    Dim ListOfImports As String
    Dim S           As String: S = Code_All
    Dim element     As Variant
    For Each element In LinkedUserforms
        If InStr(1, S, "@INCLUDE USERFORM " & element) = 0 And InStr(1, ListOfImports, "@INCLUDE USERFORM " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE USERFORM " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE USERFORM " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        oModule.CodeModule.InsertLines Line_Body_FirstAfterComments, ListOfImports
    End If
End Sub

Public Sub InjectLinkedDeclarations()
    Dim ListOfImports As String
    Dim S           As String: S = Code_All
    Dim coll        As New Collection
    Dim element     As Variant
    For Each element In LinkedDeclarations
        If InStr(1, S, "'@INCLUDE DECLARATION " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE DECLARATION " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE DECLARATION " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        oModule.CodeModule.InsertLines Line_Body_FirstAfterComments, ListOfImports
    End If
End Sub

Public Function LinkedDeclarations() As Collection
    aWorkbook.Init(oWorkbook).DeclarationsTableCreate
    Dim TargetWorksheet As Worksheet: Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    Dim coll        As New Collection
    Dim S           As String: S = Code_All
    Dim element
    For Each element In DeclarationsTableKeywords
        If RegexTest(S, "\b ?" & CStr(element) & "\b") Then
            On Error Resume Next
            coll.Add CStr(element), CStr(element)
            On Error GoTo 0
        End If
    Next
    Set LinkedDeclarations = coll
End Function

Function DeclarationsTableKeywords() As Collection
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    Dim Lr          As Long: Lr = getLastRow(TargetWorksheet)
    Dim coll        As New Collection
    Dim cell        As Range
    For Each cell In TargetWorksheet.Range("C2:C" & Lr)
        On Error Resume Next
        coll.Add cell.TEXT, cell.TEXT
        On Error GoTo 0
    Next
    Set DeclarationsTableKeywords = coll
End Function

Public Function LinkedUserforms()
    Dim coll        As New Collection
    Dim S           As String: S = Code_All
    Dim FormName
    For Each FormName In aModules.Init(oWorkbook).UserformNames
        If RegexTest(S, "\W" & FormName & "[.(\W]") = True Then coll.Add FormName
    Next
    Set LinkedUserforms = coll
End Function




Public Function LinkedClasses() As Collection
    Dim coll        As New Collection
    Dim var         As Variant
    var = aModule.Init(oModule).ClassCalls
    Dim S           As String: S = Code_All
    Dim Keyword     As String
    Dim ClassName   As String
    Dim element     As Variant
    Dim i           As Long
    On Error Resume Next
    For i = LBound(var, 1) To UBound(var, 1)
        If InStr(1, S, var(i, 1)) > 0 Or InStr(1, S, var(i, 2)) > 0 Then
            coll.Add var(i, 1), var(i, 1)
        End If
    Next
    For Each element In aModules.Init(oWorkbook).ClassNames
        If InStr(1, S, element) > 0 Then
            coll.Add element, CStr(element)
        End If
    Next
    On Error GoTo 0
    Set LinkedClasses = coll
End Function


Function LinkedSheets() As Collection
    Dim SheetNames  As New Collection
    Dim TargetWorksheet As Worksheet
    For Each TargetWorksheet In oWorkbook.Worksheets
        On Error Resume Next
        SheetNames.Add TargetWorksheet.Name
        SheetNames.Add oWorkbook.Worksheets(TargetWorksheet.Name).codeName
        On Error GoTo 0
    Next
    Dim element
    Dim S           As String: S = Code_All
    Dim coll        As New Collection
    For Each element In SheetNames
        If InStr(1, CStr(element), " ") > 0 And InStr(1, S, CStr(element)) > 0 Then
            coll.Add CStr(element)
        ElseIf RegexTest(S, "\W" & CStr(element) & "[.(\W]") = True Then
            coll.Add CStr(element), CStr(element)
        End If
    Next
    Set LinkedSheets = coll
End Function


Public Sub ExportLinkedCode()
    Dim Code        As String: Code = LinkedCode
    Dim MergedName  As String: MergedName = "Merged_" & oProcedure
    Dim fileName    As String: fileName = LOCAL_LIBRARY_PROCEDURES & MergedName & ".txt"
    Debug.Print "OVERWROTE " & MergedName
    TxtOverwrite fileName, Code
    TxtPrependContainedProcedures fileName
    FollowLink fileName
End Sub

Private Function LinkedCode() As String
    Dim MergedString As String: MergedString = Code_All
    Dim Procedure
    For Each Procedure In collLinkedProceduresDeep
        MergedString = MergedString & vbNewLine & aProcedure.Init(oWorkbook, , CStr(Procedure)).Code_All
    Next
    LinkedCode = MergedString
End Function

Public Sub ImportDependencies(Optional Overwrite As Boolean)
    ImportProcedureDependencies oProcedure, Overwrite
End Sub

Private Sub ImportProcedureDependencies( _
        Procedure As String, _
        Overwrite As Boolean)

    Dim module      As VBComponent
    On Error Resume Next
    Set module = ModuleOfProcedure(oWorkbook, Procedure)
    If module Is Nothing Then Exit Sub
    On Error GoTo 0
    Dim obj         As String
    Dim Code        As String
    Code = aProcedure.Init(oWorkbook, module, Procedure).Code_All
    Dim var
    var = Split(Code, vbNewLine)
    var = Filter(var, "'@INCLUDE ")
    Dim TextLine    As Variant
    For Each TextLine In var
        TextLine = UCase(Trim(TextLine))
        If TextLine Like "'@INCLUDE *" Then
            obj = Split(TextLine, " ")(2)
            obj = VBA.Replace(obj, vbNewLine, "")
            Select Case True
                Case TextLine Like "'@INCLUDE PROCEDURE *": ImportProcedure obj, Overwrite
                Case TextLine Like "'@INCLUDE CLASS *": ImportClass obj, Overwrite
                Case TextLine Like "'@INCLUDE USERFORM *": ImportUserform obj, Overwrite
                Case TextLine Like "'@INCLUDE DECLARATION *": ImportDeclaration obj
            End Select
        End If
    Next
End Sub

Sub Update()
    ImportProcedure oProcedure, True
End Sub

Public Sub ImportProcedure( _
        Procedure As String, _
        Overwrite As Boolean)
    '
    Dim ProcedurePath As String: ProcedurePath = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"
    Dim result      As String
    On Error Resume Next
    result = TxtRead(ProcedurePath)
    On Error GoTo 0

    If Len(result) = 0 Then
        On Error Resume Next
        result = TXTReadFromUrl(GITHUB_LIBRARY_PROCEDURES & Procedure & ".txt")
        On Error GoTo 0
        If Len(result) > 0 And Not UCase(result) Like ("*NOT FOUND*") Then
            TxtOverwrite ProcedurePath, result
        Else
            Debug.Print "File " & Procedure & ".txt not found neither localy nor online"
            Exit Sub
        End If
    End If

    Dim filelastmod: filelastmod = StringLastModified(result)
    Dim proclastmod

    Dim module      As VBComponent
    If ProcedureExists(oWorkbook, Procedure) = True Then
        Set module = ModuleOfProcedure(oWorkbook, Procedure)
        proclastmod = ProcedureLastModified(oWorkbook, module, Procedure)
        If Overwrite = True Then
            If proclastmod = 0 Or proclastmod < filelastmod Then
                ProcedureReplace module, Procedure, TxtRead(ProcedurePath)
            End If
        End If
    Else
        '        Dim ModuleName As String
        '            ModuleName = StringProcedureAssignedModule(Result)
        '        If ModuleName = "" Then ModuleName = "vbArcImports"
        '        Set Module = ModuleAddOrSet(oWorkbook, ModuleName, vbext_ct_StdModule)
        Set module = ModuleAddOrSet(oWorkbook, "vbArcImports", vbext_ct_StdModule)
        module.CodeModule.AddFromFile ProcedurePath
    End If

    ImportProcedureDependencies Procedure, Overwrite
End Sub

Sub ImportDeclaration(DeclarationName As String)
    Dim FilePath    As String
    FilePath = LOCAL_LIBRARY_DECLARATIONS & DeclarationName & ".txt"
    Dim result      As String
    On Error Resume Next
    result = TxtRead(FilePath)
    On Error GoTo 0

    If Len(result) = 0 Then    'CheckPath(filePath) = "I" Then
        On Error Resume Next
        result = TXTReadFromUrl(GITHUB_LIBRARY_DECLARATIONS & DeclarationName & ".txt")
        On Error GoTo 0
        If Len(result) > 0 And Not UCase(result) Like ("*NOT FOUND*") Then
            TxtOverwrite FilePath, result
        Else
            Debug.Print "File " & DeclarationName & ".txt not found localy or online"
            Exit Sub
        End If
    Else

    End If
    If InStr(1, aWorkbook.Init(oWorkbook).Code, result, vbTextCompare) > 0 Then Exit Sub
    Dim module      As VBComponent
    Set module = ModuleAddOrSet(oWorkbook, "vbArcImports", vbext_ct_StdModule)
    module.CodeModule.AddFromString FormatVBA7(result)

End Sub

Sub ImportUserform(UserformName As String, _
        Overwrite As Boolean)

    Dim FilePathFrM As String
    FilePathFrM = LOCAL_LIBRARY_USERFORMS & UserformName & ".frm"
    Dim FilePathFrX As String
    FilePathFrX = LOCAL_LIBRARY_USERFORMS & UserformName & ".frx"

    If CheckPath(FilePathFrM) = "I" Then
        On Error Resume Next
        Dim codeFrM As String
        codeFrM = TXTReadFromUrl(GITHUB_LIBRARY_USERFORMS & UserformName & ".frm")
        Dim codeFrX As String
        codeFrX = TXTReadFromUrl(GITHUB_LIBRARY_USERFORMS & UserformName & ".frx")
        On Error GoTo 0
        If Len(codeFrM) > 0 And Len(codeFrX) > 0 Then
            TxtOverwrite FilePathFrM, codeFrM
            TxtOverwrite FilePathFrX, codeFrX
        Else
            Debug.Print "File " & UserformName & ".frm/.frx not found neither localy nor online"
            Exit Sub
        End If
    End If

    If ModuleExists(UserformName, oWorkbook) Then
        If Overwrite = True Then
            oWorkbook.VBProject.VBComponents.Remove oWorkbook.VBProject.VBComponents(UserformName)
        Else
            Exit Sub
        End If
    End If
    oWorkbook.VBProject.VBComponents.Import FilePathFrM
End Sub

Sub ImportClass(ClassName As String, _
        Overwrite As Boolean)

    Dim FilePath    As String
    FilePath = LOCAL_LIBRARY_CLASSES & ClassName & ".cls"
    If CheckPath(FilePath) = "I" Then
        On Error Resume Next
        Dim Code    As String
        Code = TXTReadFromUrl(GITHUB_LIBRARY_CLASSES & ClassName & ".cls")
        On Error GoTo 0
        If Len(Code) > 0 And Not UCase(Code) Like ("*NOT FOUND*") Then
            TxtOverwrite FilePath, Code
        Else
            MsgBox "File " & ClassName & ".cls not found neither localy nor online"
            Exit Sub
        End If
    End If

    If ModuleExists(ClassName, oWorkbook) Then
        If Overwrite = True Then
            oWorkbook.VBProject.VBComponents.Remove oWorkbook.VBProject.VBComponents(ClassName)
        Else
            Exit Sub
        End If
    End If
    oWorkbook.VBProject.VBComponents.Import FilePath
End Sub

Public Sub Replace(result As String)

    Dim startLine   As Integer
    Dim NumLines    As Integer
    With oModule.CodeModule
        startLine = .procStartLine(oProcedure, oProcKind)
        NumLines = .ProcCountLines(oProcedure, oProcKind)
        .DeleteLines startLine, NumLines
        .InsertLines startLine, result
    End With
End Sub

Private Function EnumOptionToString(PropertyType As Property_Type) As String
    Select Case PropertyType
        Case Property_Type.Modified: EnumOptionToString = "LastModified"
        Case Property_Type.Ignore: EnumOptionToString = "Ignore"
        Case Property_Type.Parent_Assigned: EnumOptionToString = "AssignedModule"

    End Select
End Function

Private Function PropertyGet(PropertyType As Property_Type) As String
    PropertyGet = vbNullString
    Dim line        As String: line = PropertyLine(PropertyType)
    If line = 0 Then Exit Function
    Dim targetLine  As String: targetLine = Trim(oModule.CodeModule.Lines(line, 1))
    PropertyGet = Split(targetLine, " ")(1)
End Function

Public Function PropertySet(PropertyType As Property_Type, Value As String)
    Dim line        As Long: line = PropertyLine(PropertyType)
    If line = 0 Then
        oModule.CodeModule.InsertLines Line_Body_First, "'@" & EnumOptionToString(PropertyType) & " " & Value
    Else
        Dim targetLine As String: targetLine = Trim(oModule.CodeModule.Lines(line, 1))
        oModule.CodeModule.ReplaceLine line, Split(targetLine)(0) & " " & Value
    End If

End Function

Private Function PropertyLine(PropertyType As Property_Type) As Long
    Dim this        As String: this = EnumOptionToString(PropertyType)
    Dim targetLine  As String
    Dim i           As Long
    For i = Line_First To Line_Last
        targetLine = Trim(oModule.CodeModule.Lines(i, 1))
        If UCase(targetLine) Like UCase("'@" & this & " *") Then
            PropertyLine = i
            Exit Function
        End If
    Next
End Function

Public Function ObjectsReleaseText() As String
    Dim LineStart As Long, LineEnd As Long
    AssignLineStartEnd Body_Code, LineStart, LineEnd
    Dim i           As Long
    Dim S As String, this As String, Keyword As String
    For i = LineStart To LineEnd
        S = Trim(oModule.CodeModule.Lines(i, 1))
        If (S Like "Set * = *" Or S Like "Dim*As New*") _
                And Not S Like "*= Nothing*" Then
            Keyword = Split(S, " ")(1)
            this = IIf(this <> "", this & vbNewLine, "") & "Set " & Keyword & " = Nothing"
        End If
    Next
    If this = "" Then Exit Function
    Dim var: var = Split(this, vbNewLine)
    Dim sCode As String: sCode = Code_All
    Dim out, el
    ReDim out(0 To 0)
    For Each el In var
        If InStr(1, sCode, el) = 0 Then
            If out(0) <> "" Then ReDim Preserve out(0 To UBound(out) + 1)
            out(UBound(out)) = el
        End If
    Next
    this = Join(out, vbNewLine)
    ObjectsReleaseText = this
End Function
Public Sub InjectObjectsReleaseAtEnd()
    Dim this As String: this = ObjectsReleaseText
    If this <> "" Then oModule.CodeModule.InsertLines Line_Last, this
End Sub
Public Sub InjectObjectsReleaseHere()
    aCodeModule.Active.Insert Space(4) & Join(Split(ObjectsReleaseText, vbNewLine), vbNewLine & Space(4))
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 18-08-2023 08:31    Alex                (aProcedure.cls > InjectModification)
'* Updated    : 26-10-2023 00:32    Alex                (aProcedure.cls > InjectModification : added argument for description)

Public Sub InjectModification(Optional Desc As String = "")
'@LastModified 2310260032
    PropertySet Modified, Format(Now, "yymmddhhnn")

    Const vbTab2 = vbTab & vbTab
    Const vbTab4 = vbTab2 & vbTab2
    Dim sTime       As String: sTime = Format(Now, "dd-mm-yyyy hh:nn")
    Dim sProc       As String: sProc = oProcedure

    Dim sUser       As String: sUser = "Alex"
    If sUser = vbNullString Then sUser = Environ("UserName")
    Const sUPDATE   As String = "'* Updated    :"
    Dim sFirstLine  As String
    sFirstLine = "'* Modified   :" & vbTab & "Date and Time" & vbTab2 & _
            "Author" & vbTab4 & "Description" & vbCrLf
    Dim sSecondLine As String
    sSecondLine = sUPDATE & vbTab & sTime & vbTab & sUser & vbTab4 & "(" & oModule.Name & aModule.Init(oModule).Extension & " > " & sProc & " : " & Desc & ")"

    If InStr(1, Code_Header, "'* Created    :") Then
        Dim lineNo  As Long
        Dim i       As Long
        For i = Line_Header_First To Line_Header_Last
            If oModule.CodeModule.Lines(i, 1) Like "'* Created    :*" Then
                lineNo = i
                Exit For
            End If
        Next
        If oModule.CodeModule.Lines(lineNo + 1, 1) Like sUPDATE & "*" Then
            Do
                lineNo = lineNo + 1
            Loop While oModule.CodeModule.Lines(lineNo + 1, 1) Like sUPDATE & "*"
        End If
        oModule.CodeModule.InsertLines lineNo + 1, sSecondLine
        Exit Sub
    End If

    Dim nLine       As Long: nLine = Line_Declaration_First
    If nLine < 4 Then
        sSecondLine = vbLf & sFirstLine & sSecondLine
    ElseIf Not oModule.CodeModule.Lines(nLine - 2, 1) Like sUPDATE & "*" Then
        sSecondLine = vbLf & sFirstLine & sSecondLine
    End If
    If Len(oModule.CodeModule.Lines(nLine - 1, 1)) = 0 And nLine > 1 Then
        oModule.CodeModule.InsertLines nLine - 1, sSecondLine
    Else
        oModule.CodeModule.InsertLines nLine, sSecondLine & vbNewLine
    End If
End Sub

Public Function IndexInModule() As Long
    Dim Procedures  As New Collection
    Set Procedures = aModule.Init(oModule).ProceduresNames
    Dim index       As Long
    index = aCollection.Init(Procedures).IndexOf(oProcedure)
    IndexInModule = index
End Function

Public Sub MoveToTop()
    Dim Procedures  As New Collection: Set Procedures = aModule.Init(oModule).ProceduresNames
    Dim idx         As Long: idx = IndexInModule

    If idx = 1 Then Exit Sub

    Dim S           As String: S = Code_All
    Dim TargetProcedure As String: TargetProcedure = Procedures(1)

    Delete

    With oModule.CodeModule
        .InsertLines .procStartLine(TargetProcedure, oProcKind), S
    End With

    idx = Line_Declaration_First
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1
End Sub

Public Sub MoveToBottom()
    Dim Procedures  As New Collection: Set Procedures = aModule.Init(oModule).ProceduresNames
    Dim idx         As Long: idx = IndexInModule

    If idx = Procedures.Count Then Exit Sub

    Dim S           As String: S = Code_All
    Dim TargetProcedure As String: TargetProcedure = Procedures(Procedures.Count)

    Delete

    With oModule.CodeModule
        .InsertLines .procStartLine(TargetProcedure, oProcKind) + _
                .ProcCountLines(TargetProcedure, oProcKind) + 1, _
                S
    End With

    idx = Line_Declaration_First
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1
End Sub

Public Sub MoveDown()
    Dim Procedures  As New Collection: Set Procedures = aModule.Init(oModule).ProceduresNames
    Dim idx         As Long: idx = IndexInModule

    If idx = Procedures.Count Then Exit Sub

    Dim S           As String: S = Code_All
    Dim TargetProcedure As String: TargetProcedure = Procedures(idx + 1)

    Delete

    With oModule.CodeModule
        .InsertLines .procStartLine(TargetProcedure, oProcKind) + _
                .ProcCountLines(TargetProcedure, oProcKind) + 1, _
                S
    End With

    idx = Line_Declaration_First
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1

End Sub

Public Sub MoveUp()
    Dim Procedures  As New Collection: Set Procedures = aModule.Init(oModule).ProceduresNames
    Dim idx         As Long: idx = IndexInModule

    If idx = 1 Then Exit Sub

    Dim S           As String: S = Code_All
    Dim TargetProcedure As String: TargetProcedure = Procedures(idx - 1)

    Delete

    With oModule.CodeModule
        .InsertLines .procStartLine(TargetProcedure, oProcKind), S
    End With

    idx = Line_Declaration_First
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1

End Sub

Public Sub MoveToAssignedModule()
    Dim module      As VBComponent
    Set module = ParentAssigned
    If module Is Nothing Or module.Name = oModule.Name Then Exit Sub
    MoveToModule module
End Sub

Public Sub MoveToModule(module As VBComponent)
    Dim S           As String: S = Code_All
    Delete
    With module.CodeModule
        .InsertLines .CountOfLines + 1, vbNewLine & S
    End With
    Set oModule = module
End Sub

Public Sub Delete()
    TxtOverwrite Environ("TEMP") & "\" & oProcedure & ".txt", Code_All
    Debug.Print "made a backup of " & oProcedure & " in " & Environ("TEMP")
    oModule.CodeModule.DeleteLines Line_First, Line_Count
End Sub

Public Sub TestCreate()
    Dim arg         As String
    arg = ArgumentStyleFolded
    Dim Procedures  As Collection
    Set Procedures = aWorkbook.Init(oWorkbook).Procedures(True, False, False, False)

    Dim S           As String: S = vbNewLine
    If Not aCollection.Init(Procedures).Contains(, "test" & oProcedure) Then
        If KindAsString = "Sub" Then
            S = S & "Sub test" & oProcedure & "()" & vbNewLine
            S = S & "call " & arg & vbNewLine
            S = S & "End Sub"
        Else
            S = S & "Function test" & oProcedure & "()" & vbNewLine
            S = S & "Something = " & arg & vbNewLine
            S = S & "End Function"
        End If
    Else

    End If
    oModule.CodeModule.InsertLines Line_First, S
    Debug.Print "Created test" & oProcedure & " in " & oModule.Name
End Sub

Public Function Arguments() As String
    Dim str
    str = Code_Declaration_Clean
    str = Right(str, Len(str) - InStr(1, str, "("))
    str = Left(str, InStrRev(str, ")") - 1)
    If InStr(1, str, Chr(34) & "," & Chr(34)) > 0 Then
        str = VBA.Replace(str, Chr(34) & "," & Chr(34), Chr(34) & "|" & Chr(34))
    End If
    str = Split(str, ",")
    Dim i As Long
    Dim Output As String
    For i = LBound(str) To UBound(str)
        str(i) = VBA.Replace(str(i), Chr(34) & "|" & Chr(34), Chr(34) & "," & Chr(34))
        str(i) = Trim(Split(str(i), " As ")(0))
        str(i) = VBA.Replace(str(i), "Optional ", "")
        str(i) = VBA.Replace(str(i), "Astr(i) ", "")
        str(i) = VBA.Replace(str(i), "ByVal ", "")
        str(i) = VBA.Replace(str(i), "ByRef ", "")
        str(i) = VBA.Replace(str(i), "ParamArray ", "")
        str(i) = VBA.Replace(str(i), "_", "")
        Output = Output & IIf((i > LBound(str)) And (i <= UBound(str)), ", ", "") & str(i)
    Next
    Arguments = Output
End Function

Public Function ArgumentStyleFolded() As String
    If ArgumentCount = 0 Then Exit Function
    Dim str As Variant, S As String
    Dim firstPart As String, secondPart As String, Output As String
    str = Code_Declaration_Clean
    Output = oProcedure & "( _"
    Dim indentation As String
    indentation = Space(Len(Output) - 1)
    str = Right(str, Len(str) - InStr(1, str, "("))
    str = Left(str, InStrRev(str, ")") - 1)
    If InStr(1, str, Chr(34) & "," & Chr(34)) > 0 Then
        str = VBA.Replace(str, Chr(34) & "," & Chr(34), Chr(34) & "|" & Chr(34))
    End If
    str = Split(str, ",")
    Dim i           As Long
    For i = LBound(str) To UBound(str)
        str(i) = VBA.Replace(str(i), Chr(34) & "|" & Chr(34), Chr(34) & "," & Chr(34))
        
    Next
    For i = LBound(str) To UBound(str)
        S = Trim(str(i))

        S = VBA.Replace(S, "Optional ", "")
        S = VBA.Replace(S, "As ", "")
        S = VBA.Replace(S, "ByVal ", "")
        S = VBA.Replace(S, "ByRef ", "")
        S = VBA.Replace(S, "ParamArray ", "")
        S = VBA.Replace(S, "_", "")

        firstPart = Split(S, " ")(0)
        If InStr(1, S, " ") Then
            secondPart = Split(S, " ")(1)
        Else
            secondPart = "Variant"
        End If
        Output = Output & vbNewLine & indentation & firstPart & ":= " & "as" & secondPart & IIf(i <> UBound(str), ", _", ")")
    Next

    ArgumentStyleFolded = Output
End Function

Public Function ArgumentStyleClean() As String
    Dim var         As Variant
    Dim S           As String
    var = Split(ArgumentStyleFolded, vbNewLine)
    var = ArrayTrim(var)
    If UBound(var) = -1 Then
        ArgumentStyleClean = Me.Name & "()"
        Exit Function
    End If
    S = Join(var, vbNewLine)
    S = VBA.Replace(S, " _" & vbNewLine, "")
    ArgumentStyleClean = S
End Function

Sub PrintDims()
    dp Dims
End Sub

Public Property Get Dims() As Variant
    Dim S           As Variant: S = Split(Code_All, vbNewLine)
    Dim tmp: ReDim tmp(0 To 0)

    Dim i           As Long
    For i = LBound(S) To UBound(S)
        S(i) = Trim(S(i))
        If S(i) Like "*:*" Then S(i) = Split(S(i), ":")(0)
    Next

    S = ArrayFilterLike(S, "Dim *", True)
    Dim h           As String
    Dim element
    For i = LBound(S) To UBound(S)
        S(i) = Trim(StringCommentsRemove(S(i), False))
        If S(i) Like "Dim *,* As *" Then
            For Each element In Split(S(i), ", ")
                If InStr(1, element, "Dim ") = 0 Then element = "Dim " & element
                If InStr(1, element, " As ") = 0 Then element = element & " As "
                tmp(UBound(tmp)) = element
                ReDim Preserve tmp(0 To UBound(tmp) + 1)
            Next
        Else
            For Each element In Split(S(i), ", ")
                tmp(UBound(tmp)) = element
                ReDim Preserve tmp(0 To UBound(tmp) + 1)
            Next
        End If
    Next

    S = ArrayFilterLike(tmp, "Dim *", True)

    Dim var
    ReDim var(0 To 1, 0 To 0)
    For i = LBound(S) To UBound(S)
        element = S(i)
        If element Like "Dim * As *" Then

            element = Mid(element, 5)
            var(0, UBound(var, 2)) = Split(element, " As ")(0)
            var(1, UBound(var, 2)) = Split(element, " As ")(1)
            If Split(element, " As ")(1) Like "New *" Then
                var(1, UBound(var, 2)) = Split(var(1, UBound(var, 2)), " ")(1)
            End If
            If i < UBound(S) Then
                ReDim Preserve var(0 To 1, 0 To UBound(var, 2) + 1)
            End If
        ElseIf element Like "Dim *" Then

            Select Case Right(element, 1)
                Case "$": h = "String"
                Case "%": h = "Integer"
                Case "&": h = "Long"
                Case "!": h = "Single"
                Case "#": h = "Double"
                Case "@": h = "Currency"
                Case "^": h = "LongLong"
                Case Else: h = "Variant"
            End Select
            If h = "Variant" Then
                var(0, UBound(var, 2)) = Mid(element, 5)
            Else
                var(0, UBound(var, 2)) = Mid(element, 5, Len(element) - 1)
                Select Case Right(element, 1)
                    Case "$", "%", "&", "!", "#", "@", "^"
                        var(0, UBound(var, 2)) = Left(var(0, UBound(var, 2)), Len(var(0, UBound(var, 2))) - 1)
                End Select
            End If
            var(1, UBound(var, 2)) = h
            If i < UBound(S) Then
                ReDim Preserve var(0 To 1, 0 To UBound(var, 2) + 1)
            End If
        End If
    Next

    Dims = WorksheetFunction.Transpose(var)

End Property

Public Sub InjectToHeaderTop( _
        this As String, _
        SkipIfExists As Boolean)
    If SkipIfExists Then
        If InStr(1, Code_Body, this) > 0 Then Exit Sub
    End If
    oModule.CodeModule.InsertLines Line_First, _
            IIf(Len(oModule.CodeModule.Lines(Line_First, 1)) = 0, vbNewLine, "") & _
            this
End Sub

Public Sub InjectToHeaderBottom( _
        this As String, _
        SkipIfExists As Boolean)
    If SkipIfExists Then
        If InStr(1, Code_Body, this) > 0 Then Exit Sub
    End If
    oModule.CodeModule.InsertLines Line_Declaration_First, this
End Sub

Public Sub InjectToBodyTop( _
        this As String, _
        SkipIfExists As Boolean)
    If SkipIfExists Then
        If InStr(1, Code_Body, this) > 0 Then Exit Sub
    End If
    oModule.CodeModule.InsertLines Line_Body_First, this
End Sub

Public Sub InjectToBodyAfterComments( _
        this As String, _
        SkipIfExists As Boolean)
    If SkipIfExists Then
        If InStr(1, Code_Body, this) > 0 Then Exit Sub
    End If
    oModule.CodeModule.InsertLines Line_Body_FirstAfterComments, this
End Sub

Public Sub InjectToBodyBottom( _
        this As String, _
        SkipIfExists As Boolean)
    If SkipIfExists Then
        If InStr(1, Code_Body, this) > 0 Then Exit Sub
    End If
    oModule.CodeModule.InsertLines Line_Body_Last + 1, this
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 17-08-2023 08:28    Alex                (InjectDescription)

Sub InjectDescription()
    '@LastModified 2308170828

    Const vbTab2 = vbTab & vbTab
    Const vbTab4 = vbTab2 & vbTab2
    Const sFersLine As String = "'* Modified   :" & vbTab & "Date and Time" & vbTab2 & "Author" & vbTab4 & "Description" & vbCrLf

    Dim txtName     As String: txtName = AUTHOR_NAME
    If txtName = vbNullString Then txtName = Environ("UserName")

    Dim txtContacts As String: txtContacts = AUTHOR_EMAIL
    If txtContacts <> vbNullString Then txtContacts = "'* Contacts   :" & vbTab & txtContacts & vbCrLf

    Dim txtCopyright As String: txtCopyright = AUTHOR_COPYRIGHT
    If txtCopyright <> vbNullString Then txtCopyright = "'* Copyright  :" & vbTab & txtCopyright & vbCrLf

    Dim txtOther    As String: txtOther = AUTHOR_OTHERTEXT
    If txtOther <> vbNullString Then txtOther = "'* Note       :" & vbTab & txtOther & vbCrLf

    Dim txtMedia    As String: txtMedia = "'* " & vbLf & AUTHOR_MEDIA
    Dim CurentCodePane As CodePane: Set CurentCodePane = oModule.CodeModule.CodePane
    Dim i           As Byte
    Dim ProcKind    As VBIDE.vbext_ProcKind
    Dim sProc       As String: sProc = ActiveProcedure
    Dim nLine       As Long
    nLine = Line_Declaration_First
    Dim sTemp       As String
    Dim sTime       As String
    Dim sType       As String
    Dim sProcDeclartion As String
    Dim sProcArguments As String

    Dim sUser       As String

    '* @TODO Created: 01-02-2023 08:20 Author: Anastasiou Alex
    '* @TODO if name is too long have to recalculate tab length

    sUser = "Alex"
    If sUser = vbNullString Then sUser = Environ("UserName")

    On Error Resume Next

    With CurentCodePane
        sTemp = VBA.Replace(String(90, "*"), "**", "* ")
        sTime = Format(Now, "dd-mm-yyyy hh:nn")

        For i = 0 To 4
            ProcKind = i
            sProcDeclartion = Code_Declaration_Clean
            If sProcDeclartion <> vbNullString Then Exit For
        Next
        sProcArguments = AddStringParameterFromProcedureHeader(sProcDeclartion)
        sType = TypeProcedureComment(sProcDeclartion)
        sTemp = _
                "'" & sTemp & vbCrLf & _
                "'" & sType & vbTab & sProc & vbCrLf & _
                "'* Author     :" & vbTab & txtName & vbCrLf & _
                txtContacts & _
                txtCopyright & _
                txtOther & _
                "'* Purpose    :" & vbTab & vbCrLf & _
                txtMedia & _
                sFersLine & _
                "'* Created    :" & vbTab & sTime & vbTab & sUser & vbCrLf & _
                sProcArguments & _
                "'" & sTemp
        InjectToHeaderTop sTemp, False

    End With
End Sub

Private Function AddStringParameterFromProcedureHeader(ByVal sPocDeclartion As String) As String
    '@BlogPosted
    Const vbTab2 = vbTab & vbTab
    Const vbTab4 = vbTab2 & vbTab2
    Dim sDeclaration As String
    sDeclaration = Right$(sPocDeclartion, Len(sPocDeclartion) - InStr(1, sPocDeclartion, "("))
    sDeclaration = Left$(sDeclaration, InStr(1, sDeclaration, ")") - 1)
    If sDeclaration = vbNullString Then Exit Function
    Dim arStr()     As String
    arStr = Split(sDeclaration, ",")
    Dim iMaxLen     As Byte
    iMaxLen = 0
    Dim iTempLen    As Byte
    Dim sTemp       As String
    Dim i           As Byte
    For i = 0 To UBound(arStr)
        iTempLen = Len(Trim$(arStr(i)))
        If iMaxLen < iTempLen Then iMaxLen = iTempLen
    Next i
    Dim numOfSpaces As Long
    numOfSpaces = iMaxLen - Len(Trim$("'* Argument(s):"))
    If numOfSpaces < 0 Then numOfSpaces = 0
    sDeclaration = "'*" & vbLf & "'* Argument(s):" & Space(numOfSpaces) & vbTab2 & "Description" & vbCrLf & "'*" & vbCrLf
    For i = 0 To UBound(arStr)
        sTemp = "'* " & Trim$(arStr(i)) & Space(iMaxLen - Len(Trim$(arStr(i)))) & " :"
        sDeclaration = sDeclaration & sTemp & vbCrLf
    Next i
    AddStringParameterFromProcedureHeader = sDeclaration & "'* " & vbCrLf
End Function

Private Function TypeProcedureComment(ByVal StrDeclarationProcedure As String) As String
    '@BlogPosted
    If StrDeclarationProcedure Like "*Sub*" Then
        TypeProcedureComment = "* Sub        :"
    ElseIf StrDeclarationProcedure Like "*Function*" Then
        TypeProcedureComment = "* Function   :"
    ElseIf StrDeclarationProcedure Like "*Property Set*" Then
        TypeProcedureComment = "* Prop Set   :"
    ElseIf StrDeclarationProcedure Like "*Property Get*" Then
        TypeProcedureComment = "* Prop Get   :"
    ElseIf StrDeclarationProcedure Like "*Property Let*" Then
        TypeProcedureComment = "* Prop Let   :"
    Else
        TypeProcedureComment = "* Un Type    :"
    End If
End Function

Sub InjectTimer()
    Dim ProcedureText As String
    ProcedureText = Code_All
    If Contains("StartTimer", True, True, False) Then Exit Sub
    InjectToBodyAfterComments "StartTimer " & """" & oProcedure & """", True
    Sleep 200
    InjectToBodyBottom "EndTimer", True
End Sub

Function SuggestedScope() As String

    Dim result      As String

    Select Case CallerModules.Count
        Case 0
            SuggestedScope = "Unspecified"
            result = result & vbNewLine & "Procedure " & Name & " has unclarified scope in " & oWorkbook.Name
            result = result & vbNewLine & "It is called 0 times, so it may be Unused, Unassigned, or Called from elsewhere"
            Exit Function
        Case 1
            SuggestedScope = "Private"
            result = result & vbNewLine & "Suggested scope for procedure " & Name & " is Private"
        Case Is > 1
            SuggestedScope = "Public"
            result = result & vbNewLine & "Suggested scope for procedure " & Name & " is Public"
    End Select

    result = result & vbNewLine & "because it is used in " & CallerModules.Count & " module(s) of " & oWorkbook.Name

End Function

Sub ScopeSuggested()
    Dim idx         As Long: idx = Line_Declaration_First
    Dim S           As String: S = oModule.CodeModule.Lines(idx, 1)
    Dim Suggestion  As String: Suggestion = SuggestedScope

    If Suggestion = "Private" Or Suggestion = "Public" Then
        Select Case True
            Case S Like "*Public*" & Name & "*"
                ScopePublic
            Case S Like "*Private*" & Name & "*"
                ScopePrivate
            Case Else
                Dim sKind As String: sKind = KindAsString
                oModule.CodeModule.ReplaceLine idx, Trim(VBA.Replace(S, sKind, " " & Suggestion & " " & sKind, , 1, vbTextCompare))
        End Select
    End If
End Sub

Sub ScopePrivate()
    Dim idx         As Long: idx = Line_Declaration_First
    Dim S           As String: S = Code_Declaration_Clean
    Dim Suggestion  As String: Suggestion = "Private"
    Select Case True
        Case S Like "*Public*" & Name & "*"
            Debug.Print Name
            oModule.CodeModule.ReplaceLine idx, VBA.Replace(S, "Public", "Private", , 1, vbTextCompare)
        Case S Like "*Private*" & Name & "*"

        Case Else
            Dim sKind As String: sKind = KindAsString
            Select Case sKind
                Case sKind Like "*Property*"
                    Debug.Print Name
                    oModule.CodeModule.ReplaceLine idx, VBA.Replace(S, "Property", Suggestion & " " & "Property", , 1, vbTextCompare)
                Case Else
                    Debug.Print Name
                    oModule.CodeModule.ReplaceLine idx, VBA.Replace(S, sKind, Suggestion & " " & sKind, , 1, vbTextCompare)
            End Select
    End Select
End Sub

Sub ScopePublic()
    Dim idx         As Long: idx = Line_Declaration_First
    Dim S           As String: S = Code_Declaration_Clean
    Dim Suggestion  As String: Suggestion = "Public"
    Select Case True
        Case S Like "*Public*" & Name & "*"

        Case S Like "*Private*" & Name & "*"
            Debug.Print Name
            oModule.CodeModule.ReplaceLine idx, VBA.Replace(S, "Private", "Public", , 1, vbTextCompare)
        Case Else
            Dim sKind As String: sKind = KindAsString
            Debug.Print Name
            oModule.CodeModule.ReplaceLine idx, Trim(VBA.Replace(S, sKind, " " & Suggestion & " " & sKind, , 1, vbTextCompare))
    End Select
End Sub

Function CallerModules() As Collection
    Dim myName      As String: myName = Name
    Dim coll        As New Collection
    Dim module      As aModule
    Dim matchCollection As New Collection
    For Each module In aModules.Init(oWorkbook).items
        If module.Contains(myName, True, True, False) Then
            On Error Resume Next
            coll.Add module.Name, module.Name
            On Error GoTo 0
        End If
    Next
    Set CallerModules = coll
End Function

Function CallerModulesToString() As String
    Dim element
    Dim result      As String
    For Each element In CallerModules
        result = result & IIf(result <> "", vbNewLine, "") & element
    Next
    CallerModulesToString = result
End Function

Function Callers() As Collection
    Dim myName      As String: myName = Name
    Dim result      As New Collection
    Dim module      As aModule
    Dim Procedure   As aProcedure
    For Each module In CallerModules
        For Each Procedure In module.Procedures
            If Procedure.Name <> myName Then
                If Procedure.Contains(myName, True, True, False) Then
                    On Error Resume Next
                    result.Add Procedure.Name, Procedure.Name
                    Exit For
                    On Error GoTo 0
                End If
            End If
        Next
    Next
    Set Callers = result
End Function

Function CallersToString() As String
    Dim element
    Dim result      As String
    For Each element In Callers
        result = result & IIf(result <> "", vbNewLine, "") & element
    Next
    CallersToString = result
End Function

Public Sub BringProcedureHere(Procedure As String)
    Dim ap          As New aProcedure
    ap.Init , , Procedure
    Dim S           As String
    S = ap.Code_All
    If InStr(1, S, "'@AssignedModule") = 0 Then
        ap.PropertySet Parent_Assigned, ap.Parent.Name
        S = ap.Code_All
    End If
    ap.Delete
    oModule.CodeModule.InsertLines Line_Last + 1, S
End Sub

Sub BringLinkedProceduresHere()
    Dim el
    Dim ap          As aProcedure
    For Each el In LinkedProceduresDeep
        BringProcedureHere CStr(el)
    Next
End Sub



Sub AddToLinkedTable()

    Dim aw          As aWorkbook
    Set aw = aWorkbook.Init(oWorkbook)
    Dim cell        As Range
    Dim TargetWorksheet As Worksheet
    On Error Resume Next
    Set TargetWorksheet = ThisWorkbook.Sheets("Linked_Table_" & aw.NameClean)
    If TargetWorksheet Is Nothing Then Toast "Worksheet Linked_Table_" & aw.NameClean & " does not exist. Create first. Terminating": Exit Sub
    Set cell = TargetWorksheet.Columns(4).Find(aProcedure, LookAt:=xlWhole)
    If Not cell Is Nothing Then Debug.Print "Procedure " & aProcedure & " was already exporterd. Terminating.": Exit Sub
    On Error GoTo 0

    Dim var
    ReDim var(1 To 8)
    var(1) = oWorkbook.Name
    var(2) = aModule.Init(oModule).TypeToString
    var(3) = oModule.Name
    var(4) = oProcedure
    var(5) = aCollection.Init(LinkedProcedures).ToString(vbNewLine)
    var(6) = aCollection.Init(LinkedClasses).ToString(vbNewLine)
    var(7) = aCollection.Init(LinkedUserforms).ToString(vbNewLine)
    var(8) = aCollection.Init(LinkedDeclarations).ToString(vbNewLine)

    Dim Lr          As Long
    Lr = getLastRow(TargetWorksheet) + 1
    TargetWorksheet.Range("A" & Lr & ":H" & Lr).Value = var
End Sub



Public Sub Enable_DebugPrint()
    Dim N           As Long
    Dim S           As String
    With oModule.CodeModule
        For N = Line_Last To Line_First Step -1
            S = .Lines(N, 1)
            If Left(Trim(S), 6) = "'Debug" Then
                aCodeModule.Init(oModule).UncommentTargetLine N
                '                s = VBA.Replace(s, "'", "", , 1)
                '                .ReplaceLine N, s
            End If
        Next N
    End With
End Sub
Public Sub Disable_DebugPrint()
    Dim N           As Long
    Dim S           As String
    With oModule.CodeModule
        For N = Line_Last To Line_First Step -1
            S = .Lines(N, 1)
            If Left(Trim(S), 5) = "Debug" Then
                aCodeModule.Init(oModule).CommentTargetLine N
                '                .ReplaceLine N, "'" & s
            End If
        Next N
    End With
End Sub


Public Sub Enable_Stop()
    Dim N           As Long
    Dim S           As String
    Dim Keyword     As String
    Keyword = "Stop"
    With oModule.CodeModule
        For N = Line_Last To Line_First Step -1
            S = .Lines(N, 1)
            If InStrExact(1, S, Keyword) > 0 Then
                S = VBA.Replace(S, "'", "", , 1)
                .ReplaceLine N, S
            End If
        Next N
    End With
End Sub
Public Sub Disable_Stop()
    Dim N           As Long
    Dim S           As String
    Dim Keyword     As String: Keyword = "Stop"
    With oModule.CodeModule
        For N = Line_Last To Line_First Step -1
            S = .Lines(N, 1)
            If InStrExact(1, S, Keyword) > 0 Then
                .ReplaceLine N, "'" & S
            End If
        Next N
    End With
End Sub


Public Sub Comments_ReplaceQuoteWithRem()
    Dim N           As Long
    Dim S           As String
    With oModule.CodeModule
        For N = Line_Last To Line_First Step -1
            S = .Lines(N, 1)
            If Left(Trim(S), 1) = "'" Then
                .ReplaceLine N, VBA.Replace(S, "'", "Rem ", , 1)
            End If
        Next N
    End With
End Sub


Public Sub UpdatableVariable_Add(this As String)
    Dim i           As Long: i = LineLike("*UpdatableVariable = *", Body_Code)
    If i = 0 Then Exit Sub
    Dim q           As String: q = Chr(34)
    Dim line        As String: line = oModule.CodeModule.Lines(i, 1)
    If UCase(Trim(line)) Like UCase("UpdatableVariable = *") And InStr(1, line, this, vbTextCompare) = 0 Then
        oModule.CodeModule.ReplaceLine i, VBA.Replace(VBA.Replace(line, q, q & this & ",", , 1, vbTextCompare), "," & q, q)
        line = oModule.CodeModule.Lines(i, 1)
        If InStr(1, line, q & ",") > 0 Then oModule.CodeModule.ReplaceLine i, VBA.Replace(line, q & ",", q)
    End If
End Sub

Public Sub UpdatableVariable_Remove(this As String)
    Dim i           As Long: i = LineLike("*UpdatableVariable = *", Body_Code)
    If i = 0 Then Exit Sub
    Dim q           As String: q = Chr(34)
    Dim result      As String
    Dim line        As String: line = oModule.CodeModule.Lines(i, 1)
    If Trim(line) Like "UpdatableVariable = *" Then
        If InStr(1, line, q & this & q, vbTextCompare) > 0 Then
            result = VBA.Replace(line, q & this & q, "")
            result = VBA.Replace(result, ",,", ",")
            result = VBA.Replace(result, q & ",", q)
            oModule.CodeModule.ReplaceLine i, result
        End If
    End If
End Sub

Sub CreateCaller(rng As Range)
    Dim shp         As Shape
    Set shp = ActiveSheet.Shapes.AddShape _
            (msoShapeRoundedRectangle, 1, 1, 500, 10)
    With shp.ThreeD
        .BevelTopType = msoBevelCircle
        .BevelTopInset = 6
        .BevelTopDepth = 6
    End With
    With shp.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
        .Solid
    End With
    With shp.line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With

    With shp
        .OnAction = "'" & oWorkbook.Name & "'!" & oProcedure
        .Name = "Run_" & oProcedure
        .TextFrame2.TextRange.TEXT = oProcedure
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame2.WordWrap = msoFalse
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        .Left = Selection.Left
        .Top = Selection.Top
    End With
End Sub

Function LargestLineLength()
    LargestLineLength = LargestLength(Split(Code_All, vbNewLine))
End Function

Sub ConvertBlankLinesToDividers(Optional Character As String = "~")
    Dim l           As Long: l = LargestLineLength
    Dim i           As Long
    For i = Line_Last To Line_Declaration_First Step -1
        If Len(Trim(oModule.CodeModule.Lines(i, 1))) = 0 Then
            oModule.CodeModule.ReplaceLine i, "'" & String(l - 1, Character)
        End If
    Next
End Sub
