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

Private oProcedure As String
Private oWorkbook As Workbook
Private oModule As VBComponent
Private oProcKind As vbext_ProcKind

Private coll_LinkedProcedures As Collection
Private coll_LinkedProceduresDeep As Collection

Public Enum Code_Part
    All_Code
    Declaration_Code
    Declaration_Clean_Code
    Body_Code
    Header_Code
End Enum

Public Enum Line_Index
    Procedure_First
    Procedure_Last
    Procedure_Count
    Body_First
    Body_First_AfterComments
    Body_Last
    Body_Count
    Declaration_First
    Declaration_Last
    Declaration_Count
    Header_First
    Header_Last
    Header_Count
End Enum

Public Enum Property_Type
    Modified
    Ignore
    Parent_Assigned
End Enum

Private Sub Class_Initialize()
'    ProjetFoldersCreate
'@AssignedModule aProcedure
'@INCLUDE PROCEDURE ProjetFoldersCreate
'@INCLUDE CLASS aProcedure
End Sub

Public Function Init( _
                    Optional TargetWorkbook As Workbook, _
                    Optional Module As VBComponent, _
                    Optional Procedure As String, _
                    Optional ProcKind As vbext_ProcKind) As aProcedure
'@AssignedModule aProcedure
'@INCLUDE PROCEDURE AssignCPSvariables
'@INCLUDE CLASS aProcedure
    If Not AssignCPSvariables(TargetWorkbook, Module, Procedure) Then End
    oProcedure = Procedure
    Set oWorkbook = TargetWorkbook
    Set oModule = Module
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
    Dim Module As VBComponent: Set Module = ActiveModule
    Dim Procedure As String
    Dim i As Long
    Dim ProcKind As vbext_ProcKind
    Dim lineNum As Long: lineNum = aCodeModule.Init(Module).RowFirst
    For i = 0 To 4
        ProcKind = i
        Procedure = Module.CodeModule.ProcOfLine(lineNum, ProcKind)
        If Procedure <> vbNullString Then
            Init ActiveCodepaneWorkbook, Module, Procedure, ProcKind
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
    Dim Module As VBComponent
    Dim s As String
    Dim Procedure As Variant
    Dim AP As aProcedure
    For Each Procedure In LinkedProceduresDeep
        Set AP = New aProcedure
        AP.Init oWorkbook, , CStr(Procedure)
        s = AP.Code_All
        If Not ProcedureExists(TargetWorkbook, CStr(Procedure)) Then
            Set Module = ModuleAddOrSet(TargetWorkbook, "vbArcImports", vbext_ct_StdModule)
            Module.CodeModule.AddFromString s
        Else
            Set Module = ModuleOfProcedure(TargetWorkbook, CStr(Procedure))
            If Overwrite = True Then ProcedureReplace Module, CStr(Procedure), s
        End If
    Next
End Sub

Public Function Contains( _
                        This As String, _
                        WholeWord As Boolean, _
                        MatchCase As Boolean, _
                        PatternSearch As Boolean) As String
'@AssignedModule aProcedure
'@INCLUDE CLASS aProcedure

    Contains = oModule.CodeModule.Find( _
                                        This, _
                                        1, _
                                        LineIndex(Procedure_First), _
                                        LineIndex(Procedure_Last), _
                                        10000, _
                                        WholeWord, _
                                        MatchCase, _
                                        PatternSearch)
End Function

Public Function Contains_InHeader( _
                                This As String, _
                                WholeWord As Boolean, _
                                MatchCase As Boolean, _
                                PatternSearch As Boolean) As String
'@AssignedModule aProcedure
'@INCLUDE CLASS aProcedure

    Contains_InHeader = oModule.CodeModule.Find( _
                                        This, _
                                        1, _
                                        LineIndex(Header_First), _
                                        LineIndex(Header_Last), _
                                        10000, _
                                        WholeWord, _
                                        MatchCase, _
                                        PatternSearch)
End Function

Public Function Contains_InBody( _
                                This As String, _
                                WholeWord As Boolean, _
                                MatchCase As Boolean, _
                                PatternSearch As Boolean) As String
'@AssignedModule aProcedure
'@INCLUDE CLASS aProcedure

    Contains_InBody = oModule.CodeModule.Find( _
                                        This, _
                                        1, _
                                        LineIndex(Body_First), _
                                        LineIndex(Body_Last), _
                                        10000, _
                                        WholeWord, _
                                        MatchCase, _
                                        PatternSearch)
End Function

Public Function Scope() As String
'@AssignedModule aProcedure
'@INCLUDE CLASS aProcedure
    Dim sTitle As String
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
        LineStart = LineIndex(Procedure_First)
        LineEnd = LineIndex(Procedure_Last)
    Case Code_Part.Body_Code
        LineStart = LineIndex(Body_First)
        LineEnd = LineIndex(Body_Last)
    Case Code_Part.Header_Code
        LineStart = LineIndex(Header_First)
        LineEnd = LineIndex(Header_Last)
    End Select
End Sub

Public Sub CommentsToOwnLine()
'@AssignedModule aProcedure
'@INCLUDE PROCEDURE CommentsMoveToOwnLine
'@INCLUDE CLASS aProcedure
    Dim s As String:    s = CommentsMoveToOwnLine(Code_All)
    Dim i As Long:      i = LineIndex(Procedure_First)
    Delete
    oModule.CodeModule.InsertLines i, s
End Sub

Public Function Property(PropertyType As Property_Type) As Variant
'@AssignedModule aProcedure
'@INCLUDE CLASS aProcedure
    Select Case PropertyType
    
        Case Property_Type.Modified:           Property = PropertyGet(Modified)
        Case Property_Type.Parent_Assigned:    Property = PropertyGet(Parent_Assigned)
        
    End Select
End Function

Public Property Get ArgumentCount() As Long
    Dim s As String: s = Code_Declaration_Clean
    If s Like "*" & oProcedure & "()*" Then Exit Property
    ArgumentCount = UBound(Split(s, ",")) + 1
End Property

Public Property Get Name() As String
    Name = oProcedure
End Property

Public Property Get ReturnType() As String
    ReturnType = "Invalid"
    If KindAsString <> "Function" Then Exit Property
    Dim Title As String: Title = Code_Declaration_Clean
    If Title Like "*) As *" Then
        ReturnType = Split(Title, ") As ")(1)
    Else
        ReturnType = "Unspecified"
    End If
End Property

Public Property Get Parent() As VBComponent
    Set Parent = oModule
End Property

Public Property Get ParentAssigned() As VBComponent
    Dim moduleName As String: moduleName = Property(Parent_Assigned)
    If moduleName = "" Then Exit Property
    Set ParentAssigned = ModuleAddOrSet(oWorkbook, moduleName, vbext_ct_StdModule)
'End Property

'Public Property Set ParentAssigned(Module As VBComponent)
'
'    Set oModule = Module
End Property

Public Property Get KindAsLong() As Long
    Dim s As String: s = Code_Declaration_Clean
    Select Case True
        Case InStr(1, s, "Get " & oProcedure) > 0:  KindAsLong = vbext_pk_Get '3
        Case InStr(1, s, "Let " & oProcedure) > 0:  KindAsLong = vbext_pk_Let '1
        Case InStr(1, s, "Set " & oProcedure) > 0:  KindAsLong = vbext_pk_Set '2
        Case Else:                                  KindAsLong = vbext_pk_Proc '0
    End Select
End Property

Public Property Get KindAsString() As String
    Dim s As String: s = Code_Declaration_Clean
    Select Case True
        Case InStr(1, s, "Get " & oProcedure) > 0:  KindAsString = "Get"
        Case InStr(1, s, "Let " & oProcedure) > 0:  KindAsString = "Let"
        Case InStr(1, s, "Set " & oProcedure) > 0:  KindAsString = "Set"
        Case Else
            If InStr(1, s, "Function " & oProcedure) > 0 Then
                KindAsString = "Function"
            ElseIf InStr(1, s, "Sub " & oProcedure) > 0 Then
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

Public Function LineIndex(idx As Line_Index) As Long
    Select Case idx
        Case Line_Index.Procedure_First:            LineIndex = LineFirst
        Case Line_Index.Procedure_Last:             LineIndex = LineLast
        Case Line_Index.Procedure_Count:            LineIndex = LineCount
        Case Line_Index.Declaration_First:          LineIndex = DeclarationLineFirst
        Case Line_Index.Declaration_Last:           LineIndex = DeclarationLineLast
        Case Line_Index.Declaration_Count:          LineIndex = DeclarationLineCount
        Case Line_Index.Body_First:                 LineIndex = BodyLineFirst
        Case Line_Index.Body_First_AfterComments:   LineIndex = BodyLineFirstAfterComments
        Case Line_Index.Body_Last:                  LineIndex = BodyLineLast
        Case Line_Index.Body_Count:                 LineIndex = BodyLineCount
        Case Line_Index.Header_First:               LineIndex = HeaderLineFirst
        Case Line_Index.Header_Last:                LineIndex = HeaderLineLast
        Case Line_Index.Header_Count:               LineIndex = HeaderLineCount
                
    End Select
End Function

Public Function LineLike(This As String, CodePart As Code_Part) As Long
    Dim LineStart As Long, LineEnd As Long
    AssignLineStartEnd CodePart, LineStart, LineEnd
    Dim targetLine As String
    Dim i As Long
    For i = LineStart To LineEnd
        targetLine = Trim(oModule.CodeModule.Lines(i, 1))
        If UCase(targetLine) Like UCase(This) Then
            LineLike = i
            Exit Function
        End If
    Next
End Function

Public Sub EnsureBlankLineBeforeProcedure()
    Dim Line As Long: Line = DeclarationLineFirst
    If Line = 1 Then oModule.CodeModule.InsertLines 1, "": Exit Sub
    Dim check As String
    On Error Resume Next
    check = oModule.CodeModule.ProcOfLine(Line - 1, oProcKind)
    On Error GoTo 0
    If check <> Name Then oModule.CodeModule.InsertLines Line, ""
End Sub

Public Function HeaderLineFirst() As Long
    EnsureBlankLineBeforeProcedure
    HeaderLineFirst = LineFirst
End Function

Public Function HeaderLineLast() As Long
    EnsureBlankLineBeforeProcedure
    HeaderLineLast = DeclarationLineFirst - 1
End Function

Public Function HeaderLineCount()
    HeaderLineCount = HeaderLineLast - HeaderLineFirst + 1
End Function

Public Function LineFirst() As Long
    '@todo Fix
    LineFirst = oModule.CodeModule.procStartLine(oProcedure, oProcKind) 'findProcedureKind(omodule,oprocedure)
    If LineFirst = 0 Then Debug.Print "error in firstline (wrong procKind)": End
End Function

Public Function LineCount() As Long
    LineCount = oModule.CodeModule.ProcCountLines(oProcedure, oProcKind)
End Function

Public Function LineLast() As Long
    LineLast = LineFirst + LineCount - 1
End Function

Public Function BodyLineFirst() As Long
    BodyLineFirst = DeclarationLineFirst + DeclarationLineCount
End Function

Public Function BodyLineFirstAfterComments() As Long
    Dim i As Long
    Dim s As String
    For i = LineIndex(Body_First) To LineIndex(Body_Last)
        s = Trim(oModule.CodeModule.Lines(i, 1))
        If s = vbNullString Then
            Exit For
        ElseIf Left(s, 1) = "'" Then
        ElseIf Left(s, 3) = "Rem" Then
        ElseIf Right(Trim(oModule.CodeModule.Lines(i - 1, 1)), 1) = "_" Then
        ElseIf Right(s, 1) = "_" Then
        Else
            Exit For
        End If
    Next
    BodyLineFirstAfterComments = i
End Function

Public Function BodyLineLast() As Long
    BodyLineLast = LineLast - 1
End Function

Public Function BodyLineCount() As Long
    BodyLineCount = BodyLineLast - BodyLineFirst + 1
End Function

Public Function DeclarationLineFirst() As Long

    DeclarationLineFirst = oModule.CodeModule.ProcBodyLine(oProcedure, oProcKind)
End Function


Public Function DeclarationLineLast() As Long
    DeclarationLineLast = DeclarationLineFirst + DeclarationLineCount - 1
End Function

Public Function DeclarationLineCount() As Long
    Dim targetLine As Long: targetLine = DeclarationLineFirst
    Dim counter As Long: counter = 1
    Dim s As String: s = oModule.CodeModule.Lines(targetLine, counter)
    Do While Right(s, 1) = "_"
        counter = counter + 1
        s = oModule.CodeModule.Lines(targetLine, counter)
    Loop
    DeclarationLineCount = counter
End Function

Public Function Code_All()
    Code_All = oModule.CodeModule.Lines(LineFirst, LineCount)
End Function
Public Function Code_Body()
    Code_Body = oModule.CodeModule.Lines(BodyLineFirst, BodyLineCount)
End Function
Public Function Code_Header()
    Code_Header = oModule.CodeModule.Lines(HeaderLineFirst, HeaderLineCount)
End Function
Public Function Code_Declaration()
    Code_Declaration = oModule.CodeModule.Lines(DeclarationLineFirst, DeclarationLineCount)
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
    Dim Procedures As Collection
    Set Procedures = aWorkbook.Init(oWorkbook).ProceduresNames(True, False, False, False)
    Dim s As String: s = Code_All
    Dim coll As New Collection
    Dim Procedure As Variant
    For Each Procedure In Procedures
        If UCase(CStr(Procedure)) <> UCase(CStr(oProcedure)) Then
            If RegexTest(s, "\W" & CStr(Procedure) & "[.(\W]") = True Then
                coll.Add Procedure, CStr(Procedure)
            End If
        End If
    Next
    Set LinkedProcedures = coll
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 20-07-2023 10:45    Alex                code fix

Public Function LinkedProceduresDeep() As Collection
'@LastModified 2307201045
    Dim AllProcedures As Collection:       Set AllProcedures = aWorkbook.Init(oWorkbook).ProceduresNames(True, False, False, False)
    Dim Processed As Collection:           Set Processed = New Collection
    Dim CalledProcedures As Collection:    Set CalledProcedures = New Collection

    Dim Procedure As Variant
    Dim Module As VBComponent

    Processed.Add CStr(oProcedure), CStr(oProcedure)
    CalledProcedures.Add CStr(oProcedure), CStr(oProcedure)
    On Error Resume Next
    For Each Procedure In LinkedProcedures
    CalledProcedures.Add CStr(Procedure), CStr(Procedure)
    Next
    On Error GoTo 0
    
    Dim AP As aProcedure
    Dim CalledProceduresCount As Long
        CalledProceduresCount = CalledProcedures.count
    Dim element
repeat:
    For Each element In CalledProcedures
        If Not aCollection.Init(Processed).Contains(, CStr(element)) Then
            On Error Resume Next
            Set AP = New aProcedure
            For Each Procedure In AP.Init(oWorkbook, , CStr(element)).LinkedProcedures
            CalledProcedures.Add CStr(Procedure), CStr(Procedure)
            Next
            On Error GoTo 0
            Processed.Add CStr(element), CStr(element)
        End If
    Next
    If CalledProcedures.count > CalledProceduresCount Then
        CalledProceduresCount = CalledProcedures.count
        GoTo repeat
    End If

    Set LinkedProceduresDeep = aCollection.Init(CalledProcedures).Sort.Items
End Function

Sub UnfoldDeclaration()
    Dim Result As String: Result = Code_Declaration_Clean & vbLf & Code_Body & vbLf & oModule.CodeModule.Lines(LineIndex(Procedure_Last), 1)
    Replace Result
End Sub
Sub FoldDeclaration()
    Dim startLine As Long
    Dim LastLine As Long
    Dim FirstColumn As Long
    Dim LastColumn As Long

    If ArgumentCount = 0 Then Exit Sub
    Dim cp As aCodeModule
    Set cp = aCodeModule.Init(oModule)
    
    startLine = LineIndex(Declaration_First)
    FirstColumn = InStr(1, oModule.CodeModule.Lines(startLine, 1), "(")
    LastLine = LineIndex(Declaration_Last)
'    Do While InStr(1, oModule.CodeModule.Lines(LastLine, 1), ")") = 0
'        LastLine = LastLine - 1
'    Loop
    LastColumn = InStr(1, oModule.CodeModule.Lines(LastLine, 1), ")") + 1
    If LastColumn = 0 Then Exit Sub
    
    cp.SetSelection startLine, FirstColumn, LastLine, LastColumn
    cp.FoldLine
    cp.SetSelection startLine, FirstColumn, LineIndex(Declaration_Last), 10000
End Sub

Public Sub Export()
    Dim Procedure
    For Each Procedure In collLinkedProceduresDeep
        LinkedExport CStr(Procedure)
    Next
End Sub

Private Sub LinkedExport(Procedure As String)
    Dim Module As VBComponent
    Set Module = ModuleOfProcedure(oWorkbook, Procedure)
    Dim c As New aProcedure
    c.Init oWorkbook, Module, Procedure
    Dim proclastmod     As String:        proclastmod = c.Property(Modified)
    If proclastmod = vbNullString Then
        c.InjectLinkedLists
        c.PropertySet Modified, Format(Now, "yymmddhhnn")
        proclastmod = c.Property(Modified)
    End If
    Dim s               As String:      s = c.Code_All
    Dim FileFullName    As String:      FileFullName = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"
    
    If FileExists(FileFullName) Then
        Dim filelastmod
            filelastmod = StringLastModified(TxtRead(FileFullName))
        If proclastmod > filelastmod Then
            Debug.Print "OVERWROTE " & Procedure
            TxtOverwrite FileFullName, s
        Else
            Debug.Print "Skipping " & Procedure & " because:"
            Debug.Print vbTab & "StringLastModified(TxtRead(FileFullName)) >= Property(Modified)"
        End If
    Else
        Debug.Print "NEW " & Procedure
        TxtOverwrite FileFullName, s
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

    Dim codeName As String
    Dim codeText As String
    Dim cell As Range
    On Error Resume Next
    Set cell = TargetWorksheet.Columns(3).Find(DeclarationName, LookAt:=xlWhole)
    On Error GoTo 0
    If cell Is Nothing Then Exit Sub

    codeName = DeclarationName
    codeText = cell.offset(0, 1).text
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
    Dim s As String
    Dim nIndent As Long
    Dim i As Long
    For i = LineIndex(Procedure_First) To LineIndex(Procedure_Last)
        s = LTrim$(oModule.CodeModule.Lines(i, 1))
        If Trim(s) <> "" Then
            If IsBlockEnd(s) Then nIndent = nIndent - 1
            If nIndent < 0 Then nIndent = 0
            s = Space$(nIndent * 4) & s
            oModule.CodeModule.ReplaceLine i, s
            If IsBlockStart(LTrim$(s)) Then nIndent = nIndent + 1
        End If
    Next
End Sub

Public Sub NumbersAdd()
    Dim counter As Long: counter = 1
    Dim LineStart As Long, LineEnd As Long
    AssignLineStartEnd Body_Code, LineStart, LineEnd
    Dim s As String
    Dim i As Long
    For i = LineStart To LineEnd
        s = oModule.CodeModule.Lines(i, 1)
        If IsLineNumberAble(s) _
        And Right(Trim(oModule.CodeModule.Lines(i - 1, 1)), 1) <> "_" Then
            oModule.CodeModule.ReplaceLine i, counter & ":" & s
            counter = counter + 1
        End If
    Next i
End Sub

Public Sub NumbersRemove()
    Dim counter As Long: counter = 1
    Dim LineStart As Long, LineEnd As Long
    AssignLineStartEnd Body_Code, LineStart, LineEnd
    Dim s As String
    Dim i As Long
    For i = LineStart To LineEnd
        s = oModule.CodeModule.Lines(i, 1)
        If IsNumeric(Left(Trim(s), 1)) Then
            oModule.CodeModule.ReplaceLine i, _
                Mid(s, InStr(s, ":") + 1)
                'Space(InStr(s, ":"))
        End If
    Next i
End Sub

Public Sub RemoveEmptyLines()
    Dim LineStart As Long, LineEnd As Long
    AssignLineStartEnd Body_Code, LineStart, LineEnd
    Dim i As Long
    For i = LineEnd To LineStart Step -1
        If Len(Trim(oModule.CodeModule.Lines(i, 1))) = 0 Then
            oModule.CodeModule.DeleteLines i, 1
        End If
    Next
End Sub

Public Sub CommentsRemove(RemoveRem As Boolean, CodePart As Code_Part)
    Replace StringCommentsRemove(Code_All, RemoveRem)

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

Public Sub RemoveLinesLike(This As String)
    Dim s As String
    Dim i As Long
    For i = LineIndex(Body_Last) To LineIndex(Body_First) Step -1
        s = Trim(oModule.CodeModule.Lines(i, 1))
        If s Like This Then oModule.CodeModule.DeleteLines i
    Next
End Sub

Public Sub InjectLinkedClasses()
    Dim ListOfImports As String
    Dim s As String: s = Code_All
    Dim element As Variant
    For Each element In LinkedClasses
        If InStr(1, s, "@INCLUDE CLASS " & element) = 0 _
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
            LineIndex(Body_First_AfterComments), ListOfImports
    End If
End Sub

Public Sub InjectLinkedProcedures()
    Dim Procedures As Collection
    Set Procedures = collLinkedProcedures
    Dim ListOfImports As String
    Dim s As String: s = Code_All
    Dim Procedure As Variant
    For Each Procedure In Procedures
        If InStr(1, s, "@INCLUDE PROCEDURE " & Procedure) = 0 And InStr(1, ListOfImports, "@INCLUDE PROCEDURE " & Procedure) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE PROCEDURE " & Procedure
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE PROCEDURE " & Procedure
            End If
        End If
    Next
    If ListOfImports <> "" Then
        oModule.CodeModule.InsertLines LineIndex(Body_First_AfterComments), ListOfImports
    End If
End Sub

Public Sub InjectLinkedUserforms()
    Dim ListOfImports As String
    Dim s As String: s = Code_All
    Dim element As Variant
    For Each element In LinkedUserforms
        If InStr(1, s, "@INCLUDE USERFORM " & element) = 0 And InStr(1, ListOfImports, "@INCLUDE USERFORM " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE USERFORM " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE USERFORM " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        oModule.CodeModule.InsertLines LineIndex(Body_First_AfterComments), ListOfImports
    End If
End Sub

Public Sub InjectLinkedDeclarations()
    Dim ListOfImports As String
    Dim s As String: s = Code_All
    Dim coll As New Collection
    Dim element As Variant
    For Each element In LinkedDeclarations
        If InStr(1, s, "'@INCLUDE DECLARATION " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE DECLARATION " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE DECLARATION " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        oModule.CodeModule.InsertLines LineIndex(Body_First_AfterComments), ListOfImports
    End If
End Sub

Public Function LinkedDeclarations() As Collection
    aWorkbook.Init(oWorkbook).DeclarationsTableCreate
    Dim TargetWorksheet As Worksheet: Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    Dim coll As New Collection
    Dim s As String: s = Code_All
    Dim element
    For Each element In DeclarationsTableKeywords
        If RegexTest(s, "\b ?" & CStr(element) & "\b") Then
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
    Dim Lr As Long: Lr = getLastRow(TargetWorksheet)
    Dim coll As New Collection
    Dim cell As Range
    For Each cell In TargetWorksheet.Range("C2:C" & Lr)
        On Error Resume Next
        coll.Add cell.text, cell.text
        On Error GoTo 0
    Next
    Set DeclarationsTableKeywords = coll
End Function

Public Function LinkedUserforms()
    Dim coll As New Collection
    Dim s As String: s = Code_All
    Dim formName
    For Each formName In aModules.Init(oWorkbook).UserformNames
        If RegexTest(s, "\W" & formName & "[.(\W]") = True Then coll.Add formName
    Next
    Set LinkedUserforms = coll
End Function




Public Function LinkedClasses() As Collection
    Dim coll As New Collection
    Dim var As Variant
        var = aModule.Init(oModule).ClassCalls
    Dim s As String: s = Code_All
    Dim Keyword As String
    Dim ClassName As String
    Dim element As Variant
    Dim i As Long
    On Error Resume Next
    For i = LBound(var, 1) To UBound(var, 1)
        If InStr(1, s, var(i, 1)) > 0 Or InStr(1, s, var(i, 2)) > 0 Then
            coll.Add var(i, 1), var(i, 1)
        End If
    Next
    For Each element In aModules.Init(oWorkbook).ClassNames
        If InStr(1, s, element) > 0 Then
            coll.Add element, CStr(element)
        End If
    Next
    On Error GoTo 0
    Set LinkedClasses = coll
End Function


Function LinkedSheets() As Collection
    Dim SheetNames As New Collection
    Dim TargetWorksheet As Worksheet
    For Each TargetWorksheet In oWorkbook.Worksheets
        On Error Resume Next
        SheetNames.Add TargetWorksheet.Name
        SheetNames.Add oWorkbook.Worksheets(TargetWorksheet.Name).codeName
        On Error GoTo 0
    Next
    Dim element
    Dim s As String: s = Code_All
    Dim coll As New Collection
    For Each element In SheetNames
        If InStr(1, CStr(element), " ") > 0 And InStr(1, s, CStr(element)) > 0 Then
            coll.Add CStr(element)
        ElseIf RegexTest(s, "\W" & CStr(element) & "[.(\W]") = True Then
            coll.Add CStr(element), CStr(element)
        End If
    Next
    Set LinkedSheets = coll
End Function


Public Sub ExportLinkedCode()
    Dim Code       As String:   Code = LinkedCode
    Dim MergedName As String:   MergedName = "Merged_" & oProcedure
    Dim fileName   As String:   fileName = LOCAL_LIBRARY_PROCEDURES & MergedName & ".txt"
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

    Dim Module As VBComponent
    On Error Resume Next
    Set Module = ModuleOfProcedure(oWorkbook, Procedure)
    If Module Is Nothing Then Exit Sub
    On Error GoTo 0
    Dim obj As String
    Dim Code As String
        Code = aProcedure.Init(oWorkbook, Module, Procedure).Code_All
    Dim var
        var = Split(Code, vbNewLine)
        var = Filter(var, "'@INCLUDE ")
    Dim TextLine As Variant
    For Each TextLine In var
        TextLine = UCase(Trim(TextLine))
        If TextLine Like "'@INCLUDE *" Then
            obj = Split(TextLine, " ")(2)
            obj = VBA.Replace(obj, vbNewLine, "")
            Select Case True
                Case TextLine Like "'@INCLUDE PROCEDURE *":     ImportProcedure obj, Overwrite
                Case TextLine Like "'@INCLUDE CLASS *":         ImportClass obj, Overwrite
                Case TextLine Like "'@INCLUDE USERFORM *":      ImportUserform obj, Overwrite
                Case TextLine Like "'@INCLUDE DECLARATION *":   ImportDeclaration obj
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
    Dim ProcedurePath   As String:  ProcedurePath = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"
    Dim Result As String
    On Error Resume Next
    Result = TxtRead(ProcedurePath)
    On Error GoTo 0

    If Len(Result) = 0 Then
        On Error Resume Next
        Result = TXTReadFromUrl(GITHUB_LIBRARY_PROCEDURES & Procedure & ".txt")
        On Error GoTo 0
        If Len(Result) > 0 And Not UCase(Result) Like ("*NOT FOUND*") Then
            TxtOverwrite ProcedurePath, Result
        Else
            Debug.Print "File " & Procedure & ".txt not found neither localy nor online"
            Exit Sub
        End If
    End If

    Dim filelastmod:    filelastmod = StringLastModified(Result)
    Dim proclastmod
    
    Dim Module As VBComponent
    If ProcedureExists(oWorkbook, Procedure) = True Then
        Set Module = ModuleOfProcedure(oWorkbook, Procedure)
        proclastmod = ProcedureLastModified(oWorkbook, Module, Procedure)
        If Overwrite = True Then
            If proclastmod = 0 Or proclastmod < filelastmod Then
                ProcedureReplace Module, Procedure, TxtRead(ProcedurePath)
            End If
        End If
    Else
'        Dim ModuleName As String
'            ModuleName = StringProcedureAssignedModule(Result)
'        If ModuleName = "" Then ModuleName = "vbArcImports"
'        Set Module = ModuleAddOrSet(oWorkbook, ModuleName, vbext_ct_StdModule)
    Set Module = ModuleAddOrSet(oWorkbook, "vbArcImports", vbext_ct_StdModule)
        Module.CodeModule.AddFromFile ProcedurePath
    End If

    ImportProcedureDependencies Procedure, Overwrite
End Sub

Sub ImportDeclaration(DeclarationName As String)
    Dim filepath As String
    filepath = LOCAL_LIBRARY_DECLARATIONS & DeclarationName & ".txt"
    Dim Result As String
    On Error Resume Next
    Result = TxtRead(filepath)
    On Error GoTo 0

    If Len(Result) = 0 Then 'CheckPath(filePath) = "I" Then
        On Error Resume Next
        Result = TXTReadFromUrl(GITHUB_LIBRARY_DECLARATIONS & DeclarationName & ".txt")
        On Error GoTo 0
        If Len(Result) > 0 And Not UCase(Result) Like ("*NOT FOUND*") Then
            TxtOverwrite filepath, Result
        Else
            Debug.Print "File " & DeclarationName & ".txt not found localy or online"
            Exit Sub
        End If
    Else

    End If
    If InStr(1, aWorkbook.Init(oWorkbook).Code, Result, vbTextCompare) > 0 Then Exit Sub
    Dim Module As VBComponent
    Set Module = ModuleAddOrSet(oWorkbook, "vbArcImports", vbext_ct_StdModule)
    Module.CodeModule.AddFromString FormatVBA7(Result)

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

    Dim filepath As String
    filepath = LOCAL_LIBRARY_CLASSES & ClassName & ".cls"
    If CheckPath(filepath) = "I" Then
        On Error Resume Next
        Dim Code As String
        Code = TXTReadFromUrl(GITHUB_LIBRARY_CLASSES & ClassName & ".cls")
        On Error GoTo 0
        If Len(Code) > 0 And Not UCase(Code) Like ("*NOT FOUND*") Then
            TxtOverwrite filepath, Code
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
    oWorkbook.VBProject.VBComponents.Import filepath
End Sub

Public Sub Replace(Result As String)

    Dim startLine As Integer
    Dim NumLines As Integer
    With oModule.CodeModule
        startLine = .procStartLine(oProcedure, oProcKind)
        NumLines = .ProcCountLines(oProcedure, oProcKind)
        .DeleteLines startLine, NumLines
        .InsertLines startLine, Result
    End With
End Sub

Private Function EnumOptionToString(PropertyType As Property_Type) As String
    Select Case PropertyType
        Case Property_Type.Modified:          EnumOptionToString = "LastModified"
        Case Property_Type.Ignore:            EnumOptionToString = "Ignore"
        Case Property_Type.Parent_Assigned:   EnumOptionToString = "AssignedModule"
        
    End Select
End Function

Private Function PropertyGet(PropertyType As Property_Type) As String
    PropertyGet = vbNullString
    Dim Line As String: Line = PropertyLine(PropertyType)
    If Line = 0 Then Exit Function
    Dim targetLine As String: targetLine = Trim(oModule.CodeModule.Lines(Line, 1))
    PropertyGet = Split(targetLine, " ")(1)
End Function

Public Function PropertySet(PropertyType As Property_Type, Value As String)
    Dim Line As Long: Line = PropertyLine(PropertyType)
    If Line = 0 Then
        oModule.CodeModule.InsertLines BodyLineFirst, "'@" & EnumOptionToString(PropertyType) & " " & Value
    Else
        Dim targetLine As String: targetLine = Trim(oModule.CodeModule.Lines(Line, 1))
        oModule.CodeModule.ReplaceLine Line, Split(targetLine)(0) & " " & Value
    End If

End Function

Private Function PropertyLine(PropertyType As Property_Type) As Long
    Dim This As String: This = EnumOptionToString(PropertyType)
    Dim targetLine  As String
    Dim i As Long
    For i = LineFirst To LineLast
        targetLine = Trim(oModule.CodeModule.Lines(i, 1))
        If UCase(targetLine) Like UCase("'@" & This & " *") Then
            PropertyLine = i
            Exit Function
        End If
    Next
End Function

Public Sub InjectObjectsRelease()
    Dim LineStart As Long, LineEnd As Long
    AssignLineStartEnd Body_Code, LineStart, LineEnd
    Dim i As Long
    Dim s As String, This As String, Keyword As String
    For i = LineStart To LineEnd
        s = Trim(oModule.CodeModule.Lines(i, 1))
        If (s Like "Set * = *" Or s Like "Dim*As New*") _
        And Not s Like "*= Nothing*" Then
            Keyword = Split(s, " ")(1)
            This = IIf(This <> "", This & vbNewLine, "") & "Set " & Keyword & " = Nothing"
        End If
    Next
    If This = "" Then Exit Sub
    Dim var: var = Split(This, vbNewLine)
    Dim sCode As String: sCode = Code_All
    Dim out, el
    ReDim out(0 To 0)
    For Each el In var
        If InStr(1, sCode, el) = 0 Then
            If out(0) <> "" Then ReDim Preserve out(0 To UBound(out) + 1)
            out(UBound(out)) = el
        End If
    Next
    This = Join(out, vbNewLine)
    If This <> "" Then oModule.CodeModule.InsertLines LineIndex(Procedure_Last), This
End Sub

Public Sub InjectModification()
    PropertySet Modified, Format(Now, "yymmddhhnn")
    
    Const vbTab2 = vbTab & vbTab
    Const vbTab4 = vbTab2 & vbTab2
    Dim sTime As String: sTime = Format(Now, "dd-mm-yyyy hh:nn")
    Dim sProc As String: sProc = oProcedure

    '* @TODO Created: 01-02-2023 07:46 Author: Anastasiou Alex
    '* @TODO if name is too long have to recalculate tab length

    Dim sUser As String: sUser = "Alex"
    If sUser = vbNullString Then sUser = Environ("UserName")
    Const sUPDATE    As String = "'* Updated    :"
    Dim sFirstLine As String
        sFirstLine = "'* Modified   :" & vbTab & "Date and Time" & vbTab2 & _
                    "Author" & vbTab4 & "Description" & vbCrLf
    Dim sSecondLine  As String
        sSecondLine = sUPDATE & vbTab & sTime & vbTab & sUser & vbTab4 & "(" & sProc & ")"

    If InStr(1, Code_Header, "'* Created    :") Then
        Dim lineNo As Long
        Dim i As Long
        For i = LineIndex(Header_First) To LineIndex(Header_Last)
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

    Dim nLine As Long:   nLine = LineIndex(Declaration_First)
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
    Dim Procedures As New Collection
    Set Procedures = aModule.Init(oModule).ProceduresNames
    Dim index As Long
        index = aCollection.Init(Procedures).IndexOf(oProcedure)
    IndexInModule = index
End Function

Public Sub MoveToTop()
    Dim Procedures As New Collection:   Set Procedures = aModule.Init(oModule).ProceduresNames
    Dim idx As Long:                    idx = IndexInModule
    
    If idx = 1 Then Exit Sub
    
    Dim s As String:                    s = Code_All
    Dim TargetProcedure As String:      TargetProcedure = Procedures(1)
    
    Delete
    
    With oModule.CodeModule
        .InsertLines .procStartLine(TargetProcedure, oProcKind), s
    End With
    
    idx = LineIndex(Declaration_First)
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1
End Sub

Public Sub MoveToBottom()
    Dim Procedures As New Collection:   Set Procedures = aModule.Init(oModule).ProceduresNames
    Dim idx As Long:                  idx = IndexInModule
    
    If idx = Procedures.count Then Exit Sub
    
    Dim s As String:                    s = Code_All
    Dim TargetProcedure As String:      TargetProcedure = Procedures(Procedures.count)
        
    Delete
    
    With oModule.CodeModule
        .InsertLines .procStartLine(TargetProcedure, oProcKind) + _
                     .ProcCountLines(TargetProcedure, oProcKind) + 1, _
                     s
    End With
    
    idx = LineIndex(Declaration_First)
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1
End Sub

Public Sub MoveDown()
    Dim Procedures As New Collection:   Set Procedures = aModule.Init(oModule).ProceduresNames
    Dim idx As Long:                    idx = IndexInModule
    
    If idx = Procedures.count Then Exit Sub
    
    Dim s As String:                    s = Code_All
    Dim TargetProcedure As String:      TargetProcedure = Procedures(idx + 1)
    
    Delete
    
    With oModule.CodeModule
        .InsertLines .procStartLine(TargetProcedure, oProcKind) + _
                     .ProcCountLines(TargetProcedure, oProcKind) + 1, _
                     s
    End With
    
    idx = LineIndex(Declaration_First)
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1

End Sub

Public Sub MoveUp()
    Dim Procedures As New Collection:   Set Procedures = aModule.Init(oModule).ProceduresNames
    Dim idx As Long:                    idx = IndexInModule
    
    If idx = 1 Then Exit Sub
    
    Dim s As String:                    s = Code_All
    Dim TargetProcedure As String:      TargetProcedure = Procedures(idx - 1)
    
    Delete
    
    With oModule.CodeModule
        .InsertLines .procStartLine(TargetProcedure, oProcKind), s
    End With
    
    idx = LineIndex(Declaration_First)
    Application.VBE.ActiveCodePane.SetSelection idx, 1, idx, 1

End Sub

Public Sub MoveToAssignedModule()
    Dim Module As VBComponent
    Set Module = ParentAssigned
    If Module Is Nothing Or Module.Name = oModule.Name Then Exit Sub
    MoveToModule Module
End Sub

Public Sub MoveToModule(Module As VBComponent)
    Dim s As String: s = Code_All
    Delete
    With Module.CodeModule
        .InsertLines .CountOfLines + 1, vbNewLine & s
    End With
    Set oModule = Module
End Sub

Public Sub Delete()
    TxtOverwrite Environ("TEMP") & "\" & oProcedure & ".txt", Code_All
    Debug.Print "made a backup of " & oProcedure & " in " & Environ("TEMP")
    oModule.CodeModule.DeleteLines LineIndex(Procedure_First), LineIndex(Procedure_Count)
End Sub

Public Sub TestCreate()
    Dim arg As String
    arg = ArgumentStyle
    Dim Procedures As Collection
    Set Procedures = aWorkbook.Init(oWorkbook).Procedures(True, False, False, False)
    
    Dim s As String: s = vbNewLine
    If Not aCollection.Init(Procedures).Contains(, "test" & oProcedure) Then
        If KindAsString = "Sub" Then
            s = s & "Sub test" & oProcedure & "()" & vbNewLine
            s = s & "call " & arg & vbNewLine
            s = s & "End Sub"
        Else
            s = s & "Function test" & oProcedure & "()" & vbNewLine
            s = s & "Something = " & arg & vbNewLine
            s = s & "End Function"
        End If
    Else

    End If
    oModule.CodeModule.InsertLines LineIndex(Procedure_First), s
    Debug.Print "Created test" & oProcedure & " in " & oModule.Name
End Sub

Public Function ArgumentStyle() As String
    If ArgumentCount = 0 Then Exit Function
    Dim str As Variant, s As String
    Dim firstPart As String, secondPart As String, output As String
        str = Code_Declaration_Clean
    output = oProcedure & "( _"
    Dim indentation As String
        indentation = Space(Len(output) - 1)
        str = Right(str, Len(str) - InStr(1, str, "("))
        str = Left(str, InStrRev(str, ")") - 1)
    If InStr(1, str, Chr(34) & "," & Chr(34)) > 0 Then
        str = VBA.Replace(str, Chr(34) & "," & Chr(34), Chr(34) & "|" & Chr(34))
    End If
        str = Split(str, ",")
    Dim i As Long
    For i = LBound(str) To UBound(str)
        str(i) = VBA.Replace(str(i), Chr(34) & "|" & Chr(34), Chr(34) & "," & Chr(34))
    Next
    For i = LBound(str) To UBound(str)
        s = Trim(str(i))

        s = VBA.Replace(s, "Optional ", "")
        s = VBA.Replace(s, "As ", "")
        s = VBA.Replace(s, "ByVal ", "")
        s = VBA.Replace(s, "ByRef ", "")
        s = VBA.Replace(s, "ParamArray ", "")
        s = VBA.Replace(s, "_", "")

        firstPart = Split(s, " ")(0)
        If InStr(1, s, " ") Then
            secondPart = Split(s, " ")(1)
        Else
            secondPart = "Variant"
        End If
        output = output & vbNewLine & indentation & firstPart & ":= " & "as" & secondPart & IIf(i <> UBound(str), ", _", ")")
    Next

    ArgumentStyle = output
End Function

Public Function ArgumentStyleClean() As String
    Dim var As Variant
    Dim s As String
    var = Split(ArgumentStyle, vbNewLine)
    var = ArrayTrim(var)
    If UBound(var) = -1 Then
'        ProcedureArgumentStyleClean = Procedure & "()"
        Exit Function
    End If
    s = Join(var, vbNewLine)
    s = VBA.Replace(s, " _" & vbNewLine, "")
    ArgumentStyleClean = s
End Function

Sub PrintDims()
    dp Dims
End Sub
    
Public Property Get Dims() As Variant
    Dim s As Variant:   s = Split(Code_All, vbNewLine)
    Dim tmp:            ReDim tmp(0 To 0)
    
    Dim i As Long
    For i = LBound(s) To UBound(s)
        s(i) = Trim(s(i))
        If s(i) Like "*:*" Then s(i) = Split(s(i), ":")(0)
    Next

    s = ArrayFilterLike(s, "Dim *", True)
    Dim h As String
    Dim element
    For i = LBound(s) To UBound(s)
        s(i) = Trim(StringCommentsRemove(s(i), False))
        If s(i) Like "Dim *,* As *" Then
            For Each element In Split(s(i), ", ")
                If InStr(1, element, "Dim ") = 0 Then element = "Dim " & element
                If InStr(1, element, " As ") = 0 Then element = element & " As "
                tmp(UBound(tmp)) = element
                ReDim Preserve tmp(0 To UBound(tmp) + 1)
            Next
        Else
            For Each element In Split(s(i), ", ")
                tmp(UBound(tmp)) = element
                ReDim Preserve tmp(0 To UBound(tmp) + 1)
            Next
        End If
    Next

    s = ArrayFilterLike(tmp, "Dim *", True)

    Dim var
    ReDim var(0 To 1, 0 To 0)
    For i = LBound(s) To UBound(s)
        element = s(i)
        If element Like "Dim * As *" Then

            element = Mid(element, 5)
            var(0, UBound(var, 2)) = Split(element, " As ")(0)
            var(1, UBound(var, 2)) = Split(element, " As ")(1)
            If Split(element, " As ")(1) Like "New *" Then
                var(1, UBound(var, 2)) = Split(var(1, UBound(var, 2)), " ")(1)
            End If
            If i < UBound(s) Then
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
            If i < UBound(s) Then
                ReDim Preserve var(0 To 1, 0 To UBound(var, 2) + 1)
            End If
        End If
    Next

    Dims = WorksheetFunction.Transpose(var)

End Property

Public Sub InjectToHeaderTop( _
                    This As String, _
                    SkipIfExists As Boolean)
    If SkipIfExists Then
        If InStr(1, Code_Body, This) > 0 Then Exit Sub
    End If
    oModule.CodeModule.InsertLines LineIndex(Procedure_First), _
                                   IIf(Len(oModule.CodeModule.Lines(LineIndex(Procedure_First), 1)) = 0, vbNewLine, "") & _
                                   This
End Sub

Public Sub InjectToHeaderBottom( _
                        This As String, _
                        SkipIfExists As Boolean)
    If SkipIfExists Then
        If InStr(1, Code_Body, This) > 0 Then Exit Sub
    End If
    oModule.CodeModule.InsertLines LineIndex(Declaration_First), This
End Sub

Public Sub InjectToBodyTop( _
                    This As String, _
                    SkipIfExists As Boolean)
    If SkipIfExists Then
        If InStr(1, Code_Body, This) > 0 Then Exit Sub
    End If
    oModule.CodeModule.InsertLines LineIndex(Body_First), This
End Sub

Public Sub InjectToBodyAfterComments( _
                                    This As String, _
                                    SkipIfExists As Boolean)
    If SkipIfExists Then
        If InStr(1, Code_Body, This) > 0 Then Exit Sub
    End If
    oModule.CodeModule.InsertLines LineIndex(Body_First_AfterComments), This
End Sub

Public Sub InjectToBodyBottom( _
                             This As String, _
                             SkipIfExists As Boolean)
    If SkipIfExists Then
        If InStr(1, Code_Body, This) > 0 Then Exit Sub
    End If
    oModule.CodeModule.InsertLines LineIndex(Body_Last) + 1, This
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 17-08-2023 08:28    Alex                (InjectDescription)

Sub InjectDescription()
'@LastModified 2308170828

    Const vbTab2 = vbTab & vbTab
    Const vbTab4 = vbTab2 & vbTab2
    Const sFersLine     As String = "'* Modified   :" & vbTab & "Date and Time" & vbTab2 & "Author" & vbTab4 & "Description" & vbCrLf

    Dim txtName         As String: txtName = AUTHOR_NAME
    If txtName = vbNullString Then txtName = Environ("UserName")

    Dim txtContacts     As String: txtContacts = AUTHOR_EMAIL
    If txtContacts <> vbNullString Then txtContacts = "'* Contacts   :" & vbTab & txtContacts & vbCrLf

    Dim txtCopyright    As String:  txtCopyright = AUTHOR_COPYRIGHT
    If txtCopyright <> vbNullString Then txtCopyright = "'* Copyright  :" & vbTab & txtCopyright & vbCrLf

    Dim txtOther        As String:   txtOther = AUTHOR_OTHERTEXT
    If txtOther <> vbNullString Then txtOther = "'* Note       :" & vbTab & txtOther & vbCrLf

    Dim txtMedia        As String:   txtMedia = "'* " & vbLf & AUTHOR_MEDIA
    Dim CurentCodePane  As CodePane: Set CurentCodePane = oModule.CodeModule.CodePane
    Dim i               As Byte
    Dim ProcKind        As VBIDE.vbext_ProcKind
    Dim sProc           As String: sProc = ActiveProcedure
    Dim nLine           As Long
        nLine = LineIndex(Declaration_First)
    Dim sTemp           As String
    Dim sTime           As String
    Dim sType           As String
    Dim sProcDeclartion As String
    Dim sProcArguments  As String

    Dim sUser As String

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
    Dim arStr() As String
        arStr = Split(sDeclaration, ",")
    Dim iMaxLen As Byte
        iMaxLen = 0
    Dim iTempLen As Byte
    Dim sTemp  As String
    Dim i      As Byte
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
    
    Dim Result As String

    Select Case CallerModules.count
    Case 0
        SuggestedScope = "Unspecified"
        Result = Result & vbNewLine & "Procedure " & Name & " has unclarified scope in " & oWorkbook.Name
        Result = Result & vbNewLine & "It is called 0 times, so it may be Unused, Unassigned, or Called from elsewhere"
        Exit Function
    Case 1
        SuggestedScope = "Private"
        Result = Result & vbNewLine & "Suggested scope for procedure " & Name & " is Private"
    Case Is > 1
        SuggestedScope = "Public"
        Result = Result & vbNewLine & "Suggested scope for procedure " & Name & " is Public"
    End Select

    Result = Result & vbNewLine & "because it is used in " & CallerModules.count & " module(s) of " & oWorkbook.Name
    
End Function

Sub ScopeSuggested()
    Dim idx         As Long:        idx = LineIndex(Declaration_First)
    Dim s           As String:      s = oModule.CodeModule.Lines(idx, 1)
    Dim Suggestion  As String:      Suggestion = SuggestedScope
    
    If Suggestion = "Private" Or Suggestion = "Public" Then
        Select Case True
        Case s Like "*Public*" & Name & "*"
            ScopePublic
        Case s Like "*Private*" & Name & "*"
            ScopePrivate
        Case Else
            Dim sKind As String:    sKind = KindAsString
            oModule.CodeModule.ReplaceLine idx, Trim(VBA.Replace(s, sKind, " " & Suggestion & " " & sKind, , 1, vbTextCompare))
        End Select
    End If
End Sub

Sub ScopePrivate()
    Dim idx         As Long:        idx = LineIndex(Declaration_First)
    Dim s           As String:      s = Code_Declaration_Clean
    Dim Suggestion  As String:      Suggestion = "Private"
    Select Case True
    Case s Like "*Public*" & Name & "*"
        Debug.Print Name
        oModule.CodeModule.ReplaceLine idx, VBA.Replace(s, "Public", "Private", , 1, vbTextCompare)
    Case s Like "*Private*" & Name & "*"
        
    Case Else
        Dim sKind As String:    sKind = KindAsString
        Select Case sKind
            Case sKind Like "*Property*"
                Debug.Print Name
                oModule.CodeModule.ReplaceLine idx, VBA.Replace(s, "Property", Suggestion & " " & "Property", , 1, vbTextCompare)
            Case Else
                Debug.Print Name
                oModule.CodeModule.ReplaceLine idx, VBA.Replace(s, sKind, Suggestion & " " & sKind, , 1, vbTextCompare)
        End Select
    End Select
End Sub

Sub ScopePublic()
    Dim idx         As Long:        idx = LineIndex(Declaration_First)
    Dim s           As String:      s = Code_Declaration_Clean
    Dim Suggestion  As String:      Suggestion = "Public"
    Select Case True
    Case s Like "*Public*" & Name & "*"
        
    Case s Like "*Private*" & Name & "*"
        Debug.Print Name
        oModule.CodeModule.ReplaceLine idx, VBA.Replace(s, "Private", "Public", , 1, vbTextCompare)
    Case Else
        Dim sKind As String:    sKind = KindAsString
        Debug.Print Name
        oModule.CodeModule.ReplaceLine idx, Trim(VBA.Replace(s, sKind, " " & Suggestion & " " & sKind, , 1, vbTextCompare))
    End Select
End Sub

Function CallerModules() As Collection
    Dim myName As String: myName = Name
    Dim coll As New Collection
    Dim Module As aModule
    Dim matchCollection As New Collection
    For Each Module In aModules.Init(oWorkbook).Items
        If Module.Contains(myName, True, True, False) Then
            On Error Resume Next
            coll.Add Module.Name, Module.Name
            On Error GoTo 0
        End If
    Next
    Set CallerModules = coll
End Function

Function CallerModulesToString() As String
    Dim element
    Dim Result As String
    For Each element In CallerModules
        Result = Result & IIf(Result <> "", vbNewLine, "") & element
    Next
    CallerModulesToString = Result
End Function

Function Callers() As Collection
    Dim myName As String: myName = Name
    Dim Result As New Collection
    Dim Module As aModule
    Dim Procedure As aProcedure
    For Each Module In CallerModules
        For Each Procedure In Module.Procedures
            If Procedure.Name <> myName Then
                If Procedure.Contains(myName, True, True, False) Then
                    On Error Resume Next
                    Result.Add Procedure.Name, Procedure.Name
                    Exit For
                    On Error GoTo 0
                End If
            End If
        Next
    Next
    Set Callers = Result
End Function

Function CallersToString() As String
    Dim element
    Dim Result As String
    For Each element In Callers
        Result = Result & IIf(Result <> "", vbNewLine, "") & element
    Next
    CallersToString = Result
End Function

Public Sub BringProcedureHere(Procedure As String)
    Dim AP As New aProcedure
    AP.Init , , Procedure
    Dim s As String
        s = AP.Code_All
        If InStr(1, s, "'@AssignedModule") = 0 Then
            AP.PropertySet Parent_Assigned, AP.Parent.Name
            s = AP.Code_All
        End If
    AP.Delete
    oModule.CodeModule.InsertLines LineIndex(Procedure_Last) + 1, s
End Sub

Sub BringLinkedProceduresHere()
    Dim el
    Dim AP As aProcedure
    For Each el In LinkedProceduresDeep
        BringProcedureHere CStr(el)
    Next
End Sub



Sub AddToLinkedTable()
    
    Dim aw As aWorkbook
    Set aw = aWorkbook.Init(oWorkbook)
    Dim cell As Range
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

    Dim Lr As Long
    Lr = getLastRow(TargetWorksheet) + 1
    TargetWorksheet.Range("A" & Lr & ":H" & Lr).Value = var
End Sub



Public Sub Enable_DebugPrint()
    Dim N As Long
    Dim s As String
    With oModule.CodeModule
        For N = LineIndex(Procedure_Last) To LineIndex(Procedure_First) Step -1
            s = .Lines(N, 1)
            If Left(Trim(s), 6) = "'Debug" Then
                aCodeModule.Init(oModule).UncommentTargetLine N
'                s = VBA.Replace(s, "'", "", , 1)
'                .ReplaceLine N, s
            End If
        Next N
    End With
End Sub
Public Sub Disable_DebugPrint()
    Dim N As Long
    Dim s As String
    With oModule.CodeModule
        For N = LineIndex(Procedure_Last) To LineIndex(Procedure_First) Step -1
            s = .Lines(N, 1)
            If Left(Trim(s), 5) = "Debug" Then
                aCodeModule.Init(oModule).CommentTargetLine N
'                .ReplaceLine N, "'" & s
            End If
        Next N
    End With
End Sub


Public Sub Enable_Stop()
    Dim N As Long
    Dim s As String
    Dim Keyword As String
    Keyword = "Stop"
    With oModule.CodeModule
        For N = LineIndex(Procedure_Last) To LineIndex(Procedure_First) Step -1
            s = .Lines(N, 1)
            If InStrExact(1, s, Keyword) > 0 Then
                s = VBA.Replace(s, "'", "", , 1)
                .ReplaceLine N, s
            End If
        Next N
    End With
End Sub
Public Sub Disable_Stop()
    Dim N As Long
    Dim s As String
    Dim Keyword As String: Keyword = "Stop"
    With oModule.CodeModule
        For N = LineIndex(Procedure_Last) To LineIndex(Procedure_First) Step -1
            s = .Lines(N, 1)
            If InStrExact(1, s, Keyword) > 0 Then
                .ReplaceLine N, "'" & s
            End If
        Next N
    End With
End Sub


Public Sub Comments_ReplaceQuoteWithRem()
    Dim N As Long
    Dim s As String
    With oModule.CodeModule
        For N = LineIndex(Procedure_Last) To LineIndex(Procedure_First) Step -1
            s = .Lines(N, 1)
            If Left(Trim(s), 1) = "'" Then
                .ReplaceLine N, VBA.Replace(s, "'", "Rem ", , 1)
            End If
        Next N
    End With
End Sub


Public Sub UpdatableVariable_Add(This As String)
    Dim i As Long: i = LineLike("*UpdatableVariable = *", Body_Code)
    If i = 0 Then Exit Sub
    Dim Q As String: Q = Chr(34)
    Dim Line As String: Line = oModule.CodeModule.Lines(i, 1)
    If UCase(Trim(Line)) Like UCase("UpdatableVariable = *") And InStr(1, Line, This, vbTextCompare) = 0 Then
        oModule.CodeModule.ReplaceLine i, VBA.Replace(VBA.Replace(Line, Q, Q & This & ",", , 1, vbTextCompare), "," & Q, Q)
        Line = oModule.CodeModule.Lines(i, 1)
        If InStr(1, Line, Q & ",") > 0 Then oModule.CodeModule.ReplaceLine i, VBA.Replace(Line, Q & ",", Q)
    End If
End Sub

Public Sub UpdatableVariable_Remove(This As String)
    Dim i As Long: i = LineLike("*UpdatableVariable = *", Body_Code)
    If i = 0 Then Exit Sub
    Dim Q As String: Q = Chr(34)
    Dim Result As String
    Dim Line As String: Line = oModule.CodeModule.Lines(i, 1)
    If Trim(Line) Like "UpdatableVariable = *" Then
        If InStr(1, Line, Q & This & Q, vbTextCompare) > 0 Then
            Result = VBA.Replace(Line, Q & This & Q, "")
            Result = VBA.Replace(Result, ",,", ",")
            Result = VBA.Replace(Result, Q & ",", Q)
            oModule.CodeModule.ReplaceLine i, Result
        End If
    End If
End Sub

Sub CreateCaller(rng As Range)
    Dim shp As Shape
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
    With shp.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    
    With shp
        .OnAction = "'" & oWorkbook.Name & "'!" & oProcedure
        .Name = "Run_" & oProcedure
        .TextFrame2.TextRange.text = oProcedure
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
    Dim L As Long: L = LargestLineLength
    Dim i As Long
    For i = LineIndex(Procedure_Last) To LineIndex(Declaration_First) Step -1
        If Len(Trim(oModule.CodeModule.Lines(i, 1))) = 0 Then
            oModule.CodeModule.ReplaceLine i, "'" & String(L - 1, Character)
        End If
    Next
End Sub
