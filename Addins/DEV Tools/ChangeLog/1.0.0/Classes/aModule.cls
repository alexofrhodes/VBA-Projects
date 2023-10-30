VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aModule
'* Purpose    :
'* Copyright  :
'*
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 29-06-2023 14:39    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private oModule     As VBComponent
Private oWorkbook   As Workbook

Public Sub PredeclaredIDenable()
    SetPredeclaredId True
End Sub

Public Sub SetPredeclaredId(Toggle As Boolean)
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE TxtOverwrite
    '@INCLUDE PROCEDURE TxtRead
    '@INCLUDE CLASS aModule

    Dim FName1      As String: FName1 = ThisWorkbook.path & "\" & "TempClass.cls"
    oWorkbook.VBProject.VBComponents(oModule.Name).Export FName1
    TxtOverwrite FName1, VBA.Replace( _
            VBA.Replace(TxtRead(FName1), _
            IIf(Toggle = True, _
            "Attribute VB_PredeclaredId = False", _
            "Attribute VB_PredeclaredId = True"), _
            IIf(Toggle = True, _
            "Attribute VB_PredeclaredId = True", _
            "Attribute VB_PredeclaredId = False"), _
            1, _
            1), _
            vbLf, vbCrLf, _
            1, _
            -1)

    Delete
    oWorkbook.VBProject.VBComponents.Import FName1
    Kill FName1

End Sub

Public Function ClassCalls() As Variant
    'classCallsOfModule(0) is the class name
    'classCallsOfModule(1) is the keyword for the class name (eg dim clsCal as new classCalendar)
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE ClassNames
    '@INCLUDE PROCEDURE classCallsOfModule
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules

    Dim Code        As Variant
    Dim element     As Variant
    Dim Keyword     As Variant
    Dim var         As Variant
    ReDim var(1 To 2, 1 To 1)
    Dim counter     As Long: counter = 0
    Dim ClassName   As Variant
    If oModule.CodeModule.CountOfDeclarationLines > 0 Then
        Code = oModule.CodeModule.Lines(1, oModule.CodeModule.CountOfDeclarationLines)
        Code = Replace(Code, "_" & vbNewLine, "")
        Code = Split(Code, vbNewLine)
        Code = Filter(Code, " As ", , vbTextCompare)
        For Each element In Code
            element = Trim(element)
            If element Like "* As *" Then
                Keyword = Split(element, " As ")(0)
                Keyword = Split(Keyword, " ")(UBound(Split(Keyword, " ")))
                element = Split(element, " As ")(1)
                element = Replace(element, "New ", "")
                For Each ClassName In aModules.Init(oWorkbook).ClassNames
                    If element = ClassName Then
                        ReDim Preserve var(1 To 2, 1 To counter + 1)
                        var(1, UBound(var, 2)) = element
                        var(2, UBound(var, 2)) = Keyword
                        counter = counter + 1
                        '                        Exit For
                    End If
                Next
            End If
        Next
        If var(1, 1) <> "" Then
            If UBound(var, 2) > 1 Then
                ClassCalls = WorksheetFunction.Transpose(var)
            Else
                Dim VAR2(1 To 1, 1 To 2)
                VAR2(1, 1) = var(1, 1)
                VAR2(1, 2) = var(2, 1)
                ClassCalls = VAR2
            End If
        End If
    End If

End Function
Public Function Active() As aModule
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE ActiveModule
    '@INCLUDE CLASS aModule
    Init ActiveModule
    Set Active = Me
End Function

Public Function Init(module As Variant) As aModule
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    If TypeName(module) = "VBComponent" Then
        Set oModule = module
    Else
        Set module = module.Parent.VBProject.VBComponents(module.codeName)
    End If
    Set oWorkbook = Parent
    Set Init = Me
End Function

Sub PrintListOfInclude()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE dp
    '@INCLUDE CLASS aModule
    dp ListOfInclude
End Sub

Function ListOfInclude() As Collection
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE ArrayQuickSort
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aCollection

    Dim var: var = Split(Code, vbNewLine)
    ArrayQuickSort Filter(Filter(var, "'@INCLUDE ", True), """", False)
    Dim coll        As New Collection
    Dim el
    On Error Resume Next
    For Each el In var
        el = Trim(el)
        If el Like "'@INCLUDE *" Then
            coll.Add el, CStr(el)
        End If
    Next
    On Error GoTo 0
    var = aCollection.Init(coll).ToArray
    Dim matches     As String: matches = Join(var, vbNewLine)
    Dim out         As String: out = out & vbNewLine & "'=== Module: " & oModule.Name & "===" & vbNewLine & matches
    Set ListOfInclude = coll
End Function

Public Function Duplicate( _
        Optional NewName As String, _
        Optional ActivateNewModule As Boolean = True) As Boolean
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule

    If NewName = "" Then NewName = Name & "_Copy"
    On Error GoTo Error_Handler
    Dim TargetFile  As String
    TargetFile = Environ("Temp") & "\" & NewName & Extension
    oModule.Name = NewName
    oModule.Export TargetFile
    oModule.Name = Name
    oWorkbook.VBProject.VBComponents.Import TargetFile
    If Len(Dir(TargetFile)) > 0 Then Kill Replace(TargetFile, Extension, ".*")
    If ActivateNewModule = True Then oWorkbook.VBProject.VBComponents(NewName).Activate
    '    DuplicateUserForm = True
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: aModule.Duplicate" & vbCrLf & _
            "Error Description: " & Err.Description & _
            Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
            , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

Sub PrintTodoList()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE dp
    '@INCLUDE CLASS aModule
    dp TodoList
End Sub

Function TodoList()
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    Dim out         As String
    Dim matches     As String
    matches = Join(Filter(Split(Code, vbNewLine), "@TODO", True), vbNewLine)
    out = out & vbNewLine & "'---------"
    out = out & vbNewLine & "'Module: " & Name
    out = out & vbNewLine & "'---------"
    '    out = out & vbNewLine & "'" & Replace(Matches, vbNewLine, vbNewLine & "'")
    out = out & vbNewLine & matches    ' "'" & Replace(Matches, vbNewLine, vbNewLine & "'")
    TodoList = out
End Function


Sub HeaderAdd()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE ActiveProcedure
    '@INCLUDE PROCEDURE PadRight
    '@INCLUDE CLASS aModule
    '@INCLUDE DECLARATION AUTHOR_COPYRIGHT
    '@INCLUDE DECLARATION AUTHOR_EMAIL
    '@INCLUDE DECLARATION AUTHOR_NAME
    '@INCLUDE DECLARATION AUTHOR_OTHERTEXT

    Const vbTab2 = vbTab & vbTab
    Const vbTab4 = vbTab & vbTab & vbTab & vbTab
    Const sFersLine As String = "'* Modified   :" & vbTab & "Date and Time" & vbTab & vbTab & "Author" & vbTab4 & "Description" & vbCrLf

    Dim txtName     As String: txtName = AUTHOR_NAME
    If txtName = vbNullString Then txtName = Environ("UserName")
    txtName = "'* Author     :" & vbTab & txtName & vbCrLf

    Dim txtContacts As String: txtContacts = AUTHOR_EMAIL
    If txtContacts <> vbNullString Then txtContacts = "'* Contacts   :" & vbTab & txtContacts & vbCrLf

    Dim txtCopyright As String: txtCopyright = AUTHOR_COPYRIGHT
    '    If txtCopyright <> vbNullString Then txtCopyright = "'* Copyright  :" & vbTab & txtCopyright & vbCrLf
    txtCopyright = "'* Copyright  :" & vbTab & txtCopyright & vbCrLf

    Dim txtOther    As String: txtOther = AUTHOR_OTHERTEXT
    If txtOther <> vbNullString Then txtOther = "'* Note       :" & vbTab & txtOther & vbCrLf

    Dim txtMedia    As String: txtMedia = "'* " & vbLf & AUTHOR_MEDIA
    Dim CurentCodePane As CodePane: Set CurentCodePane = oModule.CodeModule.CodePane
    Dim i           As Byte
    Dim ProcKind    As VBIDE.vbext_ProcKind
    Dim sProc       As String: sProc = ActiveProcedure
    Dim nLine       As Long: nLine = 1
    Dim sTemp       As String
    Dim sTime       As String
    Dim sType       As String
    Dim sProcDeclartion As String
    Dim sProcArguments As String

    Dim sUser       As String
    On Error GoTo ErrorHandler

    sUser = "Alex"
    If sUser = vbNullString Then sUser = Environ("UserName")

    sTemp = Replace(String(90, "*"), "**", "* ")
    sTime = Format(Now, "dd-mm-yyyy hh:nn")
    sProc = oModule.Name
    sType = "'* " & PadRight(Me.TypeToString, 11) & ":" & vbTab & sProc & vbCrLf

    If LineLike("* " & PadRight(Me.TypeToString, 11) & ": " & oModule.Name, True, True) > 0 Then GoTo UpdateHeader

    With oModule.CodeModule.CodePane
        sTemp = vbLf & "'" & sTemp & vbCrLf & _
                sType & _
                "'* Purpose    :" & vbCrLf & _
                txtCopyright & "'* " & vbCrLf & _
                txtName & _
                txtContacts & _
                txtOther & _
                txtMedia & _
                sFersLine & _
                "'* Created    :" & vbTab & sTime & vbTab & sUser & vbCrLf & _
                sProcArguments & _
                "'" & sTemp
        sTemp = sTemp & vbNewLine & IIf(HeaderContains("Option Explicit"), "", "Option Explicit" & vbNewLine) & _
                                    IIf(HeaderContains("Option Private Module"), "", "Option Private Module" & vbNewLine) & _
                                    IIf(HeaderContains("Option Base 1"), "", "Option Base 1") & vbNewLine & _
                                    IIf(HeaderContains("C_MODULE_NAME"), "", "C_MODULE_NAME = " & """" & Me.Name & """" & vbNewLine)
         
'         sTemp = sTemp & "'You may later want to use: amodule.active.ListProceduresPublic"
        .CodeModule.InsertLines 1, sTemp & vbNewLine
    End With
    
NormalExit:
    Exit Sub
    
UpdateHeader:
    nLine = LineLike("'* Created    : *", True, True) + 1
    'put update to bottom of changes list
    Do While Trim(oModule.CodeModule.Lines(nLine, 1)) Like "'* Modified*)"
        nLine = nLine + 1
    Loop
    oModule.CodeModule.InsertLines nLine, "'* Modified   :" & vbTab & sTime & vbTab & sUser
    GoTo NormalExit
    
ErrorHandler:
    GoTo NormalExit
End Sub

Function item() As VBComponent
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    Set item = oModule
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 09:21    Alex                (aModule.cls > ListProcedures)

Public Sub ListProcedures()
'@LastModified 2308220921
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aCollection
    Dim coll        As New Collection
    Set coll = ProceduresNames
    If coll.Count = 0 Then Exit Sub
    
    Dim txt         As String
    txt = "'@FOLDER ListOfProcedures" & vbNewLine & "'" & String(100, "-") & vbNewLine
    
    Dim tmp
    Dim arr:    arr = aCollection.Init(coll).ToArray
    Dim ll:     ll = LargestLength(coll)
    Dim ap As aProcedure
    For Each ap In Procedures
        Dim ks
        ks = ap.KindAsString
        Select Case ks
        Case "Get"
            tmp = "'" & Space(4) & "|> " & ks & " | --- " & ap.Name
        Case "Let", "Set"
            tmp = "'" & Space(4) & "|> --- | " & ks & " " & ap.Name
        Case "Sub"
            tmp = "'" & Space(4) & "|> Sub " & String(5, "-") & " " & ap.Name
        Case "Function"
            tmp = "'" & Space(4) & "|> Fun " & String(5, "-") & " " & ap.Name
        End Select
        tmp = tmp & Space(4 + (ll - Len(ap.Name))) & ":"
        
        Dim Desc As String, index As Long
        index = ap.LineLike("*'@Description*", Body_Code)
        If index > 0 Then
            Desc = ap.Parent.CodeModule.Lines(index, 1)
            tmp = tmp & " " & Desc
        End If
        txt = txt & tmp & vbNewLine
    Next
    txt = txt & "'" & String(100, "-") & vbNewLine
    txt = txt & "'@EndFolder ListOfProcedures" & vbNewLine
    FolderDelete "ListOfProcedures"
    oModule.CodeModule.InsertLines 1, txt
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 09:21    Alex                (aModule.cls > ListProceduresPublic)

Public Sub ListProceduresPublic()
'@LastModified 2308220921
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aCollection
    Dim coll        As New Collection
    Set coll = PublicProceduresNames
    If coll.Count = 0 Then Exit Sub
    
    Dim txt         As String
    
    txt = "'@FOLDER ListOfProcedures" & vbNewLine & "'" & String(100, "-") & vbNewLine
    
    Dim tmp
    Dim arr:    arr = aCollection.Init(coll).ToArray
    Dim ll:     ll = LargestLength(coll)
    Dim ap As aProcedure
    For Each ap In Procedures
        Dim ks
        ks = ap.KindAsString
        Select Case ks
        Case "Get"
            tmp = "'" & Space(4) & "|> " & ks & " | --- " & ap.Name
        Case "Let", "Set"
            tmp = "'" & Space(4) & "|> --- | " & ks & " " & ap.Name
        Case "Sub"
            tmp = "'" & Space(4) & "|> Sub " & String(5, "-") & " " & ap.Name
        Case "Function"
            tmp = "'" & Space(4) & "|> Fun " & String(5, "-") & " " & ap.Name
        End Select
        tmp = tmp & Space(4 + (ll - Len(ap.Name))) & ":"
        
        Dim Desc As String, index As Long
        index = ap.LineLike("*'@Description*", Body_Code)
        If index > 0 Then
            Desc = ap.Parent.CodeModule.Lines(index, 1)
            tmp = tmp & " " & Desc
        End If
        txt = txt & tmp & vbNewLine
    Next
    txt = txt & "'" & String(100, "-") & vbNewLine
    txt = txt & "'@EndFolder ListOfProcedures" & vbNewLine
    FolderDelete "ListOfProcedures"
    oModule.CodeModule.InsertLines 1, txt
End Sub

Public Sub FolderDelete(FolderName As String)
'@LastModified 2310251309
    '@AssignedModule aModule
'@INCLUDE DECLARATION oModule
'@INCLUDE CLASS aModule
    Dim startLine As Long, endLine As Long
    Dim sLine       As String
    Dim i           As Long
    For i = 1 To oModule.CodeModule.CountOfLines
        sLine = Trim(oModule.CodeModule.Lines(i, 1))
        If sLine = "'@FOLDER " & FolderName Then startLine = i
        If startLine <> 0 Then
            If sLine = "'@EndFolder " & FolderName Then endLine = i
        End If
        If startLine <> 0 And endLine <> 0 Then Exit For
    Next
'
    If startLine = 0 Or endLine = 0 Then Exit Sub
    oModule.CodeModule.DeleteLines startLine, endLine - startLine + 1
End Sub

Public Sub ProcedureSortByKind()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE ArrayQuickSort
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    If oModule.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim coll        As New Collection: Set coll = Procedures
    Dim var         As Variant: ReDim var(1 To coll.Count)
    Dim p           As aProcedure
    Dim startLine   As Long
    Dim i           As Long
    For i = 1 To coll.Count
        Set p = coll(i)
        var(i) = p.Name
        If i = 1 Then startLine = p.Line_First
    Next
    ArrayQuickSort var
    Dim h           As String: h = Header
    Dim S           As String
    Dim TheSubs As String, TheFunctions As String
    For i = LBound(var) To UBound(var)
        Set p = aProcedure.Init(Parent, oModule, CStr(var(i)))
        S = p.Code_All
        If p.KindAsString = "Sub" Then
            TheSubs = TheSubs & IIf(TheSubs <> "", vbNewLine, "") & S
        Else
            TheFunctions = TheFunctions & IIf(TheFunctions <> "", vbNewLine, "") & S
        End If
    Next i
    CodeRemove
    oModule.CodeModule.AddFromString h & vbLf & TheSubs & vbLf & TheFunctions

End Sub
Public Sub ProcedureSortByScope()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE ArrayQuickSort
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    If oModule.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim coll        As New Collection: Set coll = Procedures
    Dim var         As Variant: ReDim var(1 To coll.Count)
    Dim p           As aProcedure
    Dim startLine   As Long
    Dim i           As Long
    For i = 1 To coll.Count
        Set p = coll(i)
        var(i) = p.Name
        If i = 1 Then startLine = p.Line_First
    Next
    ArrayQuickSort var
    Dim h           As String: h = Header
    Dim S           As String
    Dim ThePublic As String, ThePrivate As String
    For i = LBound(var) To UBound(var)
        Set p = aProcedure.Init(Parent, oModule, CStr(var(i)))
        S = p.Code_All
        If p.Scope = "Public" Then
            ThePublic = ThePublic & IIf(ThePublic <> "", vbNewLine, "") & S
        Else
            ThePrivate = ThePrivate & IIf(ThePublic <> "", vbNewLine, "") & S
        End If
    Next i
    CodeRemove
    oModule.CodeModule.AddFromString h & vbLf & ThePrivate & vbLf & ThePublic
End Sub

Public Sub ProcedureSortAZ()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE ArrayQuickSort
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    If oModule.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim coll        As New Collection: Set coll = Procedures
    Dim var         As Variant: ReDim var(1 To coll.Count)
    Dim p           As aProcedure
    Dim startLine   As Long
    Dim i           As Long
    For i = 1 To coll.Count
        Set p = coll(i)
        var(i) = p.Name
        If i = 1 Then startLine = p.Line_First
    Next
    ArrayQuickSort var
    Dim S           As String: S = Header
    For i = LBound(var) To UBound(var)
        Set p = aProcedure.Init(Parent, oModule, CStr(var(i)))
        S = S & IIf(S <> "", vbNewLine, "") & p.Code_All
    Next i
    CodeRemove
    oModule.CodeModule.AddFromString S
End Sub

Public Function Component() As VBComponent
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    Set Component = oModule
End Function

Public Function LineLike( _
        this As String, _
        IgnoreIndentation As Boolean, _
        IgnoreCase As Boolean) As Long
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    If oModule.CodeModule.CountOfLines = 0 Then Exit Function
    If IgnoreCase Then this = UCase(this)
    Dim sLine       As String
    Dim i           As Long
    For i = 1 To oModule.CodeModule.CountOfLines
        sLine = oModule.CodeModule.Lines(i, 1)
        If IgnoreIndentation Then sLine = Trim(sLine)
        If IgnoreCase Then sLine = UCase(sLine)
        If sLine Like this Then
            LineLike = i
            Exit Function
        End If
    Next
End Function

Public Function LinesLike( _
        this As String, _
        IgnoreIndentation As Boolean, _
        IgnoreCase As Boolean) As Collection
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    If oModule.CodeModule.CountOfLines = 0 Then Exit Function
    If IgnoreCase Then this = UCase(this)
    Dim coll        As New Collection
    Dim sLine       As String
    Dim i           As Long
    For i = 1 To oModule.CodeModule.CountOfLines
        sLine = oModule.CodeModule.Lines(i, 1)
        If IgnoreIndentation Then sLine = Trim(sLine)
        If IgnoreCase Then sLine = UCase(sLine)
        If sLine Like this Then
            coll.Add i
        End If
    Next
    Set LinesLike = coll

End Function


'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 13:56    Alex                (aModule.cls > RemoveLinesLike)

Public Sub RemoveLinesLike(this As String)
'@LastModified 2308221356
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.RemoveLinesLike this
    Next
End Sub

Public Sub ProcedureScopePrivate()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.ScopePrivate
    Next
End Sub

Public Sub ProcedureScopePublic()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.ScopePublic
    Next
End Sub
Sub EnableDebugPrint()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.Enable_DebugPrint
    Next
End Sub
Sub DisableDebugPrint()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.Disable_DebugPrint
    Next
End Sub
Public Sub EnableStop()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.Enable_Stop
    Next
End Sub
Public Sub DisableStop()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.Disable_Stop
    Next
End Sub

Public Sub ProcedureFoldDeclarations()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.FoldDeclaration
    Next
End Sub


Public Function Procedures() As Collection  'ExcludeUnderscore as Boolean '@TODO
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim lineNum     As Long
    Dim ProcedureName As String
    Dim ProcKind    As Long
    Dim coll        As New Collection
    Dim i           As Long
    Dim cProcedure  As aProcedure
    With oModule.CodeModule

        lineNum = .CountOfDeclarationLines + 1
        Do Until lineNum >= .CountOfLines
            On Error Resume Next
            For i = 0 To 4
                ProcKind = i
                ProcedureName = .ProcOfLine(lineNum, ProcKind)
                If ProcedureName <> vbNullString Then Exit For
            Next
            On Error GoTo 0
            'Err.clear
            If ProcedureName = vbNullString Then Exit Function

            'Possible Error:
            '---------------
            ' Procedures with an underscore _ which is used in events
            ' and Private Procedures
            ' may have the same name in different components

            If InStr(1, ProcedureName, "_") = 0 Then    '>0 and ExcludeUnderscore then goto SKIP

                Set cProcedure = New aProcedure
                cProcedure.Init oWorkbook, oModule, ProcedureName, ProcKind
                On Error Resume Next
                coll.Add cProcedure    ', cProcedure.Name
                '                    coll.Add ProcedureName, ProcedureName
                '                    If Err.Number > 0 Then Debug.Print ProcedureName & " exists multiple times, added only once, may lead to error"
                '                    Err.clear
                On Error GoTo 0
                lineNum = .procStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
            End If
            lineNum = .procStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
            ProcedureName = vbNullString
        Loop
    End With
    Set Procedures = coll
End Function

Function ProceduresNames() As Collection
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    Dim coll        As New Collection
    Dim ProcedureName As String
    For Each ap In Procedures
        ProcedureName = ap.Name
        On Error Resume Next
        coll.Add ProcedureName, ProcedureName
        On Error GoTo 0
        If Err.Number > 0 Then Debug.Print ProcedureName & " exists multiple times, added only once, may lead to error"
        Err.clear
    Next
    Set ProceduresNames = coll
End Function
Function PublicProceduresNames() As Collection
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim ap          As aProcedure
    Dim coll        As New Collection
    Dim ProcedureName As String
    For Each ap In Procedures
        ProcedureName = ap.Name
        On Error Resume Next
        If ap.Scope = "Public" Then coll.Add ProcedureName, ProcedureName
        On Error GoTo 0
        If Err.Number > 0 Then Debug.Print ProcedureName & " exists multiple times, added only once, may lead to error"
        Err.clear
    Next
    Set PublicProceduresNames = coll
End Function
Sub UpdateProcedures()
    '@AssignedModule aModule
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aModule
    Dim Procedure   As aProcedure
    For Each Procedure In Procedures
        Procedure.Update
    Next
End Sub

Sub CommentsToOwnLine()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE CommentsMoveToOwnLine
    '@INCLUDE CLASS aModule
    Dim S           As String
    With oModule.CodeModule
        If .CountOfLines = 0 Then Exit Sub
        S = CommentsMoveToOwnLine(Code)
        CodeRemove
        .AddFromString S
    End With
End Sub

Public Sub Indent()
    '@AssignedModule aModule
    '@INCLUDE PROCEDURE IsBlockEnd
    '@INCLUDE PROCEDURE IsBlockStart
    '@INCLUDE CLASS aModule
    If oModule.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim LineStart As Long, LineEnd As Long
    LineStart = 1
    LineEnd = oModule.CodeModule.CountOfLines
    Dim nIndent     As Long
    Dim S           As String
    Dim i           As Long
    For i = LineStart To LineEnd
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

Sub Activate()
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    With Application.VBE.MainWindow
        .Visible = True
        .WindowState = vbext_ws_Maximize
    End With
    With oModule.CodeModule.CodePane
        .Show
        .Window.Visible = True
        .Window.WindowState = vbext_ws_Maximize
        .Window.SetFocus
        .SetSelection 1, 1, 1, 1
    End With
End Sub

Function Code() As String
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    With oModule.CodeModule
        If .CountOfLines = 0 Then Code = "": Exit Function
        Code = .Lines(1, .CountOfLines)
    End With
End Function

Public Function body() As String
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    With oModule.CodeModule
        If .CountOfLines - .CountOfDeclarationLines = 0 Then Exit Function
        body = .Lines(.CountOfDeclarationLines + 1, .CountOfLines - .CountOfDeclarationLines)
    End With
End Function

Function Contains( _
        this As String, _
        WholeWord As Boolean, _
        MatchCase As Boolean, _
        PatternSearch As Boolean)
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule

    If oModule.CodeModule.CountOfLines = 0 Then Exit Function
    Contains = oModule.CodeModule.Find( _
            this, _
            1, _
            1, _
            oModule.CodeModule.CountOfLines, _
            10000, _
            WholeWord, _
            MatchCase, _
            PatternSearch)
End Function

Function Header() As String
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    If oModule.CodeModule.CountOfDeclarationLines = 0 Then Exit Function
    With oModule.CodeModule
        Header = .Lines(1, .CountOfDeclarationLines)
    End With
End Function

Function HeaderContains(this As String) As Boolean
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    HeaderContains = (InStr(1, Header, this) > 0)
End Function

Function Ignore() As Boolean
    '@AssignedModule aModule
    '@INCLUDE CLASS aModule
    Ignore = InStr(1, Header, "@IGNORE", vbTextCompare) > 0
End Function

Public Property Get Parent() As Workbook
    Dim S           As String
    S = oModule.Collection.Parent.fileName
    S = Right(S, Len(S) - InStrRev(S, "\"))
    Set Parent = Workbooks(S)
End Property

Function Name() As String
    Name = oModule.Name
    If oModule.Type = vbext_ct_Document And oModule.Name <> "ThisWorkbook" Then
        On Error GoTo HELL
        Name = GetSheetByCodeName(Parent, oModule.Name).Name
        On Error GoTo 0
    End If
    Exit Function
HELL:
    Debug.Print "Probably the file is in another language"
    Name = "ThisWorkbook"
End Function

Function TypeToLong() As vbext_ComponentType
    TypeToLong = oModule.Type
End Function

Function TypeToString() As String
    Select Case oModule.Type
        Case vbext_ct_ActiveXDesigner
            TypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            TypeToString = "Class"
        Case vbext_ct_Document
            TypeToString = "Document"
        Case vbext_ct_MSForm
            TypeToString = "UserForm"
        Case vbext_ct_StdModule
            TypeToString = "Module"
        Case Else
            TypeToString = "Unknown Type: " & CStr(oModule.Type)
    End Select
End Function

Sub ExportProcedures(targetPath As String)
    If targetPath = "" Then Exit Sub    '<- FOR when called with FolderPicker
    targetPath = Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    Dim ap          As aProcedure
    Dim coll        As New Collection: Set coll = Procedures
    If coll.Count > 0 Then
        For Each ap In coll
            TxtOverwrite targetPath & ap.Name & ".txt", ap.Code_All
        Next
    End If
End Sub

Sub Export(targetPath As String)
    If targetPath = "" Then Exit Sub    '<- FOR when called with FolderPicker
    targetPath = Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    oModule.Export targetPath & Name & Extension
End Sub


Public Sub Rename(NewName As String)
    If Not ModuleExists(NewName, oWorkbook) Then oModule.Name = NewName
End Sub

Function Extension() As String
    Dim ext         As String
    Select Case oModule.Type
        Case vbext_ct_MSForm
            ext = ".frm"
        Case vbext_ct_ClassModule, vbext_ct_Document
            ext = ".cls"
            '        Case vbext_ct_Document
            '            If oModule.Name = "ThisWorkbook" Then
            '                Extension = ".WBcls"
            '            Else
            '                Extension = ".DOCcls"
            '            End If
        Case vbext_ct_StdModule
            ext = ".bas"
        Case Else
            ext = ".txt"
    End Select
    Extension = ext
End Function

Public Function Copy( _
        TargetWorkbook As Workbook, _
        OverwriteExisting As Boolean) As Boolean

    If Name = "ThisWorkbook" Then Exit Function
    If oModule.Type = vbext_ct_Document Then Exit Function
    If Parent.Name = TargetWorkbook.Name Then Exit Function
    Dim TempModule  As VBIDE.VBComponent

    If ModuleExists(Name, TargetWorkbook) Then
        If OverwriteExisting = True Then
            With TargetWorkbook.VBProject
                .VBComponents.Remove .VBComponents(Name)
            End With
        Else
            Exit Function
        End If
    End If

    Dim fname       As String
    fname = Environ("Temp") & "\" & Name & Extension
    oModule.Export fileName:=fname

    TargetWorkbook.VBProject.VBComponents.Import fileName:=fname
    Kill fname
    Copy = True
End Function

Sub Delete()
    If Ignore Then Exit Sub
    Application.DisplayAlerts = False
    If oModule.Type = vbext_ct_Document Then
        If oModule.Name = "ThisWorkbook" Then
            oModule.CodeModule.DeleteLines 1, oModule.CodeModule.CountOfLines
        Else
            If Parent.Sheets.Count = 1 Then Parent.Sheets.Add
            GetSheetByCodeName(oWorkbook, Name).Delete
        End If
    Else
        Parent.VBProject.VBComponents.Remove oModule
    End If
    Application.DisplayAlerts = True
End Sub

Sub CodeRemove()
    If Ignore Then Exit Sub
    If oModule.CodeModule.CountOfLines = 0 Then Exit Sub
    oModule.CodeModule.DeleteLines 1, oModule.CodeModule.CountOfLines    '+ 1
End Sub

Sub CodeMove(ToModule As VBComponent)
    If Ignore Then Exit Sub
    With ToModule.CodeModule
        .InsertLines 1, Header
        .InsertLines .CountOfLines + 1, body
    End With
    CodeRemove
End Sub

Sub DeleteIfEmpty()
    If Ignore Then Exit Sub
    If Name = "ThisWorkbook" Then Exit Sub
    If oModule.CodeModule.CountOfLines = 0 Then Delete
End Sub

Public Sub RemoveEmptyLinesButLeaveOne()
    Dim Code        As String
    Dim currentline As String, nextline As String
    Dim i           As Long
    With oModule.CodeModule
        If .CountOfLines > 0 Then
            For i = .CountOfLines - 1 To 1 Step -1
                currentline = Trim(.Lines(i, 1))
                nextline = Trim(.Lines(i + 1, 1))
                If Len(currentline) = 0 Then
                    If Len(nextline) = 0 Then
                        .DeleteLines i + 1, 1
                    End If
                End If
            Next
        End If
    End With
End Sub
Public Sub RemoveEmptyLines()
    Dim Code        As String
    Dim currentline As String, nextline As String
    Dim i           As Long
    With oModule.CodeModule
        If .CountOfLines > 0 Then
            For i = .CountOfLines - 1 To 1 Step -1
                currentline = Trim(.Lines(i, 1))
                If Len(currentline) = 0 Then
                    .DeleteLines i, 1
                End If
            Next
        End If
    End With
End Sub

Public Sub CommentsRemove()
    Dim ap          As aProcedure
    For Each ap In Procedures
        ap.CommentsRemove False, Body_Code
    Next ap
End Sub
