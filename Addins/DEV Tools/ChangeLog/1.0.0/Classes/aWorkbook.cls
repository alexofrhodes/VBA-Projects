VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aWorkbook
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 04-05-2023 12:19    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private oWorkbook   As Workbook

Public Function Init(TargetWorkbook As Workbook) As aWorkbook
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    Set oWorkbook = TargetWorkbook
    Set Init = Me
End Function

Public Function item() As Workbook
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    Set item = oWorkbook
End Function

Sub ReferencesList()
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE DECLARATION GUID
    Dim ws          As Worksheet
    Set ws = ThisWorkbook.Sheets("References")
    ws.Cells(1, 1).value = "Reference Description"
    ws.Cells(1, 2).value = "GUID"
    ws.Cells(1, 3).value = "Path"
    ws.Cells(1, 4).value = "Version"
    Dim myRef       As Reference
    Dim refs        As VBIDE.REFERENCES
    Set refs = oWorkbook.VBProject.REFERENCES
    Dim i           As Long
    i = 2
    For Each myRef In refs
        ws.Cells(i, 1) = IIf(myRef.Description <> "", myRef.Description, myRef.Name)
        ws.Cells(i, 2) = myRef.GUID
        ws.Cells(i, 3) = myRef.fullPath
        ws.Cells(i, 4) = myRef.major & "." & myRef.minor
        i = i + 1
    Next myRef
End Sub

Sub AddReferenceFromFile(FilePath As String)
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    Dim oRefs       As REFERENCES: Set oRefs = oWorkbook.VBProject.REFERENCES
    oRefs.AddFromFile FilePath
End Sub

Sub AddReferenceFromGUID(GUID As String, major As String, minor As String)
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE DECLARATION GUID
    Dim oRefs       As REFERENCES: Set oRefs = oWorkbook.VBProject.REFERENCES
    oRefs.AddFromGuid GUID, major, minor
End Sub

Sub RemoveReferenceByName(RefName As String)
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    Dim oRef        As Reference
    Dim oRefs       As REFERENCES: Set oRefs = oWorkbook.VBProject.REFERENCES
    For Each oRef In oRefs
        If oRef.Name = RefName Then
            oRefs.Remove oRef
            Exit For
        End If
    Next oRef
End Sub

Sub RemoveReferenceByDescription(RefDescription As String)
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    Dim oRef        As Reference
    Dim oRefs       As REFERENCES: Set oRefs = oWorkbook.VBProject.REFERENCES
    For Each oRef In oRefs
        If oRef.Description = RefDescription Then
            oRefs.Remove oRef
            Exit For
        End If
    Next oRef
End Sub

Sub RemoveReferenceByGUID(refGUID As String)
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE DECLARATION GUID
    Dim oRefs       As REFERENCES: Set oRefs = oWorkbook.VBProject.REFERENCES
    Dim oRef        As Reference
    For Each oRef In oRefs
        '        debug.print oRef.Name
        If oRef.GUID = refGUID Then
            oRefs.Remove oRef
            Exit For
        End If
    Next oRef
End Sub

Public Sub ExportReferences(FilePath As String)
    '@AssignedModule aWorkbook
    '@INCLUDE PROCEDURE TxtOverwrite
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE DECLARATION GUID
    Dim result      As String
    Dim oRef        As Reference
    For Each oRef In oWorkbook.VBProject.REFERENCES
        result = result & IIf(result <> "", vbNewLine, "") & oRef.Name & vbTab & oRef.GUID & vbTab & oRef.major & vbTab & oRef.minor
    Next
    TxtOverwrite FilePath & NameClean & "_References.txt", result
End Sub


Public Sub ImportReferences(FilePath As String)
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    Dim FSO         As New Scripting.FileSystemObject
    With FSO.OpenTextFile(FilePath, ForReading, True)
        Dim line    As Long
        Do While Not .AtEndOfStream
            Dim values As Variant
            values = Split(.ReadLine, vbTab)
            On Error Resume Next
            oWorkbook.VBProject.REFERENCES.AddFromGuid values(1), values(2), values(3)
        Loop
    End With
End Sub

Public Sub ReferenceAdd_VBIDE()
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    On Error Resume Next
    ActiveWorkbook.VBProject.REFERENCES.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
End Sub

Sub ReferenceAdd_ScriptControl()
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    On Error Resume Next
    Application.VBE.ActiveVBProject.REFERENCES.AddFromGuid "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}", 1, 0
End Sub


Function TodoList()
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim out         As String
    Dim am          As aModule
    For Each am In aModules.Init(oWorkbook).items
        If InStr(1, am.Code, "@TODO", vbTextCompare) > 0 Then
            out = out & IIf(out <> "", vbNewLine, "") & am.TodoList
        End If
    Next am
    TodoList = out
End Function

Function ProceduresLike(this As String) As Collection
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aProcedure
    '@INCLUDE CLASS aWorkbook
    Dim Procedure   As aProcedure
    Dim out         As New Collection
    For Each Procedure In Procedures(True, False, False, False)
        If UCase(Procedure.Name) Like UCase(this) Then
            out.Add Procedure
        End If
    Next
    Set ProceduresLike = out
End Function


'* Modified   : Date and Time       Author              Description
'* Updated    : 18-08-2023 12:43    Alex                fixed line aModules.Init(oWorkbook).Items (aWorkbook.cls > Code)

Public Function Code() As String
    '@LastModified 2308181243
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim module      As aModule
    Dim txt
    For Each module In aModules.Init(oWorkbook).items
        If module.item.CodeModule.CountOfLines > 0 Then
            txt = txt & _
                    vbNewLine & _
                    "'" & String(10, "=") & " (" & module.TypeToString & ") " & module.Name & String(10, "=") & _
                    vbNewLine & _
                    module.Code
        End If
    Next
    Code = txt
End Function

Sub ModulesMerge(TargetModule As VBComponent, ParamArray Modules() As Variant)
    '@AssignedModule aWorkbook
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE CLASS aModule
    Dim element     As Variant
    Dim module      As VBComponent
    For Each element In Modules
        Set module = element
        If module.Type = vbext_ct_StdModule Then
            If module.Name <> TargetModule.Name Then
                aModule.Init(module).CodeMove TargetModule
            End If
        End If
    Next
End Sub


Public Property Get ProceduresNames( _
        includeModule As Boolean, _
        includeDocument As Boolean, _
        includeClass As Boolean, _
        includeForm As Boolean) As Collection

    Dim module      As VBComponent
    Dim ProcKind    As VBIDE.vbext_ProcKind
    Dim lineNum     As Long
    Dim coll        As New Collection
    Dim ProcedureName As String
    Dim i           As Long
    For Each module In oWorkbook.VBProject.VBComponents
        If Not includeModule And module.Type = vbext_ct_StdModule Then GoTo SKIP
        If Not includeClass And module.Type = vbext_ct_ClassModule Then GoTo SKIP
        If Not includeDocument And module.Type = vbext_ct_Document Then GoTo SKIP
        If Not includeForm And module.Type = vbext_ct_MSForm Then GoTo SKIP
        With module.CodeModule
            lineNum = .CountOfDeclarationLines + 1
            Do Until lineNum >= .CountOfLines

                For i = 0 To 4
                    ProcKind = i
                    ProcedureName = .ProcOfLine(lineNum, ProcKind)
                    If ProcedureName <> vbNullString Then

                        'Possible Error:
                        '---------------
                        ' Procedures with an underscore _ which is used in events
                        ' and Private Procedures
                        ' may have the same name in different components

                        If InStr(1, ProcedureName, "_") = 0 Then
                            ProcedureName = .ProcOfLine(lineNum, ProcKind)
                            On Error Resume Next
                            coll.Add ProcedureName, module.Name & "|" & ProcedureName & "|" & ProcKind
                            On Error GoTo 0
                        End If

                        Exit For

                    End If
                Next
                lineNum = .procStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
            Loop
        End With
SKIP:
    Next module

    Set ProceduresNames = coll
End Property

Public Property Get Procedures( _
                            includeModule As Boolean, _
                            includeDocument As Boolean, _
                            includeClass As Boolean, _
                            includeForm As Boolean) As Collection
    'as aProcedure objects
    Dim module      As VBComponent
    Dim ProcKind    As VBIDE.vbext_ProcKind
    Dim lineNum     As Long
    Dim coll        As New Collection
    Dim ProcedureName As String
    Dim cProcedure  As aProcedure
    Dim i           As Long
    For Each module In oWorkbook.VBProject.VBComponents
        If Not includeModule And module.Type = vbext_ct_StdModule Then GoTo SKIP
        If Not includeClass And module.Type = vbext_ct_ClassModule Then GoTo SKIP
        If Not includeDocument And module.Type = vbext_ct_Document Then GoTo SKIP
        If Not includeForm And module.Type = vbext_ct_MSForm Then GoTo SKIP
        With module.CodeModule
            lineNum = .CountOfDeclarationLines + 1
            Do Until lineNum >= .CountOfLines

                For i = 0 To 4
                    ProcKind = i
                    ProcedureName = .ProcOfLine(lineNum, ProcKind)
                    If ProcedureName <> vbNullString Then

                        'Possible Error:
                        '---------------
                        ' Procedures with an underscore _ which is used in events
                        ' and Private Procedures
                        ' may have the same name in different components

                        If InStr(1, ProcedureName, "_") = 0 Then
                            ProcedureName = .ProcOfLine(lineNum, ProcKind)
                            Set cProcedure = New aProcedure
                            cProcedure.Init oWorkbook, module, ProcedureName, ProcKind
                            On Error Resume Next
                            coll.Add cProcedure, cProcedure.Name & ProcKind
                            'coll.Add ProcedureName, ProcedureName
                            'If Err.Number > 0 Then Debug.Print ProcedureName & " exists multiple times, added only once, may lead to error"
                            'Err.clear
                            On Error GoTo 0
                        End If

                        Exit For

                    End If
                Next
                lineNum = .procStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
            Loop
        End With
SKIP:
    Next module

    Set Procedures = coll
End Property


Function ProceduresArray()
    Dim module      As VBComponent
    Dim ProcKind    As VBIDE.vbext_ProcKind
    Dim lineNum     As Long
    Dim coll        As New Collection
    Dim ProcedureName As String
    Dim cProcedure  As aProcedure
    Dim arr
    Dim i           As Long
    Dim BodyLine    As Long
    ReDim arr(1 To 4, 1 To 1)
    For Each module In oWorkbook.VBProject.VBComponents
        With module.CodeModule
            lineNum = .CountOfDeclarationLines + 1
            Do Until lineNum >= .CountOfLines
                If arr(UBound(arr, 1), 1) <> "" Then
                    ReDim Preserve arr(1 To 4, 1 To UBound(arr, 2) + 1)
                End If
                On Error Resume Next
                For i = 0 To 4
                    ProcKind = i
                    ProcedureName = .ProcOfLine(lineNum, ProcKind)
                    If ProcedureName <> vbNullString Then
                        BodyLine = module.CodeModule.ProcBodyLine(ProcedureName, ProcKind)
                        arr(1, UBound(arr, 2)) = Switch(module.Type = vbext_ct_StdModule, "Module", _
                                module.Type = vbext_ct_ClassModule, "Class", _
                                module.Type = 3, "Userform", _
                                module.Type = 100, "Document", _
                                module.Type = 11, "ActiveXDesigner")
                        arr(2, UBound(arr, 2)) = module.Name
                        arr(3, UBound(arr, 2)) = Switch(ProcKind = vbext_pk_Proc, IIf(module.CodeModule.Lines(BodyLine, 1) Like "*Sub " & ProcedureName & "*", "Sub", "Function"), _
                                ProcKind = vbext_pk_Let, "Let", _
                                ProcKind = vbext_pk_Set, "Set", _
                                ProcKind = vbext_pk_Get, "Get")
                        arr(4, UBound(arr, 2)) = ProcedureName
                        Exit For
                    End If
                Next
                On Error GoTo 0
                lineNum = .procStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
            Loop
        End With
SKIP:
    Next module
    ProceduresArray = WorksheetFunction.Transpose(arr)
End Function

Sub Export()
    uProjectManager.Show
End Sub

Sub Backup(targetFolder As String)
    oWorkbook.SaveCopyAs targetFolder & _
            Format(Now, "yyyy-mm-dd hh-nn") & " " & oWorkbook.Name
End Sub

Function Name()
    Name = oWorkbook.Name
End Function
Function Extension()
    Extension = Mid(Name, InStr(1, Name, "."))
End Function

Function NameClean()
    NameClean = Left(Name, InStrRev(Name, ".") - 1)
End Function

Public Sub ExportModules(targetPath As String)
    targetPath = VBA.Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    Dim ams As aModules: Set ams = aModules.Init(oWorkbook)
    ams.Export targetPath, CreateSubFolders:=True
End Sub
Public Sub ExportProcedures(targetPath As String)
    targetPath = VBA.Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    Dim ams As aModules: Set ams = aModules.Init(oWorkbook)
    ams.ExportProcedures targetPath
End Sub
Public Sub ExportXML(targetPath As String)
    targetPath = Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    Dim TmpFile     As String
    TmpFile = oWorkbook.path & "\temp_workbook_file" & Extension
    oWorkbook.SaveCopyAs TmpFile
    Dim c           As New clsEditOpenXML
    c.ExtractRibbonX TmpFile, targetPath & "customUI.xml"
    Kill TmpFile
    FolderDelete oWorkbook.path & "\Unzipped " & "temp_workbook_file" & Extension & ".zip"
    Set c = Nothing
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 25-10-2023 23:22    Alex                (aWorkbook.cls > ExportCodeUnified)

Public Sub ExportCodeUnified(targetPath As String)
'@LastModified 2310252322
    targetPath = Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath
    Dim result As String, tmp As String
    Dim am          As aModule
    For Each am In aModules.Init(oWorkbook).items
        tmp = "'" & am.Name & vbTab & am.TypeToString & vbNewLine & vbNewLine & am.Code
        result = IIf(result = "", tmp, result & vbNewLine & vbNewLine & tmp)
    Next
    TxtOverwrite targetPath & "#UnifiedProject.txt", result
End Sub

Function HasProject() As Boolean
    Dim WbProjComp  As Object
    On Error Resume Next
    Set WbProjComp = oWorkbook.VBProject.VBComponents
    HasProject = Not WbProjComp Is Nothing
End Function

Sub CreateLinkedTable_Sheet()
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = CreateOrSetSheet("Linked_Table_" & NameClean, ThisWorkbook)
    If TargetWorksheet.Range("A1") <> "" Then Exit Sub
    Dim var
    ReDim var(1 To 8)

    var(1) = "WORKBOOK"
    var(2) = "MODULE TYPE"
    var(3) = "MODULE NAME"
    var(4) = "PROCEDURE"
    var(5) = "LINKED PROCEDURES"
    var(6) = "LINKED CLASSES"
    var(7) = "LINKED USERFORMS"
    var(8) = "LINKED DECLARATIONS"

    With TargetWorksheet
        .Range("A1:H1").value = var
        With .rows(1).Cells.Font
            .Bold = True
            .Size = 14
        End With
    End With

End Sub

Sub CreateLinkedTable()

    StartTimer "tableOfProcedures"

    aWorkbook.Init(oWorkbook).DeclarationsTableCreate

    CreateLinkedTable_Sheet

    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Linked_Table_" & NameClean)
    On Error Resume Next
    TargetWorksheet.ListObjects(1).Unlist
    On Error GoTo 0
    TargetWorksheet.Range("A1").CurrentRegion.offset(1).clear

    Dim var
    ReDim var(1 To 8, 1 To 1)

    Dim counter     As Long
    counter = 1

    Dim module      As VBComponent
    Dim Procedures  As Collection

    Dim Procedure   As aProcedure
    For Each module In oWorkbook.VBProject.VBComponents
        Select Case module.Type
            Case vbext_ct_StdModule, vbext_ct_Document, vbext_ct_MSForm    '@TODO check if adding these is ok
                Set Procedures = aModule.Init(module).Procedures
                For Each Procedure In Procedures
                    ReDim Preserve var(1 To 8, 1 To counter)
                    var(1, counter) = oWorkbook.Name
                    var(2, counter) = aModule.Init(module).TypeToString
                    var(3, counter) = module.Name
                    var(4, counter) = Procedure.Name
                    var(5, counter) = aCollection.Init(Procedure.LinkedProcedures).ToString(vbNewLine)
                    var(6, counter) = aCollection.Init(Procedure.LinkedClasses).ToString(vbNewLine)
                    var(7, counter) = aCollection.Init(Procedure.LinkedUserforms).ToString(vbNewLine)
                    var(8, counter) = aCollection.Init(Procedure.LinkedDeclarations).ToString(vbNewLine)

                    counter = counter + 1
                Next
        End Select
    Next

    var = WorksheetFunction.Transpose(var)
    ArrayToRange2D var, TargetWorksheet.Range("A2")

    LinkedProceduresTableSort
    TargetWorksheet.ListObjects.Add xlSrcRange, TargetWorksheet.Range("A1").CurrentRegion, , xlYes
    TargetWorksheet.Cells.VerticalAlignment = xlVAlignTop

    EndTimer

End Sub

Private Sub LinkedProceduresTableSort()
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Worksheets("Linked_Table_" & NameClean)
    Dim sort1       As String
    sort1 = "A1"
    Dim sort2       As String
    sort2 = "C1"
    Dim sort3       As String
    sort3 = "D1"

    With TargetWorksheet.Sort
        .SortFields.clear
        .SortFields.Add Key:=TargetWorksheet.Range(sort1), Order:=xlAscending
        If Not sort2 = "" Then
            .SortFields.Add Key:=TargetWorksheet.Range(sort2), Order:=xlAscending
        End If
        If Not sort3 = "" Then
            .SortFields.Add Key:=TargetWorksheet.Range(sort3), Order:=xlAscending
        End If

        .SetRange TargetWorksheet.Range("A1").CurrentRegion
        '        .header = xlYes
        .Apply
    End With
End Sub

'Sub LinkedProceduresTableFilter()
'    Dim TargetWorksheet As Worksheet: Set TargetWorksheet = ThisWorkbook.SHEETS("LinkedProcedures_Table")
'    TargetWorksheet.rows.Hidden = False
'    On Error Resume Next
'    TargetWorksheet.ShowAllData
'    On Error GoTo 0
'    TargetWorksheet.Range("A6").CurrentRegion.AdvancedFilter _
     '        action:=xlFilterInPlace, _
     '        criteriaRange:=TargetWorksheet.Range("A1").CurrentRegion
'    LinkedProceduresTableSort
'End Sub


Sub ExportDeclarations(targetPath As String)
    targetPath = Replace(targetPath & "\", "\\", "\")
    FoldersCreate targetPath

    Dim DeclarationArray As Variant
    DeclarationArray = aCollection.CollectionsToArray2D(getDeclarations)
    If TypeName(DeclarationArray) <> "Empty" Then
        TxtOverwrite targetPath & "Declarations.txt", ArrayToString(DeclarationArray)
    End If
End Sub

Private Function DeclarationsWorksheetCreate() As Boolean
    If WorksheetExists("Declarations_Table", ThisWorkbook) Then Exit Function
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets.Add
    With TargetWorksheet
        .Name = "Declarations_Table"
        .Cells.VerticalAlignment = xlVAlignTop
        .Range("A1:F1").value = Split("SCOPE,TYPE,NAME,CODE,MODULE TYPE,MODULE NAME", ",")
        .rows(1).Cells.Font.Bold = True
        .rows(1).Cells.Font.Size = 14
    End With
End Function

Public Sub DeclarationsTableCreate()

    DeclarationsWorksheetCreate

    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    'if sheet was created within the hour, you probably don't have new declarations
    If Format(Now, "YYMMDDHHNN") - TargetWorksheet.Range("Z1").value < 60 Then Exit Sub

    TargetWorksheet.Range("A2").CurrentRegion.offset(1).clear
    ArrayToRange2D aCollection.CollectionsToArray2D( _
            getDeclarations( _
            includeScope:=True, _
            includeType:=True, _
            includeKeywords:=True, _
            includeDeclarations:=True, _
            includeComponentName:=True, _
            includeComponentType:=True)), _
            TargetWorksheet.Range("A2")

    TargetWorksheet.Range("Z1").value = Format(Now, "YYMMDDHHNN")

    DeclarationsTableSort
End Sub

Private Sub DeclarationsTableSort()
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Worksheets("Declarations_Table")

    Dim sort1       As String: sort1 = "B1"
    Dim sort2       As String: sort2 = "C1"
    Dim sort3       As String    ': sort3 = "D1"

    With TargetWorksheet.Sort
        .SortFields.clear
        .SortFields.Add Key:=TargetWorksheet.Range(sort1), Order:=xlAscending

        If Not sort2 = "" Then
            .SortFields.Add Key:=TargetWorksheet.Range(sort2), Order:=xlAscending
        End If
        If Not sort3 = "" Then
            .SortFields.Add Key:=TargetWorksheet.Range(sort3), Order:=xlAscending
        End If

        .SetRange TargetWorksheet.Range("A1").CurrentRegion
        .Header = xlYes
        .Apply
    End With
End Sub


Public Function DeclarationsTableKeywords() As Collection
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

Public Function getDeclarations( _
        Optional includeScope As Boolean, _
        Optional includeType As Boolean, _
        Optional includeKeywords As Boolean, _
        Optional includeDeclarations As Boolean, _
        Optional includeComponentName As Boolean, _
        Optional includeComponentType As Boolean) As Collection

    Dim ComponentCollection As New Collection
    Dim ComponentTypecollection As New Collection
    Dim DeclarationsCollection As New Collection
    Dim KeywordsCollection As New Collection
    Dim Output      As New Collection
    Dim ScopeCollection As New Collection
    Dim TypeCollection As New Collection

    Dim element     As Variant
    Dim OriginalDeclarations As Variant
    Dim str         As Variant

    Dim tmp         As String
    Dim helper      As String
    Dim i           As Long

    Dim module      As VBComponent
    For Each module In oWorkbook.VBProject.VBComponents
        '        If Module.Type = vbext_ct_StdModule Or Module.Type = vbext_ct_MSForm Then
        If module.CodeModule.CountOfDeclarationLines > 0 Then
            str = module.CodeModule.Lines(1, module.CodeModule.CountOfDeclarationLines)
            str = Replace(str, "_" & vbNewLine, "")
            OriginalDeclarations = str
            tmp = str
            Do While InStr(1, str, "End Type") > 0
                tmp = Mid(str, InStr(1, str, "Type "), InStr(1, str, "End Type") - InStr(1, str, "Type ") + 8)
                str = Replace(str, tmp, Split(tmp, vbNewLine)(0))
            Loop
            Do While InStr(1, str, "End Enum") > 0
                tmp = Mid(str, InStr(1, str, "Enum "), InStr(1, str, "End Enum") - InStr(1, str, "Enum ") + 8)
                str = Replace(str, tmp, Split(tmp, vbNewLine)(0))
            Loop
            Do While InStr(1, str, "  ") > 0
                str = Replace(str, "  ", " ")
            Loop

            str = Split(str, vbNewLine)
            tmp = OriginalDeclarations

            For Each element In str
                If Len(CStr(element)) > 0 And Not Trim(CStr(element)) Like "'*" And Not Trim(CStr(element)) Like "Rem*" Then
                    If RegexTest(CStr(element), "\b ?Enum \b") Then
                        KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Enum")
                        DeclarationsCollection.Add DeclarationsKeywordSubstring(tmp, , "Enum " & KeywordsCollection.item(KeywordsCollection.Count), "End Enum", , , True)
                        TypeCollection.Add "Enum"
                        ComponentCollection.Add module.Name
                        ComponentTypecollection.Add aModule.Init(module).TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf RegexTest(CStr(element), "\b ?Type \b") Then
                        KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Type")
                        DeclarationsCollection.Add DeclarationsKeywordSubstring(tmp, , "Type " & KeywordsCollection.item(KeywordsCollection.Count), "End Type", , , True)
                        TypeCollection.Add "Type"
                        ComponentCollection.Add module.Name
                        ComponentTypecollection.Add aModule.Init(module).TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf InStr(1, CStr(element), "Const ", vbTextCompare) > 0 Then
                        KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Const")
                        DeclarationsCollection.Add CStr(element)
                        TypeCollection.Add "Const"
                        ComponentCollection.Add module.Name
                        ComponentTypecollection.Add aModule.Init(module).TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf RegexTest(CStr(element), "\b ?Sub \b") Then
                        KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Sub")
                        DeclarationsCollection.Add CStr(element)
                        TypeCollection.Add "Sub"
                        ComponentCollection.Add module.Name
                        ComponentTypecollection.Add aModule.Init(module).TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf RegexTest(CStr(element), "\b ?Function \b") Then
                        KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Function")
                        DeclarationsCollection.Add CStr(element)
                        TypeCollection.Add "Function"
                        ComponentCollection.Add module.Name
                        ComponentTypecollection.Add aModule.Init(module).TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf element Like "*(*) As *" Then
                        helper = Left(element, InStr(1, CStr(element), "(") - 1)
                        helper = Mid(helper, InStrRev(helper, " ") + 1)
                        KeywordsCollection.Add helper
                        DeclarationsCollection.Add CStr(element)
                        TypeCollection.Add "Other"
                        ComponentCollection.Add module.Name
                        ComponentTypecollection.Add aModule.Init(module).TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf element Like "* As *" Then
                        KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", , "As")
                        DeclarationsCollection.Add CStr(element)
                        TypeCollection.Add "Other"
                        ComponentCollection.Add module.Name
                        ComponentTypecollection.Add aModule.Init(module).TypeToString
                        ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                    Else
                    End If
                End If
            Next element
        End If
        '        End If
    Next module

    If includeScope = True Then Output.Add ScopeCollection
    If includeType = True Then Output.Add TypeCollection
    If includeKeywords = True Then Output.Add KeywordsCollection
    If includeDeclarations = True Then Output.Add DeclarationsCollection
    If includeComponentType = True Then Output.Add ComponentTypecollection
    If includeComponentName = True Then Output.Add ComponentCollection

    Set getDeclarations = Output
End Function

Private Function DeclarationsKeywordSubstring( _
        str As Variant, Optional delim As String _
        , Optional afterWord As String _
        , Optional beforeWord As String _
        , Optional counter As Integer _
        , Optional outer As Boolean _
        , Optional includeWords As Boolean) As String

    Dim i           As Long
    If afterWord = "" And beforeWord = "" And counter = 0 Then
        MsgBox ("Pass at least 1 parameter betweenn -AfterWord- , -BeforeWord- , -counter-")
        Exit Function
    End If
    If TypeName(str) = "String" Then
        If delim <> "" Then
            str = Split(str, delim)
            If UBound(str) <> 0 Then
                If afterWord = "" And beforeWord = "" And counter <> 0 Then
                    If counter - 1 <= UBound(str) Then
                        DeclarationsKeywordSubstring = str(counter - 1)
                        Exit Function
                    End If
                End If
                For i = LBound(str) To UBound(str)
                    If afterWord <> "" And beforeWord = "" Then
                        If i <> 0 Then
                            If str(i - 1) = afterWord Or str(i - 1) = "#" & afterWord Then
                                DeclarationsKeywordSubstring = str(i)
                                Exit Function
                            End If
                        End If
                    ElseIf afterWord = "" And beforeWord <> "" Then
                        If i <> UBound(str) Then
                            If str(i + 1) = beforeWord Or str(i + 1) = "#" & beforeWord Then
                                DeclarationsKeywordSubstring = str(i)
                                Exit Function
                            End If
                        End If
                    ElseIf afterWord <> "" And beforeWord <> "" Then
                        If i <> 0 And i <> UBound(str) Then
                            If (str(i - 1) = afterWord Or str(i - 1) = "#" & afterWord) And (str(i + 1) = beforeWord Or str(i + 1) = "#" & beforeWord) Then
                                DeclarationsKeywordSubstring = str(i)
                                Exit Function
                            End If
                        End If
                    End If
                Next i
            End If
        Else
            If InStr(1, str, afterWord) > 0 And InStr(1, str, beforeWord) > 0 Then
                If includeWords = False Then
                    DeclarationsKeywordSubstring = Mid(str, InStr(1, str, afterWord) + Len(afterWord))
                Else
                    DeclarationsKeywordSubstring = Mid(str, InStr(1, str, afterWord))
                End If
                If outer = True Then
                    If includeWords = False Then
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStrRev(DeclarationsKeywordSubstring, beforeWord) - 1)
                    Else
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStrRev(DeclarationsKeywordSubstring, beforeWord) + Len(beforeWord) - 1)
                    End If
                Else
                    If includeWords = False Then
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStr(1, DeclarationsKeywordSubstring, beforeWord) - 1)
                    Else
                        DeclarationsKeywordSubstring = Left(DeclarationsKeywordSubstring, InStr(1, DeclarationsKeywordSubstring, beforeWord) + Len(beforeWord) - 1)
                    End If
                End If
                Exit Function
            End If
        End If
    Else
        '
    End If
    DeclarationsKeywordSubstring = vbNullString
End Function

Public Sub DeclarationsTableExport()
    Dim ws          As Worksheet
    Set ws = ThisWorkbook.Sheets("Declarations_Table")
    Dim codeName    As String
    Dim codeText    As String
    Dim cell        As Range
    Dim Lr          As Long
    Lr = getLastRow(ws)
    Dim codeCollection As Collection
    For Each cell In ws.Range("C2:C" & Lr)
        codeName = cell.TEXT
        Set codeCollection = New Collection
        On Error Resume Next
        codeCollection.Add cell.offset(0, 1).TEXT, cell.offset(0, 1).TEXT
        On Error GoTo 0
        Do While cell.offset(1, 0).TEXT = cell.TEXT
            Set cell = cell.offset(1, 0)
            On Error Resume Next
            codeCollection.Add cell.offset(0, 1).TEXT, cell.offset(0, 1).TEXT
            On Error GoTo 0
        Loop
        codeText = aCollection.Init(codeCollection).ToString(IIf(codeCollection.Count > 1, vbNewLine, ""))
        TxtOverwrite LOCAL_LIBRARY_DECLARATIONS & codeName & ".txt", codeText
    Next
End Sub

Function getDeclaredKeywordsOfWorkbook() As Variant
    getDeclaredKeywordsOfWorkbook = WorksheetFunction.Transpose(aCollection.CollectionsToArray2D(getDeclarations(, , True)))
End Function

Function getDeclaredEnumOfWorkbook() As String
    Dim c           As Variant
    Dim out         As String
    For Each c In aCollection.CollectionsToArray2D(getDeclarations(, , , True))
        If InStr(1, CStr(c), "Enum ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredEnumOfWorkbook = out
End Function

Function getDeclaredTypeOfWorkbook() As String
    Dim c           As Variant
    Dim out         As String
    For Each c In aCollection.CollectionsToArray2D(getDeclarations(, , , True))
        If InStr(1, CStr(c), "Type ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredTypeOfWorkbook = out
End Function

Function getDeclaredSubOfWorkbook(TargetWorkbook As Workbook) As String
    Dim c           As Variant
    Dim out         As String
    For Each c In aCollection.CollectionsToArray2D(getDeclarations(, , , True))
        If InStr(1, CStr(c), "Sub ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredSubOfWorkbook = out
End Function

Function getDeclaredFunctionOfWorkbook() As String
    Dim c           As Variant
    Dim out         As String
    For Each c In aCollection.CollectionsToArray2D(getDeclarations(, , , True))
        If InStr(1, CStr(c), "Function ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredFunctionOfWorkbook = out
End Function


