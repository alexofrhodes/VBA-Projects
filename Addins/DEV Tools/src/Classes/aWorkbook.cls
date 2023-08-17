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

Private oWorkbook As Workbook

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
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("References")
    ws.Cells(1, 1).Value = "Reference Description"
    ws.Cells(1, 2).Value = "GUID"
    ws.Cells(1, 3).Value = "Path"
    ws.Cells(1, 4).Value = "Version"
    Dim myRef As Reference
    Dim refs As VBIDE.REFERENCES
    Set refs = oWorkbook.VBProject.REFERENCES
    Dim i As Long
    i = 2
    For Each myRef In refs
        ws.Cells(i, 1) = IIf(myRef.Description <> "", myRef.Description, myRef.Name)
        ws.Cells(i, 2) = myRef.GUID
        ws.Cells(i, 3) = myRef.fullPath
        ws.Cells(i, 4) = myRef.Major & "." & myRef.Minor
        i = i + 1
    Next myRef
End Sub

Sub AddReferenceFromFile(filepath As String)
'@AssignedModule aWorkbook
'@INCLUDE CLASS aWorkbook
    Dim oRefs As REFERENCES: Set oRefs = oWorkbook.VBProject.REFERENCES
    oRefs.AddFromFile filepath
End Sub

Sub AddReferenceFromGUID(GUID As String, Major As String, Minor As String)
'@AssignedModule aWorkbook
'@INCLUDE CLASS aWorkbook
'@INCLUDE DECLARATION GUID
    Dim oRefs As REFERENCES: Set oRefs = oWorkbook.VBProject.REFERENCES
    oRefs.AddFromGuid GUID, Major, Minor
End Sub

Sub RemoveReferenceByName(RefName As String)
'@AssignedModule aWorkbook
'@INCLUDE CLASS aWorkbook
    Dim oRef As Reference
    Dim oRefs As REFERENCES: Set oRefs = oWorkbook.VBProject.REFERENCES
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
    Dim oRef As Reference
    Dim oRefs As REFERENCES: Set oRefs = oWorkbook.VBProject.REFERENCES
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
    Dim oRefs As REFERENCES: Set oRefs = oWorkbook.VBProject.REFERENCES
    Dim oRef As Reference
    For Each oRef In oRefs
'        debug.print oRef.Name
        If oRef.GUID = refGUID Then
            oRefs.Remove oRef
            Exit For
        End If
    Next oRef
End Sub

Public Sub ExportReferences(filepath As String)
'@AssignedModule aWorkbook
'@INCLUDE PROCEDURE TxtOverwrite
'@INCLUDE CLASS aWorkbook
'@INCLUDE DECLARATION GUID
    Dim Result As String
    Dim oRef As Reference
    For Each oRef In oWorkbook.VBProject.REFERENCES
        Result = Result & IIf(Result <> "", vbNewLine, "") & oRef.Name & vbTab & oRef.GUID & vbTab & oRef.Major & vbTab & oRef.Minor
    Next
    TxtOverwrite filepath, Result
End Sub


Public Sub ImportReferences(filepath As String)
'@AssignedModule aWorkbook
'@INCLUDE CLASS aWorkbook
    Dim FSO As New Scripting.FileSystemObject
    With FSO.OpenTextFile(filepath, ForReading, True)
        Dim Line As Long
        Do While Not .AtEndOfStream
            Dim Values As Variant
            Values = Split(.ReadLine, vbTab)
            On Error Resume Next
            oWorkbook.VBProject.REFERENCES.AddFromGuid Values(1), Values(2), Values(3)
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
    Dim out As String
    Dim am As aModule
    For Each am In aModules.Init(oWorkbook).Items
        out = out & IIf(out <> "", vbNewLine, "") & am.TodoList
    Next am
    TodoList = out
End Function

Function ProceduresLike(This As String) As Collection
'@AssignedModule aWorkbook
'@INCLUDE CLASS aProcedure
'@INCLUDE CLASS aWorkbook
    Dim Procedure As aProcedure
    Dim out As New Collection
    For Each Procedure In Procedures(True, False, False, False)
        If UCase(Procedure.Name) Like UCase(This) Then
            out.Add Procedure
        End If
    Next
    Set ProceduresLike = out
End Function


Public Function Code() As String
'@AssignedModule aWorkbook
'@INCLUDE CLASS aWorkbook
'@INCLUDE CLASS aModule
'@INCLUDE CLASS aModules
    Dim Module As aModule
    Dim txt
    For Each Module In aModules.Items
        If Module.item.CodeModule.CountOfLines > 0 Then
            txt = txt & _
                  vbNewLine & _
                  "'" & String(10, "=") & " (" & Module.TypeToString & ") " & Module.Name & String(10, "=") & _
                  vbNewLine & _
                  Module.Code
        End If
    Next
    Code = txt
End Function

Sub ModulesMerge(TargetModule As VBComponent, ParamArray Modules() As Variant)
'@AssignedModule aWorkbook
'@INCLUDE CLASS aWorkbook
'@INCLUDE CLASS aModule
    Dim element As Variant
    Dim Module As VBComponent
    For Each element In Modules
        Set Module = element
        If Module.Type = vbext_ct_StdModule Then
            If Module.Name <> TargetModule.Name Then
                aModule.Init(Module).CodeMove TargetModule
            End If
        End If
    Next
End Sub


Public Property Get ProceduresNames( _
                              includeModule As Boolean, _
                              includeDocument As Boolean, _
                              includeClass As Boolean, _
                              includeForm As Boolean) As Collection

    Dim Module As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim lineNum As Long
    Dim coll As New Collection
    Dim ProcedureName As String
    Dim i As Long
    For Each Module In oWorkbook.VBProject.VBComponents
        If Not includeModule And Module.Type = vbext_ct_StdModule Then GoTo SKIP
        If Not includeClass And Module.Type = vbext_ct_ClassModule Then GoTo SKIP
        If Not includeDocument And Module.Type = vbext_ct_Document Then GoTo SKIP
        If Not includeForm And Module.Type = vbext_ct_MSForm Then GoTo SKIP
        With Module.CodeModule
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
                            coll.Add ProcedureName, Module.Name & "|" & ProcedureName & "|" & ProcKind
                            On Error GoTo 0
                        End If
                        
                        Exit For
                        
                    End If
                Next
                lineNum = .procStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
            Loop
        End With
SKIP:
    Next Module
    
    Set ProceduresNames = coll
End Property

Public Property Get Procedures( _
                              includeModule As Boolean, _
                              includeDocument As Boolean, _
                              includeClass As Boolean, _
                              includeForm As Boolean) As Collection
'as aProcedure objects
    Dim Module As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim lineNum As Long
    Dim coll As New Collection
    Dim ProcedureName As String
    Dim cProcedure As aProcedure
    Dim i As Long
    For Each Module In oWorkbook.VBProject.VBComponents
        If Not includeModule And Module.Type = vbext_ct_StdModule Then GoTo SKIP
        If Not includeClass And Module.Type = vbext_ct_ClassModule Then GoTo SKIP
        If Not includeDocument And Module.Type = vbext_ct_Document Then GoTo SKIP
        If Not includeForm And Module.Type = vbext_ct_MSForm Then GoTo SKIP
        With Module.CodeModule
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
                            cProcedure.Init oWorkbook, Module, ProcedureName, ProcKind
                            On Error Resume Next
                            coll.Add cProcedure, cProcedure.Name & ProcKind
'                            coll.Add ProcedureName, ProcedureName
'                            If Err.Number > 0 Then Debug.Print ProcedureName & " exists multiple times, added only once, may lead to error"
'                            Err.clear
                            On Error GoTo 0
                        End If
                        
                        Exit For
                        
                    End If
                Next
                lineNum = .procStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
            Loop
        End With
SKIP:
    Next Module
    
    Set Procedures = coll
End Property


Function ProceduresArray()
    Dim Module As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim lineNum As Long
    Dim coll As New Collection
    Dim ProcedureName As String
    Dim cProcedure As aProcedure
    Dim arr
    Dim i As Long
    Dim BodyLine As Long
    ReDim arr(1 To 4, 1 To 1)
    For Each Module In oWorkbook.VBProject.VBComponents
        With Module.CodeModule
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
                        BodyLine = Module.CodeModule.ProcBodyLine(ProcedureName, ProcKind)
                        arr(1, UBound(arr, 2)) = Switch(Module.Type = vbext_ct_StdModule, "Module", _
                                                        Module.Type = vbext_ct_ClassModule, "Class", _
                                                        Module.Type = 3, "Userform", _
                                                        Module.Type = 100, "Document", _
                                                        Module.Type = 11, "ActiveXDesigner")
                        arr(2, UBound(arr, 2)) = Module.Name
                        arr(3, UBound(arr, 2)) = Switch(ProcKind = vbext_pk_Proc, IIf(Module.CodeModule.Lines(BodyLine, 1) Like "*Sub " & ProcedureName & "*", "Sub", "Function"), _
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
    Next Module
    ProceduresArray = WorksheetFunction.Transpose(arr)
End Function

Sub Export()
    uProjectManager.Show
End Sub

Sub Backup(TargetFolder As String)
    oWorkbook.SaveCopyAs TargetFolder & _
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

Public Sub ExportProject( _
            Optional bSeparateProcedures As Boolean, _
            Optional bExportComponents As Boolean, _
            Optional bWorkbookBackup As Boolean, _
            Optional bExportUnified As Boolean, _
            Optional bExportDeclarations As Boolean, _
            Optional bExportReferences As Boolean, _
            Optional bExportXML As Boolean, _
            Optional UseWorkbookFolder As Boolean, _
            Optional OpenFolderAfterExport As Boolean)
    
    Dim exportPath As String:
    If UseWorkbookFolder Then
        exportPath = oWorkbook.Path & "\"
    Else
        exportPath = Environ("USERprofile") & "\Documents\" & "vbArc\Backups\" & NameClean & "\"
    End If
    exportPath = exportPath & "src " & NameClean & " " & Format(Now, "YY-MM-DD HH-NN") & "\"
    
    FoldersCreate exportPath
    Dim ams As aModules:    Set ams = aModules.Init(oWorkbook)
    If bWorkbookBackup = True Then Backup exportPath
    If bExportUnified Then ExportCode exportPath
    If bExportReferences = True Then ExportReferences exportPath & "References.txt"
    If bExportComponents = True Then ams.Export exportPath, CreateSubFolders:=True
    If bExportDeclarations = True Then ExportDeclarations exportPath
    If bExportXML = True Then ExportXML exportPath
    If bSeparateProcedures = True Then ams.ExportProcedures exportPath
    On Error Resume Next
    Toast , "Export complete"
    If OpenFolderAfterExport Then FollowLink exportPath
End Sub

Sub ExportXML(TargetPath As String)
    TargetPath = Replace(TargetPath & "\", "\\", "\")
    FoldersCreate TargetPath
    Dim TmpFile As String
    TmpFile = oWorkbook.Path & "\temp_workbook_file" & Extension
    oWorkbook.SaveCopyAs TmpFile
    Dim c As New clsEditOpenXML
    c.ExtractRibbonX TmpFile, TargetPath & "customUI.xml"
    Kill TmpFile
    Set c = Nothing
End Sub
Sub ExportCode(TargetPath As String)
    TargetPath = Replace(TargetPath & "\", "\\", "\")
    FoldersCreate TargetPath
    Dim Result As String, tmp As String
    Dim am As aModule
    For Each am In aModules.Init(oWorkbook).Items
        tmp = "'" & am.Name & vbTab & am.TypeToString & vbNewLine & vbNewLine & am.Code
        Result = IIf(Result = "", tmp, Result & vbNewLine & vbNewLine & tmp)
    Next
    TxtAppend TargetPath & "#UnifiedProject.txt", Result
End Sub

Sub ExportHalf()
    ExportProject _
                 bSeparateProcedures:=False, _
                 bExportComponents:=True, _
                 bWorkbookBackup:=True, _
                 bExportUnified:=True, _
                 bExportDeclarations:=False, _
                 bExportReferences:=True, _
                 bExportXML:=False
End Sub

Function HasProject() As Boolean
    Dim WbProjComp As Object
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
        .Range("A1:H1").Value = var
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
    TargetWorksheet.Range("A1").CurrentRegion.offset(1).Clear

    Dim var
    ReDim var(1 To 8, 1 To 1)
    
    Dim counter As Long
    counter = 1

    Dim Module As VBComponent
    Dim Procedures As Collection
    
    Dim Procedure As aProcedure
    For Each Module In oWorkbook.VBProject.VBComponents
        Select Case Module.Type
        Case vbext_ct_StdModule, vbext_ct_Document, vbext_ct_MSForm  '@TODO check if adding these is ok
            Set Procedures = aModule.Init(Module).Procedures
            For Each Procedure In Procedures
                ReDim Preserve var(1 To 8, 1 To counter)
                var(1, counter) = oWorkbook.Name
                var(2, counter) = aModule.Init(Module).TypeToString
                var(3, counter) = Module.Name
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
    Dim sort1 As String
    sort1 = "A1"
    Dim sort2 As String
    sort2 = "C1"
    Dim sort3 As String
    sort3 = "D1"

    With TargetWorksheet.Sort
        .SortFields.Clear
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


Sub ExportDeclarations(TargetPath As String)
    TargetPath = Replace(TargetPath & "\", "\\", "\")
    FoldersCreate TargetPath

    Dim DeclarationArray As Variant
    DeclarationArray = aCollection.CollectionsToArray2D(getDeclarations)
    If TypeName(DeclarationArray) <> "Empty" Then
        TxtOverwrite TargetPath & "Declarations.txt", ArrayToString(DeclarationArray)
    End If
End Sub

Private Function DeclarationsWorksheetCreate() As Boolean
    If WorksheetExists("Declarations_Table", ThisWorkbook) Then Exit Function
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets.Add
    With TargetWorksheet
        .Name = "Declarations_Table"
        .Cells.VerticalAlignment = xlVAlignTop
        .Range("A1:F1").Value = Split("SCOPE,TYPE,NAME,CODE,MODULE TYPE,MODULE NAME", ",")
        .rows(1).Cells.Font.Bold = True
        .rows(1).Cells.Font.Size = 14
    End With
End Function

Public Sub DeclarationsTableCreate()

    DeclarationsWorksheetCreate

    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    'if sheet was created within the hour, you probably don't have new declarations
    If Format(Now, "YYMMDDHHNN") - TargetWorksheet.Range("Z1").Value < 60 Then Exit Sub

    TargetWorksheet.Range("A2").CurrentRegion.offset(1).Clear
    ArrayToRange2D aCollection.CollectionsToArray2D( _
                        getDeclarations( _
                            includeScope:=True, _
                            includeType:=True, _
                            includeKeywords:=True, _
                            includeDeclarations:=True, _
                            includeComponentName:=True, _
                            includeComponentType:=True)), _
                    TargetWorksheet.Range("A2")

    TargetWorksheet.Range("Z1").Value = Format(Now, "YYMMDDHHNN")

    DeclarationsTableSort
End Sub

Private Sub DeclarationsTableSort()
    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Worksheets("Declarations_Table")

    Dim sort1 As String: sort1 = "B1"
    Dim sort2 As String: sort2 = "C1"
    Dim sort3 As String ': sort3 = "D1"

    With TargetWorksheet.Sort
        .SortFields.Clear
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

Public Function getDeclarations( _
                        Optional includeScope As Boolean, _
                        Optional includeType As Boolean, _
                        Optional includeKeywords As Boolean, _
                        Optional includeDeclarations As Boolean, _
                        Optional includeComponentName As Boolean, _
                        Optional includeComponentType As Boolean) As Collection

    Dim ComponentCollection     As New Collection
    Dim ComponentTypecollection As New Collection
    Dim DeclarationsCollection  As New Collection
    Dim KeywordsCollection      As New Collection
    Dim output                  As New Collection
    Dim ScopeCollection         As New Collection
    Dim TypeCollection          As New Collection

    Dim element                 As Variant
    Dim OriginalDeclarations    As Variant
    Dim str                     As Variant

    Dim tmp                     As String
    Dim helper                  As String
    Dim i                       As Long

    Dim Module                  As VBComponent
    For Each Module In oWorkbook.VBProject.VBComponents
'        If Module.Type = vbext_ct_StdModule Or Module.Type = vbext_ct_MSForm Then
            If Module.CodeModule.CountOfDeclarationLines > 0 Then
                str = Module.CodeModule.Lines(1, Module.CodeModule.CountOfDeclarationLines)
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
                            DeclarationsCollection.Add DeclarationsKeywordSubstring(tmp, , "Enum " & KeywordsCollection.item(KeywordsCollection.count), "End Enum", , , True)
                            TypeCollection.Add "Enum"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add aModule.Init(Module).TypeToString
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(element), "\b ?Type \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Type")
                            DeclarationsCollection.Add DeclarationsKeywordSubstring(tmp, , "Type " & KeywordsCollection.item(KeywordsCollection.count), "End Type", , , True)
                            TypeCollection.Add "Type"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add aModule.Init(Module).TypeToString
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf InStr(1, CStr(element), "Const ", vbTextCompare) > 0 Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Const")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Const"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add aModule.Init(Module).TypeToString
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(element), "\b ?Sub \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Sub")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Sub"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add aModule.Init(Module).TypeToString
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(element), "\b ?Function \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Function")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Function"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add aModule.Init(Module).TypeToString
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf element Like "*(*) As *" Then
                            helper = Left(element, InStr(1, CStr(element), "(") - 1)
                            helper = Mid(helper, InStrRev(helper, " ") + 1)
                            KeywordsCollection.Add helper
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Other"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add aModule.Init(Module).TypeToString
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf element Like "* As *" Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", , "As")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Other"
                            ComponentCollection.Add Module.Name
                            ComponentTypecollection.Add aModule.Init(Module).TypeToString
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                        Else
                        End If
                    End If
                Next element
            End If
'        End If
    Next Module

    If includeScope = True Then output.Add ScopeCollection
    If includeType = True Then output.Add TypeCollection
    If includeKeywords = True Then output.Add KeywordsCollection
    If includeDeclarations = True Then output.Add DeclarationsCollection
    If includeComponentType = True Then output.Add ComponentTypecollection
    If includeComponentName = True Then output.Add ComponentCollection

    Set getDeclarations = output
End Function

Private Function DeclarationsKeywordSubstring( _
                                            str As Variant, Optional delim As String _
                                            , Optional afterWord As String _
                                            , Optional beforeWord As String _
                                            , Optional counter As Integer _
                                            , Optional outer As Boolean _
                                            , Optional includeWords As Boolean) As String

    Dim i As Long
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
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Declarations_Table")
    Dim codeName As String
    Dim codeText As String
    Dim cell As Range
    Dim Lr As Long
    Lr = getLastRow(ws)
    Dim codeCollection As Collection
    For Each cell In ws.Range("C2:C" & Lr)
        codeName = cell.text
        Set codeCollection = New Collection
        On Error Resume Next
        codeCollection.Add cell.offset(0, 1).text, cell.offset(0, 1).text
        On Error GoTo 0
        Do While cell.offset(1, 0).text = cell.text
            Set cell = cell.offset(1, 0)
            On Error Resume Next
            codeCollection.Add cell.offset(0, 1).text, cell.offset(0, 1).text
            On Error GoTo 0
        Loop
        codeText = aCollection.Init(codeCollection).ToString(IIf(codeCollection.count > 1, vbNewLine, ""))
        TxtOverwrite LOCAL_LIBRARY_DECLARATIONS & codeName & ".txt", codeText
    Next
End Sub

Function getDeclaredKeywordsOfWorkbook() As Variant
    getDeclaredKeywordsOfWorkbook = WorksheetFunction.Transpose(aCollection.CollectionsToArray2D(getDeclarations(, , True)))
End Function

Function getDeclaredEnumOfWorkbook() As String
    Dim c As Variant
    Dim out As String
    For Each c In aCollection.CollectionsToArray2D(getDeclarations(, , , True))
        If InStr(1, CStr(c), "Enum ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredEnumOfWorkbook = out
End Function

Function getDeclaredTypeOfWorkbook() As String
    Dim c As Variant
    Dim out As String
    For Each c In aCollection.CollectionsToArray2D(getDeclarations(, , , True))
        If InStr(1, CStr(c), "Type ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredTypeOfWorkbook = out
End Function

Function getDeclaredSubOfWorkbook(TargetWorkbook As Workbook) As String
    Dim c As Variant
    Dim out As String
    For Each c In aCollection.CollectionsToArray2D(getDeclarations(, , , True))
        If InStr(1, CStr(c), "Sub ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredSubOfWorkbook = out
End Function

Function getDeclaredFunctionOfWorkbook() As String
    Dim c As Variant
    Dim out As String
    For Each c In aCollection.CollectionsToArray2D(getDeclarations(, , , True))
        If InStr(1, CStr(c), "Function ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredFunctionOfWorkbook = out
End Function

