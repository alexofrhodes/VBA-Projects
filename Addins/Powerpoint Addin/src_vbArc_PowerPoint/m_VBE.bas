Attribute VB_Name = "m_VBE"
Option Explicit

Function ThisProject() As vbProject
    Set ThisProject = Application.VBE.VBProjects("vbArc")
End Function

Function ThisProjectPath() As String
    ThisProjectPath = getFileFolder(ThisProject.FileName)
End Function

Function getProjectFromModuleName(UniqueModuleName As String) As vbProject
    Dim vbProj As Object
    Dim vbProject As Object
    For Each vbProject In Application.VBE.VBProjects
        If vbProject.Protection = 0 Then
            Dim component As Object
            For Each component In vbProject.VBComponents
                If component.Name = UniqueModuleName Then
                    Set getProjectFromModuleName = vbProject
                    Exit Function
                End If
            Next component
        End If
    Next vbProject
End Function

Sub ListVBProjects(lbox As ListBox)
    lbox.ColumnCount = 2
'    lbox.Font.Name = "Elephant"
    lbox.Font.Size = 9
    lbox.ColumnWidths = "200;0"
    Dim Project As vbProject
    Dim ProjectFileName As String
    For Each Project In Application.VBE.VBProjects
        ProjectFileName = Project.FileName
        If Project.Protection = 0 Then
            If InStr(1, ProjectFileName, ".ppa") > 0 Then
                If AddIns(ProjectFileName).Loaded = 0 Then GoTo ResumeNext
            End If
            lbox.AddItem
            lbox.List(UBound(lbox.List), 0) = Mid(ProjectFileName, InStrRev(ProjectFileName, "\") + 1)
            lbox.List(UBound(lbox.List), 1) = ProjectFileName
        End If
ResumeNext:
    Next
End Sub

Function getProjectByPath(FilePath As String)
    Dim Project As vbProject
    For Each Project In Application.VBE.VBProjects
        If Project.FileName = FilePath Then
            Set getProjectByPath = Project
            Exit Function
        End If
    Next
End Function

Sub EditAddin(Project As vbProject)
    Dim Ad               As AddIn:  Set Ad = AddIns(Project.FileName)
    Dim AddinName        As String: AddinName = Mid(Ad.FullName, InStrRev(Ad.FullName, "\") + 1)
    Dim extension        As String: extension = Right(AddinName, InStrRev(AddinName, "."))
    Dim PresentationPath As String: PresentationPath = Ad.Path & "\" & Replace(AddinName, extension, ".pptm")
    Dim ProjectName      As String: ProjectName = Project.Name
    Dim Pres As Presentation
    If FileExists(PresentationPath) Then
        Ad.Loaded = False
        Ad.Registered = msoFalse
        Set Pres = Presentations.Open(PresentationPath)
    Else
        ExportModules Project
        Set Pres = Presentations.Add
        Pres.vbProject.Name = ProjectName
        Pres.SaveAs PresentationPath, ppSaveAsOpenXMLPresentationMacroEnabled
        ImportModules Pres.vbProject, False
        Pres.SaveAs PresentationPath, ppSaveAsOpenXMLPresentationMacroEnabled
        Ad.Loaded = False
        Ad.Registered = msoFalse
    End If
    Set Ad = Nothing
End Sub

Sub FinishEditing(Project As vbProject)
    Dim Pres As Presentation:   Set Pres = Presentations(Project.FileName)
    Pres.SaveAs Project.FileName
    ExportModules Project
    Dim savePath As String:     savePath = Replace(Project.FileName, ".ppt", ".ppa")
    AddIns.Remove AddIns(savePath).Name
    KillFile savePath
    Pres.SaveCopyAs savePath, ppSaveAsOpenXMLAddin
    AddIns.Add savePath
    AddIns(savePath).Loaded = True
    Pres.Close
    Set Pres = Nothing
End Sub

Sub KillFile(FilePath As String)
    On Error GoTo ErrorHandler
    Dim fso As New FileSystemObject, aFile As file
    If (fso.FileExists(FilePath)) Then
        SetAttr FilePath, vbNormal
        Set aFile = fso.GetFile(FilePath)
        aFile.Delete
    End If
    Exit Sub
ErrorHandler:
    If MsgBox("Could not delete " & vbNewLine & FilePath & vbNewLine & _
              "Cancel operation?", vbYesNo) = vbYes Then
        End
    Else
        Exit Sub
    End If
End Sub

Public Function FileExists(ByVal FileName As String) As Boolean
    If InStr(1, FileName, "\") = 0 Then Exit Function
    If Right(FileName, 1) = "\" Then FileName = Left(FileName, Len(FileName) - 1)
    FileExists = (Dir(FileName, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "")
End Function

Public Sub ExportModules(Project As vbProject)
    If Project.Protection = 1 Then
        MsgBox "The VBA in this project is protected," & _
               "not possible to Export the code"
        Exit Sub
    End If
    Dim bExport As Boolean
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    If FolderWithVBAProjectFiles(Project) = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    On Error Resume Next
    Kill FolderWithVBAProjectFiles(Project) & "\*.*"
    On Error GoTo 0
    
    szExportPath = FolderWithVBAProjectFiles(Project)
    For Each cmpComponent In Project.VBComponents
        bExport = True
        szFileName = cmpComponent.Name
        Select Case cmpComponent.Type
        Case vbext_ct_ClassModule
            szFileName = szFileName & ".cls"
        Case vbext_ct_MSForm
            szFileName = szFileName & ".frm"
        Case vbext_ct_StdModule
            szFileName = szFileName & ".bas"
        Case vbext_ct_Document
            ''' This is a worksheet or workbook object. Don't try to export.
            bExport = False
        End Select
        If bExport Then
            cmpComponent.Export szExportPath & szFileName
        End If
    Next cmpComponent
    
'    MsgBox "Export is ready"
End Sub

Public Sub ImportModules(Project As vbProject, alert As Boolean)
    'WARNING!
    'DELETES OLD MODULES AND USERFORMS BEFORE IMPORTING NEW

    If Project.Protection = 1 Then
        MsgBox "The VBA in this project is protected," & _
               "not possible to Import the code"
        Exit Sub
    End If
        
    Dim szFileName As String
    Dim szImportPath As String
        szImportPath = FolderWithVBAProjectFiles(Project)
        
    Dim objFSO As Scripting.FileSystemObject
    Set objFSO = New Scripting.FileSystemObject
    
    If szImportPath = "Error" Then
        MsgBox "Import Folder could not be created"
        Exit Sub
    Else
        Dim fileCount As Long
        fileCount = objFSO.GetFolder(szImportPath).Files.count
        If fileCount = 0 Then
            MsgBox "No modules found in " & szImportPath
            Exit Sub
        Else
            If MsgBox("This will remove " & fileCount & " Modules from" & vbNewLine & _
            Project.FileName & vbNewLine & _
            "and import " & Project.VBComponents.count & " from the folder" & vbNewLine & _
            szImportPath, vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    
    Call DeleteVBAModulesAndUserForms(Project)
    
    Dim objFile As Scripting.file
    For Each objFile In objFSO.GetFolder(szImportPath).Files
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            Project.VBComponents.Import objFile.Path
        End If
    Next objFile
End Sub

Function FolderWithVBAProjectFiles(Project As vbProject) As String
    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath As String
    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")
    SpecialPath = Left(Project.FileName, InStrRev(Project.FileName, "\"))
    SpecialPath = SpecialPath & "src_" & Mid(Project.FileName, InStrRev(Project.FileName, "\") + 1)
    SpecialPath = Left(SpecialPath, InStrRev(SpecialPath, ".") - 1) & "\"
    If fso.FolderExists(SpecialPath) = False Then
        On Error Resume Next
        MkDir SpecialPath
        On Error GoTo 0
    End If
    If fso.FolderExists(SpecialPath) = True Then
        FolderWithVBAProjectFiles = SpecialPath
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
End Function

Function DeleteVBAModulesAndUserForms(Project As vbProject)
    Dim vbComp As VBIDE.VBComponent
    For Each vbComp In Project.VBComponents
        '    Debug.Print VBComp.Name
        If vbComp.Type = vbext_ct_Document Then
            'Thisworkbook or worksheet module
            'We do nothing
        Else
            Project.VBComponents.Remove vbComp
        End If
    Next vbComp
End Function


Function ProceduresOfProject(TargetProject As vbProject) As Collection
    Dim Module As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim lineNum As Long
    Dim coll As New Collection
    Dim ProcedureName As String
    For Each Module In TargetProject.VBComponents
        If Module.Type = vbext_ct_StdModule Then
            With Module.CodeModule
                lineNum = .CountOfDeclarationLines + 1
                Do Until lineNum >= .CountOfLines
                    ProcedureName = .ProcOfLine(lineNum, ProcKind)
                    ' _ is used in events. Events may have the same name in different components
                    If InStr(1, ProcedureName, "_") = 0 Then
                        coll.Add ProcedureName
                    End If
                    lineNum = .ProcStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
                Loop
            End With
        End If
    Next Module
    Set ProceduresOfProject = coll
End Function

Function ProcedureExists( _
                        TargetProject As vbProject, _
                        ProcedureName As Variant) As Boolean
    Dim Procedures As Collection
    Set Procedures = ProceduresOfProject(TargetProject)
    Dim Procedure As Variant
    For Each Procedure In Procedures
        If UCase(CStr(Procedure)) = UCase(ProcedureName) Then
            ProcedureExists = True
            Exit Function
        End If
    Next
End Function

