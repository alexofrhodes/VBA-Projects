Attribute VB_Name = "Dependencies"


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Dependencies
'* Purpose    : list dependencies, export/import procedures/components
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
'* Created    : some time ago       Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


#If VBA7 Then
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
#Else
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
#End If

Rem ___CHANGE THESE TO MATCH YOUR FOLDER AND REPO____
'------------------------------------------------------------------------------
Public Const GITHUB_LIBRARY = "https://raw.githubusercontent.com/alexofrhodes/VBA-Library/"
'------------------------------------------------------------------------------
Public Const GITHUB_LIBRARY_DECLARATIONS = GITHUB_LIBRARY & "Declarations/"
Public Const GITHUB_LIBRARY_PROCEDURES = GITHUB_LIBRARY & "Procedures/"
Public Const GITHUB_LIBRARY_USERFORMS = GITHUB_LIBRARY & "Userforms/"
Public Const GITHUB_LIBRARY_CLASSES = GITHUB_LIBRARY & "Classes/"
'------------------------------------------------------------------------------
Public Const GITHUB_BLOG = "https://alexofrhodes.github.io/"
Public Const GITHUB_URL = "https://github.com/alexofrhodes/"
'------------------------------------------------------------------------------
Public Const AUTHOR_YOUTUBE = "https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg"
Public Const AUTHOR_VK = "https://vk.com/video/playlist/735281600_1"
Public Const AUTHOR_NAME = "Anastasiou Alex"
Public Const AUTHOR_EMAIL = "AnastasiouAlex@gmail.com"
Public Const AUTHOR_COPYRIGHT = ""
Public Const AUTHOR_OTHERTEXT = ""

Public Const GUID = "NBGD41100771701DDE7600"

Public ShowInVBE    As Boolean

'------------------------------------------------------------------------------
Public Function LOCAL_LIBRARY(): LOCAL_LIBRARY = LIBRARY_FOLDER: End Function
'------------------------------------------------------------------------------
Public Function LOCAL_LIBRARY_DECLARATIONS(): LOCAL_LIBRARY_DECLARATIONS = LOCAL_LIBRARY & "Declarations\": End Function
Public Function LOCAL_LIBRARY_PROCEDURES(): LOCAL_LIBRARY_PROCEDURES = LOCAL_LIBRARY & "Procedures\": End Function
Public Function LOCAL_LIBRARY_USERFORMS(): LOCAL_LIBRARY_USERFORMS = LOCAL_LIBRARY & "Userforms\": End Function
Public Function LOCAL_LIBRARY_CLASSES(): LOCAL_LIBRARY_CLASSES = LOCAL_LIBRARY & "Classes\": End Function

Function LIBRARY_FOLDER() As String
    If GetMotherBoardProp = GUID Then
        LIBRARY_FOLDER = "C:\Users\acer\Documents\GitHub\VBA-Library\"
    Else
        LIBRARY_FOLDER = Environ$("USERPROFILE") & "\Documents\vbArc\Library\"
    End If
End Function

Public Function AUTHOR_MEDIA() As String
    AUTHOR_MEDIA = "'* BLOG       : " & GITHUB_BLOG & vbNewLine & _
                   "'* GITHUB     : " & GITHUB_URL & vbNewLine & _
                   "'* YOUTUBE    : " & AUTHOR_YOUTUBE & vbNewLine & _
                   "'* VK         : " & AUTHOR_VK & vbNewLine & "'*" & vbNewLine
End Function

Function DevInfo() As String
    DevInfo = Join( _
                Array( _
                    "AUTHOR     " & AUTHOR_NAME, _
                    "EMAIL      " & AUTHOR_EMAIL, _
                    "BLOG       " & GITHUB_BLOG, _
                    "GITHUB     " & GITHUB_URL, _
                    "YOUTUBE    " & AUTHOR_YOUTUBE, _
                    "VK         " & AUTHOR_VK), _
                vbNewLine)
End Function


'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 08:29    Alex                (z_zTest.bas > ListOfIncludes)

Function ListOfIncludes(TargetWorkbook As Workbook)
'@LastModified 2308220829
'@INCLUDE CLASS aWorkbook
'@INCLUDE PROCEDURE ArrayQuickSort
'@INCLUDE PROCEDURE ArrayDuplicatesRemove
'@INCLUDE PROCEDURE ArrayReplace
'@INCLUDE PROCEDURE ArrayTrim
    Dim arr
    arr = ArrayTrim( _
            Split( _
                aWorkbook.Init(ThisWorkbook).Code, _
                vbNewLine))
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        arr(i) = WorksheetFunction.Proper(arr(i))
    Next
    arr = ArrayDuplicatesRemove( _
            Filter( _
                Filter( _
                    arr, _
                    "'@INCLUDE PROCEDURE ", _
                    True, _
                    vbTextCompare), _
                """", _
                False))
    ArrayReplace arr, "'@INCLUDE PROCEDURE ", ""
    ArrayQuickSort arr
    
    ListOfIncludes = arr
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 08:28    Alex                (Dependencies.bas > ProceduresNotExported)

Function ProceduresNotExported(TargetWorkbook As Workbook)
'@LastModified 2308220828
'@INCLUDE PROCEDURE FileExists
'@INCLUDE PROCEDURE ListOfIncludes
    Dim out
    Dim TargetFile
    Dim Procedure
    For Each Procedure In ListOfIncludes(TargetWorkbook)
        TargetFile = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"
        If Not FileExists(TargetFile) Then
            out = out & IIf(out <> "", vbNewLine, "") & Procedure
        End If
    Next
    out = Split(out, vbNewLine)
    ProceduresNotExported = out
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 08:28    Alex                (Dependencies.bas > IncludeNotImported)

Function IncludeNotImported(TargetWorkbook As Workbook)
'@LastModified 2308220828
'@INCLUDE PROCEDURE FileExists
'@INCLUDE PROCEDURE ListOfIncludes
    Dim out
    Dim TargetFile
    Dim Procedure
    For Each Procedure In ListOfIncludes(TargetWorkbook)
        If Not ProcedureExists(TargetWorkbook, Procedure) Then
            out = out & IIf(out <> "", vbNewLine, "") & Procedure
        End If
    Next
    out = Split(out, vbNewLine)
    IncludeNotImported = out
End Function

'--------------------------------------------
Sub AddLinkedListsToActiveProcedure()
    AddLinkedLists ThisWorkbook, ActiveModule, ActiveProcedure
End Sub

Sub ExportActiveProcedure()
    ExportProcedure ThisWorkbook, ActiveModule, ActiveProcedure, ExportMergedTxt:=True
End Sub

Sub ExportAllProceduresOfThisWorkbook()
    ExportAllProcedures ThisWorkbook
End Sub

Sub ImportActiveProcedureDependencies()
    ImportProcedureDependencies ActiveProcedure, ThisWorkbook, ActiveModule, Overwrite:=True
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 30-10-2023 12:23    Alex                (Dependencies.bas > AddLinkedListsToAllProcedures : )

Sub AddLinkedListsToAllProcedures(TargetWorkbook As Workbook)
'@LastModified 2310301223
    Dim Procedure
    Dim module      As VBComponent
    For Each module In TargetWorkbook.VBProject.VBComponents
        If module.Type = vbext_ct_StdModule And module.Name <> "Dependencies" Then
            For Each Procedure In ProceduresOfModule(module)
                AddLinkedLists TargetWorkbook, module, CStr(Procedure)
            Next Procedure
        End If
    Next module
    Toast "Done"
End Sub

Sub AddLinkedListsToActiveWorkbook()
    AddLinkedListsToAllProcedures ActiveCodepaneWorkbook
End Sub

Sub AddLinkedListsToProceduresOfModule(module As VBComponent)
    Dim Procedure
    On Error GoTo EH
    For Each Procedure In ProceduresOfModule(module)
        Debug.Print Procedure
        AddLinkedLists , module, CStr(Procedure)
    Next Procedure
    Debug.Print vbNewLine & "---" & "Done"
    Exit Sub
EH:
    Debug.Print "Error at: " & module.Name & vbTab & Procedure
    Resume Next
End Sub

Sub ExportAllProcedures(TargetWorkbook As Workbook)
    Dim Procedure
    Dim module      As VBComponent
    For Each module In TargetWorkbook.VBProject.VBComponents
        If module.Type = vbext_ct_StdModule Then
            For Each Procedure In ProceduresOfModule(module)
                ExportProcedure TargetWorkbook, module, CStr(Procedure), False
            Next Procedure
        End If
    Next module
End Sub

Sub RemoveComments(TargetWorkbook As Workbook)
    Dim module      As VBComponent
    Dim S           As String
    Dim i           As Long
    For Each module In TargetWorkbook.VBProject.VBComponents
        For i = module.CodeModule.CountOfLines To 1 Step -1
            S = Trim(module.CodeModule.Lines(i, 1))
            If S Like "'*" Then module.CodeModule.DeleteLines i, 1
        Next i
    Next
End Sub

Function ArrayAppend(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    Dim holdarr     As Variant
    Dim ub1         As Long
    Dim ub2         As Long
    Dim i           As Long
    Dim newind      As Long
    If IsEmpty(arr1) Or Not IsArray(arr1) Then
        arr1 = Array()
    End If
    If IsEmpty(arr2) Or Not IsArray(arr2) Then
        arr2 = Array()
    End If
    ub1 = UBound(arr1)
    ub2 = UBound(arr2)
    If ub1 = -1 Then
        ArrayAppend = arr2
        Exit Function
    End If
    If ub2 = -1 Then
        ArrayAppend = arr1
        Exit Function
    End If
    holdarr = arr1
    ReDim Preserve holdarr(ub1 + ub2 + 1)
    newind = UBound(arr1) + 1
    For i = 0 To ub2
        If VarType(arr2(i)) = vbObject Then
            Set holdarr(newind) = arr2(i)
        Else
            holdarr(newind) = arr2(i)
        End If
        newind = newind + 1
    Next i
    ArrayAppend = holdarr
End Function

Public Sub ArrayQuickSort(ByRef SortableArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1)
    On Error Resume Next
    Dim i           As Long
    Dim j           As Long
    Dim varMid      As Variant
    Dim varX        As Variant
    If IsEmpty(SortableArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortableArray), "()") < 1 Then
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortableArray)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortableArray)
    End If
    If lngMin >= lngMax Then
        Exit Sub
    End If
    i = lngMin
    j = lngMax
    varMid = Empty
    varMid = SortableArray((lngMin + lngMax) \ 2)
    If IsObject(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If
    While i <= j
        While SortableArray(i) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortableArray(j) And j > lngMin
            j = j - 1
        Wend
        If i <= j Then
            varX = SortableArray(i)
            SortableArray(i) = SortableArray(j)
            SortableArray(j) = varX
            i = i + 1
            j = j - 1
        End If
    Wend
    If (lngMin < j) Then Call ArrayQuickSort(SortableArray, lngMin, j)
    If (i < lngMax) Then Call ArrayQuickSort(SortableArray, i, lngMax)
End Sub

Public Function cleanArray(varArray As Variant) As Variant()
    Dim TempArray() As Variant
    Dim OldIndex    As Integer
    Dim NewIndex    As Integer
    Dim Output      As String
    If Not ArrayAllocated(varArray) Then Exit Function
    ReDim TempArray(LBound(varArray) To UBound(varArray))
    For OldIndex = LBound(varArray) To UBound(varArray)
        Output = CleanTrim(varArray(OldIndex))
        If Not Output = "" Then
            TempArray(NewIndex) = Output
            NewIndex = NewIndex + 1
        End If
    Next OldIndex
    ReDim Preserve TempArray(LBound(varArray) To NewIndex - 1)
    cleanArray = TempArray
End Function

Function ArrayDuplicatesRemove(myArray As Variant) As Variant
    Dim nFirst As Long, nLast As Long, i As Long
    Dim item        As String

    Dim arrTemp()   As String
    Dim coll        As New Collection
    If Not ArrayAllocated(myArray) Then Exit Function
    nFirst = LBound(myArray)
    nLast = UBound(myArray)
    ReDim arrTemp(nFirst To nLast)

    For i = nFirst To nLast
        arrTemp(i) = CStr(myArray(i))
    Next i

    On Error Resume Next
    For i = nFirst To nLast
        coll.Add arrTemp(i), arrTemp(i)
    Next i
    Err.clear
    On Error GoTo 0

    nLast = coll.Count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)

    For i = nFirst To nLast
        arrTemp(i) = coll(i - nFirst + 1)
    Next i

    ArrayDuplicatesRemove = arrTemp

End Function

Public Function ArrayToCollection(items As Variant) As Collection
    If Not ArrayAllocated(items) Then Exit Function
    Dim coll        As New Collection
    Dim i           As Integer
    For i = LBound(items) To UBound(items)
        coll.Add items(i)
    Next
    Set ArrayToCollection = coll
End Function

Function CleanTrim(ByVal S As String, Optional ConvertNonBreakingSpace As Boolean = True) As String
    Dim X As Long, CodesToClean As Variant
    CodesToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
            21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157)
    If ConvertNonBreakingSpace Then S = Replace(S, Chr(160), " ")
    S = Replace(S, vbCr, "")
    For X = LBound(CodesToClean) To UBound(CodesToClean)
        If InStr(S, Chr(CodesToClean(X))) Then
            S = Replace(S, Chr(CodesToClean(X)), vbNullString)
        End If
    Next
    CleanTrim = S
    CleanTrim = Trim(S)
End Function

Sub AddLinkedLists(Optional TargetWorkbook As Workbook, _
                    Optional module As VBComponent, _
                    Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, Procedure) Then Exit Sub
    ProcedureLinesRemoveInclude TargetWorkbook, module, Procedure
    ProcedureAssignedModuleAdd TargetWorkbook, module, Procedure
    AddListOfLinkedProceduresToProcedure TargetWorkbook, module, Procedure
    AddListOfLinkedClassesToProcedure TargetWorkbook, module, Procedure
    AddListOfLinkedUserformsToProcedure TargetWorkbook, module, Procedure
    AddListOfLinkedDeclarationsToProcedure TargetWorkbook, module, Procedure
End Sub


Sub AddListOfLinkedClassesToProcedure( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional ProcedureName As String)

    If Not AssignCPSvariables(TargetWorkbook, module, ProcedureName) Then Stop
    Dim ListOfImports As String
    Dim Code        As String
    Code = ProcedureCode(TargetWorkbook, module, ProcedureName)
    Dim myClasses   As Collection
    Set myClasses = LinkedClasses(TargetWorkbook, module, ProcedureName)
    Dim element     As Variant
    For Each element In myClasses
        If InStr(1, Code, "@INCLUDE CLASS " & element) = 0 _
                And InStr(1, ListOfImports, "@INCLUDE CLASS " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE CLASS " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE CLASS " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        module.CodeModule.InsertLines _
                ProcedureBodyLineFirstAfterComments(module, ProcedureName), ListOfImports
    End If
End Sub

Sub AddListOfLinkedDeclarationsToProcedure( _
                                            Optional TargetWorkbook As Workbook, _
                                            Optional module As VBComponent, _
                                            Optional ProcedureName As String)

    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim ListOfImports As String
    If module Is Nothing Then Set module = ModuleOfProcedure(TargetWorkbook, ProcedureName)
    Dim ProcedureText As String
    ProcedureText = ProcedureCode(TargetWorkbook, module, ProcedureName)
    Dim myDeclarations As Collection
    Set myDeclarations = LinkedDeclarations(TargetWorkbook, module, ProcedureName)
    Dim coll        As New Collection
    Dim element     As Variant
    For Each element In myDeclarations
        If InStr(1, ProcedureText, "'@INCLUDE DECLARATION " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE DECLARATION " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE DECLARATION " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(module, ProcedureName), ListOfImports
    End If
End Sub

Sub AddListOfLinkedProceduresToProcedure( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional ProcedureName As String)

    If Not AssignCPSvariables(TargetWorkbook, module, ProcedureName) Then Stop
    Dim Procedures  As Collection
    Set Procedures = LinkedProcedures(TargetWorkbook, module, ProcedureName)
    Dim ListOfImports As String
    Dim Code        As String
    Code = ProcedureCode(TargetWorkbook, module, ProcedureName)
    Dim Procedure   As Variant
    For Each Procedure In Procedures
        If InStr(1, Code, "@INCLUDE PROCEDURE " & Procedure) = 0 And InStr(1, ListOfImports, "@INCLUDE PROCEDURE " & Procedure) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE PROCEDURE " & Procedure
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE PROCEDURE " & Procedure
            End If
        End If
    Next
    If ListOfImports <> "" Then
        module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(module, ProcedureName), ListOfImports
    End If
End Sub

Sub AddListOfLinkedUserformsToProcedure( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional ProcedureName As String)

    If Not AssignCPSvariables(TargetWorkbook, module, ProcedureName) Then Stop

    Dim ListOfImports As String
    Dim Code        As String
    Code = ProcedureCode(TargetWorkbook, module, ProcedureName)
    Dim myClasses   As Collection
    Set myClasses = LinkedUserforms(TargetWorkbook, module, ProcedureName)
    Dim element     As Variant
    For Each element In myClasses
        If InStr(1, Code, "@INCLUDE USERFORM " & element) = 0 And InStr(1, ListOfImports, "@INCLUDE USERFORM " & element) = 0 Then
            If ListOfImports = "" Then
                ListOfImports = "'@INCLUDE USERFORM " & element
            Else
                ListOfImports = ListOfImports & vbNewLine & "'@INCLUDE USERFORM " & element
            End If
        End If
    Next
    If ListOfImports <> "" Then
        module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(module, ProcedureName), ListOfImports
    End If
End Sub

Public Function ActiveProcedure() As String
    Application.VBE.ActiveCodePane.GetSelection L1&, c1&, L2&, c2&
    ActiveProcedure = Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(L1&, vbext_pk_Proc)
End Function

Public Function ActiveModule() As VBComponent
    '@LastModified 2308171258
    
    Dim Module1 As VBComponent
    'may erroneously return userform or worksheet

    Set Module1 = Application.VBE.SelectedVBComponent
    Dim Module2 As VBComponent
    If Not Application.VBE.ActiveCodePane Is Nothing Then Set Module2 = Application.VBE.ActiveCodePane.CodeModule.Parent

    If Module1.Name = Module2.Name Then
        Set ActiveModule = Module1
    Else
        Dim ans As Long
        ans = MsgBox("SelectedVBComponent <> ActiveCodePane.CodeModule.Parent" & vbLf & _
                      "Choose " & Module1.Name & " ?    Click no to choose " & Module2.Name, _
                      vbExclamation + vbYesNoCancel)
        Select Case ans
            Case vbYes: Set ActiveModule = Module1
            Case vbNo: Set ActiveModule = Module2
            Case vbCancel: Stop
        End Select
    End If
End Function


Public Function ActiveCodepaneWorkbook() As Workbook
    On Error GoTo ErrorHandler
    Dim WorkbookName As String
    WorkbookName = Application.VBE.SelectedVBComponent.Collection.Parent.fileName
    WorkbookName = Right(WorkbookName, Len(WorkbookName) - InStrRev(WorkbookName, "\"))
    Set ActiveCodepaneWorkbook = Workbooks(WorkbookName)
    Exit Function
ErrorHandler:
    MsgBox "doesn't work on new-unsaved workbooks"
End Function

Public Function ArrayAllocated(ByVal arr As Variant) As Boolean
    On Error Resume Next
    ArrayAllocated = IsArray(arr) And (Not IsError(LBound(arr, 1))) And LBound(arr, 1) <= UBound(arr, 1)
End Function

Public Function ArrayDimensionLength(SourceArray As Variant) As Integer
    Dim i           As Integer
    Dim test        As Long
    On Error GoTo catch
    Do
        i = i + 1
        test = UBound(SourceArray, i)
    Loop
catch:
    ArrayDimensionLength = i - 1
End Function

Public Sub ArrayToRange2D(arr2d As Variant, cell As Range)

    If ArrayDimensionLength(arr2d) = 1 Then arr2d = WorksheetFunction.Transpose(arr2d)
    Dim dif         As Long
    dif = IIf(LBound(arr2d, 1) = 0, 1, 0)
    Dim rng         As Range
    Set rng = cell.Resize(UBound(arr2d, 1) + dif, UBound(arr2d, 2) + dif)

    If Application.WorksheetFunction.CountA(rng) > 0 Then
        Exit Sub
    End If

    rng.Value = arr2d
End Sub

Function AssignCPSvariables( _
        ByRef TargetWorkbook As Workbook, _
        ByRef module As VBComponent, _
        ByRef Procedure As String) As Boolean

    If Not AssignWorkbookVariable(TargetWorkbook) Then Exit Function
    If Not AssignModuleVariable(TargetWorkbook, module, Procedure) Then Exit Function
    If Not AssignProcedureVariable(TargetWorkbook, Procedure) Then Exit Function
    AssignCPSvariables = True

End Function

Function AssignModuleVariable( _
        ByVal TargetWorkbook As Workbook, _
        ByRef module As VBComponent, _
        Optional ByVal Procedure As String) As Boolean
    If module Is Nothing Then
        If Procedure = "" Then
            Set module = ActiveModule
        End If
        On Error Resume Next
        Set module = ModuleOfProcedure(TargetWorkbook, Procedure)
        On Error GoTo 0
    End If
    AssignModuleVariable = Not module Is Nothing
End Function

Function AssignProcedureVariable(TargetWorkbook As Workbook, ByRef Procedure As String) As Boolean
    If Procedure = "" Then
        Dim cps     As String
        cps = CodepaneSelection
        If Len(cps) > 0 Then
            Procedure = cps
        Else
            Procedure = ActiveProcedure
        End If
        If Not ProcedureExists(TargetWorkbook, Procedure) Then
            Debug.Print Procedure & " not found in Workbook " & TargetWorkbook.Name
        End If
    End If
    AssignProcedureVariable = Not Procedure = ""
End Function

Function AssignWorkbookVariable(ByRef TargetWorkbook As Workbook) As Boolean
    If TargetWorkbook Is Nothing Then
        On Error Resume Next
        Set TargetWorkbook = ActiveCodepaneWorkbook
        On Error GoTo 0
    End If
    AssignWorkbookVariable = Not TargetWorkbook Is Nothing
End Function

Function CheckPath(path) As String
    Dim RetVal
    RetVal = "I"
    If (RetVal = "I") And FileExists(path) Then RetVal = "F"
    If (RetVal = "I") And FolderExists(path) Then RetVal = "D"
    If (RetVal = "I") And URLExists(path) Then RetVal = "U"
    CheckPath = RetVal
End Function

Function ClassNames(Optional TargetWorkbook As Workbook)
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Set ClassNames = ComponentNames(vbext_ct_ClassModule, TargetWorkbook)
End Function

Public Function CodepaneSelection() As String
    Dim startLine As Long, StartColumn As Long, endLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection startLine, StartColumn, endLine, EndColumn
    If endLine - startLine = 0 Then
        CodepaneSelection = Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(startLine, 1), StartColumn, EndColumn - StartColumn)
        Exit Function
    End If
    Dim str         As String
    Dim i           As Long
    For i = startLine To endLine
        If str = "" Then
            str = Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1), StartColumn)
        ElseIf i < endLine Then
            str = str & vbNewLine & Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1)
        Else
            str = str & vbNewLine & Left(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1), EndColumn - 1)
        End If
    Next
    CodepaneSelection = str
End Function

Public Function CollectionContains( _
        Kollection As Collection, _
        Optional Key As Variant, _
        Optional item As Variant) As Boolean
    Dim strKey      As String
    Dim var         As Variant
    If Not IsMissing(Key) Then
        strKey = CStr(Key)
        On Error Resume Next
        CollectionContains = True
        var = Kollection(strKey)
        If Err.Number = 91 Then GoTo CheckForObject
        If Err.Number = 5 Then GoTo NotFound
        On Error GoTo 0
        Exit Function
CheckForObject:
        If IsObject(Kollection(strKey)) Then
            CollectionContains = True
            On Error GoTo 0
            Exit Function
        End If
NotFound:
        CollectionContains = False
        On Error GoTo 0
        Exit Function
    ElseIf Not IsMissing(item) Then
        CollectionContains = False
        For Each var In Kollection
            If var = item Then
                CollectionContains = True
                Exit Function
            End If
        Next var
    Else
        CollectionContains = False
    End If
End Function

Public Function CollectionSort(colInput As Collection) As Collection
    Dim iCounter    As Integer
    Dim iCounter2   As Integer
    Dim Temp        As Variant
    Set CollectionSort = New Collection
    For iCounter = 1 To colInput.Count - 1
        For iCounter2 = iCounter + 1 To colInput.Count
            If colInput(iCounter) > colInput(iCounter2) Then
                Temp = colInput(iCounter2)
                colInput.Remove iCounter2
                colInput.Add Temp, , iCounter
            End If
        Next iCounter2
    Next iCounter
    Set CollectionSort = colInput
End Function

Function CollectionsToArray2D(collections As Collection) As Variant
    If collections.Count = 0 Then Exit Function
    Dim columnCount As Long
    columnCount = collections.Count
    Dim rowCount    As Long
    rowCount = collections.item(1).Count
    Dim var         As Variant
    ReDim var(1 To rowCount, 1 To columnCount)
    Dim cols        As Long
    Dim rows        As Long
    For rows = 1 To rowCount
        For cols = 1 To collections.Count
            var(rows, cols) = collections(cols).item(rows)
        Next cols
    Next rows
    CollectionsToArray2D = var
End Function

Function ComponentNames( _
        moduleType As vbext_ComponentType, _
        Optional TargetWorkbook As Workbook)
    Dim coll        As New Collection
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim module      As VBComponent
    For Each module In TargetWorkbook.VBProject.VBComponents
        If module.Type = moduleType Then
            coll.Add module.Name
        End If
    Next
    Set ComponentNames = coll
End Function

Function DeclarationsKeywordSubstring(str As Variant, Optional delim As String _
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
    End If
    DeclarationsKeywordSubstring = vbNullString
End Function

Sub DeclarationsTableCreate(TargetWorkbook As Workbook)

    DeclarationsWorksheetCreate

    Dim TargetWorksheet As Worksheet
    Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    If Format(Now, "YYMMDDHHNN") - TargetWorksheet.Range("Z1").Value < 60 Then Exit Sub

    TargetWorksheet.Range("A2").CurrentRegion.offset(1).clear
    ArrayToRange2D CollectionsToArray2D( _
            getDeclarations( _
            wb:=TargetWorkbook, _
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


Sub DeclarationsTableSort()

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



Function DeclarationsWorksheetCreate() As Boolean
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

Sub ExportLinkedDeclaration(TargetWorkbook As Workbook, DeclarationName As String)
    DeclarationsTableCreate TargetWorkbook
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



Sub ExportProcedure( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional ProcedureName As String, _
        Optional ExportMergedTxt As Boolean)

    If Not AssignCPSvariables(TargetWorkbook, module, ProcedureName) Then Exit Sub

    ProjetFoldersCreate

    Dim ExportedProcedures As New Collection
    On Error GoTo ErrorHandler

    ExportedProcedures.Add CStr(ProcedureName), CStr(ProcedureName)

    Dim Procedure
    For Each Procedure In LinkedProceduresDeep(ProcedureName, TargetWorkbook)
        ExportedProcedures.Add CStr(Procedure), CStr(Procedure)
    Next

    If ExportedProcedures.Count > 1 Then
        For Each Procedure In ExportedProcedures
            ExportTargetProcedure TargetWorkbook, , CStr(Procedure)
        Next
        If ExportMergedTxt Then
            Dim MergedName As String: MergedName = "Merged_" & ProcedureName
            Dim fileName As String: fileName = LOCAL_LIBRARY_PROCEDURES & MergedName & ".txt"
            Dim MergedString As String

            For Each Procedure In ExportedProcedures
                MergedString = MergedString & vbNewLine & ProcedureCode(TargetWorkbook, , Procedure)
            Next
            Debug.Print "OVERWROTE " & MergedName
            TxtOverwrite fileName, MergedString
            TxtPrependContainedProcedures fileName
        End If
    End If

    FollowLink LOCAL_LIBRARY_PROCEDURES

    Exit Sub
ErrorHandler:
    MsgBox "An error occured in Sub ExportProcedure"
End Sub

Sub ExportTargetProcedure( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Procedure As String)

    If Not AssignCPSvariables(TargetWorkbook, module, Procedure) Then Exit Sub

    Dim proclastmod
    proclastmod = ProcedureLastModified(TargetWorkbook, module, Procedure)
    If proclastmod = 0 Then
        AddLinkedLists TargetWorkbook, module, Procedure
        proclastmod = ProcedureLastModAdd(TargetWorkbook, module, Procedure)
    End If

    Dim Code        As String
    Code = ProcedureCode(TargetWorkbook, module, CStr(Procedure))
    Dim FileFullName As String
    FileFullName = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"
    If FileExists(FileFullName) Then
        Dim filelastmod
        filelastmod = StringLastModified(TxtRead(FileFullName))
        If proclastmod > filelastmod Then
            Debug.Print "OVERWROTE " & Procedure
            TxtOverwrite FileFullName, Code
        End If
    Else
        Debug.Print "NEW " & Procedure
        TxtOverwrite FileFullName, Code
    End If

    Dim element
    For Each element In LinkedUserforms(TargetWorkbook, module, CStr(Procedure))
        TargetWorkbook.VBProject.VBComponents(element).Export LOCAL_LIBRARY_USERFORMS & element & ".frm"
    Next
    For Each element In LinkedClasses(TargetWorkbook, module, CStr(Procedure))
        TargetWorkbook.VBProject.VBComponents(element).Export LOCAL_LIBRARY_CLASSES & element & ".cls"
    Next
    For Each element In LinkedDeclarations(TargetWorkbook, module, CStr(Procedure))
        ExportLinkedDeclaration TargetWorkbook, CStr(element)
    Next
End Sub

Public Function FileExists(ByVal fileName As String) As Boolean
    If InStr(1, fileName, "\") = 0 Then Exit Function
    If Right(fileName, 1) = "\" Then fileName = Left(fileName, Len(fileName) - 1)
    FileExists = (Dir(fileName, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "")
End Function

Function FolderExists(ByVal strPath As String) As Boolean
'@LastModified 2310251309
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Sub FoldersCreate(FolderPath As String)
    On Error Resume Next
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim ArrayElement As Variant
    individualFolders = Split(FolderPath, "\")
    For Each ArrayElement In individualFolders
        tempFolderPath = tempFolderPath & ArrayElement & "\"
        If FolderExists(tempFolderPath) = False Then
            MkDir tempFolderPath
        End If
    Next ArrayElement
End Sub

Sub FollowLink(FolderPath As String)
    If Right(FolderPath, 1) = "\" Then FolderPath = Left(FolderPath, Len(FolderPath) - 1)
    On Error Resume Next
    Dim oShell      As Object
    Dim Wnd         As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.Name = "File Explorer" Then
            If Wnd.document.Folder.Self.path = FolderPath Then Exit Sub
        End If
    Next Wnd
    Application.ThisWorkbook.FollowHyperlink Address:=FolderPath, NewWindow:=True
End Sub

Function FormatVBA7(str As String) As String
'FormatVBA7(join(filter(filter(split(aworkbook.Init(thisworkbook).Code,vbnewline),"Declare ",True ,vbTextCompare),"""" & "Declare", False,vbTextCompare),vbnewline))
    Dim selectedText
    selectedText = str
    selectedText = Replace(selectedText, " _" & vbNewLine, "")
    selectedText = Split(selectedText, vbNewLine)
    Dim IsVba7      As String
    Dim NotVba7     As String
    Dim colIsVBA7   As New Collection
    Dim colNotVBA7  As New Collection
    Dim i           As Long
    For i = LBound(selectedText) To UBound(selectedText)
        If InStr(1, selectedText(i), "PtrSafe", vbTextCompare) Then
            IsVba7 = selectedText(i)
            NotVba7 = Replace(selectedText(i), "Declare ptrsafe ", "Declare ", , , vbTextCompare)
        Else
            IsVba7 = Replace(selectedText(i), "Declare ", "Declare PtrSafe ")
            NotVba7 = selectedText(i)
        End If
        colIsVBA7.Add IsVba7
        colNotVBA7.Add NotVba7
    Next
    Set colIsVBA7 = CollectionSort(colIsVBA7)
    Set colNotVBA7 = CollectionSort(colNotVBA7)
    Dim out         As String
    out = "#If VBA7 then" & vbNewLine & _
            collectionToString(colIsVBA7, vbNewLine) & vbNewLine & _
            "#Else" & vbNewLine & _
            collectionToString(colNotVBA7, vbNewLine) & vbNewLine & _
            "#End If"
    FormatVBA7 = out

End Function

Function GetMotherBoardProp() As String

    Dim strComputer As String
    Dim objSvcs     As Object
    Dim objItms As Object, objItm As Object
    Dim vItem
    strComputer = "."
    Set objSvcs = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set objItms = objSvcs.execquery("Select * from Win32_BaseBoard")
    For Each objItm In objItms
        GetMotherBoardProp = objItm.SerialNumber
    Next

    Set objSvcs = Nothing
End Function

Public Function GetSheetByCodeName(wb As Workbook, codeName As String) As Worksheet
    Dim sh          As Worksheet
    For Each sh In wb.Worksheets
        If UCase(sh.codeName) = UCase(codeName) Then Set GetSheetByCodeName = sh: Exit For
    Next sh
End Function

Sub ImportClass( _
        Optional ClassName As String, _
        Optional TargetWorkbook As Workbook, _
        Optional Overwrite As Boolean)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If ClassName = "" Then ClassName = CodepaneSelection
    If ClassName = "" Or InStr(1, ClassName, " ") > 0 Then Exit Sub
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

    If ModuleExists(ClassName, TargetWorkbook) Then
        If Overwrite = True Then
            TargetWorkbook.VBProject.VBComponents.Remove TargetWorkbook.VBProject.VBComponents(ClassName)
        Else
            Exit Sub
        End If
    End If
    TargetWorkbook.VBProject.VBComponents.Import FilePath
End Sub


Sub ImportDeclaration( _
        Optional DeclarationName As String, _
        Optional module As VBComponent, _
        Optional TargetWorkbook As Workbook)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If DeclarationName = "" Then DeclarationName = CodepaneSelection
    If DeclarationName = "" Or InStr(1, DeclarationName, " ") > 0 Then Exit Sub
    Dim FilePath    As String
    FilePath = LOCAL_LIBRARY_DECLARATIONS & DeclarationName & ".txt"
    Dim Code        As String
    On Error Resume Next
    Code = TxtRead(FilePath)
    On Error GoTo 0

    If Len(Code) = 0 Then    'CheckPath(filePath) = "I" Then
        On Error Resume Next
        Code = TXTReadFromUrl(GITHUB_LIBRARY_DECLARATIONS & DeclarationName & ".txt")
        On Error GoTo 0
        If Len(Code) > 0 And Not UCase(Code) Like ("*NOT FOUND*") Then
            Code = FormatVBA7(Code)
            TxtOverwrite FilePath, Code
        Else
            Debug.Print "File " & DeclarationName & ".txt not found localy or online"
            Exit Sub
        End If
    Else

    End If
    If InStr(1, WorkbookCode(TargetWorkbook), Code, vbTextCompare) > 0 Then Exit Sub
    If module Is Nothing Then Set module = ModuleAddOrSet(TargetWorkbook, "vbArcImports", vbext_ct_StdModule)
    module.CodeModule.AddFromString Code

End Sub

Sub ImportProcedure( _
        Optional Procedure As String, _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Overwrite As Boolean)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If Procedure = "" Then Procedure = CodepaneSelection
    If Procedure = "" Or InStr(1, Procedure, " ") > 0 Then Exit Sub
    Dim ProcedurePath As String
    ProcedurePath = LOCAL_LIBRARY_PROCEDURES & Procedure & ".txt"

    Dim Code        As String
    On Error Resume Next
    Code = TxtRead(ProcedurePath)
    On Error GoTo 0

    If Len(Code) = 0 Then
        On Error Resume Next
        Code = TXTReadFromUrl(GITHUB_LIBRARY_PROCEDURES & Procedure & ".txt")
        On Error GoTo 0
        If Len(Code) > 0 And Not UCase(Code) Like ("*NOT FOUND*") Then
            TxtOverwrite ProcedurePath, Code
        Else
            Debug.Print "File " & Procedure & ".txt not found neither localy nor online"
            Exit Sub
        End If
    End If

    Dim filelastmod
    filelastmod = StringLastModified(Code)
    Dim proclastmod

    If ProcedureExists(TargetWorkbook, Procedure) = True Then
        Set module = ModuleOfProcedure(TargetWorkbook, Procedure)
        proclastmod = ProcedureLastModified(TargetWorkbook, module, Procedure)
        If Overwrite = True Then
            If proclastmod = 0 Or proclastmod < filelastmod Then
                ProcedureReplace module, Procedure, TxtRead(ProcedurePath)
            End If
        End If
    Else
        If module Is Nothing Then
            '            Dim ModuleName As String
            '                ModuleName = StringProcedureAssignedModule(Code)
            '            If ModuleName = "" Then ModuleName = "vbArcImports"
            '            Set Module = ModuleAddOrSet(TargetWorkbook, ModuleName, vbext_ct_StdModule)
            Set module = ModuleAddOrSet(TargetWorkbook, "vbArcImports", vbext_ct_StdModule)
        End If
        module.CodeModule.AddFromFile ProcedurePath
    End If

    ImportProcedureDependencies Procedure, TargetWorkbook, module, Overwrite
    '    ProcedureMoveToAssignedModule TargetWorkbook, Module, Procedure
End Sub

Sub ImportProcedureDependencies( _
        Optional Procedure As String, _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Overwrite As Boolean)

    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If Procedure = "" Then
        Dim cps     As String
        cps = CodepaneSelection
        If Len(cps) > 0 Then
            Procedure = cps
        Else
            Procedure = ActiveProcedure
        End If
        If Not ProcedureExists(TargetWorkbook, Procedure) Then Exit Sub
    End If
    On Error Resume Next
    If module Is Nothing Then Set module = ModuleOfProcedure(TargetWorkbook, Procedure)
    If module Is Nothing Then Exit Sub
    On Error GoTo 0
    Dim var
    Dim importfile  As String
    var = Split(ProcedureCode(TargetWorkbook, module, Procedure), vbNewLine)
    var = Filter(var, "'@INCLUDE ")
    Dim TextLine    As Variant
    For Each TextLine In var
        TextLine = Trim(TextLine)
        If TextLine Like "'@INCLUDE *" Then
            importfile = Split(TextLine, " ")(2)
            importfile = Replace(importfile, vbNewLine, "")
            If TextLine Like "'@INCLUDE PROCEDURE *" Then
                ImportProcedure importfile, TargetWorkbook, module, Overwrite
            ElseIf TextLine Like "'@INCLUDE CLASS *" Then
                ImportClass importfile, TargetWorkbook, Overwrite
            ElseIf TextLine Like "'@INCLUDE USERFORM *" Then
                ImportUserform importfile, TargetWorkbook, Overwrite
            ElseIf TextLine Like "'@INCLUDE DECLARATION *" Then
                ImportDeclaration importfile, module, TargetWorkbook
            End If
        End If
    Next
End Sub

Sub ImportUserform( _
        Optional UserformName As String, _
        Optional TargetWorkbook As Workbook, _
        Optional Overwrite As Boolean)
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If UserformName = "" Then UserformName = CodepaneSelection
    If UserformName = "" Or InStr(1, UserformName, " ") > 0 Then Exit Sub
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
            MsgBox "File " & UserformName & ".frm/.frx not found neither localy nor online"
            Exit Sub
        End If
    End If

    If ModuleExists(UserformName, TargetWorkbook) Then
        If Overwrite = True Then
            TargetWorkbook.VBProject.VBComponents.Remove TargetWorkbook.VBProject.VBComponents(UserformName)
        Else
            Exit Sub
        End If
    End If
    TargetWorkbook.VBProject.VBComponents.Import FilePathFrM
End Sub

Public Function getLastCell(TargetWorksheet As Worksheet)
    Dim cell As Range
    Set cell = TargetWorksheet.Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
    If cell Is Nothing Then Set cell = TargetWorksheet.Range("A1")
    Set getLastCell = cell
End Function

Public Function Len2( _
        ByVal val As Variant) _
        As Integer
    If IsArray(val) And Right(TypeName(val), 2) = "()" Then
        Len2 = UBound(val) - LBound(val) + 1
    ElseIf TypeName(val) = "String" Then
        Len2 = Len(val)
    ElseIf IsNumeric(val) Then
        Len2 = Len(CStr(val))
    Else
        Len2 = val.Count
    End If
End Function

Function LinkedClasses( _
        TargetWorkbook As Workbook, _
        module As VBComponent, _
        Procedure As String) As Collection

    Dim coll        As New Collection
    Dim var         As Variant
    var = classCallsOfModule(module)
    Dim Code        As String
    Code = ProcedureCode(TargetWorkbook, module, Procedure)
    Dim Keyword     As String
    Dim ClassName   As String
    Dim element     As Variant
    Dim i           As Long
    On Error Resume Next
    For i = LBound(var, 1) To UBound(var, 1)
        If InStr(1, Code, var(i, 1)) > 0 Or InStr(1, Code, var(i, 2)) > 0 Then
            coll.Add var(i, 1), var(i, 1)
        End If
    Next
    For Each element In ClassNames
        If InStr(1, Code, element) > 0 Then
            coll.Add element, CStr(element)
        End If
    Next
    On Error GoTo 0
    Set LinkedClasses = coll
End Function

Function LinkedDeclarations( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Procedure As String) As Collection

    If Not AssignCPSvariables(TargetWorkbook, module, Procedure) Then Stop

    DeclarationsTableCreate TargetWorkbook

    Dim TargetWorksheet As Worksheet: Set TargetWorksheet = ThisWorkbook.Sheets("Declarations_Table")
    Dim coll        As New Collection
    Dim Code        As String: Code = ProcedureCode(TargetWorkbook, module, Procedure)
    Dim element
    For Each element In DeclarationsTableKeywords
        If RegexTest(Code, "\b ?" & CStr(element) & "\b") Then
            On Error Resume Next
            coll.Add CStr(element), CStr(element)
            On Error GoTo 0
        End If
    Next
    Set LinkedDeclarations = coll
End Function

Function LinkedProcedures( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional ProcedureName As String) As Collection
    If Not AssignCPSvariables(TargetWorkbook, module, ProcedureName) Then Stop
    Dim Procedures  As Collection
    Set Procedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Code        As String
    Code = ProcedureCode(TargetWorkbook, module, ProcedureName)
    Dim coll        As New Collection
    Dim Procedure   As Variant
    For Each Procedure In Procedures
        If UCase(CStr(Procedure)) <> UCase(CStr(ProcedureName)) Then
            If RegexTest(Code, "\W" & CStr(Procedure) & "[.(\W]") = True Then
                coll.Add Procedure, CStr(Procedure)
            End If
        End If
    Next
    Set LinkedProcedures = coll
End Function

Function LinkedProceduresDeep( _
        ProcedureName As Variant, _
        TargetWorkbook As Workbook) As Collection

    Dim AllProcedures As Collection: Set AllProcedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Processed   As Collection: Set Processed = New Collection
    Dim CalledProcedures As Collection: Set CalledProcedures = New Collection

    Dim Procedure   As Variant
    Dim module      As VBComponent

    Processed.Add CStr(ProcedureName), CStr(ProcedureName)
    On Error Resume Next
    For Each Procedure In LinkedProcedures(TargetWorkbook, , CStr(ProcedureName))
        CalledProcedures.Add CStr(Procedure), CStr(Procedure)
    Next
    On Error GoTo 0

    Dim CalledProceduresCount As Long
    CalledProceduresCount = CalledProcedures.Count
    Dim element
repeat:
    For Each element In CalledProcedures
        If Not CollectionContains(Processed, , CStr(element)) Then
            On Error Resume Next
            For Each Procedure In LinkedProcedures(TargetWorkbook, , CStr(element))
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

    Set LinkedProceduresDeep = CollectionSort(CalledProcedures)
End Function


Sub LinkedProceduresMoveHere(Optional Procedure As String)
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = ActiveCodepaneWorkbook
    If Not AssignProcedureVariable(TargetWorkbook, Procedure) Then Exit Sub
    Dim el
    For Each el In LinkedProceduresDeep(Procedure, TargetWorkbook)
        ProcedureMoveHere CStr(el)
    Next
End Sub




Function LinkedUserforms( _
        TargetWorkbook As Workbook, _
        module As VBComponent, _
        Procedure As String) As Collection
    Dim coll        As New Collection
    Dim Code        As String
    Code = ProcedureCode(TargetWorkbook, module, Procedure)
    Dim FormName
    For Each FormName In UserformNames(TargetWorkbook)
        If RegexTest(Code, "\W" & FormName & "[.(\W]") = True Then coll.Add FormName    '& " " & "(Userform)"
    Next
    Set LinkedUserforms = coll
End Function

Function ModuleAddOrSet( _
        TargetWorkbook As Workbook, _
        TargetName As String, _
        moduleType As VBIDE.vbext_ComponentType) As VBComponent


    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim module      As VBComponent
    On Error Resume Next
    Set module = TargetWorkbook.VBProject.VBComponents(TargetName)
    On Error GoTo 0
    If module Is Nothing Then
        Set module = TargetWorkbook.VBProject.VBComponents.Add(moduleType)
        module.Name = TargetName
    End If
    Set ModuleAddOrSet = module
End Function




Function ModuleCode(module As VBComponent) As String
    With module.CodeModule
        If .CountOfLines = 0 Then ModuleCode = "": Exit Function
        ModuleCode = .Lines(1, .CountOfLines)
    End With
End Function

Public Function ModuleExists( _
        TargetName As String, _
        TargetWorkbook As Workbook) As Boolean
    Dim module      As VBComponent
    On Error Resume Next
    Set module = TargetWorkbook.VBProject.VBComponents(TargetName)
    On Error GoTo 0
    ModuleExists = Not module Is Nothing
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 25-10-2023 13:24    Alex                (Dependencies.bas > ModuleOfProcedure)

Public Function ModuleOfProcedure( _
        TargetWorkbook As Workbook, _
        ProcedureName As Variant) As VBComponent
'@LastModified 2310251324
    Dim ProcKind    As VBIDE.vbext_ProcKind
    Dim lineNum As Long, NumProc As Long
    Dim Procedure   As String
    Dim module      As VBComponent
    For Each module In TargetWorkbook.VBProject.VBComponents
        If module.Type = vbext_ct_StdModule Then
            With module.CodeModule
                lineNum = .CountOfDeclarationLines + 1
                Do Until lineNum >= .CountOfLines
                    Procedure = .ProcOfLine(lineNum, ProcKind)
                    If UCase(Procedure) = UCase(ProcedureName) Then
                        Set ModuleOfProcedure = module
                        Exit Function
                    End If
                    lineNum = .procStartLine(Procedure, ProcKind) + .ProcCountLines(Procedure, ProcKind) + 1
                Loop
            End With
        End If
    Next module
End Function

Function ModuleOrSheetName(module As VBComponent) As String
    If module.Type = vbext_ct_Document Then
        If module.Name = "ThisWorkbook" Then
            ModuleOrSheetName = module.Name
        Else
            ModuleOrSheetName = GetSheetByCodeName(WorkbookOfModule(module), module.Name).Name
        End If
    Else
        ModuleName = module.Name
    End If
End Function

Function ModuleTypeToString(componentType As VBIDE.vbext_ComponentType) As String
    Select Case componentType
        Case vbext_ct_ActiveXDesigner
            ModuleTypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            ModuleTypeToString = "Class"
        Case vbext_ct_Document
            ModuleTypeToString = "Document"
        Case vbext_ct_MSForm
            ModuleTypeToString = "UserForm"
        Case vbext_ct_StdModule
            ModuleTypeToString = "Module"
        Case Else
            ModuleTypeToString = "Unknown Type: " & CStr(componentType)
    End Select
End Function

Function ProcedureAssignedModule( _
        TargetWorkbook As Workbook, _
        module As VBComponent, _
        Procedure As String) As VBComponent
    Dim ComponentName As Variant
    ComponentName = Split(ProcedureCode(TargetWorkbook, module, Procedure), vbNewLine)
    ComponentName = Filter(ComponentName, "'@AssignedModule")
    If Len2(ComponentName) <> 1 Then Exit Function
    Dim ub          As Long
    ub = UBound(Split(ComponentName(0), " "))
    ComponentName = Split(ComponentName(0), " ")(ub)
    Set ProcedureAssignedModule = ModuleAddOrSet(TargetWorkbook, CStr(ComponentName), vbext_ct_StdModule)
End Function

Sub ProcedureAssignedModuleAdd( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, Procedure) Then Stop
    ProcedureLinesRemove "'@AssignedModule *", TargetWorkbook, module, Procedure
    module.CodeModule.InsertLines ProcedureBodyLineFirstAfterComments(module, Procedure), _
            "'@AssignedModule " & module.Name
End Sub

Function ProcedureBodyLineFirst( _
        module As VBComponent, _
        Procedure As String) As Long
    ProcedureBodyLineFirst = ProcedureTitleLineFirst(module, Procedure) + ProcedureTitleLineCount(module, Procedure)
End Function

Function ProcedureBodyLineFirstAfterComments( _
        module As VBComponent, _
        Procedure As String) As Long
    Dim N           As Long
    Dim S           As String
    For N = ProcedureBodyLineFirst(module, Procedure) To module.CodeModule.CountOfLines
        S = Trim(module.CodeModule.Lines(N, 1))
        If S = vbNullString Then
            Exit For
        ElseIf Left(S, 1) = "'" Then
        ElseIf Left(S, 3) = "Rem" Then
        ElseIf Right(Trim(module.CodeModule.Lines(N - 1, 1)), 1) = "_" Then
        ElseIf Right(S, 1) = "_" Then
        Else
            Exit For
        End If
    Next N
    ProcedureBodyLineFirstAfterComments = N
End Function



Public Function ProcedureCode( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Procedure As Variant, _
        Optional IncludeHeader As Boolean = True) As String
    If Not AssignCPSvariables(TargetWorkbook, module, CStr(Procedure)) Then Exit Function
    Dim lProcStart  As Long
    Dim lProcBodyStart As Long
    Dim lProcNoLines As Long
    Const vbext_pk_Proc = 0
    On Error GoTo Error_Handler
    lProcStart = module.CodeModule.procStartLine(Procedure, vbext_pk_Proc)
    lProcBodyStart = module.CodeModule.ProcBodyLine(Procedure, vbext_pk_Proc)
    lProcNoLines = module.CodeModule.ProcCountLines(Procedure, vbext_pk_Proc)
    If IncludeHeader = True Then
        ProcedureCode = module.CodeModule.Lines(lProcStart, lProcNoLines)
    Else
        lProcNoLines = lProcNoLines - (lProcBodyStart - lProcStart)
        ProcedureCode = module.CodeModule.Lines(lProcBodyStart, lProcNoLines)
    End If
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
Error_Handler:
    Debug.Print "Error Source: ProcedureCode" & vbCrLf & _
            "Error Description: " & Err.Description & _
            Switch(Erl = 0, vbNullString, Erl <> 0, vbCrLf & "Line No: " & Erl)
    Resume Error_Handler_Exit
End Function

Function ProcedureExists( _
        TargetWorkbook As Workbook, _
        ProcedureName As Variant) As Boolean
    Dim Procedures  As Collection
    Set Procedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Procedure   As Variant
    For Each Procedure In Procedures
        If UCase(CStr(Procedure)) = UCase(ProcedureName) Then
            ProcedureExists = True
            Exit Function
        End If
    Next
End Function

Function ProcedureLastModAdd( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Procedure As String, _
        Optional ModificationDate As Double)



    If Not AssignCPSvariables(TargetWorkbook, module, Procedure) Then Exit Function
    If ModificationDate = 0 Then ModificationDate = Format(Now, "yymmddhhnn")
    Dim LastModLine As Long
    LastModLine = ProcedureLineContaining(module, Procedure, "'@LastModified *")
    If LastModLine = 0 Then GoTo PASS
    Dim LDate       As Double
    LDate = Split(module.CodeModule.Lines(LastModLine, 1), " ")(1)
    ProcedureLinesRemove "'@LastModified *", TargetWorkbook, module, Procedure
PASS:
    module.CodeModule.InsertLines ProcedureBodyLineFirst(module, Procedure), _
            "'@LastModified " & ModificationDate

    ProcedureLastModAdd = ModificationDate
End Function

Function ProcedureLastModified( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, Procedure) Then Stop
    ProcedureLastModified = StringLastModified(ProcedureCode(TargetWorkbook, module, Procedure))
End Function

Function ProcedureLinesCount( _
        module As VBComponent, _
        Procedure As String) As Long
    ProcedureLinesCount = module.CodeModule.ProcCountLines(Procedure, vbext_pk_Proc)
End Function

Public Function ProcedureLinesFirst( _
        module As VBComponent, _
        Procedure As String) As Long
    Dim ProcKind    As VBIDE.vbext_ProcKind
    ProcKind = vbext_pk_Proc
    ProcedureLinesFirst = module.CodeModule.procStartLine(Procedure, ProcKind)
End Function


Public Function ProcedureLinesLast( _
        module As VBComponent, _
        Procedure As String, _
        Optional IncludeTail As Boolean) As Long
    Dim ProcKind    As VBIDE.vbext_ProcKind
    ProcKind = vbext_pk_Proc
    Dim startAt     As Long
    startAt = module.CodeModule.procStartLine(Procedure, ProcKind)
    Dim CountOf     As Long
    CountOf = module.CodeModule.ProcCountLines(Procedure, ProcKind)
    Dim endAt       As Long
    endAt = startAt + CountOf - 1
    If Not IncludeTail Then
        Do While Not Trim(module.CodeModule.Lines(endAt, 1)) Like "End *"
            endAt = endAt - 1
        Loop
    End If
    ProcedureLinesLast = endAt
End Function

Sub ProcedureLinesRemove( _
        myCriteria As String, _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, Procedure) Then Stop

    Dim Code        As String
    Dim i           As Long
    For i = ProcedureLinesLast(module, Procedure) To ProcedureLinesFirst(module, Procedure) Step -1
        Code = Trim(module.CodeModule.Lines(i, 1))
        If Code Like myCriteria Then module.CodeModule.DeleteLines i
    Next
End Sub

Sub ProcedureLinesRemoveInclude( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, Procedure) Then Stop
    ProcedureLinesRemove "'@INCLUDE", TargetWorkbook, module, Procedure
End Sub


Sub ProcedureMoveHere( _
        Optional Procedure As String)


    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = ActiveCodepaneWorkbook
    If Not AssignProcedureVariable(TargetWorkbook, Procedure) Then Exit Sub
    Dim module      As VBComponent
    Set module = ModuleOfProcedure(TargetWorkbook, Procedure)
    Dim S           As String
    S = ProcedureCode(TargetWorkbook, module, Procedure)

    If InStr(1, S, "'@AssignedModule") = 0 Then
        ProcedureAssignedModuleAdd TargetWorkbook, module, Procedure
        S = ProcedureCode(TargetWorkbook, module, Procedure)
    End If

    Dim sl As Long, cl As Long
    sl = ProcedureLinesFirst(module, Procedure)
    cl = ProcedureLinesLast(module, Procedure, False) - sl + 1
    ActiveModule.CodeModule.InsertLines ProcedureLinesLast(module, ActiveProcedure, True) + 1, S
    module.CodeModule.DeleteLines sl, cl
End Sub

Sub ProcedureMoveToAssignedModule( _
        Optional TargetWorkbook As Workbook, _
        Optional module As VBComponent, _
        Optional Procedure As String)
    If Not AssignCPSvariables(TargetWorkbook, module, Procedure) Then Exit Sub
    Dim MoveToModule As VBComponent
    Set MoveToModule = ProcedureAssignedModule(TargetWorkbook, module, Procedure)
    If MoveToModule Is Nothing Then Exit Sub
    ProcedureMoveToModule TargetWorkbook, module, Procedure, MoveToModule
End Sub

Sub ProcedureMoveToModule( _
        TargetWorkbook As Workbook, _
        module As VBComponent, _
        Procedure As String, _
        MoveToModule As VBComponent)
    Dim Code        As String
    Code = ProcedureCode(TargetWorkbook, module, Procedure)
    Dim startLine   As Long
    startLine = ProcedureLinesFirst(module, Procedure)
    Dim CountLines  As Long
    CountLines = ProcedureLinesCount(module, Procedure)
    MoveToModule.CodeModule.InsertLines MoveToModule.CodeModule.CountOfLines + 1, vbNewLine & Code
    module.CodeModule.DeleteLines startLine, CountLines

End Sub

Public Sub ProcedureReplace( _
        module As VBComponent, _
        Procedure As String, _
        Code As String)

    Dim startLine   As Integer
    Dim NumLines    As Integer
    With module.CodeModule
        startLine = .procStartLine(Procedure, vbext_pk_Proc)
        NumLines = .ProcCountLines(Procedure, vbext_pk_Proc)
        .DeleteLines startLine, NumLines
        .InsertLines startLine, Code
    End With
End Sub

Function ProcedureTitle( _
        module As VBComponent, _
        Procedure As String) As String
    Dim titleLine   As Long
    titleLine = ProcedureTitleLineFirst(module, Procedure)
    Dim Title       As String
    Title = module.CodeModule.Lines(titleLine, 1)
    Dim counter     As Long
    counter = 1
    Do While Right(Title, 1) = "_"
        counter = counter + 1
        Title = module.CodeModule.Lines(titleLine, counter)
    Loop

    ProcedureTitle = Title
End Function

Function ProcedureTitleLineCount( _
        module As VBComponent, _
        Procedure As String) As Long

    ProcedureTitleLineCount = ProcedureTitleLineLast(module, Procedure) - ProcedureTitleLineFirst(module, Procedure) + 1
End Function



Public Function ProcedureTitleLineFirst( _
        module As VBComponent, _
        Procedure As String) As Long
    ProcedureTitleLineFirst = module.CodeModule.ProcBodyLine(Procedure, vbext_pk_Proc)
End Function

Function ProcedureTitleLineLast( _
        module As VBComponent, _
        Procedure As String) As Long
    ProcedureTitleLineLast = ProcedureTitleLineFirst(module, Procedure) + UBound(Split(ProcedureTitle(module, Procedure), vbNewLine))
End Function

Public Function ProceduresOfModule( _
        module As VBComponent) As Collection
    Dim ProcKind    As VBIDE.vbext_ProcKind
    Dim lineNum     As Long
    Dim coll        As New Collection
    Dim Procedure   As String
    With module.CodeModule
        lineNum = .CountOfDeclarationLines + 1
        Do Until lineNum >= .CountOfLines
            ProcedureAs = .ProcOfLine(lineNum, ProcKind)
            coll.Add ProcedureAs
            lineNum = .procStartLine(ProcedureAs, ProcKind) + .ProcCountLines(ProcedureAs, ProcKind) + 1
        Loop
    End With
    Set ProceduresOfModule = coll
End Function

Function ProceduresOfTXT( _
        Code As String) As Collection


    Code = Replace(Code, vbNewLine, vbLf)
    Dim var
    var = Split(Code, vbLf)

    Dim out
    out = ArrayAppend(Filter(var, "Sub" & Space(1), True, vbBinaryCompare), Filter(var, "Function ", True, vbBinaryCompare))
    If TypeName(out) = "Empty" Then Exit Function
    out = Filter(out, "(", True)
    out = Filter(out, "Declare", False)
    out = Filter(out, Chr(34) & "Sub", False)
    out = Filter(out, Chr(34) & "Function", False)
    out = Filter(out, "End Sub", False)
    out = Filter(out, "End Function", False)

    Dim i           As Long
    For i = LBound(out) To UBound(out)
        out(i) = Left(out(i), InStr(1, out(i), "(") - 1)
        out(i) = Replace(out(i), "Private ", "")
        out(i) = Replace(out(i), "Public ", "")
        out(i) = Replace(out(i), "Sub ", "")
        out(i) = Replace(out(i), "Function ", "")
        If UBound(Split(out(i), " ")) > 0 Then
            out(i) = ""
        End If
    Next

    ArrayQuickSort out
    out = cleanArray(out)
    out = ArrayDuplicatesRemove(out)
    Set ProceduresOfTXT = ArrayToCollection(out)
End Function

Sub CallSeparateProcedures()
    Dim FilePath    As Variant
    FilePath = DataFilePicker("*.txt", False)
    If FilePath = vbNullString Then Exit Sub
    Dim OutputFolder As Variant
    OutputFolder = SelectFolder(Left(FilePath, InStrRev(FilePath, "\")))

    TxtSeparateProcedures FilePath, OutputFolder

End Sub

Sub TxtSeparateProcedures(FilePath As Variant, Optional OutputFolder As Variant)

    '@AssignedModule F_FileFolder
    '@INCLUDE PROCEDURE TxtOverwrite
    '@INCLUDE PROCEDURE TxtRead
    Dim fname       As String
    If OutputFolder = "" Then
        OutputFolder = Left(FilePath, InStrRev(FilePath, "\"))
    Else
        FoldersCreate CStr(OutputFolder)
    End If
    Dim Code        As Variant
    Code = Split(TxtRead(FilePath), vbLf)
    Dim out         As String
    Dim i           As Long
    For i = LBound(Code) To UBound(Code)

        out = IIf(out = "", Code(i), out & Code(i)) & vbNewLine
        If RegexTest(Code(i), "Sub ") _
                And Not Code(i) Like Chr(34) & "*Sub*" Then
            fname = Split(Code(i), "Sub ")(1)
            fname = Trim(Split(fname, "(")(0)) & ".txt"
        ElseIf RegexTest(Code(i), "Function ") _
                And Not Code(i) Like Chr(34) & "*Function*" Then
            fname = Split(Code(i), "Function ")(1)
            fname = Trim(Split(fname, "(")(0)) & ".txt"
        End If
        If Trim(Code(i)) = "End Sub" Or Trim(Code(i)) = "End Function" Then
            TxtOverwrite OutputFolder & fname, out
            out = ""
            fname = ""
        End If
    Next

End Sub

Function ProceduresOfWorkbook( _
        TargetWorkbook As Workbook, _
        Optional ExcludeDocument As Boolean = True, _
        Optional ExcludeClass As Boolean = True, _
        Optional ExcludeForm As Boolean = True) As Collection
    Dim module      As VBComponent
    Dim ProcKind    As VBIDE.vbext_ProcKind
    Dim lineNum     As Long
    Dim coll        As New Collection
    Dim ProcedureName As String
    For Each module In TargetWorkbook.VBProject.VBComponents
        If ExcludeClass = True And module.Type = vbext_ct_ClassModule Then GoTo SKIP
        If ExcludeDocument = True And module.Type = vbext_ct_Document Then GoTo SKIP
        If ExcludeForm = True And module.Type = vbext_ct_MSForm Then GoTo SKIP
        With module.CodeModule
            lineNum = .CountOfDeclarationLines + 1
            Do Until lineNum >= .CountOfLines
                ProcedureName = .ProcOfLine(lineNum, ProcKind)
                If InStr(1, ProcedureName, "_") = 0 Then
                    coll.Add ProcedureName
                End If
                lineNum = .procStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
            Loop
        End With
SKIP:
    Next module
    Set ProceduresOfWorkbook = coll
End Function

Sub ProjetFoldersCreate()
    Dim element
    For Each element In vbarcFolders
        FoldersCreate CStr(element)
    Next
End Sub

Public Function RegexTest( _
        ByVal string1 As String, _
        ByVal stringPattern As String, _
        Optional ByVal globalFlag As Boolean, _
        Optional ByVal ignoreCaseFlag As Boolean, _
        Optional ByVal multilineFlag As Boolean) As Boolean
    Dim REGEX       As Object
    Set REGEX = CreateObject("VBScript.RegExp")
    With REGEX
        .Global = globalFlag
        .IgnoreCase = ignoreCaseFlag
        .MultiLine = multilineFlag
        .pattern = stringPattern
    End With
    RegexTest = REGEX.test(string1)
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 11:03    Alex                (Dependencies.bas > StringLastModified)

Function StringLastModified(txt As String)
'@LastModified 2308221103

    Dim Code        As Variant
    Code = Filter(Filter(Split(txt, vbLf), "'@LastModified ", True), """", False)
    If ArrayAllocated(Code) Then
        Dim lastDate As Variant
        If Trim(Code(0)) Like "'@LastModified *" Then
            lastDate = Split(Code(0), " ")(1)
            lastDate = DateSerial(Left(lastDate, 2), Mid(lastDate, 3, 2), Mid(lastDate, 5, 2)) _
                    & " " & TimeSerial(Mid(lastDate, 7, 2), Mid(lastDate, 9, 2), 0)
            StringLastModified = Split(Code(0), " ")(1)
        End If
    Else

    End If
End Function

Function StringProcedureAssignedModule(txt As String) As String
    Dim ComponentName As Variant
    ComponentName = Split(txt, vbLf)
    ComponentName = Filter(ComponentName, "'@AssignedModule")
    If Not ArrayAllocated(ComponentName) Then Exit Function
    Dim ub          As Long
    ub = UBound(Split(ComponentName(0), " "))
    ComponentName = Split(ComponentName(0), " ")(ub)
    StringProcedureAssignedModule = ComponentName
End Function



Function TXTReadFromUrl(url As String) As String
    On Error GoTo Err_GetFromWebpage
    Dim objWeb      As Object
    Dim strXML      As String
    Set objWeb = CreateObject("Msxml2.ServerXMLHTTP")
    objWeb.Open "GET", url, False
    objWeb.setRequestHeader "Content-Type", "text/xml"
    objWeb.setRequestHeader "Cache-Control", "no-cache"
    objWeb.setRequestHeader "Pragma", "no-cache"
    objWeb.send
    Do While objWeb.readyState <> 4
        DoEvents
    Loop
    strXML = objWeb.responseText
    TXTReadFromUrl = strXML
End_GetFromWebpage:
    Set objWeb = Nothing
    Exit Function
Err_GetFromWebpage:
    MsgBox Err.Description & " (" & Err.Number & ")"
    Resume End_GetFromWebpage
End Function

Sub TxtOverwrite(sFile As String, sText As String)
    On Error GoTo ERR_HANDLER
    Dim FileNumber  As Integer
    FileNumber = FreeFile
    Open sFile For Output As #FileNumber
    Print #FileNumber, sText
    Close #FileNumber
Exit_Err_Handler:
    Exit Sub
ERR_HANDLER:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: TxtOverwrite" & vbCrLf & _
            "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Sub

Sub TxtPrepend(FilePath As String, txt As String)
    Dim S           As String
    S = TxtRead(FilePath)
    TxtOverwrite FilePath, txt & vbNewLine & S
End Sub


Sub CallTxtPrependContainedProcedures()
    Dim FilePath    As Variant
    FilePath = DataFilePicker("*.txt", False)
    If FilePath = vbNullString Then Exit Sub
    TxtPrependContainedProcedures CStr(FilePath)
End Sub

Sub TxtPrependContainedProcedures(fileName As String)
    Dim S           As String: S = TxtRead(fileName)
    Dim V           As New Collection
    Set V = ProceduresOfTXT(S)
    If V.Count = 0 Then Exit Sub
    Dim line        As String: line = String(30, "'")
    TxtPrepend fileName, _
            "'Contains the following " & "#" & V.Count & " procedures " & vbNewLine & line & vbNewLine & _
            "'" & collectionToString(V, vbNewLine & "'") & vbNewLine & line & vbNewLine & vbNewLine
End Sub

Function TxtRead(sPath As Variant) As String
    Dim sTXT        As String
    If Dir(sPath) = "" Then
        Debug.Print "File was not found."
        Debug.Print sPath
        Exit Function
    End If
    Open sPath For Input As #1
    Do Until EOF(1)
        Line Input #1, sTXT
        TxtRead = TxtRead & sTXT & vbLf
    Loop
    Close
    If Len(TxtRead) = 0 Then
        TxtRead = ""
    Else
        TxtRead = Left(TxtRead, Len(TxtRead) - 1)
    End If
End Function

Function URLExists(url) As Boolean
    Dim Request     As Object
    Dim FF          As Integer
    Dim rc          As Variant

    On Error GoTo EndNow
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")

    With Request
        .Open "GET", url, False
        .send
        rc = .statusText
    End With
    Set Request = Nothing
    If rc = "OK" Then URLExists = True

    Exit Function
EndNow:
End Function

Function UserformNames(TargetWorkbook As Workbook)
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Set UserformNames = ComponentNames(vbext_ct_MSForm, TargetWorkbook)
End Function






Function WorkbookCode(TargetWorkbook) As String
    If TypeName(TargetWorkbook) <> "Workbook" Then Stop
    Dim module      As VBComponent
    Dim txt
    For Each module In TargetWorkbook.VBProject.VBComponents
        If module.CodeModule.CountOfLines > 0 Then
            txt = txt & _
                    vbNewLine & _
                    "'" & String(10, "=") & ModuleOrSheetName(module) & " (" & module.Type & ") " & String(10, "=") & _
                    vbNewLine & _
                    ModuleCode(module)
        End If
    Next
    WorkbookCode = txt
End Function


Function WorkbookOfModule(vbComp As VBComponent) As Workbook
    Set WorkbookOfModule = WorkbookOfProject(vbComp.Collection.Parent)
End Function

Function WorkbookOfProject(vbProj As VBProject) As Workbook
    tmpStr = vbProj.fileName
    tmpStr = Right(tmpStr, Len(tmpStr) - InStrRev(tmpStr, "\"))
    Set WorkbookOfProject = Workbooks(tmpStr)
End Function



Function WorksheetExists(SheetName As String, TargetWorkbook As Workbook) As Boolean
    Dim TargetWorksheet As Worksheet
    On Error Resume Next
    Set TargetWorksheet = TargetWorkbook.Sheets(SheetName)
    On Error GoTo 0
    WorksheetExists = Not TargetWorksheet Is Nothing
End Function

Function classCallsOfModule(module As VBComponent) As Variant


    Dim Code        As Variant
    Dim element     As Variant
    Dim Keyword     As Variant
    Dim var         As Variant
    ReDim var(1 To 2, 1 To 1)
    Dim counter     As Long
    counter = 0
    If module.CodeModule.CountOfDeclarationLines > 0 Then
        Code = module.CodeModule.Lines(1, module.CodeModule.CountOfDeclarationLines)
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

                For Each ClassName In ClassNames
                    If element = ClassName Then

                        ReDim Preserve var(1 To 2, 1 To counter + 1)
                        var(1, UBound(var, 2)) = element
                        var(2, UBound(var, 2)) = Keyword
                        counter = counter + 1
                    End If
                Next
            End If
        Next
        If var(1, 1) <> "" Then

            If UBound(var, 2) > 1 Then
                classCallsOfModule = WorksheetFunction.Transpose(var)
            Else
                Dim VAR2(1 To 1, 1 To 2)
                VAR2(1, 1) = var(1, 1)
                VAR2(1, 2) = var(2, 1)
                classCallsOfModule = VAR2
            End If
        End If
    End If

End Function

Function collectionToString(coll As Collection, delim As String) As String
    Dim element
    Dim out         As String
    For Each element In coll
        out = IIf(out = "", element, out & delim & element)
    Next
    collectionToString = out
End Function

Function getDeclarations( _
        wb As Workbook, _
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
    For Each module In wb.VBProject.VBComponents
        If module.Type = vbext_ct_StdModule Or module.Type = vbext_ct_MSForm Then
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
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(element), "\b ?Type \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Type")
                            DeclarationsCollection.Add DeclarationsKeywordSubstring(tmp, , "Type " & KeywordsCollection.item(KeywordsCollection.Count), "End Type", , , True)
                            TypeCollection.Add "Type"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf InStr(1, CStr(element), "Const ", vbTextCompare) > 0 Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Const")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Const"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(element), "\b ?Sub \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Sub")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Sub"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf RegexTest(CStr(element), "\b ?Function \b") Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", "Function")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Function"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf element Like "*(*) As *" Then
                            helper = Left(element, InStr(1, CStr(element), "(") - 1)
                            helper = Mid(helper, InStrRev(helper, " ") + 1)
                            KeywordsCollection.Add helper
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Other"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        ElseIf element Like "* As *" Then
                            KeywordsCollection.Add DeclarationsKeywordSubstring(CStr(element), " ", , "As")
                            DeclarationsCollection.Add CStr(element)
                            TypeCollection.Add "Other"
                            ComponentCollection.Add module.Name
                            ComponentTypecollection.Add ModuleTypeToString(module.Type)
                            ScopeCollection.Add IIf(InStr(1, DeclarationsCollection.item(DeclarationsCollection.Count), "Public", vbTextCompare), "Public", "Private")
                        Else
                        End If
                    End If
                Next element
            End If
        End If
    Next module

    If includeScope = True Then Output.Add ScopeCollection
    If includeType = True Then Output.Add TypeCollection
    If includeKeywords = True Then Output.Add KeywordsCollection
    If includeDeclarations = True Then Output.Add DeclarationsCollection
    If includeComponentType = True Then Output.Add ComponentTypecollection
    If includeComponentName = True Then Output.Add ComponentCollection

    Set getDeclarations = Output
End Function

Function getLastRow(TargetSheet As Worksheet)
    Dim LastCell    As Range
    Set LastCell = TargetSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    getLastRow = LastCell.Row
End Function

Function vbarcFolders() As Collection
    Dim coll        As New Collection
    coll.Add LOCAL_LIBRARY_PROCEDURES
    coll.Add LOCAL_LIBRARY_CLASSES
    coll.Add LOCAL_LIBRARY_USERFORMS
    coll.Add LOCAL_LIBRARY_DECLARATIONS
    Set vbarcFolders = coll
End Function

Function ProcedureLineContaining(module As VBComponent, Procedure As String, this As String) As Long
    Dim i           As Long
    For i = ProcedureLinesFirst(module, Procedure) To ProcedureLinesLast(module, Procedure)
        If module.CodeModule.Lines(i, 1) Like this Then
            ProcedureLineContaining = i
            Exit Function
        End If
    Next
End Function

Public Function ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Function

Public Function CLIP(Optional StoreText As String) As String
    Dim X           As Variant
    X = StoreText
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(StoreText)
                    .SetData "text", X
                Case Else
                    CLIP = .GetData("text")
            End Select
        End With
    End With
End Function


