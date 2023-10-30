VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uSkeleton 
   Caption         =   "github.com/alexofrhodes"
   ClientHeight    =   9600.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   18960
   OleObjectBlob   =   "uSkeleton.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uSkeleton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Userform   : uSkeleton
'* Created    : 06-10-2022 10:40
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Option Explicit


Dim skellyBook      As Workbook
Dim skellyModule    As VBComponent
Dim skellyProcName

Private Sub exportDeclarationsAndCalls_Click()
    '@AssignedModule uSkeleton
    '@INCLUDE PROCEDURE ArrayToRange2D
    '@INCLUDE PROCEDURE CollectionsToArray2D
    '@INCLUDE PROCEDURE getDeclarations
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE CLASS aCollection
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyBook
    setUp
    Dim var         As Variant
    Dim collections As New Collection
    Set collections = FindCalls(skellyBook)
    If collections(1).Count > 0 Then
        var = aCollection.CollectionsToArray2D(collections)
        ArrayToRange2D var, dataToSheet(skellyBook, "exportCalls", "A2")
        With skellyBook.Sheets("exportCalls").Range("A1:B1")
            .value = Array("Procedure", "Calls")
            .Font.Bold = True
            .Font.Size = 14
        End With
        With skellyBook.Sheets("exportCalls").Cells
            .WrapText = False
            .Columns.AutoFit
            .WrapText = True
            .Columns.AutoFit
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
    Set collections = aWorkbook.Init(skellyBook).getDeclarations(True, True, True, True, True, True)
    If collections(1).Count > 0 Then
        var = aCollection.CollectionsToArray2D(collections)
        ArrayToRange2D var, dataToSheet(skellyBook, "exportDeclarations", "A2")
        With skellyBook.Sheets("exportDeclarations").Range("A1:F1")
            .value = Array("Component Type", "Component Name", "Declaration Scope", "Declaration Type", "Declaration Keyword", "Declaration Code")
            .Font.Bold = True
            .Font.Size = 14
        End With
        With skellyBook.Sheets("exportDeclarations").Cells
            .WrapText = False
            .Columns.AutoFit
            .WrapText = True
            .Columns.AutoFit
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
End Sub

Function AddProcedureCallsToCollectionSkeleton(skellyBook As Workbook, skellyModule As VBComponent, ProcName As String) As Collection
    '@AssignedModule uSkeleton
    '@INCLUDE PROCEDURE ProcedureCode
    '@INCLUDE PROCEDURE ProceduresOfWorkbook
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyBook
    '@INCLUDE DECLARATION skellyModule
    Dim coll        As Collection
    Set coll = New Collection
    Dim WorkbookProcedure As Variant
    Dim AllProcs    As Collection
    Set AllProcs = ProceduresOfWorkbook(skellyBook)
    Dim procText    As String
    procText = ProcedureCode(skellyBook, skellyModule, ProcName)
    For Each WorkbookProcedure In AllProcs
        If CStr(WorkbookProcedure) <> ProcName Then
            If InStr(1, procText, CStr(WorkbookProcedure)) Then
                coll.Add CStr(WorkbookProcedure)
            End If
        End If
    Next WorkbookProcedure
    Set AddProcedureCallsToCollectionSkeleton = coll
End Function



Private Sub UserForm_Initialize()
    '@AssignedModule uSkeleton
    '@INCLUDE USERFORM uSkeleton
    LoadProjects
End Sub
Private Sub GetInfo_Click()
    '@AssignedModule uSkeleton
    '@INCLUDE PROCEDURE PlayTheSound
    '@INCLUDE CLASS aUserform
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE USERFORM uAuthor
    With aUserform.Init(Me)
        .Transition .Effect(GetInfo, "Top", GetInfo.Top - 10, 150)

        .Transition .Effect(GetInfo, "Top", GetInfo.Top + 10, 150)
    End With
    uAuthor.Show
End Sub

Sub LoadProjects()
    '@INCLUDE WorkbookProjectProtected
    '@AssignedModule uSkeleton
    '@INCLUDE PROCEDURE WorkbookProjectProtected
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyBook
    For Each skellyBook In Workbooks
        If Not WorkbookProjectProtected(skellyBook) Then LProjects.AddItem skellyBook.Name
    Next
    On Error Resume Next
    Dim ad
    For Each ad In AddIns
        If Not WorkbookProjectProtected(Workbooks(ad.Name)) Then
            If Err = 0 Then LProjects.AddItem ad.Name
            Err.clear
        End If
    Next
End Sub

Private Sub LProjects_Click()
    '@AssignedModule uSkeleton
    '@INCLUDE USERFORM uSkeleton
    loadComponents
End Sub

Sub loadComponents()
    '@INCLUDE ModuleTypeToString
    '@INCLUDE SortListboxOnColumn
    '@INCLUDE setUp
    '@INCLUDE ReleaseMe
    '@INCLUDE ControlsResizeColumns
    '@AssignedModule uSkeleton
    '@INCLUDE PROCEDURE ModuleTypeToString
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyBook
    '@INCLUDE DECLARATION skellyModule
    LComponents.clear: LProcedures.clear: TPROCS.TEXT = "": LCalls.clear: TCalls.TEXT = "": LDeclarations.clear: TDeclarations.TEXT = "":
    setUp
    For Each skellyModule In skellyBook.VBProject.VBComponents
        LComponents.AddItem
        LComponents.list(LComponents.ListCount - 1, 0) = aModule.Init(skellyModule).TypeToString
        LComponents.list(LComponents.ListCount - 1, 1) = skellyModule.Name
    Next
    ReleaseMe
    aListBox.Init(LComponents).SortOnColumn 0
    ControlsResizeColumns LComponents
End Sub

Private Sub LComponents_Click()
    '@AssignedModule uSkeleton
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyModule
    LProcedures.clear: TPROCS.TEXT = "": LCalls.clear: TCalls.TEXT = "": LDeclarations.clear: TDeclarations.TEXT = "":
    setUp
    For Each skellyProcName In ProcList(skellyModule)
        LProcedures.AddItem skellyProcName
    Next
    TComps.TEXT = aModule.Init(skellyModule).Code
    aListBox.Init(LProcedures).SortOnColumn 0
End Sub

Private Sub LProcedures_Click()
    '@AssignedModule uSkeleton
    '@INCLUDE PROCEDURE tmp
    '@INCLUDE PROCEDURE ModuleOfProcedure
    '@INCLUDE PROCEDURE ProcedureCode
    '@INCLUDE PROCEDURE getDeclarations
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE CLASS aListBox
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyBook
    '@INCLUDE DECLARATION skellyModule
    LCalls.clear: TCalls.TEXT = "": LDeclarations.clear: TDeclarations.TEXT = "":
    setUp
    TPROCS.TEXT = ProcedureCode(skellyBook, skellyModule, CStr(skellyProcName))
    Do While InStr(1, TPROCS.TEXT, "  ") > 0
        TPROCS.TEXT = Replace(TPROCS.TEXT, "  ", " ")
    Loop
    Dim element
    For Each element In AddProcedureCallsToCollectionSkeleton(skellyBook, ModuleOfProcedure(skellyBook, CStr(skellyProcName)), CStr(skellyProcName))
        LCalls.AddItem element
    Next
    aListBox.Init(LCalls).SortOnColumn 1
    'test
    Dim coll        As Collection: Set coll = aWorkbook.Init(skellyBook).getDeclarations(True, True, True, True, True, True)
    Dim keyCol      As Collection: Set keyCol = coll.item(5)
    Dim decCol      As Collection: Set decCol = coll.item(6)
    Dim i           As Long
    Dim tmp         As String
    For i = 1 To keyCol.Count
        'if the DECLARATION keyword exists inside the procedure
        If InStr(1, TPROCS.TEXT, keyCol.item(i)) > 0 Then
            'and if it is not a VARIABLE inside the procedure
            If InStr(1, TPROCS.TEXT, keyCol.item(i) & " As") = 0 Then
                'avoid duplicates

                If aListBox.Init(LDeclarations).Contains(keyCol.item(i)) = False Then
                    LDeclarations.AddItem keyCol.item(i)
                End If
            End If
        End If
    Next i
    aListBox.Init(LDeclarations).SortOnColumn 0
    aListBox.Init(LCalls).SortOnColumn 0
    ReleaseMe
End Sub

Private Sub LCalls_Click()
    '@AssignedModule uSkeleton
    '@INCLUDE PROCEDURE ModuleOfProcedure
    '@INCLUDE PROCEDURE ProcedureCode
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyBook
    '@INCLUDE DECLARATION skellyModule
    setUp
    skellyProcName = LCalls.list(LCalls.ListIndex)
    Set skellyModule = ModuleOfProcedure(skellyBook, CStr(skellyProcName))
    TCalls.TEXT = ProcedureCode(skellyBook, skellyModule, CStr(skellyProcName))
    ReleaseMe
End Sub

Private Sub LDeclarations_Click()
    '@AssignedModule uSkeleton
    '@INCLUDE PROCEDURE getDeclarations
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyBook
    setUp
    Dim coll        As Collection: Set coll = aWorkbook.Init(skellyBook).getDeclarations(True, True, True, True, True, True)
    Dim keyCol      As Collection: Set keyCol = coll.item(5)
    Dim decCol      As Collection: Set decCol = coll.item(6)
    Dim i           As Long
    For i = 1 To keyCol.Count
        If keyCol.item(i) = LDeclarations.list(LDeclarations.ListIndex) Then
            TDeclarations.TEXT = decCol.item(i)
        End If
    Next i
End Sub

Sub setUp()
    '@AssignedModule uSkeleton
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyBook
    '@INCLUDE DECLARATION skellyModule
    On Error Resume Next
    Set skellyBook = Workbooks(LProjects.list(LProjects.ListIndex))
    Set skellyModule = skellyBook.VBProject.VBComponents(LComponents.list(LComponents.ListIndex, 1))
    skellyProcName = LProcedures.list(LProcedures.ListIndex)
End Sub

Sub ReleaseMe()
    '@AssignedModule uSkeleton
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyBook
    '@INCLUDE DECLARATION skellyModule
    Set skellyModule = Nothing
    Set skellyBook = Nothing
End Sub

Function FindCalls(skellyBook As Workbook) As Collection
    '@INCLUDE ProceduresOfWorkbook
    '@INCLUDE ModuleOfProcedure
    '@INCLUDE CollectionToArray
    '@INCLUDE AddProcedureCallsToCollectionSkeleton
    '@AssignedModule uSkeleton
    '@INCLUDE PROCEDURE tmp
    '@INCLUDE PROCEDURE ModuleOfProcedure
    '@INCLUDE PROCEDURE ProceduresOfWorkbook
    '@INCLUDE CLASS aCollection
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyBook
    Dim Procedure   As Variant
    Dim Output      As New Collection
    Dim Procedures  As New Collection
    Dim calls       As New Collection
    Dim element     As Variant
    Dim tmp         As New Collection
    For Each Procedure In ProceduresOfWorkbook(skellyBook)
        Set tmp = AddProcedureCallsToCollectionSkeleton(skellyBook, ModuleOfProcedure(skellyBook, CStr(Procedure)), CStr(Procedure))
        If tmp.Count > 0 Then
            Procedures.Add Procedure
            calls.Add aCollection.Init(tmp).ToString(vbNewLine)
        End If
    Next
    Output.Add Procedures
    Output.Add calls
    Set FindCalls = Output
End Function

Function dataToSheet(Optional skellyBook As Workbook, Optional wsName As String, Optional rngAddress As String, Optional confirmClear As Boolean) As Range
    '@INCLUDE answer
    '@INCLUDE sheetExists
    '@AssignedModule uSkeleton
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyBook
    If skellyBook Is Nothing Then Set skellyBook = Workbooks.Add
    Dim ws          As Worksheet
    If sheetExists(wsName, skellyBook) Then
        If confirmClear = True Then
            Dim Answer As Integer
            Answer = MsgBox("Sheet " & wsName & " already exists. Cells will be cleared. Proceed?", vbYesNo)
            If Answer = vbNo Then Exit Function
        End If
        Set ws = skellyBook.Sheets(wsName)
        ws.Cells.clear
    Else
        If wsName = "" Then
            Set ws = skellyBook.Sheets(1)
        Else
            Set ws = skellyBook.Sheets.Add
            ws.Name = wsName
        End If
    End If
    If rngAddress <> "" Then
        Set dataToSheet = ws.Range(rngAddress)
    Else
        Set dataToSheet = ws.Range("A1")
    End If
End Function

Function ProcList(skellyModule As VBComponent) As Collection
    '@AssignedModule uSkeleton
    '@INCLUDE USERFORM uSkeleton
    '@INCLUDE DECLARATION skellyModule
    Dim codeMod     As CodeModule
    Set codeMod = skellyModule.CodeModule
    Dim coll        As Collection
    Set coll = New Collection
    Dim lineNum     As Long
    Dim NumLines    As Long
    Dim ProcName    As String
    Dim ProcKind    As VBIDE.vbext_ProcKind
    lineNum = codeMod.CountOfDeclarationLines + 1
    Do Until lineNum >= codeMod.CountOfLines
        ProcName = codeMod.ProcOfLine(lineNum, ProcKind)
        coll.Add ProcName
        lineNum = codeMod.procStartLine(ProcName, ProcKind) + codeMod.ProcCountLines(ProcName, ProcKind) + 1
    Loop
    Set ProcList = coll
End Function

Function ControlsResizeColumns(LBox As msforms.control, Optional ResizeListbox As Boolean)
    '@INCLUDE sheetExists
    '@AssignedModule uSkeleton
    '@INCLUDE USERFORM uSkeleton
    If LBox.ListCount = 0 Then Exit Function
    Application.ScreenUpdating = False
    Dim ws          As Worksheet
    If sheetExists("ListboxColumnWidth", ThisWorkbook) = False Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ListboxColumnwidth"
    Else
        Set ws = ThisWorkbook.Worksheets("ListboxColumnwidth")
        ws.Cells.clear
    End If
    ws.Cells.Font.Size = 12
    ws.Cells.Font.Name = "Calibri"
    '---Listbox to range-----
    Dim rng         As Range
    Set rng = ThisWorkbook.Sheets("ListboxColumnwidth").Range("A1")
    Set rng = rng.Resize(UBound(LBox.list) + 1, LBox.columnCount)
    rng = LBox.list
    '---Get ColumnWidths------
    rng.Columns.AutoFit
    Dim sWidth      As String
    Dim vR()        As Variant
    Dim n           As Integer
    Dim cell        As Range
    For Each cell In rng.Resize(1)
        n = n + 1
        ReDim Preserve vR(1 To n)
        vR(n) = cell.EntireColumn.Width
    Next cell
    sWidth = Join(vR, ";")
    'Debug.Print sWidth
    '---assign ColumnWidths----
    With LBox
        .ColumnWidths = sWidth
        '.RowSource = "A1:A3"
        .BorderStyle = fmBorderStyleSingle
    End With
    'Remove worksheet
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    '----Resize Listbox--------
    If ResizeListbox = False Then Exit Function
    Dim w           As Long
    Dim i           As Long
    For i = LBound(vR) To UBound(vR)
        w = w + vR(i)
    Next
    DoEvents
    LBox.Width = w + 10
End Function

Function sheetExists(sheetToFind As String, Optional InWorkbook As Workbook) As Boolean
    '@AssignedModule uSkeleton
    '@INCLUDE USERFORM uSkeleton
    If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook
    On Error Resume Next
    sheetExists = Not InWorkbook.Sheets(sheetToFind) Is Nothing
End Function

