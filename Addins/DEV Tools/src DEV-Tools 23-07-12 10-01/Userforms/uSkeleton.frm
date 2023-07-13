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


Dim skellyBook As Workbook
Dim skellyModule As VBComponent
Dim skellyProcName

Private Sub exportDeclarationsAndCalls_Click()
    setUp
    Dim var As Variant
    Dim collections As New Collection
    Set collections = FindCalls(skellyBook)
    If collections(1).count > 0 Then
        var = aCollection.CollectionsToArray2D(collections)
        ArrayToRange2D var, dataToSheet(skellyBook, "exportCalls", "A2")
        With skellyBook.Sheets("exportCalls").Range("A1:B1")
            .Value = Array("Procedure", "Calls")
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
    If collections(1).count > 0 Then
        var = aCollection.CollectionsToArray2D(collections)
        ArrayToRange2D var, dataToSheet(skellyBook, "exportDeclarations", "A2")
        With skellyBook.Sheets("exportDeclarations").Range("A1:F1")
            .Value = Array("Component Type", "Component Name", "Declaration Scope", "Declaration Type", "Declaration Keyword", "Declaration Code")
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

Function AddProcedureCallsToCollectionSkeleton(skellyBook As Workbook, skellyModule As VBComponent, procName As String) As Collection
    '@INCLUDE ProceduresOfWorkbook
    '@INCLUDE ProcedureCode
    Dim coll As Collection
    Set coll = New Collection
    Dim WorkbookProcedure As Variant
    Dim AllProcs As Collection
    Set AllProcs = ProceduresOfWorkbook(skellyBook)
    Dim procText As String
        procText = ProcedureCode(skellyBook, skellyModule, procName)
    For Each WorkbookProcedure In AllProcs
        If CStr(WorkbookProcedure) <> procName Then
            If InStr(1, procText, CStr(WorkbookProcedure)) Then
                coll.Add CStr(WorkbookProcedure)
            End If
        End If
    Next WorkbookProcedure
    Set AddProcedureCallsToCollectionSkeleton = coll
End Function



Private Sub UserForm_Initialize()
    LoadProjects
End Sub
Private Sub GetInfo_Click()
    uAuthor.Show
End Sub
Sub LoadProjects()
    '@INCLUDE WorkbookProjectProtected
    For Each skellyBook In Workbooks
        If Not WorkbookProjectProtected(skellyBook) Then LProjects.AddItem skellyBook.Name
    Next
    On Error Resume Next
    Dim ad
    For Each ad In AddIns
        If Not WorkbookProjectProtected(Workbooks(ad.Name)) Then
            If Err = 0 Then LProjects.AddItem ad.Name
            Err.Clear
        End If
    Next
End Sub

Private Sub LProjects_Click()
    loadComponents
End Sub

Sub loadComponents()
    '@INCLUDE ModuleTypeToString
    '@INCLUDE SortListboxOnColumn
    '@INCLUDE setUp
    '@INCLUDE ReleaseMe
    '@INCLUDE ControlsResizeColumns
    LComponents.Clear: LProcedures.Clear: TPROCS.TEXT = "": LCalls.Clear: TCalls.TEXT = "": LDeclarations.Clear: TDeclarations.TEXT = "":
    setUp
    For Each skellyModule In skellyBook.VBProject.VBComponents
        LComponents.AddItem
        LComponents.List(LComponents.ListCount - 1, 0) = aModule.Init(skellyModule).TypeToString
        LComponents.List(LComponents.ListCount - 1, 1) = skellyModule.Name
    Next
    ReleaseMe
    aListBox.Init(LComponents).SortOnColumn 0
    ControlsResizeColumns LComponents
End Sub

Private Sub LComponents_Click()
    LProcedures.Clear: TPROCS.TEXT = "": LCalls.Clear: TCalls.TEXT = "": LDeclarations.Clear: TDeclarations.TEXT = "":
    setUp
    For Each skellyProcName In ProcList(skellyModule)
        LProcedures.AddItem skellyProcName
    Next
    TComps.TEXT = aModule.Init(skellyModule).Code
    aListBox.Init(LProcedures).SortOnColumn 0
End Sub

Private Sub LProcedures_Click()
    LCalls.Clear: TCalls.TEXT = "": LDeclarations.Clear: TDeclarations.TEXT = "":
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
    Dim coll As Collection:     Set coll = aWorkbook.Init(skellyBook).getDeclarations(True, True, True, True, True, True)
    Dim keyCol As Collection:   Set keyCol = coll.item(5)
    Dim decCol As Collection:   Set decCol = coll.item(6)
    Dim i As Long
    Dim tmp As String
    For i = 1 To keyCol.count
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
    setUp
    skellyProcName = LCalls.List(LCalls.ListIndex)
    Set skellyModule = ModuleOfProcedure(skellyBook, CStr(skellyProcName))
    TCalls.TEXT = ProcedureCode(skellyBook, skellyModule, CStr(skellyProcName))
    ReleaseMe
End Sub

Private Sub LDeclarations_Click()
    setUp
    Dim coll As Collection:     Set coll = aWorkbook.Init(skellyBook).getDeclarations(True, True, True, True, True, True)
    Dim keyCol As Collection:   Set keyCol = coll.item(5)
    Dim decCol As Collection:   Set decCol = coll.item(6)
    Dim i As Long
    For i = 1 To keyCol.count
        If keyCol.item(i) = LDeclarations.List(LDeclarations.ListIndex) Then
            TDeclarations.TEXT = decCol.item(i)
        End If
    Next i
End Sub

Sub setUp()
    On Error Resume Next
    Set skellyBook = Workbooks(LProjects.List(LProjects.ListIndex))
    Set skellyModule = skellyBook.VBProject.VBComponents(LComponents.List(LComponents.ListIndex, 1))
    skellyProcName = LProcedures.List(LProcedures.ListIndex)
End Sub

Sub ReleaseMe()
    Set skellyModule = Nothing
    Set skellyBook = Nothing
End Sub

Function FindCalls(skellyBook As Workbook) As Collection
    '@INCLUDE ProceduresOfWorkbook
    '@INCLUDE ModuleOfProcedure
    '@INCLUDE CollectionToArray
    '@INCLUDE AddProcedureCallsToCollectionSkeleton
    Dim Procedure As Variant
    Dim output As New Collection
    Dim Procedures As New Collection
    Dim calls As New Collection
    Dim element As Variant
    Dim tmp As New Collection
    For Each Procedure In ProceduresOfWorkbook(skellyBook)
        Set tmp = AddProcedureCallsToCollectionSkeleton(skellyBook, ModuleOfProcedure(skellyBook, CStr(Procedure)), CStr(Procedure))
        If tmp.count > 0 Then
            Procedures.Add Procedure
            calls.Add aCollection.Init(tmp).ToString(vbNewLine)
        End If
    Next
    output.Add Procedures
    output.Add calls
    Set FindCalls = output
End Function

Function dataToSheet(Optional skellyBook As Workbook, Optional wsName As String, Optional rngAddress As String, Optional confirmClear As Boolean) As Range
    '@INCLUDE answer
    '@INCLUDE sheetExists
    If skellyBook Is Nothing Then Set skellyBook = Workbooks.Add
    Dim ws As Worksheet
    If sheetExists(wsName, skellyBook) Then
        If confirmClear = True Then
            Dim answer As Integer
            answer = MsgBox("Sheet " & wsName & " already exists. Cells will be cleared. Proceed?", vbYesNo)
            If answer = vbNo Then Exit Function
        End If
        Set ws = skellyBook.Sheets(wsName)
        ws.Cells.Clear
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
    Dim codeMod As CodeModule
    Set codeMod = skellyModule.CodeModule
    Dim coll As Collection
    Set coll = New Collection
    Dim lineNum As Long
    Dim NumLines As Long
    Dim procName As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    lineNum = codeMod.CountOfDeclarationLines + 1
    Do Until lineNum >= codeMod.CountOfLines
        procName = codeMod.ProcOfLine(lineNum, ProcKind)
        coll.Add procName
        lineNum = codeMod.procStartLine(procName, ProcKind) + codeMod.ProcCountLines(procName, ProcKind) + 1
    Loop
    Set ProcList = coll
End Function

Function ControlsResizeColumns(LBox As MSForms.control, Optional ResizeListbox As Boolean)
    '@INCLUDE sheetExists
    If LBox.ListCount = 0 Then Exit Function
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    If sheetExists("ListboxColumnWidth", ThisWorkbook) = False Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ListboxColumnwidth"
    Else
        Set ws = ThisWorkbook.Worksheets("ListboxColumnwidth")
        ws.Cells.Clear
    End If
    ws.Cells.Font.Size = 12
    ws.Cells.Font.Name = "Calibri"
    '---Listbox to range-----
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("ListboxColumnwidth").Range("A1")
    Set rng = rng.Resize(UBound(LBox.List) + 1, LBox.ColumnCount)
    rng = LBox.List
    '---Get ColumnWidths------
    rng.Columns.AutoFit
    Dim sWidth As String
    Dim vR() As Variant
    Dim n As Integer
    Dim cell As Range
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
    Dim w As Long
    Dim i As Long
    For i = LBound(vR) To UBound(vR)
        w = w + vR(i)
    Next
    DoEvents
    LBox.Width = w + 10
End Function

Function sheetExists(sheetToFind As String, Optional InWorkbook As Workbook) As Boolean
    If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook
    On Error Resume Next
    sheetExists = Not InWorkbook.Sheets(sheetToFind) Is Nothing
End Function

