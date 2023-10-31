VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uChangeLog 
   Caption         =   "Changelog Manager"
   ClientHeight    =   9432.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11316
   OleObjectBlob   =   "uChangeLog.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uChangeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Table As ListObject
Private TargetWorksheet As Worksheet
Private targetVersion As String
Private TargetWorkbook As Workbook


Sub DeleteLastMod()
    Dim cell As Range
    On Error Resume Next
    Set cell = Table.DataBodyRange.Cells(1, 2)
    targetVersion = cell.Value
    On Error GoTo 0
    If targetVersion = "" Then Exit Sub
    Application.ScreenUpdating = False
    Dim rng As Range
    Set rng = VersionRange(targetVersion)
    If rng Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To rng.rows.Count
        Table.ListRows(1).Delete
    Next
    
    Dim targetFolder As String
    targetFolder = TargetWorkbook.path & "\ChangeLog\" & targetVersion
    FolderDelete targetFolder
    
    If targetVersion = "1.0.0" Then PushVersionInitial
    Application.ScreenUpdating = True
End Sub

Public Function PushVersionInitial()
'@LastModified 2310102011
    PushVersionInitial = ListModifications(False, False, False)
End Function
Public Function PushVersionMajor()
'@LastModified 2310102011
    PushVersionMajor = ListModifications(True, False, False)
End Function
Public Function PushVersionMinor()
'@LastModified 2310102011
    PushVersionMinor = ListModifications(False, True, False)
End Function
Public Function PushVersionPatch()
'@LastModified 2310102012
    PushVersionPatch = ListModifications(False, False, True)
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 28-08-2023 07:36    Alex                (uChangeLog.frm > ListModifications)
'* Updated    : 10-10-2023 20:12    Alex                (uChangeLog.frm > ListModifications : Now we can create a mod table for any project)
'* Updated    : 30-10-2023 10:58    Alex                (uChangeLog.frm > ListModifications : + changelog.txt | + additional export options)
'* Updated    : 30-10-2023 22:18    Alex                (uChangeLog.frm > ListModifications : in Normal_Exit added important TargetWorkbook.Save)

Function ListModifications( _
                     major As Boolean, _
                     minor As Boolean, _
                     patch As Boolean) As Boolean
'@LastModified 2310302218
'@INCLUDE PROCEDURE ListModificationsAfterLastPush
    Dim paramIndex As Long: paramIndex = Abs(major) + Abs(minor) + Abs(patch)
    Select Case paramIndex
    Case 0, 1
        
    Case Else
        Toast "Choose only 0 or 1 semantic versioning category as True"
        Exit Function
    End Select
  
    Dim TableRow As ListRow
    
    Dim newVersion As String
    If major + minor + patch = 0 Or Table.ListRows.Count = 0 Then
        newVersion = "1.0.0"
        If Table.ListRows.Count = 0 Then 'Table.DataBodyRange.Cells(1, 1).value = vbNullString Then
            Dim ans As String
retry:
            ans = InputboxString("Initializing", "Select initialization date for this project." & vbLf & _
                "Code modifications marked up to this date" & vbLf & _
                "by aProcedure.active.InjectModification won't be noted." & vbLf & vbLf & _
                "Type in as YY-MM-DD", _
                Format(Date - 1, "YY-MM-DD"))
            If Not ans Like "??-??-??" Then
                If MsgBox("Abort?", vbYesNo) = vbYes Then
                    ToggleControls False
                    Exit Function
                Else
                    GoTo retry
                End If
            End If
            Set TableRow = Table.ListRows.Add '(1)
            TableRow.Range(1, 1) = Format(Date - 1, "YY-MM-DD")
            TableRow.Range(1, 2) = newVersion
            TableRow.Range(1, 3) = "Initial Release"
            TableRow.Range.Cells.Font.Bold = False
            GoTo Normal_Exit
        Else
            Toast "Already initialized"
            GoTo CLEANUP
        End If
    End If
    
    Dim previousVersion
    previousVersion = Table.DataBodyRange.Cells(1, 2).Value
      
    'repush in the same day
    If Table.ListRows(1).Range(1, 1) = Format(Date, "YY-MM-DD") Then
        DeleteLastMod
        ListModifications major, minor, patch
        Exit Function
    End If
    
    newVersion = Split(previousVersion, ".")(0) + Abs(major) & "." & _
                 IIf(major, 0, Split(previousVersion, ".")(1) + Abs(minor)) & "." & _
                 IIf(minor, 0, IIf(major, 0, Split(previousVersion, ".")(2) + Abs(patch)))

    Application.ScreenUpdating = False
    
    Dim var
        var = ListModificationsAfterLastPush
        
    Dim customMessage As String
        customMessage = InputboxString("Custom message", _
                        UBound(var) + 1 & " code modifications found since " & Table.ListRows(1).Range(1, 1) & _
                        vbLf & vbLf & _
                        "Optional description for version " & newVersion)
    
    If Not ArrayAllocated(var) And customMessage = "" Then
        Toast "No mods found after " & Table.DataBodyRange.Cells(1, 1).Value & vbLf & _
                "and no custom description." & vbLf & _
                "Aborting operation"
        GoTo CLEANUP
    Else
        If customMessage <> "" Then
            Set TableRow = Table.ListRows.Add(1, True)
            TableRow.Range(1, 1) = Format(Date, "YY-MM-DD")
            TableRow.Range(1, 2) = newVersion
            TableRow.Range(1, 3) = customMessage
            TableRow.Range.Cells.Font.Bold = False
        End If
        Dim dif As Long:  dif = IIf(LBound(var) = 0, 1, 0)
        Dim i As Long
        For i = 1 To UBound(var) + dif
            Set TableRow = Table.ListRows.Add(i + IIf(customMessage = "", 0, 1), True)
            If customMessage = "" And i = 1 Then
                TableRow.Range(1, 1) = Format(Date, "YY-MM-DD")
                TableRow.Range(1, 2) = newVersion
            End If
            TableRow.Range(1, 3) = var(i - dif)
            TableRow.Range.Cells.Font.Bold = False
        Next
    End If
    
Normal_Exit:
    TargetWorksheet.Columns.AutoFit
    TxtOverwrite TargetWorkbook.path & "\" & aWorkbook.Init(TargetWorkbook).NameClean & "_ChangeLog.txt", PrettyPrint.ArrayToTable(Table.Range.Value, True)
    Dim targetFolder As String
    targetFolder = TargetWorkbook.path & "\ChangeLog\" & newVersion & "\"
    FoldersCreate targetFolder
    TargetWorkbook.Save
    TargetWorkbook.SaveCopyAs targetFolder & TargetWorkbook.Name
    With aWorkbook.Init(TargetWorkbook)
        If chWorkbookBackup Then TargetWorkbook.SaveCopyAs targetFolder & .Name
        If chExportReferences Then .ExportReferences targetFolder
        If chExportUnified Then .ExportCodeUnified targetFolder
        If chExportComponents Then .ExportModules targetFolder
        If chExportProcedures Then .ExportProcedures targetFolder & "PROCEDURES\", createModuleSubfolders:=True
        If chExportXML Then .ExportXML targetFolder
    End With
    ListModifications = True
CLEANUP:
    Application.ScreenUpdating = True
    Toast "Complete"
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 13:48    Alex                (Dependencies.bas > ListModificationsAfterLastPush)
'* Updated    : 10-10-2023 20:14    Alex                (Dependencies.bas > ListModificationsAfterLastPush)

Function ListModificationsAfterLastPush() As Variant
'@LastModified 2310260039
'@INCLUDE PROCEDURE ListModificationsBetween
    Dim out
    If Table.ListRows.Count = 0 Then
        out = ListModificationsBetween
    Else
        out = ListModificationsBetween(afterYYMMDD:=dateAfterLastPush)
    End If
    Dim i As Long
    For i = LBound(out) To UBound(out)
'        out(i) = Replace(Split(out(i), "(")(1), ")", "")
        out(i) = Replace(Split(Split(out(i), ": ")(1), "  ")(0) & " " & Split(out(i), "(")(1), ")", "")
    Next
    ListModificationsAfterLastPush = out
End Function

Public Function dateAfterLastPush()
    Dim dateCell As Range: Set dateCell = Table.DataBodyRange.Cells(1, 1)
    dateAfterLastPush = CStr( _
                        Mid(dateCell.TEXT, 1, 2) & _
                        Mid(dateCell.TEXT, 4, 2) & _
                        Mid(dateCell.TEXT, 7, 2) + 1)
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 22-08-2023 09:51    Alex                (Dependencies.bas > ListModificationsBetween)

Function ListModificationsBetween(Optional afterYYMMDD = 200000, Optional beforeYYMMDD = 300000) As Variant
'@LastModified 2308220951
'@INCLUDE CLASS aWorkbook
'@INCLUDE PROCEDURE ArrayQuickSort
'@INCLUDE PROCEDURE ArrayTrim
    Dim arr
    arr = ArrayTrim( _
            Split( _
                aWorkbook.Init(TargetWorkbook).Code, _
                vbNewLine))
    arr = ArrayTrim( _
            Filter( _
                Filter( _
                    arr, _
                    "'* Updated    : ", _
                    True, _
                    vbTextCompare), _
                """", _
                False) _
            )
    ArrayTrim arr
    ArrayQuickSort arr
    Dim out
    Dim line, element
    For Each line In arr
        element = Split(Split(line, ": ")(1), " ")(0)
        element = Mid(element, 9, 2) & Mid(element, 4, 2) & Mid(element, 1, 2)
        If (afterYYMMDD <= CLng(element)) And (CLng(element) <= beforeYYMMDD) Then
            out = out & IIf(out <> "", vbNewLine, "") & line
        End If
    Next
'    out = Split(out, vbNewLine)
    ListModificationsBetween = Filter(Split(out, vbNewLine), "(", True)
End Function



Private Sub GetInfo_Click()
    With aUserform.Init(Me)
        .Transition .Effect(GetInfo, "Top", GetInfo.Top - 10, 150)
        .Transition .Effect(GetInfo, "Top", GetInfo.Top + 10, 150)
    End With
    uAuthor.Show
End Sub

Private Sub Label1_Click()
    If MsgBox("Irreversible proccess. Proceed?", vbYesNo) = vbNo Then Exit Sub
    DeleteLastMod
    UpdateVersions
End Sub

Private Sub Label2_Click()
    PushVersionMajor
    UpdateVersions
End Sub

Private Sub Label3_Click()
    PushVersionMinor
    UpdateVersions
End Sub

Private Sub Label4_Click()
    PushVersionPatch
    UpdateVersions
End Sub

Sub UpdateVersions()
    Application.EnableEvents = False
    LB_Versions.clear
    TextBox1.Value = ""
    ListVersions
    Application.EnableEvents = True
    If LB_Versions.ListCount > 0 Then LB_Versions.ListIndex = 0
End Sub

Private Sub Label5_Click()
    If Not TargetWorkbook Is Nothing Then FollowLink TargetWorkbook.path
End Sub

Private Sub Label6_Click()
    FollowLink "https://semver.org/"
End Sub

Private Sub UserForm_Initialize()
    aListBox.Init(LB_Books).LoadVBProjects
    LB_Versions.columnCount = 2
    LB_Versions.ColumnWidths = "40;"
    TextBox1.Font.Name = "Consolas"
    TextBox1.ScrollBars = fmScrollBarsBoth
    
    TextBox1.Font.Size = 10
    LB_Versions.Font.Size = 10
    LB_Books.Font.Size = 10
    
    ToggleControls False

    chWorkbookBackup.Value = True
    chWorkbookBackup.Enabled = False
    
    chExportUnified.Value = True
    chExportUnified.Enabled = False
    
    chExportReferences.Value = True
    chExportReferences.Enabled = False
    
    chExportComponents.Value = True
    
    TextBox1.Tag = "wh"
    With aUserform.Init(Me)
        .Resizable
        .ShowMaximizeButton
        .ShowMinimizeButton
    End With
End Sub


Private Sub LB_Books_Change()
    ClearPreviouslyLoaded
    Set TargetWorkbook = Workbooks(LB_Books.list(LB_Books.ListIndex))
    
    If WorkbookProjectProtected(TargetWorkbook) Then
        
    End If
    
    If Not aWorkbook.Init(TargetWorkbook).HasProject Then
        
    End If
    
    If Not CanCreateAndEditWorksheet Then
        Toast "Can't create and/or edit worksheets"
        ToggleControls False
        Exit Sub
    End If
    
    If Not CheckIfInitialized(TargetWorkbook) Then Exit Sub
    ListVersions
    If LB_Versions.ListCount > 0 Then LB_Versions.ListIndex = 0
    ToggleControls True
End Sub

Sub ToggleControls(targetStatus As Boolean)
    Label1.Enabled = targetStatus
    Label2.Enabled = targetStatus
    Label3.Enabled = targetStatus
    Label4.Enabled = targetStatus
End Sub

Sub ListVersions()
    Dim i As Long
    Dim var
    var = myVersions
    If Not ArrayAllocated(var) Then Exit Sub
    For i = LBound(var) To UBound(var)
        LB_Versions.AddItem
        LB_Versions.list(LB_Versions.ListCount - 1, 0) = var(i, 1)
        LB_Versions.list(LB_Versions.ListCount - 1, 1) = var(i, 2)
    Next
End Sub
Sub ClearPreviouslyLoaded()
    LB_Versions.clear
    TextBox1.Value = ""
    Set TargetWorkbook = Nothing
    Set TargetWorksheet = Nothing
    Set Table = Nothing
End Sub

Private Sub LB_Versions_Change()
    If LB_Versions.ListIndex = -1 Then Exit Sub
    targetVersion = LB_Versions.list(LB_Versions.ListIndex)
    Dim out As String
    Dim rng As Range
    Set rng = VersionRange(targetVersion)
    Dim i As Long
    For i = 1 To rng.rows.Count
        out = out & IIf(out <> "", vbNewLine, "") & rng.Cells(i, 3).Value
    Next
    Dim var
    var = Split(out, vbNewLine)
'    var = Filter(var, ">", True)
'    var = Filter(var, ":", True)
    out = Join(var, vbNewLine)
    out = StringFormatAlignRowsElements(out, ">", True)
    out = StringFormatAlignRowsElements(out, ":", False)
    TextBox1.Value = out
End Sub

Function VersionRange(targetVersion As String) As Range
    If Table.ListRows.Count = 0 Then Exit Function
    Dim cell As Range
    Dim rng As Range
    For Each cell In Table.DataBodyRange.Columns(2).Cells
        If cell.Value = targetVersion Then
            Do While (cell.Value = targetVersion Or cell.Value = "") And (Not cell.ListObject Is Nothing)
                If rng Is Nothing Then
                    Set rng = cell.offset(0, -1).Resize(1, 3)
                Else
                    Set rng = rng.Resize(rng.rows.Count + 1)
                End If
                Set cell = cell.offset(1, 0)
            Loop
            Exit For
        End If
    Next
    Set VersionRange = rng
End Function

Function myVersions() As Variant
    Dim cell As Range
    Dim rng As Range
    On Error Resume Next
    Set rng = Table.ListColumns(2).DataBodyRange
    Set rng = Intersect(rng, rng.SpecialCells(xlCellTypeConstants))
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    Dim out
    Dim lim As Long: lim = rng.Count
    ReDim out(1 To lim, 1 To 2)
    Dim i As Long: i = 0
    For Each cell In rng
        i = i + 1
        If cell <> "" Then out(i, 1) = cell.Value
        If cell <> "" Then out(i, 2) = cell.offset(0, -1).Value
    Next
    myVersions = out
End Function

Function CheckIfInitialized(TargetWorkbook As Workbook) As Boolean
    On Error Resume Next
    Set TargetWorksheet = TargetWorkbook.Sheets("ChangeLog")
    If TargetWorksheet Is Nothing Then
        TargetWorkbook.Sheets.Add(Before:=Sheets(1)).Name = "ChangeLog"
        Set TargetWorksheet = TargetWorkbook.Sheets("ChangeLog")
    End If
    Set Table = TargetWorksheet.ListObjects("TB_ChangeLog")
    On Error GoTo 0
    If Table Is Nothing Then
        Dim cell As Range
        Set cell = getLastCell(TargetWorksheet)
        Set cell = TargetWorksheet.Cells(cell.Row + 3, 1)
        Dim rng As Range
        Set rng = TargetWorksheet.Range(cell, cell.offset(0, 2))
        rng.Value = Array("Date", "Version", "Changes")
        Set Table = TargetWorksheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=rng, XlListObjecthasheaders:=xlYes)
        Table.Name = "TB_ChangeLog"
        CheckIfInitialized = PushVersionInitial
    ElseIf Table.ListRows.Count = 0 Then
        CheckIfInitialized = PushVersionInitial
    Else
        CheckIfInitialized = True
    End If
End Function
