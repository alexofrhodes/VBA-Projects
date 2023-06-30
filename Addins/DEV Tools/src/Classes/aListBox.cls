VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aListBox
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
'* Modified   : 27-06-2023 12:38    Alex                figured how to get the control's hWnd
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'TODO?
'Private oList As Variant
'Public Function List() As Variant
'    List = LB.List
'End Function
'Public Sub RememberList()
'    oList = List
'End Sub
'Public Sub LoadList()
'    LB.List = List
'End Sub

Option Explicit

Private LB As MSForms.ListBox
Private oForm As Object

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Type POINTAPI
    X               As Long
    Y               As Long
End Type

#If VBA7 Then

Private Type Msg
    hWnd            As LongPtr
    message         As Long
    wParam          As LongPtr
    lParam          As LongPtr
    time            As Long
    pt              As POINTAPI
End Type

#If Win64 Then
Private Declare PtrSafe Function WindowFromPoint Lib "USER32" (ByVal Point As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function WindowFromPoint Lib "USER32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
#End If

Private Declare PtrSafe Function GetMessage Lib "USER32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare PtrSafe Function DispatchMessage Lib "USER32" Alias "DispatchMessageA" (lpMsg As Msg) As LongPtr
Private Declare PtrSafe Function TranslateMessage Lib "USER32" (lpMsg As Msg) As Long
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As LongPtr) As Long
Private Declare PtrSafe Function IsWindow Lib "USER32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As LongPtr, ByVal fAccept As Long)
Private Declare PtrSafe Sub DragFinish Lib "shell32.dll" (ByVal HDROP As LongPtr)
Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As LongPtr, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

#Else

Private Type Msg
    hWnd            As Long
    message         As Long
    wParam          As Long
    lParam          As Long
    time            As Long
    pt              As POINTAPI
End Type

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As Long, ByVal fAccept As Long)
Private Declare Sub DragFinish Lib "shell32.dll" (ByVal HDROP As Long)
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

#End If

Function LoadCSV(TargetFile As String, Clear As Boolean)
    If Clear Then LB.Clear
    Dim ArrayOfElements
    Dim s: s = TxtRead(TargetFile)
    Dim arr
    arr = Split(s, vbLf)
    Dim iCols, iRows
    iRows = UBound(arr) + 1
    iCols = UBound(Split(arr(0), ",")) + 1
    ReDim ArrayOfElements(1 To iRows, 1 To iCols)
    Dim X, Y
    For X = 1 To iRows
        For Y = 1 To iCols
            ArrayOfElements(X, Y) = Split(arr(X - 1), ",")(Y - 1)
        Next
    Next
    LB.ColumnCount = iCols
    LB.list = ArrayOfElements
End Function

Sub ToRange(cell As Range)
    cell.Resize(LB.ListCount, LB.ColumnCount) = LB.list
End Sub

Public Function Init(ByRef TargetListBox As MSForms.ListBox) As aListBox
    Set LB = TargetListBox
    Set oForm = LB.Parent
    Set Init = Me
 
End Function

Sub ClearSelection(LBox As MSForms.ListBox)
    On Error Resume Next
    Dim i As Long
    For i = 0 To LBox.ListCount
        LBox.Selected(i) = False
    Next i
End Sub

Public Sub AcceptFiles(Optional sExpansion As String = "*.xlsm;*.xlsb;*.xlsx", Optional iDeepSubPath As Integer = 999)

'eg on userform_activate:
'aListBox.Init(ListBox1).AcceptFiles sExpansion:="*.xlsm;*.xlsb;*.xlsx",iDeepSubPath:=999

    Dim objList As MSForms.ListBox: Set objList = LB
    
    'you can use .[_GethWnd] with controls which support it
    Dim ctr As control: Set ctr = LB
    Dim lGethWnd As LongPtr: lGethWnd = ctr.[_GethWnd]

#If VBA7 Then
    Dim hWnd As LongPtr, HDROP As LongPtr
#Else
    Dim hWnd As Long, HDROP As Long
#End If

    Const WM_DROPFILES = &H233
    Dim tMsg As Msg, sFileName As String * 1000
    Dim lFilesCount As Long, i As Long
    hWnd = aUserform.Init(oForm).hWnd
    Call DragAcceptFiles(lGethWnd, True)

    Dim sPathFile   As String

    Do While GetMessage(tMsg, 0, 0, 0) And IsWindow(hWnd)
        If tMsg.message = WM_DROPFILES Then
            HDROP = tMsg.wParam
            lFilesCount = DragQueryFile(HDROP, &HFFFFFFFF, 0, 0)
            If lFilesCount Then
                For i = 0 To lFilesCount - 1
                    sPathFile = VBA.Trim$(VBA.Left(sFileName, DragQueryFile(HDROP, i, sFileName, VBA.Len(sFileName))))
                    Call getFilesFromPath(objList, sPathFile, sExpansion, iDeepSubPath)
                Next i
            End If
            Call DragFinish(HDROP)
        End If
        Call TranslateMessage(tMsg)
        Call DispatchMessage(tMsg)
    Loop
End Sub
Private Sub getFilesFromPath(ByRef objList As MSForms.ListBox, ByVal sPathFile As String, ByVal sExpansion As String, ByVal iDeepSubPath As Integer)
    Dim sExp        As String
    sExp = sGetExtensionName(sPathFile)
    If sExp = vbNullString Then
        Dim oCol    As Collection
        Set oCol = GetFiles(sPathFile, iDeepSubPath)
        Dim i       As Integer
        Dim iCount  As Integer
        iCount = oCol.count
        Dim sFile   As String
        For i = 1 To iCount
            sFile = oCol.item(i).Path
            sExp = sGetExtensionName(sFile)
            Call filterFiles(objList, sFile, sExpansion, sExp)
        Next i
    Else
        Call filterFiles(objList, sPathFile, sExpansion, sExp)
    End If
End Sub
Private Sub filterFiles(ByRef objList As MSForms.ListBox, ByVal sPathFile As String, ByVal sExpansion As String, ByVal sExp As String)
    sExp = "*." & sExp
    If sExpansion = vbNullString Or sExpansion = "*.*" Then
    ElseIf Not likeExp(sExpansion, sExp) Then
        Exit Sub
    End If
    objList.AddItem sPathFile
End Sub
Private Function likeExp(ByVal sExpansion As String, ByVal sExp As String) As Boolean
    Dim arr         As Variant
    arr = VBA.Split(sExpansion, ";")
    Dim i           As Long
    Dim iCount      As Long
    iCount = UBound(arr, 1)
    For i = 0 To iCount
        If sExp Like arr(i) Then
            likeExp = True
            Exit For
        End If
    Next i
End Function
Public Function sGetExtensionName(ByVal sPathFile As String) As String
    Dim fso         As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    sGetExtensionName = fso.GetExtensionName(sPathFile)
    Set fso = Nothing
End Function
Private Function GetFiles(ByVal Path As String, ByVal iDeepSubPath As Integer) As Collection
    Dim MainFolder  As Object
    Dim iFolder     As Object
    Dim iFile       As Object
    Dim fso         As Object
    Dim MainColl    As New Collection
    Dim iColl       As Collection
    Dim i           As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set MainFolder = fso.getFolder(Path)
    If MainFolder Is Nothing Then Exit Function

    For Each iFile In MainFolder.Files
        If VBA.InStr(1, iFile.Name, "~") = 0 Then
            MainColl.Add iFile, iFile.Path
        End If
    Next

    If iDeepSubPath > 0 Then
        For Each iFolder In MainFolder.SubFolders
            Set iColl = GetFiles(iFolder.Path, iDeepSubPath - 1)
            For i = 1 To iColl.count
                MainColl.Add iColl(i)
            Next
        Next
    End If
    Set GetFiles = MainColl
End Function

Public Sub HeightToEntries(ByVal NumberOfEntries As Long)
    Const SM_CYEDGE = 46&
    With LB
        NumberOfEntries = IIf(NumberOfEntries > .ListCount, .ListCount, NumberOfEntries)
        .Height = ((9.75 * NumberOfEntries) + IIf(.SpecialEffect = fmSpecialEffectFlat, 0, GetSystemMetrics(SM_CYEDGE)))
    End With
End Sub

Public Sub LoadVBProjects()
    Dim coll As New Collection
    Dim WB As Workbook
    For Each WB In Workbooks
        If Len(WB.Path) > 0 Then
            If WorkbookProjectProtected(WB) = False Then
                On Error Resume Next
                coll.Add WB.Name, WB.Name
                On Error GoTo 0
            End If
        End If
    Next
    Rem list addins
    Dim vbProj As VBProject
    Dim wbPath As String
    For Each vbProj In Application.VBE.VBProjects
        On Error GoTo ErrorHandler
        wbPath = vbProj.Filename
        If Right(wbPath, 4) = "xlam" Or Right(wbPath, 3) = "xla" Then
            Dim wbName As String
            wbName = Mid(wbPath, InStrRev(wbPath, "\") + 1)
            If WorkbookProjectProtected(Workbooks(wbName)) = False Then
                On Error Resume Next
                coll.Add wbName, wbName
                On Error GoTo 0
            End If
        End If
SKIP:
    Next vbProj

    Dim el As Variant
    For Each el In coll
    LB.AddItem el
    Next

    Exit Sub
ErrorHandler:
    If Err.Number = 76 Then GoTo SKIP
End Sub

Public Sub SelectItems(This As Variant, Optional ByIndex As Boolean)
    Dim i As Long
    Select Case TypeName(This)
    Case Is = "String", "Long", "Integer"
        For i = 0 To LB.ListCount - 1
            If LB.list(i) = CStr(This) Then
                LB.Selected(i) = True
                DoEvents
                If LB.multiSelect = fmMultiSelectSingle Then Exit Sub
            End If
        Next
    Case Else
        Dim el As Variant
        If ByIndex Then
            For Each el In This
                LB.Selected(el) = True
            Next
        Else
            For Each el In This
                For i = 0 To LB.ListCount - 1
                    If LB.list(i) = el Then
                        LB.Selected(i) = True
                        DoEvents
                    End If
                Next
            Next
        End If
    End Select
End Sub

Public Sub AddHeader(Header As MSForms.ListBox, arrHeaders)
    Header.Width = LB.Width
    Dim i As Long
    Header.ColumnCount = LB.ColumnCount
    Header.ColumnWidths = LB.ColumnWidths
    Header.Clear
    Header.AddItem
    If ArrayDimensions(arrHeaders) = 1 Then
        For i = 0 To UBound(arrHeaders)
            Header.list(0, i) = arrHeaders(i)
        Next i
    Else
        For i = 1 To UBound(arrHeaders, 2)
            Header.list(0, i - 1) = arrHeaders(1, i)
        Next i
    End If
'    LB.ZOrder (1)
'    Header.ZOrder (0)
    Header.SpecialEffect = fmSpecialEffectFlat
    Header.BackColor = RGB(200, 200, 200)
    Header.Height = 15
    Header.Width = LB.Width
    Header.Left = LB.Left
    Header.Top = LB.Top - Header.Height - 1
    Header.Font.Bold = True
    Header.Font.Name = "Comic Sans MS"
    Header.Font.Size = LB.Font.Size
End Sub

Public Function Contains(This As String, _
                Optional ColumnIndexZeroBased As Long = -1, _
                Optional CaseSensitive As Boolean = False) As Boolean
    Dim i      As Long
    Dim n      As Long
    Dim sTemp  As String
    If ColumnIndexZeroBased > LB.ColumnCount - 1 Or ColumnIndexZeroBased < 0 Then
        ColumnIndexZeroBased = -1
    End If
    n = LB.ListCount
    If ColumnIndexZeroBased <> -1 Then
        For i = n - 1 To 0 Step -1
            If CaseSensitive = True Then
                sTemp = LB.list(i, ColumnIndexZeroBased)
            Else
                This = LCase(This)
                sTemp = LCase(LB.list(i, ColumnIndexZeroBased))
            End If
            If InStr(1, sTemp, This) > 0 Then
                Contains = True
                Exit Function
            End If
        Next i
    Else
        Dim ColumnCount As Long
        n = LB.ListCount
        For i = n - 1 To 0 Step -1
            For ColumnCount = 0 To LB.ColumnCount - 1
                If CaseSensitive = True Then
                    sTemp = LB.list(i, ColumnCount)
                Else
                    This = LCase(This)
                    sTemp = LCase(LB.list(i, ColumnCount))
                End If
                If InStr(1, sTemp, This) > 0 Then
                    Contains = True
                    Exit Function
                End If
            Next ColumnCount
        Next i
    End If
End Function

Public Sub FilterByColumn(This As String, _
                            Optional ColumnIndexZeroBased As Long = -1, _
                            Optional CaseSensitive As Boolean = False)
    Dim i               As Long
    Dim n               As Long
    Dim sTemp           As String
    If ColumnIndexZeroBased > LB.ColumnCount - 1 Or ColumnIndexZeroBased < 0 Then
        ColumnIndexZeroBased = -1
    End If
    n = LB.ListCount
    If ColumnIndexZeroBased <> -1 Then
        For i = n - 1 To 0 Step -1
            If CaseSensitive = True Then
                sTemp = LB.list(i, ColumnIndexZeroBased)
            Else
                This = LCase(This)
                sTemp = LCase(LB.list(i, ColumnIndexZeroBased))
            End If
            If InStr(1, sTemp, This) = 0 Then
                LB.RemoveItem (i)
            End If
        Next i
    Else
        Dim ColumnCount As Long
        n = LB.ListCount
        For i = n - 1 To 0 Step -1
            For ColumnCount = 0 To LB.ColumnCount - 1
                If CaseSensitive = True Then
                    sTemp = LB.list(i, ColumnCount)
                Else
                    This = LCase(This)
                    sTemp = LCase(LB.list(i, ColumnCount))
                End If
                If InStr(1, sTemp, This) > 0 Then
                Else
                    If ColumnCount = LB.ColumnCount - 1 Then
                        LB.RemoveItem (i)
                    End If
                End If
            Next ColumnCount
        Next i
    End If
End Sub

Public Sub SortOnColumn(OnColumn As Long)
    Dim vntData As Variant
    Dim vntTempItem As Variant
    Dim lngOuterIndex As Long
    Dim lngInnerIndex As Long
    Dim lngSubItemIndex As Long
    vntData = LB.list
    For lngOuterIndex = LBound(vntData, 1) To UBound(vntData, 1) - 1
        For lngInnerIndex = lngOuterIndex + 1 To UBound(vntData, 1)
            If vntData(lngOuterIndex, OnColumn) > vntData(lngInnerIndex, OnColumn) Then
                For lngSubItemIndex = 0 To LB.ColumnCount - 1
                    vntTempItem = vntData(lngOuterIndex, lngSubItemIndex)
                    vntData(lngOuterIndex, lngSubItemIndex) = vntData(lngInnerIndex, lngSubItemIndex)
                    vntData(lngInnerIndex, lngSubItemIndex) = vntTempItem
                Next
            End If
        Next lngInnerIndex
    Next lngOuterIndex
    LB.Clear
    LB.list = vntData
End Sub

Public Function selectedIndexes() As Collection
    Dim i As Long
    Dim coll As New Collection
    If LB.ListCount > 0 Then
        For i = 0 To LB.ListCount - 1
            If LB.Selected(i) Then coll.Add i
        Next i
    End If
    Set selectedIndexes = coll
End Function

Public Function SelectedValues() As Collection
    Dim i As Long
    Dim listItem As Long
    Dim selectedCollection As Collection
    Set selectedCollection = New Collection
    Dim listboxCount As Long

        If LB.ListCount > 0 Then
            For i = 0 To LB.ListCount - 1
                If LB.Selected(i) Then
                    selectedCollection.Add LB.list(i, LB.BoundColumn - 1)
                End If
            Next i
        End If

    Set SelectedValues = selectedCollection
End Function

Public Function SelectedCount() As Long
    Dim i As Long
    Dim listItem As Long
    Dim selectedCollection As Collection
    Set selectedCollection = New Collection
    Dim listboxCount As Long
    Dim counter As Long

        If LB.ListCount > 0 Then
            For i = 0 To LB.ListCount - 1
                If LB.Selected(i) = True Then
                    counter = counter + 1
                End If
            Next i
        End If
    SelectedCount = counter
End Function

Public Sub ResizeColumns(Optional ResizeControl As Boolean, Optional ResizeListbox As Boolean)
    If LB.ListCount = 0 Then Exit Sub
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("ListboxColumnwidth", ThisWorkbook)
    Dim rng As Range
    Set rng = ws.Range("A1")
    Set rng = rng.Resize(UBound(LB.list) + 1, LB.ColumnCount)
    rng = LB.list
    rng.Font.Name = LB.Font.Name
    rng.Font.Size = LB.Font.Size + 2
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
    With LB
        .ColumnWidths = sWidth
        .BorderStyle = fmBorderStyleSingle
    End With
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If ResizeListbox = False Then Exit Sub
    Dim W As Long
    Dim i As Long
    For i = LBound(vR) To UBound(vR)
        W = W + vR(i)
    Next
    DoEvents
    LB.Width = W + 10
End Sub

Public Sub DeselectAll()
    If LB.ListCount <> 0 Then
        Dim i As Long
        For i = 0 To LB.ListCount - 1
            LB.Selected(i) = False
        Next i
    End If
End Sub
Public Sub SelectAll()
    If LB.ListCount <> 0 Then
        Dim i As Long
        For i = 0 To LB.ListCount - 1
            LB.Selected(i) = True
        Next i
    End If
End Sub
Public Sub SelectLike(This As String)
    DeselectAll
    If This = "" Then Exit Sub
    Dim i As Long
    For i = 0 To LB.ListCount - 1
        If UCase(LB.list(i, 1)) Like "*" & UCase(This) & "*" Then
            LB.Selected(i) = True
        End If
    Next i
End Sub

Public Sub SortAZ()
    Dim j As Long
    Dim i As Long
    Dim Temp As Variant
    With LB
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If LCase(.list(i)) > LCase(.list(i + 1)) Then
                    Temp = .list(i)
                    .list(i) = .list(i + 1)
                    .list(i + 1) = Temp
                End If
            Next i
        Next j
    End With
End Sub

Public Sub SortZA()
    Dim j As Long
    Dim i As Long
    Dim Temp As Variant
    With LB
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If LCase(.list(i)) < LCase(.list(i + 1)) Then
                    Temp = .list(i)
                    .list(i) = .list(i + 1)
                    .list(i + 1) = Temp
                End If
            Next i
        Next j
    End With
End Sub

'Sub ToRangeSelect(cell As Range)
'    cell.RESIZE(lBox.ListCount, LB.ColumnCount) = CollectionsToArray2D(SelectedValues)
'End Sub

Public Sub SelectedToRange()
   Dim rng As Range
    If GetInputRange(rng, "Range picker", "Select range to output listbox' list") = False Then Exit Sub
    Dim var: var = aCollection.CollectionsToArray2D(SelectedValues)
    rng.Resize(UBound(var, 1), LB.ColumnCount) = var
End Sub



Public Sub RemoveSelected()
    Dim coll As New Collection: Set coll = selectedIndexes
    If coll.count = 0 Then Exit Sub
    Dim i As Long
    For i = coll.count To 1 Step -1
        LB.RemoveItem coll(i)
    Next
End Sub



'Private Sub ListViewControl_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim dbBefore As Long: dbBefore = ListBox1.ListCount
'    Dim FileFullPath As String
'    Dim fileItem As Long
'    Dim objFSO As scripting.FileSystemObject
'    Dim objTopFolder As scripting.Folder
'    Set objFSO = CreateObject("Scripting.FileSystemObject")
'    Dim var As Variant, element As Variant
'    For fileItem = 1 To Data.Files.Count
'        FileFullPath = Data.Files(fileItem)
'        If oLogFiles = True Then
'            If LCase(IsFileFolderURL(FileFullPath)) = "f" Then
'                var = Split(TextBox2.TEXT, ",")
'                On Error Resume Next
'                If Left(Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1), 1) <> "~" And (var(0) = "*" Or var(0) = "") Then GoTo PASS
'                For Each element In var
'                    If InStr(1, FileFullPath, element, vbTextCompare) > 0 And Left(Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1), 1) <> "~" Then
'PASS:
'                        If Not aListBox.Init(ListBox1).Contains(FileFullPath) Then
'                            AddToListBox ListBox1, FileFullPath
'                        End If
'                    End If
'                Next
'            Else        'if drag dropped folder
'                Set objTopFolder = objFSO.getFolder(FileFullPath)
'                FileRecursive objTopFolder, oSearchInSubfolders.Value
'            End If
'        End If
'        If oLogFolders = True Then
'            If UCase(IsFileFolderURL(FileFullPath)) = "D" Then
'                Set objTopFolder = objFSO.getFolder(FileFullPath)
'                If Not aListBox.Init(ListBox1).Contains(objTopFolder.Path & "\") Then
'                    AddToListBox ListBox1, objTopFolder.Path
'                End If
'                FolderRecursive objTopFolder, oSearchInSubfolders.Value
'            End If
'        End If
'    Next fileItem
'
'    If ListBox1.ListCount - dbBefore > 0 Then ListboxToDatabaseSheet
'
'    Set objFSO = Nothing
'    Set objTopFolder = Nothing
'End Sub

