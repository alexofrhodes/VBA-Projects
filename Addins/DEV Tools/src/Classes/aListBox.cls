VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@TODO C:\Users\acer\Dropbox\SOFTWARE\EXCEL\Manfred van den Noort\      '-> listbox_drag_drop_advanced

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

Option Explicit


Public WithEvents ListboxEvents As MSForms.ListBox
Attribute ListboxEvents.VB_VarHelpID = -1

Private oListBox As MSForms.ListBox
Private oForm As Object

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Type POINTAPI
    x               As Long
    y               As Long
End Type

#If VBA7 Then
    
    Private Type Msg
        hwnd            As LongPtr
        message         As Long
        wParam          As LongPtr
        lParam          As LongPtr
        time            As Long
        pt              As POINTAPI
    End Type
    
    #If Win64 Then
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal POINT As LongPtr) As LongPtr
    #Else
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
    #End If
    
    Private Declare PtrSafe Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
    Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As LongPtr
    Private Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
    Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As LongPtr) As Long
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As LongPtr, ByVal fAccept As Long)
    Private Declare PtrSafe Sub DragFinish Lib "shell32.dll" (ByVal HDROP As LongPtr)
    Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As LongPtr, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

#Else
    
    Private Type Msg
        hwnd            As Long
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

Public Function Init(ByRef TargetListBox As MSForms.ListBox) As aListBox
    Set oListBox = TargetListBox
    Set oForm = oListBox.Parent
    Set Init = Me
    
'    Me.AutofitColumns  '@TODO where to put this? needed for TargetColumn, in case they were not manually set
End Function

Public Function Parent()
    Set Parent = oForm
End Function

Public Sub EnableEvents1()
        Set ListboxEvents = oListBox
End Sub

Public Sub listboxevents_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Toast SelectedRowsText
End Sub


Public Function TargetColumn(x, y)
        ' Calculate the column index based on the X-coordinate
        Dim columnIndex As Long
        Dim columnWidth As Long
        Dim totalWidth As Long
        Dim totalColumns As Long
        
        totalColumns = oListBox.ColumnCount
        
        ' Calculate the total width of all columns
        For columnIndex = 1 To totalColumns
            columnWidth = Replace(Split(oListBox.ColumnWidths, ";")(columnIndex - 1), " pt", "")
            totalWidth = totalWidth + columnWidth
        Next columnIndex
        
        ' Calculate the approximate column index based on the X-coordinate
        Dim clickedX As Long
        clickedX = x
        
        If clickedX > 0 Then
            Dim cumulativeWidth As Long
            Dim lastColumnIndex As Long
            
            For columnIndex = 1 To totalColumns
                columnWidth = Replace(Split(oListBox.ColumnWidths, ";")(columnIndex - 1), " pt", "")
                cumulativeWidth = cumulativeWidth + columnWidth
                
                If clickedX < cumulativeWidth Then
                    ' The mouse click is within this column
                    lastColumnIndex = columnIndex
                    Exit For
                End If
            Next columnIndex
            
            ' Display the clicked column index
            TargetColumn = lastColumnIndex
        End If
End Function

Sub AutofitColumns(Optional ResizeListbox As Boolean)
 Application.ScreenUpdating = False
    Dim ws As Worksheet
    If sheetExists("ListboxColumnWidth", ThisWorkbook) = False Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ListboxColumnwidth"
    Else
        Set ws = ThisWorkbook.Worksheets("ListboxColumnwidth")
        ws.Cells.Clear
    End If
    
    '---Listbox to range-----
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("ListboxColumnwidth").Range("A1")
    Set rng = rng.Resize(UBound(oListBox.List) + 1, oListBox.ColumnCount)
    rng = oListBox.List
    
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
    With oListBox
        .ColumnWidths = sWidth
        '.RowSource = "A1:A3"
        .BorderStyle = fmBorderStyleSingle
    End With
        
    'remove worksheet
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    
    '----Resize Listbox--------
    If ResizeListbox = False Then Exit Sub
    Dim w As Long
    Dim i As Long
    For i = LBound(vR) To UBound(vR)
        w = w + vR(i)
    Next
    DoEvents
    oListBox.Width = w + 10
End Sub

Private Function sheetExists(sheetToFind As String, Optional InWorkbook As Workbook) As Boolean
    If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook
    On Error Resume Next
    sheetExists = Not InWorkbook.Sheets(sheetToFind) Is Nothing
End Function

Function LoadCSV(TargetFile As String, Clear As Boolean)
    If Clear Then oListBox.Clear
    Dim ArrayOfElements
    Dim s: s = TxtRead(TargetFile)
    Dim arr
    arr = Split(s, vbLf)
    Dim iCols, iRows
    iRows = UBound(arr) + 1
    iCols = UBound(Split(arr(0), ",")) + 1
    ReDim ArrayOfElements(1 To iRows, 1 To iCols)
    Dim x, y
    For x = 1 To iRows
        For y = 1 To iCols
            ArrayOfElements(x, y) = Split(arr(x - 1), ",")(y - 1)
        Next
    Next
    oListBox.ColumnCount = iCols
    oListBox.List = ArrayOfElements
End Function

Sub ToRange(cell As Range)
    cell.Resize(oListBox.ListCount, oListBox.ColumnCount) = oListBox.List
End Sub

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
    
    'you can use .[_GethWnd] with controls which support it
    Dim ctr As control: Set ctr = oListBox
    Dim lGethWnd As LongPtr: lGethWnd = ctr.[_GethWnd]

#If VBA7 Then
    Dim hwnd As LongPtr, HDROP As LongPtr
#Else
    Dim hwnd As Long, HDROP As Long
#End If

    Const WM_DROPFILES = &H233
    Dim tMsg As Msg, sFileName As String * 1000
    Dim lFilesCount As Long, i As Long
    hwnd = aUserform.Init(oForm).hwnd
    Call DragAcceptFiles(lGethWnd, True)

    Dim sPathFile   As String

    Do While GetMessage(tMsg, 0, 0, 0) And IsWindow(hwnd)
        If tMsg.message = WM_DROPFILES Then
            HDROP = tMsg.wParam
            lFilesCount = DragQueryFile(HDROP, &HFFFFFFFF, 0, 0)
            If lFilesCount Then
                For i = 0 To lFilesCount - 1
                    sPathFile = VBA.Trim$(VBA.Left(sFileName, DragQueryFile(HDROP, i, sFileName, VBA.Len(sFileName))))
                    Call getFilesFromPath(sPathFile, sExpansion, iDeepSubPath)
                Next i
            End If
            Call DragFinish(HDROP)
        End If
        Call TranslateMessage(tMsg)
        Call DispatchMessage(tMsg)
    Loop
End Sub

Private Sub getFilesFromPath(ByVal sPathFile As String, ByVal sExpansion As String, ByVal iDeepSubPath As Integer)
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
            Call filterFiles(sFile, sExpansion, sExp)
        Next i
    Else
        Call filterFiles(sPathFile, sExpansion, sExp)
    End If
End Sub

Private Sub filterFiles(ByVal sPathFile As String, ByVal sExpansion As String, ByVal sExp As String)
    sExp = "*." & sExp
    If sExpansion = vbNullString Or sExpansion = "*.*" Then
    ElseIf Not likeExp(sExpansion, sExp) Then
        Exit Sub
    End If
    oListBox.AddItem sPathFile
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

Private Function sGetExtensionName(ByVal sPathFile As String) As String
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
    With oListBox
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
        wbPath = vbProj.FileName
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
    oListBox.AddItem el
    Next

    Exit Sub
ErrorHandler:
    If Err.Number = 76 Then GoTo SKIP
End Sub

Public Sub SelectItems(This As Variant, Optional ByIndex As Boolean)
    Dim i As Long
    Select Case TypeName(This)
    Case Is = "String", "Long", "Integer"
        For i = 0 To oListBox.ListCount - 1
            If oListBox.List(i) = CStr(This) Then
                oListBox.Selected(i) = True
                DoEvents
                If oListBox.multiSelect = fmMultiSelectSingle Then Exit Sub
            End If
        Next
    Case Else
        Dim el As Variant
        If ByIndex Then
            For Each el In This
                oListBox.Selected(el) = True
            Next
        Else
            For Each el In This
                For i = 0 To oListBox.ListCount - 1
                    If oListBox.List(i) = el Then
                        oListBox.Selected(i) = True
                        DoEvents
                    End If
                Next
            Next
        End If
    End Select
End Sub

Public Sub AddHeader(Header As MSForms.ListBox, arrHeaders)
    Header.Width = oListBox.Width
    Dim i As Long
    Header.ColumnCount = oListBox.ColumnCount
    Header.ColumnWidths = oListBox.ColumnWidths
    Header.Clear
    Header.AddItem
    If ArrayDimensions(arrHeaders) = 1 Then
        For i = 0 To UBound(arrHeaders)
            Header.List(0, i) = arrHeaders(i)
        Next i
    Else
        For i = 1 To UBound(arrHeaders, 2)
            Header.List(0, i - 1) = arrHeaders(1, i)
        Next i
    End If
'    oListBox.ZOrder (1)
'    Header.ZOrder (0)
    Header.SpecialEffect = fmSpecialEffectFlat
    Header.BackColor = RGB(200, 200, 200)
    Header.Height = 15
    Header.Width = oListBox.Width
    Header.Left = oListBox.Left
    Header.Top = oListBox.Top - Header.Height - 1
    Header.Font.Bold = True
    Header.Font.Name = "Comic Sans MS"
    Header.Font.Size = oListBox.Font.Size
End Sub

Public Function Contains(This As String, _
                Optional ColumnIndexZeroBased As Long = -1, _
                Optional CaseSensitive As Boolean = False) As Boolean
    Dim i      As Long
    Dim n      As Long
    Dim sTemp  As String
    If ColumnIndexZeroBased > oListBox.ColumnCount - 1 Or ColumnIndexZeroBased < 0 Then
        ColumnIndexZeroBased = -1
    End If
    n = oListBox.ListCount
    If ColumnIndexZeroBased <> -1 Then
        For i = n - 1 To 0 Step -1
            If CaseSensitive = True Then
                sTemp = oListBox.List(i, ColumnIndexZeroBased)
            Else
                This = LCase(This)
                sTemp = LCase(oListBox.List(i, ColumnIndexZeroBased))
            End If
            If InStr(1, sTemp, This) > 0 Then
                Contains = True
                Exit Function
            End If
        Next i
    Else
        Dim ColumnCount As Long
        n = oListBox.ListCount
        For i = n - 1 To 0 Step -1
            For ColumnCount = 0 To oListBox.ColumnCount - 1
                If CaseSensitive = True Then
                    sTemp = oListBox.List(i, ColumnCount)
                Else
                    This = LCase(This)
                    sTemp = LCase(oListBox.List(i, ColumnCount))
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
    If ColumnIndexZeroBased > oListBox.ColumnCount - 1 Or ColumnIndexZeroBased < 0 Then
        ColumnIndexZeroBased = -1
    End If
    n = oListBox.ListCount
    If ColumnIndexZeroBased <> -1 Then
        For i = n - 1 To 0 Step -1
            If CaseSensitive = True Then
                sTemp = oListBox.List(i, ColumnIndexZeroBased)
            Else
                This = LCase(This)
                sTemp = LCase(oListBox.List(i, ColumnIndexZeroBased))
            End If
            If InStr(1, sTemp, This) = 0 Then
                oListBox.RemoveItem (i)
            End If
        Next i
    Else
        Dim ColumnCount As Long
        n = oListBox.ListCount
        For i = n - 1 To 0 Step -1
            For ColumnCount = 0 To oListBox.ColumnCount - 1
                If CaseSensitive = True Then
                    sTemp = oListBox.List(i, ColumnCount)
                Else
                    This = LCase(This)
                    sTemp = LCase(oListBox.List(i, ColumnCount))
                End If
                If InStr(1, sTemp, This) > 0 Then
                Else
                    If ColumnCount = oListBox.ColumnCount - 1 Then
                        oListBox.RemoveItem (i)
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
    vntData = oListBox.List
    For lngOuterIndex = LBound(vntData, 1) To UBound(vntData, 1) - 1
        For lngInnerIndex = lngOuterIndex + 1 To UBound(vntData, 1)
            If vntData(lngOuterIndex, OnColumn) > vntData(lngInnerIndex, OnColumn) Then
                For lngSubItemIndex = 0 To oListBox.ColumnCount - 1
                    vntTempItem = vntData(lngOuterIndex, lngSubItemIndex)
                    vntData(lngOuterIndex, lngSubItemIndex) = vntData(lngInnerIndex, lngSubItemIndex)
                    vntData(lngInnerIndex, lngSubItemIndex) = vntTempItem
                Next
            End If
        Next lngInnerIndex
    Next lngOuterIndex
    oListBox.Clear
    oListBox.List = vntData
End Sub

Public Function selectedIndexes() As Collection
    Dim i As Long
    Dim coll As New Collection
    If oListBox.ListCount > 0 Then
        For i = 0 To oListBox.ListCount - 1
            If oListBox.Selected(i) Then coll.Add i
        Next i
    End If
    Set selectedIndexes = coll
End Function

Public Function SelectedValues() As Collection 'single column
    Dim i As Long
    Dim listItem As Long
    Dim selectedCollection As Collection
    Set selectedCollection = New Collection
    Dim listboxCount As Long

        If oListBox.ListCount > 0 Then
            For i = 0 To oListBox.ListCount - 1
                If oListBox.Selected(i) Then
                    selectedCollection.Add oListBox.List(i, oListBox.BoundColumn - 1)
                End If
            Next i
        End If

    Set SelectedValues = selectedCollection
End Function


Public Function SelectedRowsText() As String
    Dim x As Long, y As Long
    Dim listItem As Long
    Dim out As String
    Dim indexColl As New Collection
    Set indexColl = selectedIndexes
    For x = 1 To indexColl.count
        For y = 0 To oListBox.ColumnCount - 1
            out = out & IIf(y > 0, "|", "") & oListBox.List(indexColl(x), y)
        Next
        If x < indexColl.count Then out = out & vbNewLine
    Next
    SelectedRowsText = out
End Function


Public Function SelectedCount() As Long
    Dim i As Long
    Dim listItem As Long
    Dim selectedCollection As Collection
    Set selectedCollection = New Collection
    Dim listboxCount As Long
    Dim counter As Long

        If oListBox.ListCount > 0 Then
            For i = 0 To oListBox.ListCount - 1
                If oListBox.Selected(i) = True Then
                    counter = counter + 1
                End If
            Next i
        End If
    SelectedCount = counter
End Function

Public Sub DeselectAll()
    If oListBox.ListCount <> 0 Then
        Dim i As Long
        For i = 0 To oListBox.ListCount - 1
            oListBox.Selected(i) = False
        Next i
    End If
End Sub

Public Sub SelectAll()
    If oListBox.ListCount <> 0 Then
        Dim i As Long
        For i = 0 To oListBox.ListCount - 1
            oListBox.Selected(i) = True
        Next i
    End If
End Sub

Public Sub SelectLike(This As String)
    DeselectAll
    If This = "" Then Exit Sub
    Dim i As Long
    For i = 0 To oListBox.ListCount - 1
        If UCase(oListBox.List(i, 1)) Like "*" & UCase(This) & "*" Then
            oListBox.Selected(i) = True
        End If
    Next i
End Sub

Public Sub SortAZ()
    Dim j As Long
    Dim i As Long
    Dim Temp As Variant
    With oListBox
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If LCase(.List(i)) > LCase(.List(i + 1)) Then
                    Temp = .List(i)
                    .List(i) = .List(i + 1)
                    .List(i + 1) = Temp
                End If
            Next i
        Next j
    End With
End Sub

Public Sub SortZA()
    Dim j As Long
    Dim i As Long
    Dim Temp As Variant
    With oListBox
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If LCase(.List(i)) < LCase(.List(i + 1)) Then
                    Temp = .List(i)
                    .List(i) = .List(i + 1)
                    .List(i + 1) = Temp
                End If
            Next i
        Next j
    End With
End Sub

'Sub ToRangeSelect(cell As Range)
'    cell.RESIZE(lBox.ListCount, oListBox.ColumnCount) = CollectionsToArray2D(SelectedValues)
'End Sub

Public Sub SelectedToRange()
   Dim rng As Range
    If GetInputRange(rng, "Range picker", "Select range to output listbox' list") = False Then Exit Sub
    Dim var: var = aCollection.CollectionsToArray2D(SelectedValues)
    rng.Resize(UBound(var, 1), oListBox.ColumnCount) = var
End Sub

Public Sub RemoveSelected()
    Dim coll As New Collection: Set coll = selectedIndexes
    If coll.count = 0 Then Exit Sub
    Dim i As Long
    For i = coll.count To 1 Step -1
        oListBox.RemoveItem coll(i)
    Next
End Sub


