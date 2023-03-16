VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LB As Object
Private oForm As UserForm

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If


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

Public Sub HeightToEntries(ByVal NumberOfEntries As Long)
    Const SM_CYEDGE = 46&
    With LB
        NumberOfEntries = IIf(NumberOfEntries > .ListCount, .ListCount, NumberOfEntries)
        .Height = ((9.75 * NumberOfEntries) + IIf(.SpecialEffect = fmSpecialEffectFlat, 0, GetSystemMetrics(SM_CYEDGE)))
    End With
End Sub

Public Function Init(TargetListBox As Object) As aListBox
    Set LB = TargetListBox
    Set oForm = LB.Parent
    Set Init = Me
End Function

Public Sub LoadVBProjects()
    Dim coll As New Collection
    Dim WB As Workbook
    For Each WB In Workbooks
        If Len(WB.path) > 0 Then
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
            If LB.List(i) = CStr(This) Then
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
                    If LB.List(i) = el Then
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
    Header.clear
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
    Dim N      As Long
    Dim sTemp  As String
    If ColumnIndexZeroBased > LB.ColumnCount - 1 Or ColumnIndexZeroBased < 0 Then
        ColumnIndexZeroBased = -1
    End If
    N = LB.ListCount
    If ColumnIndexZeroBased <> -1 Then
        For i = N - 1 To 0 Step -1
            If CaseSensitive = True Then
                sTemp = LB.List(i, ColumnIndexZeroBased)
            Else
                This = LCase(This)
                sTemp = LCase(LB.List(i, ColumnIndexZeroBased))
            End If
            If InStr(1, sTemp, This) > 0 Then
                Contains = True
                Exit Function
            End If
        Next i
    Else
        Dim ColumnCount As Long
        N = LB.ListCount
        For i = N - 1 To 0 Step -1
            For ColumnCount = 0 To LB.ColumnCount - 1
                If CaseSensitive = True Then
                    sTemp = LB.List(i, ColumnCount)
                Else
                    This = LCase(This)
                    sTemp = LCase(LB.List(i, ColumnCount))
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
    Dim N               As Long
    Dim sTemp           As String
    If ColumnIndexZeroBased > LB.ColumnCount - 1 Or ColumnIndexZeroBased < 0 Then
        ColumnIndexZeroBased = -1
    End If
    N = LB.ListCount
    If ColumnIndexZeroBased <> -1 Then
        For i = N - 1 To 0 Step -1
            If CaseSensitive = True Then
                sTemp = LB.List(i, ColumnIndexZeroBased)
            Else
                This = LCase(This)
                sTemp = LCase(LB.List(i, ColumnIndexZeroBased))
            End If
            If InStr(1, sTemp, This) = 0 Then
                LB.RemoveItem (i)
            End If
        Next i
    Else
        Dim ColumnCount As Long
        N = LB.ListCount
        For i = N - 1 To 0 Step -1
            For ColumnCount = 0 To LB.ColumnCount - 1
                If CaseSensitive = True Then
                    sTemp = LB.List(i, ColumnCount)
                Else
                    This = LCase(This)
                    sTemp = LCase(LB.List(i, ColumnCount))
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
    vntData = LB.List
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
    LB.clear
    LB.List = vntData
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
                    selectedCollection.Add LB.List(i, LB.BoundColumn - 1)
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
    Dim WS As Worksheet
    Set WS = CreateOrSetSheet("ListboxColumnwidth", ThisWorkbook)
    Dim rng As Range
    Set rng = WS.Range("A1")
    Set rng = rng.RESIZE(UBound(LB.List) + 1, LB.ColumnCount)
    rng = LB.List
    rng.Font.Name = LB.Font.Name
    rng.Font.Size = LB.Font.Size + 2
    rng.Columns.AutoFit
    Dim sWidth As String
    Dim vR() As Variant
    Dim N As Integer
    Dim cell As Range
    For Each cell In rng.RESIZE(1)
        N = N + 1
        ReDim Preserve vR(1 To N)
        vR(N) = cell.EntireColumn.Width
    Next cell
    sWidth = Join(vR, ";")
    With LB
        .ColumnWidths = sWidth
        .BorderStyle = fmBorderStyleSingle
    End With
    Application.DisplayAlerts = False
    WS.Delete
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
        If UCase(LB.List(i, 1)) Like "*" & UCase(This) & "*" Then
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
    With LB
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
'    cell.RESIZE(lBox.ListCount, LB.ColumnCount) = CollectionsToArray2D(SelectedValues)
'End Sub

Public Sub SelectedToRange()
   Dim rng As Range
    If GetInputRange(rng, "Range picker", "Select range to output listbox' list") = False Then Exit Sub
    Dim var: var = aCollection.CollectionsToArray2D(SelectedValues)
    rng.RESIZE(UBound(var, 1), LB.ColumnCount) = var
End Sub





'------------------
'in userform
'------------------
'
'Private Type POINTAPI
'  X As Long
'  Y As Long
'End Type
'
'#If VBA7 Then
'    Private Type MSG
'        hWnd As LongPtr
'        message As Long
'        wParam As LongPtr
'        lParam As LongPtr
'        time As Long
'        pt As POINTAPI
'    End Type
'
'    Private Declare PtrSafe Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
'    Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As LongPtr
'    Private Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
'    Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As LongPtr) As Long
'    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
'    Private Declare PtrSafe Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As LongPtr, ByVal fAccept As Long)
'    Private Declare PtrSafe Sub DragFinish Lib "shell32.dll" (ByVal HDROP As LongPtr)
'    Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As LongPtr, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
'#Else
'
'    Private Type MSG
'        hWnd As Long
'        message As Long
'        wParam As Long
'        lParam As Long
'        time As Long
'        pt As POINTAPI
'    End Type
'
'    Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
'    Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
'    Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
'    Private Declare Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As Long) As Long
'    Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'    Private Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As Long, ByVal fAccept As Long)
'    Private Declare Sub DragFinish Lib "shell32.dll" (ByVal HDROP As Long)
'    Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
'#End If
'
'Private Sub UserForm_Activate()
'
'   #If VBA7 Then
'        Dim hWnd As LongPtr, HDROP As LongPtr
'    #Else
'        Dim hWnd As Long, HDROP As Long
'    #End If
'
'    Const WM_DROPFILES = &H233
'    Dim tMsg As MSG, sFileName As String * 256
'    Dim lFilesCount As Long, i As Long
'
'
'    Call WindowFromAccessibleObject(Me, hWnd)
'    Call DragAcceptFiles(ListBox1.[_GethWnd], True)
'
'    Do While GetMessage(tMsg, 0, 0, 0) And IsWindow(hWnd)
'        If tMsg.message = WM_DROPFILES Then
'            HDROP = tMsg.wParam
'            lFilesCount = DragQueryFile(HDROP, &HFFFFFFFF, 0, 0)
'            If lFilesCount Then
'                For i = 0 To lFilesCount - 1
'                    Dim CleanName As String
'                    CleanName = Left(sFileName, DragQueryFile(HDROP, i, sFileName, Len(sFileName)))
'                    If isFDU(CleanName) = "F" Then
'                        ListBox1.AddItem
'                        ListBox1.List(ListBox1.ListCount - 1, 0) = Mid(CleanName, InStrRev(CleanName, "\") + 1)
'                        ListBox1.List(ListBox1.ListCount - 1, 1) = CleanName
'                    Else
'                        Dim element As Variant
'                        Dim out As New Collection
'                        FilesAndOrFoldersInFolderOrZip CleanName, oLogFolders, oLogFiles, oSearchInSubfolders, out
'                        For Each element In out
'                            ListBox1.AddItem element
'                            ListBox1.List(ListBox1.ListCount - 1, 0) = Mid(element, InStrRev(element, "\") + 1)
'                            ListBox1.List(ListBox1.ListCount - 1, 1) = element
'                        Next
'                    End If
'                Next i
'            End If
'            Call DragFinish(HDROP)
'        End If
'        Call TranslateMessage(tMsg)
'        Call DispatchMessage(tMsg)
'    Loop
'End Sub

