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
'* Modified   : 21-07-2023 08:28    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit



Public WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Public WithEvents DoubleClickEvents As msforms.ListBox
Attribute DoubleClickEvents.VB_VarHelpID = -1

Public WithEvents ExtendedSelectionEvents As msforms.ListBox
Attribute ExtendedSelectionEvents.VB_VarHelpID = -1
Private ItemsSelected As Variant


Public WithEvents DragSource As msforms.ListBox
Attribute DragSource.VB_VarHelpID = -1
Public WithEvents DragTarget As msforms.ListBox
Attribute DragTarget.VB_VarHelpID = -1
Private from_lb_index As Long
Private to_lb_index As Long
Private dragindexes As Collection
Private dragArray
Private isUpdatingSelection As Boolean



Private oListBox    As msforms.ListBox
Private oList
Private oForm       As Object
Private previousFilterLength As Long

Private isDoubleClick As Boolean

#If VBA7 Then
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
Private Declare Function  GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

Private Type POINTAPI
    x               As Long
    y               As Long
End Type

#If VBA7 Then

Private Type msg
    hWnd            As LongPtr
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

Private Declare PtrSafe Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As LongPtr
Private Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As LongPtr) As Long
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As LongPtr, ByVal fAccept As Long)
Private Declare PtrSafe Sub DragFinish Lib "shell32.dll" (ByVal HDROP As LongPtr)
Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As LongPtr, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

#Else

Private Type msg
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


Public Function Init(ByRef TargetListBox As msforms.ListBox) As aListBox
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Set oListBox = TargetListBox
    Set oForm = oListBox.Parent
    Set Init = Me
End Function

Public Function Parent()
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Set Parent = oForm
End Function
Sub RememberList()
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    oList = oListBox.list
End Sub



Private Sub Emitter_Click(control As Object)
    '@AssignedModule aListBox
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aListBox
    If control.Name Like "lblHeader-*" Then
        Dim col     As Long
        col = Split(control.Name, "-")(1)
        Me.SortOnColumn col - 1
    End If
End Sub
Private Sub Emitter_Keyup(control As Object, KeyCode As msforms.ReturnInteger, Shift As Integer)
    '@AssignedModule aListBox
    '@INCLUDE PROCEDURE ArrayAllocated
    '@INCLUDE PROCEDURE ArrayFilter2D
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aListBox
    If Not ArrayAllocated(oList) Then RememberList
    If control.Name = "txtFilter" Then
        If Len(control.TEXT) = 0 Then
            oListBox.list = oList
        ElseIf Len(control.TEXT) > previousFilterLength Then
            oListBox.list = ArrayFilter2D(oListBox.list, control.TEXT, oForm.Controls("cbxFilter").TEXT)
        ElseIf Len(control.TEXT) < previousFilterLength Then
            oListBox.list = ArrayFilter2D(oList, control.TEXT, oForm.Controls("cbxFilter").TEXT)
        End If
        previousFilterLength = Len(control.TEXT)
    End If

    If control.Name = "cbxFilter" Then
        oForm.Controls("txtFilter").TEXT = ""
        previousFilterLength = 0
    End If
End Sub


Public Sub ListenToDragDrop(Optional ToLbx As msforms.ListBox)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Set DragSource = oListBox
    If Not ToLbx Is Nothing Then
        Set DragTarget = ToLbx
    Else
        Set DragTarget = oListBox
    End If
    DragTarget.columnCount = DragSource.columnCount
    DragSource.MousePointer = 5
End Sub


Private Sub DragSource_Change()
    '@AssignedModule aListBox
    '@INCLUDE PROCEDURE ArrayAllocated
    '@INCLUDE CLASS aListBox
    Dim AltPressed  As Boolean: AltPressed = CBool(GetKeyState(&H12) And &H8000)
    Dim CtrlPressed As Boolean: CtrlPressed = CBool(GetKeyState(&H11) And &H8000)
    Dim ShiftPressed As Boolean: ShiftPressed = CBool(GetKeyState(&HA0) And &HFF80)

    If AltPressed Then Exit Sub

    If Not isUpdatingSelection _
            And ArrayAllocated(ItemsSelected) _
            And Not CtrlPressed _
            And Not ShiftPressed Then
        isUpdatingSelection = True
        With DragSource
            Dim i   As Long
            For i = 0 To .ListCount - 1
                If ItemsSelected(i) <> .Selected(i) Then
                    .Selected(i) = ItemsSelected(i)
                End If
            Next i
        End With
        isUpdatingSelection = False
    End If
End Sub
Private Sub DragSource_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    '@AssignedModule aListBox
    '@INCLUDE PROCEDURE ArrayAllocated
    '@INCLUDE CLASS aListBox
    Dim AltPressed  As Boolean: AltPressed = CBool(GetKeyState(&H12) And &H8000)
    Dim CtrlPressed As Boolean: CtrlPressed = CBool(GetKeyState(&H11) And &H8000)
    Dim ShiftPressed As Boolean: ShiftPressed = CBool(GetKeyState(&HA0) And &HFF80)

    If AltPressed Then Exit Sub

    If Not CtrlPressed And Not ShiftPressed And ArrayAllocated(ItemsSelected) Then
        Dim index   As Long: index = Int(y * 0.85 / DragSource.Font.Size)
        If index > UBound(ItemsSelected) Then Exit Sub
        ItemsSelected(index) = True
        With DragSource
            Dim i   As Long
            For i = 0 To .ListCount - 1
                ItemsSelected(i) = .Selected(i)
            Next i
        End With
    End If

End Sub
Private Sub DragSource_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    With DragSource
        ReDim ItemsSelected(0 To .ListCount - 1) As Boolean
        Dim i       As Long
        For i = 0 To .ListCount - 1
            ItemsSelected(i) = .Selected(i)
        Next i
    End With
End Sub

Private Sub DragTarget_BeforeDragOver(ByVal Cancel As msforms.ReturnBoolean, ByVal Data As msforms.DataObject, ByVal x As Single, ByVal y As Single, ByVal DragState As msforms.fmDragState, ByVal Effect As msforms.ReturnEffect, ByVal Shift As Integer)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Effect = fmDropEffectMove
    Cancel = True
End Sub
Private Sub DragTarget_BeforeDropOrPaste(ByVal Cancel As msforms.ReturnBoolean, ByVal Action As msforms.fmAction, ByVal Data As msforms.DataObject, ByVal x As Single, ByVal y As Single, ByVal Effect As msforms.ReturnEffect, ByVal Shift As Integer)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE DECLARATION Action
    With DragTarget
        to_lb_index = Int(y * 0.85 / .Font.Size)
        If to_lb_index >= .ListCount Then to_lb_index = .ListCount
        Cancel = True
        Effect = fmDropEffectMove
        Dim difIndexes As Long
        Dim indexesCounter As Long
        For indexesCounter = dragindexes.Count To 1 Step -1
            DragSource.RemoveItem dragindexes(indexesCounter)
            If dragindexes(indexesCounter) < to_lb_index Then to_lb_index = to_lb_index - 1
        Next

        For indexesCounter = 1 To dragindexes.Count    'To 1 Step -1
            If DragTarget.ListCount = 0 Then
                .AddItem
            Else
                .AddItem , to_lb_index    'Data.GetText, to_lb_index
            End If
            Dim i   As Long

            Dim cellTxt As String
            For i = 1 To UBound(dragArray, 2)
                cellTxt = IIf( _
                        IsNull(dragArray(indexesCounter, i)), _
                        "", _
                        dragArray(indexesCounter, i))
                DragTarget.list(to_lb_index, i - 1) = cellTxt
            Next
            to_lb_index = to_lb_index + 1
        Next
    End With

    ClearSelection
    DoEvents
    oForm.Repaint
End Sub
Private Sub DragSource_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim obj         As DataObject
    Dim drageff     As Long
    If Button = 1 Then
        Set obj = New DataObject
        '        obj.SetText DragSource.text

        from_lb_index = DragSource.ListIndex

        Set dragindexes = selectedIndexes  'test
        dragArray = SelectedRowsArray

        drageff = obj.StartDrag
        Set obj = Nothing
    End If
End Sub




Public Sub ListenToExtendedSelection()
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Set ExtendedSelectionEvents = oListBox
End Sub
Private Sub ExtendedSelectionEvents_Change()
    'author: Petr Radimersky
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim i           As Long
    If CBool(GetKeyState(&HA0) And &HFF80) Then
        With ExtendedSelectionEvents
            For i = 0 To .ListCount - 1
                If ItemsSelected(i) = False And .Selected(i) = True Then
                    ItemsSelected(i) = True
                End If
            Next i
        End With
    End If
End Sub
Private Sub ExtendedSelectionEvents_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim i           As Long
    If CBool(GetKeyState(&HA0) And &HFF80) Then
        With ExtendedSelectionEvents
            ReDim ItemsSelected(0 To .ListCount - 1) As Boolean
            For i = 0 To .ListCount - 1
                ItemsSelected(i) = .Selected(i)
            Next i
        End With
    End If
End Sub
Private Sub ExtendedSelectionEvents_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim i           As Long
    If CBool(GetKeyState(&HA0) And &HFF80) Then
        With ExtendedSelectionEvents
            For i = 0 To .ListCount - 1
                .Selected(i) = ItemsSelected(i)
            Next i
        End With
    End If
End Sub



Public Sub ListenToDoubleClick()
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Set DoubleClickEvents = oListBox
End Sub

Private Sub DoubleClickEvents_DblClick(ByVal Cancel As msforms.ReturnBoolean)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    isDoubleClick = True
End Sub

Public Sub DoubleClickEvents_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    If Not isDoubleClick Then Exit Sub
    MsgBox oListBox.list(oListBox.ListIndex, targetColumn(x, y) - 1)
    isDoubleClick = False
End Sub

Public Function targetColumn(x, y)
    ' Calculate the column index based on the X-coordinate
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim ColumnIndex As Long
    Dim columnWidth As Long
    Dim totalWidth  As Long
    Dim totalColumns As Long

    totalColumns = oListBox.columnCount

    ' Calculate the total width of all columns
    For ColumnIndex = 1 To totalColumns
        columnWidth = Replace(Split(oListBox.ColumnWidths, ";")(ColumnIndex - 1), " pt", "")
        totalWidth = totalWidth + columnWidth
    Next ColumnIndex

    ' Calculate the approximate column index based on the X-coordinate
    Dim clickedX    As Long
    clickedX = x

    If clickedX > 0 Then
        Dim cumulativeWidth As Long
        Dim lastColumnIndex As Long

        For ColumnIndex = 1 To totalColumns
            columnWidth = Replace(Split(oListBox.ColumnWidths, ";")(ColumnIndex - 1), " pt", "")
            cumulativeWidth = cumulativeWidth + columnWidth

            If clickedX < cumulativeWidth Then
                ' The mouse click is within this column
                lastColumnIndex = ColumnIndex
                Exit For
            End If
        Next ColumnIndex

        ' Display the clicked column index
        targetColumn = lastColumnIndex
    End If
End Function


Public Sub AutofitColumns(ResizeListbox As Boolean)
    '@AssignedModule aListBox
    '@INCLUDE PROCEDURE Min
    '@INCLUDE CLASS aListBox
    '@INCLUDE CLASS aUserform
    oListBox.Font.Name = "consolas"    '<---!
    Dim lRowCt      As Long
    Dim lColCt      As Long
    Dim lLengths()
    Dim mvTable: mvTable = oListBox.list
    ReDim lLengths(UBound(mvTable, 2))
    With oListBox
        For lRowCt = LBound(mvTable, 1) To UBound(mvTable, 1)
            For lColCt = LBound(mvTable, 2) To UBound(mvTable, 2)
                lLengths(lColCt) = Application.Max(4, lLengths(lColCt), Len(mvTable(lRowCt, lColCt)))
            Next
        Next
    End With

    Dim lCt         As Long
    Dim sWidths     As String
    Dim dTotWidth   As Double
    Dim lblHidden   As msforms.control
    Set lblHidden = oForm.Controls.Add("Forms.commandbutton.1")
    With lblHidden
        .Caption = "Measure size"
        .Font.Size = oListBox.Font.Size
        .Font.Name = oListBox.Font.Name
        .WordWrap = False
        .AutoSize = True
    End With
    For lCt = LBound(lLengths) To UBound(lLengths)
        lblHidden.Caption = String(lLengths(lCt), "M")
        dTotWidth = dTotWidth + lblHidden.Width
        If Len(sWidths) = 0 Then
            sWidths = CStr(Int(lblHidden.Width))  '+ 12)
        Else
            sWidths = sWidths & ";" & CStr(Int(lblHidden.Width))  '+ 12)
        End If
    Next

    'Now set the widths of the columns
    oListBox.ColumnWidths = sWidths

    If Not ResizeListbox Then GoTo TidyUp

    'Adjust the dimensions of the listbox itself. You may want to adjust the constants I hard coded here.

    'Listbox will always be at least 200 wide
    '    oListBox.Width = Application.Min(Application.Max(200, dTotWidth + 12), oListBox.Width)
    oListBox.Width = Application.Max(TotalColumnsWidth, 200)
    aUserform.Init(oForm).ResizeToFitControls marginRight:=10, marginBottom:=10
    'Listbox will always be at least 48 high.
    '    oListBox.Height = Application.Min(Application.Max((oListBox.ListCount + 1) * 12, 48), oListBox.Height)
TidyUp:
    oForm.Controls.Remove lblHidden.Name
    On Error GoTo 0
End Sub

Function TotalColumnsWidth()
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    If oListBox.ColumnWidths = "" Then
        MsgBox "Column widths not set"
        TotalColumnsWidth = oListBox.Width
    End If
    TotalColumnsWidth = Evaluate(Replace(Join(Split(oListBox.ColumnWidths, ";"), "+"), "pt", "")) + 10
End Function

Sub AddFilter()
    '@AssignedModule aListBox
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aListBox
    '@INCLUDE DECLARATION Emitter
    If oListBox.ColumnWidths = "" Then
        MsgBox "Column widths not set"
        Exit Sub
    End If
    Dim var: var = Split(Replace(oListBox.ColumnWidths, "pt", ""), ";")
    Dim txt         As msforms.Textbox
    Dim i           As Long
    Set txt = oForm.Controls.Add("Forms.Textbox.1")
    txt.Name = "txtFilter"
    txt.Top = oListBox.Top - txt.Height - 2
    txt.Width = 150

    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll oForm

    Dim cbx         As msforms.ComboBox
    Set cbx = oForm.Controls.Add("Forms.ComboBox.1")
    cbx.Name = "cbxFilter"
    cbx.Height = txt.Height
    cbx.Width = 50
    cbx.Left = oListBox.Left
    cbx.Top = txt.Top
    For i = -1 To oListBox.columnCount - 1
        cbx.AddItem i
    Next
    cbx.ListIndex = 0

    txt.Left = oListBox.Left + cbx.Width + 6
End Sub

Public Sub FilterByColumn(this As String, _
        Optional ColumnIndexZeroBased As Long = -1, _
        Optional CaseSensitive As Boolean = False)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim i           As Long
    Dim n           As Long
    Dim sTemp       As String
    If ColumnIndexZeroBased > oListBox.columnCount - 1 Or ColumnIndexZeroBased < 0 Then
        ColumnIndexZeroBased = -1
    End If
    n = oListBox.ListCount
    If ColumnIndexZeroBased <> -1 Then
        For i = n - 1 To 0 Step -1
            If CaseSensitive = True Then
                sTemp = oListBox.list(i, ColumnIndexZeroBased)
            Else
                this = LCase(this)
                sTemp = LCase(oListBox.list(i, ColumnIndexZeroBased))
            End If
            If InStr(1, sTemp, this) = 0 Then
                oListBox.RemoveItem (i)
            End If
        Next i
    Else
        Dim columnCount As Long
        n = oListBox.ListCount
        For i = n - 1 To 0 Step -1
            For columnCount = 0 To oListBox.columnCount - 1
                If CaseSensitive = True Then
                    sTemp = oListBox.list(i, columnCount)
                Else
                    this = LCase(this)
                    sTemp = LCase(oListBox.list(i, columnCount))
                End If
                If InStr(1, sTemp, this) > 0 Then
                Else
                    If columnCount = oListBox.columnCount - 1 Then
                        oListBox.RemoveItem (i)
                    End If
                End If
            Next columnCount
        Next i
    End If
End Sub


Function LoadCSV(TargetFile As String, clear As Boolean)
    '@AssignedModule aListBox
    '@INCLUDE PROCEDURE TxtRead
    '@INCLUDE CLASS aListBox
    If clear Then oListBox.clear
    Dim ArrayOfElements
    Dim S: S = TxtRead(TargetFile)
    Dim arr
    arr = Split(S, vbLf)
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
    oListBox.columnCount = iCols
    oListBox.list = ArrayOfElements
End Function

Sub ToRange(cell As Range)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    cell.Resize(oListBox.ListCount, oListBox.columnCount) = oListBox.list
End Sub

Sub ClearSelection()
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    On Error Resume Next
    Dim i           As Long
    For i = 0 To oListBox.ListCount
        oListBox.Selected(i) = False
    Next i
End Sub


Public Sub AcceptFiles(Optional sExpansion As String = "*.xlsm;*.xlsb;*.xlsx", Optional iDeepSubPath As Integer = 999)
    'eg on userform_activate:
    'aListBox.Init(ListBox1).AcceptFiles sExpansion:="*.xlsm;*.xlsb;*.xlsx",iDeepSubPath:=999
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE CLASS aUserform

    'you can use .[_GethWnd] with controls which support it
    Dim ctr         As control: Set ctr = oListBox
    Dim lGethWnd    As LongPtr: lGethWnd = ctr.[_GethWnd]

#If VBA7 Then
    Dim hWnd As LongPtr, HDROP As LongPtr
#Else
    Dim hWnd As Long, HDROP As Long
#End If

    Const WM_DROPFILES = &H233
    Dim tMsg As msg, sFileName As String * 1000
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
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim sExp        As String
    sExp = sGetExtensionName(sPathFile)
    If sExp = vbNullString Then
        Dim oCol    As Collection
        Set oCol = GetFiles(sPathFile, iDeepSubPath)
        Dim i       As Integer
        Dim iCount  As Integer
        iCount = oCol.Count
        Dim sFile   As String
        For i = 1 To iCount
            sFile = oCol.item(i).path
            sExp = sGetExtensionName(sFile)
            Call filterFiles(sFile, sExpansion, sExp)
        Next i
    Else
        Call filterFiles(sPathFile, sExpansion, sExp)
    End If
End Sub
Private Sub filterFiles(ByVal sPathFile As String, ByVal sExpansion As String, ByVal sExp As String)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    sExp = "*." & sExp
    If sExpansion = vbNullString Or sExpansion = "*.*" Then
    ElseIf Not likeExp(sExpansion, sExp) Then
        Exit Sub
    End If
    oListBox.AddItem sPathFile
End Sub
Private Function likeExp(ByVal sExpansion As String, ByVal sExp As String) As Boolean
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
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
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim FSO         As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    sGetExtensionName = FSO.GetExtensionName(sPathFile)
    Set FSO = Nothing
End Function
Private Function GetFiles(ByVal path As String, ByVal iDeepSubPath As Integer) As Collection
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim MainFolder  As Object
    Dim iFolder     As Object
    Dim iFile       As Object
    Dim FSO         As Object
    Dim MainColl    As New Collection
    Dim iColl       As Collection
    Dim i           As Long

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set MainFolder = FSO.GetFolder(path)
    If MainFolder Is Nothing Then Exit Function

    For Each iFile In MainFolder.Files
        If VBA.InStr(1, iFile.Name, "~") = 0 Then
            MainColl.Add iFile, iFile.path
        End If
    Next

    If iDeepSubPath > 0 Then
        For Each iFolder In MainFolder.SubFolders
            Set iColl = GetFiles(iFolder.path, iDeepSubPath - 1)
            For i = 1 To iColl.Count
                MainColl.Add iColl(i)
            Next
        Next
    End If
    Set GetFiles = MainColl
End Function

Public Sub HeightToEntries(ByVal NumberOfEntries As Long)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    '@INCLUDE DECLARATION GetSystemMetrics
    Const SM_CYEDGE = 46&
    With oListBox
        NumberOfEntries = IIf(NumberOfEntries > .ListCount, .ListCount, NumberOfEntries)
        .Height = ((9.75 * NumberOfEntries) _
                + IIf(oListBox.Width < TotalColumnsWidth, 18, 0) _
                + IIf(.SpecialEffect = fmSpecialEffectFlat, 0, GetSystemMetrics(SM_CYEDGE)))
    End With
End Sub

Public Sub LoadVBProjects()
    '@AssignedModule aListBox
    '@INCLUDE PROCEDURE WorkbookProjectProtected
    '@INCLUDE CLASS aListBox
    oListBox.clear
    oListBox.columnCount = 1
    Dim coll        As New Collection
    Dim wb          As Workbook
    For Each wb In Workbooks
        If Len(wb.path) > 0 Then
            If WorkbookProjectProtected(wb) = False Then
                On Error Resume Next
                coll.Add wb.Name, wb.Name
                On Error GoTo 0
            End If
        End If
    Next
    Rem list addins
    Dim vbProj      As VBProject
    Dim wbPath      As String
    For Each vbProj In Application.VBE.VBProjects
        On Error GoTo ErrorHandler
        wbPath = vbProj.fileName
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

    Dim el          As Variant
    For Each el In coll
        oListBox.AddItem el
    Next

    Exit Sub
ErrorHandler:
    If Err.Number = 76 Then GoTo SKIP
End Sub

Public Sub SelectItems(this As Variant, Optional ByIndex As Boolean)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim i           As Long
    Select Case TypeName(this)
        Case Is = "String", "Long", "Integer"
            For i = 0 To oListBox.ListCount - 1
                If oListBox.list(i) = CStr(this) Then
                    oListBox.Selected(i) = True
                    DoEvents
                    If oListBox.multiSelect = fmMultiSelectSingle Then Exit Sub
                End If
            Next
        Case Else
            Dim el  As Variant
            If ByIndex Then
                For Each el In this
                    oListBox.Selected(el) = True
                Next
            Else
                For Each el In this
                    For i = 0 To oListBox.ListCount - 1
                        If oListBox.list(i) = el Then
                            oListBox.Selected(i) = True
                            DoEvents
                        End If
                    Next
                Next
            End If
    End Select
End Sub

Public Sub AddHeader(Optional arrHeaders)
    '@AssignedModule aListBox
    '@INCLUDE PROCEDURE ArrayDimensions
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aListBox
    '@INCLUDE DECLARATION Emitter
    If oListBox.ColumnWidths = "" Then Exit Sub
    Dim i           As Long

    If IsMissing(arrHeaders) Then
        ReDim arrHeaders(0 To oListBox.columnCount - 1)
        For i = 1 To oListBox.columnCount
            arrHeaders(i - 1) = i
        Next
    End If

    Dim Header      As msforms.Label
    For i = 1 To oListBox.columnCount
        Set Header = oForm.Controls.Add("Forms.Label.1")
        If i = 1 Then
            Header.Left = oListBox.Left
        Else
            Header.Left = oForm.Controls("lblHeader-" & i - 1).Left + oForm.Controls("lblHeader-" & i - 1).Width
        End If
        Header.Name = "lblHeader-" & i
        Header.SpecialEffect = fmSpecialEffectFlat
        Header.BackColor = RGB(200, 200, 200)
        Header.Width = Trim(Replace(Split(oListBox.ColumnWidths, ";")(i - 1), "pt", ""))
        Header.Height = 12
        Header.Font.Bold = True
        Header.Font.Name = "Segoe UI"
        Header.Top = oListBox.Top - Header.Height - 1
        Header.TextAlign = fmTextAlignCenter
        If ArrayDimensions(arrHeaders) = 1 Then
            Header.Caption = arrHeaders(i - 1)
        Else
            Header.Caption = arrHeaders(0, i)
        End If
    Next

    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll oForm

    '    oListBox.ZOrder (1)
    '    Header.ZOrder (0)

End Sub
Sub removeHeaders()
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim ctl         As msforms.control
    For Each ctl In oForm.Controls
        If ctl.Name Like "lblHeader-*" Then oForm.Controls.Remove ctl.Name
    Next
End Sub
Public Function Contains(this As String, _
        Optional ColumnIndexZeroBased As Long = -1, _
        Optional CaseSensitive As Boolean = False) As Boolean
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim i           As Long
    Dim n           As Long
    Dim sTemp       As String
    If ColumnIndexZeroBased > oListBox.columnCount - 1 Or ColumnIndexZeroBased < 0 Then
        ColumnIndexZeroBased = -1
    End If
    n = oListBox.ListCount
    If ColumnIndexZeroBased <> -1 Then
        For i = n - 1 To 0 Step -1
            If CaseSensitive = True Then
                sTemp = oListBox.list(i, ColumnIndexZeroBased)
            Else
                this = LCase(this)
                sTemp = LCase(oListBox.list(i, ColumnIndexZeroBased))
            End If
            If InStr(1, sTemp, this) > 0 Then
                Contains = True
                Exit Function
            End If
        Next i
    Else
        Dim columnCount As Long
        n = oListBox.ListCount
        For i = n - 1 To 0 Step -1
            For columnCount = 0 To oListBox.columnCount - 1
                If CaseSensitive = True Then
                    sTemp = oListBox.list(i, columnCount)
                Else
                    this = LCase(this)
                    sTemp = LCase(oListBox.list(i, columnCount))
                End If
                If InStr(1, sTemp, this) > 0 Then
                    Contains = True
                    Exit Function
                End If
            Next columnCount
        Next i
    End If
End Function

Public Sub SortOnColumn(OnColumn As Long)
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox

    Dim vntData     As Variant
    Dim vntTempItem As Variant
    Dim lngOuterIndex As Long
    Dim lngInnerIndex As Long
    Dim lngSubItemIndex As Long
    vntData = oListBox.list

    For lngOuterIndex = LBound(vntData, 1) To UBound(vntData, 1) - 1
        For lngInnerIndex = lngOuterIndex + 1 To UBound(vntData, 1)
            ' Use StrComp with vbTextCompare option to compare strings case-insensitively
            If StrComp(vntData(lngOuterIndex, OnColumn), vntData(lngInnerIndex, OnColumn), vbTextCompare) > 0 Then
                For lngSubItemIndex = 0 To oListBox.columnCount - 1
                    vntTempItem = vntData(lngOuterIndex, lngSubItemIndex)
                    vntData(lngOuterIndex, lngSubItemIndex) = vntData(lngInnerIndex, lngSubItemIndex)
                    vntData(lngInnerIndex, lngSubItemIndex) = vntTempItem
                Next lngSubItemIndex
            End If
        Next lngInnerIndex
    Next lngOuterIndex

    oListBox.clear
    oListBox.list = vntData
End Sub

'Public Sub SortOnColumn(OnColumn As Long)
'
'    Dim vntData As Variant
'    Dim vntTempItem As Variant
'    Dim lngOuterIndex As Long
'    Dim lngInnerIndex As Long
'    Dim lngSubItemIndex As Long
'    vntData = oListBox.List
'    For lngOuterIndex = LBound(vntData, 1) To UBound(vntData, 1) - 1
'        For lngInnerIndex = lngOuterIndex + 1 To UBound(vntData, 1)
'            If vntData(lngOuterIndex, OnColumn) > vntData(lngInnerIndex, OnColumn) Then
'                For lngSubItemIndex = 0 To oListBox.columnCount - 1
'                    vntTempItem = vntData(lngOuterIndex, lngSubItemIndex)
'                    vntData(lngOuterIndex, lngSubItemIndex) = vntData(lngInnerIndex, lngSubItemIndex)
'                    vntData(lngInnerIndex, lngSubItemIndex) = vntTempItem
'                Next
'            End If
'        Next lngInnerIndex
'    Next lngOuterIndex
'    oListBox.Clear
'    oListBox.List = vntData
'End Sub

Public Function selectedIndexes() As Collection
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim i           As Long
    Dim coll        As New Collection
    If oListBox.ListCount > 0 Then
        For i = 0 To oListBox.ListCount - 1
            If oListBox.Selected(i) Then coll.Add i
        Next i
    End If
    Set selectedIndexes = coll
End Function

Public Function SelectedValues() As Collection    'single column
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim i           As Long
    Dim ListItem    As Long
    Dim selectedCollection As Collection
    Set selectedCollection = New Collection
    Dim listboxCount As Long

    If oListBox.ListCount > 0 Then
        For i = 0 To oListBox.ListCount - 1
            If oListBox.Selected(i) Then
                selectedCollection.Add oListBox.list(i, oListBox.BoundColumn - 1)
            End If
        Next i
    End If

    Set SelectedValues = selectedCollection
End Function


Public Function SelectedRowsArray()
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim x As Long, y As Long
    Dim ListItem    As Long

    Dim indexColl   As New Collection
    Set indexColl = selectedIndexes
    If indexColl.Count = 0 Then SelectedRowsArray = Array(): Exit Function
    Dim out()       As Variant
    ReDim out(1 To indexColl.Count, 1 To oListBox.columnCount)
    For x = 1 To indexColl.Count
        For y = 1 To oListBox.columnCount
            out(x, y) = oListBox.list(indexColl(x), y - 1)
        Next
    Next
    SelectedRowsArray = out
End Function

Public Function SelectedRowsText() As String
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim x As Long, y As Long
    Dim ListItem    As Long
    Dim out         As String
    Dim indexColl   As New Collection
    Set indexColl = selectedIndexes
    For x = 1 To indexColl.Count
        For y = 0 To oListBox.columnCount - 1
            out = out & IIf(y > 0, "|", "") & oListBox.list(indexColl(x), y)
        Next
        If x < indexColl.Count Then out = out & vbNewLine
    Next
    SelectedRowsText = out
End Function


Public Function SelectedCount() As Long
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim i           As Long
    Dim ListItem    As Long
    Dim selectedCollection As Collection
    Set selectedCollection = New Collection
    Dim listboxCount As Long
    Dim counter     As Long

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
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    If oListBox.ListCount <> 0 Then
        Dim i       As Long
        For i = 0 To oListBox.ListCount - 1
            oListBox.Selected(i) = False
        Next i
    End If
End Sub

Public Sub SelectAll()
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    If oListBox.ListCount <> 0 Then
        Dim i       As Long
        For i = 0 To oListBox.ListCount - 1
            oListBox.Selected(i) = True
        Next i
    End If
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 24-10-2023 20:05    Alex                (aListBox.cls > SelectLike)

Public Sub SelectLike(this As String, casteSensitive As Boolean, Optional targetColumn = 1)
'@LastModified 2310242005
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    DeselectAll
    If targetColumn < 1 Or targetColumn > oListBox.columnCount Then Stop
    If this = "" Then Exit Sub
    If oListBox.ListCount = 0 Then Exit Sub
    Dim val1
    Dim i           As Long
    For i = 0 To oListBox.ListCount - 1
        val1 = oListBox.list(i, targetColumn - 1)
        If casteSensitive Then val1 = UCase(val1): this = UCase(this)
        If val1 Like this Then
            oListBox.Selected(i) = True
        End If
    Next i
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 24-10-2023 20:05    Alex                (aListBox.cls > DeselectLike)

Public Sub DeselectLike(this As String, casteSensitive As Boolean, Optional targetColumn = 1)
'@LastModified 2310242005
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    DeselectAll
    If targetColumn < 1 Or targetColumn > oListBox.columnCount Then Stop
    If this = "" Then Exit Sub
    If oListBox.ListCount = 0 Then Exit Sub
    Dim val1
    Dim i           As Long
    For i = 0 To oListBox.ListCount - 1
        val1 = oListBox.list(i, targetColumn - 1)
        If casteSensitive Then val1 = UCase(val1): this = UCase(this)
        If UCase(oListBox.list(i, targetColumn)) Like UCase(this) Then
            oListBox.Selected(i) = False
        End If
    Next i
End Sub
Public Sub SortAZ()    '1D
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim j           As Long
    Dim i           As Long
    Dim Temp        As Variant
    With oListBox
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

Public Sub SortZA()    '1D
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim j           As Long
    Dim i           As Long
    Dim Temp        As Variant
    With oListBox
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
'    cell.RESIZE(lBox.ListCount, oListBox.ColumnCount) = CollectionsToArray2D(SelectedValues)
'End Sub

Public Sub SelectedToRange()
    '@AssignedModule aListBox
    '@INCLUDE PROCEDURE CollectionsToArray2D
    '@INCLUDE PROCEDURE GetInputRange
    '@INCLUDE CLASS aCollection
    '@INCLUDE CLASS aListBox
    Dim rng         As Range
    If GetInputRange(rng, "Range picker", "Select range to output listbox' list") = False Then Exit Sub
    Dim var: var = aCollection.CollectionsToArray2D(SelectedValues)
    rng.Resize(UBound(var, 1), oListBox.columnCount) = var
End Sub

Public Sub RemoveSelected()
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim coll        As New Collection: Set coll = selectedIndexes
    If coll.Count = 0 Then Exit Sub
    Dim i           As Long
    For i = coll.Count To 1 Step -1
        oListBox.RemoveItem coll(i)
    Next
End Sub

Public Sub ShowTheseColumns(targetColumns As Variant)
    'targetColumns is array of integers, base 1
    '@AssignedModule aListBox
    '@INCLUDE CLASS aListBox
    Dim RetVal()    As Variant
    ReDim RetVal(1 To oListBox.columnCount)
    If oListBox.ColumnWidths = "" Then Exit Sub    ' AutofitColumns (False)
    Dim i As Long, ii As Long
    For i = 1 To oListBox.columnCount
        For ii = LBound(targetColumns) To UBound(targetColumns)
            If i = targetColumns(ii) Then
                RetVal(i) = Split(Replace(oListBox.ColumnWidths, "pt", ""), ";")(i - 1)
                GoTo resumeNext
            End If
        Next
        RetVal(i) = 0
resumeNext:
    Next
    oListBox.ColumnWidths = Join(RetVal, ";")
End Sub
