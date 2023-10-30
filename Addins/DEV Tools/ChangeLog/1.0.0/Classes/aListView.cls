VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aListView
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 29-06-2023 13:45    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Public WithEvents EventDragSort As ListView
Attribute EventDragSort.VB_VarHelpID = -1
Public WithEvents EventDropFiles As ListView
Attribute EventDropFiles.VB_VarHelpID = -1
Public WithEvents aListViewEvent As ListView
Attribute aListViewEvent.VB_VarHelpID = -1

#If VBA7 Then
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

'Windows API Constants
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Private Const LVM_FIRST = &H1000
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)

Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2

Private Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Private Const LVNI_SELECTED = &H2
Private Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)

Public oListView    As ListView

Private option_DD_LogFiles As Boolean
Private option_DD_LogFolders As Boolean
Private option_DD_IncludeSubfolders As Boolean
Private option_DD_CommaSeparatedFilters As String


Public LstItmObj    As ListItem
Public swapNeeded   As Boolean    'swap mode

Private Type POINTAPI
    x               As Long
    y               As Long
End Type
Private Type LVHITTESTINFO
    pt              As POINTAPI
    Flags           As Long
    iItem           As Long
    iSubItem        As Long
End Type

Private Const LVM_HITTEST As Long = &H418
Private Const LVM_SUBITEMHITTEST As Long = &H1039
Private Const LVHT_ONITEM As Long = &H2

Public Function value()
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    Dim headCount   As Long: headCount = oListView.ColumnHeaders.Count
    Dim rowCount    As Long: rowCount = oListView.ListItems.Count

    Dim arr()
    ReDim arr(1 To rowCount, 1 To headCount)
    Dim x           As Long
    Dim y           As Long
    For y = 1 To oListView.ColumnHeaders.Count
        arr(1, y) = oListView.ColumnHeaders(y)
    Next
    If oListView.ListItems.Count > 0 Then
        For x = 1 To oListView.ListItems.Count
            For y = 1 To headCount
                If y = 1 Then
                    arr(x, y) = oListView.ListItems(x)
                Else
                    arr(x, y) = oListView.ListItems(x).ListSubItems(y - 1)
                End If
            Next
        Next
        value = arr
    Else
        value = Array()
    End If
End Function

Public Function RowArray(targetRow As Long)
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    Dim headCount   As Long: headCount = oListView.ColumnHeaders.Count
    Dim arr
    ReDim arr(1 To headCount)
    Dim i           As Long
    For i = 1 To headCount
        If i = 1 Then
            arr(i) = oListView.ListItems(targetRow)
        Else
            arr(i) = oListView.ListItems(targetRow).ListSubItems(i - 1)
        End If
    Next
    RowArray = arr
End Function

Public Function SelectionArray()
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    Dim item        As MSComctlLib.ListItem
    Dim targetRow   As Long
    Dim arr()
    Dim i As Long, counter As Long
    Dim headCount   As Long: headCount = oListView.ColumnHeaders.Count
    If oListView.multiSelect = False Then
        targetRow = oListView.SelectedItem.index
        arr = RowArray(targetRow)
    Else
        ReDim arr(1 To 1, 1 To headCount)
        For Each item In oListView.ListItems
            If item.Selected Then
                targetRow = item.index
                counter = counter + 1
                ReDim Preserve arr(1 To counter, 1 To headCount)
                For i = 1 To headCount
                    If i = 1 Then
                        arr(counter, i) = oListView.ListItems(targetRow)
                    Else
                        arr(counter, i) = oListView.ListItems(targetRow).ListSubItems(i - 1)
                    End If
                Next

            End If
        Next
    End If
    SelectionArray = arr
End Function

Public Function ClickedColumn(x, y)
    'call from userform ListView1_MouseDown event and such
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    '@INCLUDE DECLARATION SendMessage
    ClickedColumn = -1

    Dim hitTestInfo As LVHITTESTINFO
    hitTestInfo.pt.x = x
    hitTestInfo.pt.y = y
    SendMessage oListView.hWnd, LVM_SUBITEMHITTEST, 0, hitTestInfo

    If hitTestInfo.Flags = 4 Then
        Dim ColumnIndex As Long
        ColumnIndex = hitTestInfo.iSubItem
        ClickedColumn = ColumnIndex
    End If

End Function

Sub DeselectAll()
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    '@INCLUDE DECLARATION SendMessage
    Dim sThisItem As Long, lLvHwnd As Long, lSelectedItems As Long, lItemIndex As Long

    On Error GoTo ErrFailed

    With oListView
        lLvHwnd = .hWnd
        lSelectedItems = SendMessage(lLvHwnd, LVM_GETSELECTEDCOUNT, 0, ByVal 0&)
        lItemIndex = -1
        For sThisItem = 1 To lSelectedItems
            lItemIndex = SendMessage(lLvHwnd, LVM_GETNEXTITEM, lItemIndex, ByVal LVNI_SELECTED)
            .ListItems(lItemIndex + 1).Selected = False
        Next
    End With
    Exit Sub

ErrFailed:
    Debug.Print Err.Description
    Debug.Assert False
End Sub

Public Sub EnableDragSort()
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    Set EventDragSort = oListView
    swapNeeded = False
End Sub

Private Sub EventDragSort_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    Set LstItmObj = EventDragSort.SelectedItem
End Sub

'when drop occurs we make mouseclick to select next item and then set swap mode on
Private Sub EventDragSort_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'that click will occur only after end of this Sub, that's why we can't make rows swaping here
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    Call SingleClick
    swapNeeded = True
End Sub

'this Sub starts after OLEDragDrop ends so new row is already selected and old row is already saved to LstItmObj so here we just need to swap those two rows
Private Sub eventdragsort_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    '@INCLUDE DECLARATION Sleep
    If LstItmObj Is Nothing Then Exit Sub
    If (swapNeeded) Then
        Sleep 30
        Dim insertedList As ListItem
        Dim SelectedIndex As Integer
        Dim newListSubItemObj As ListSubItem

        SelectedIndex = EventDragSort.SelectedItem.index
        EventDragSort.ListItems.Remove LstItmObj.index

        Set insertedList = EventDragSort.ListItems.Add(SelectedIndex, LstItmObj.Key, LstItmObj.TEXT, LstItmObj.Icon, LstItmObj.SmallIcon)
        For Each newListSubItemObj In LstItmObj.ListSubItems
            insertedList.ListSubItems.Add newListSubItemObj.index, newListSubItemObj.Key, newListSubItemObj.TEXT, newListSubItemObj.ReportIcon, newListSubItemObj.TooltipText
        Next newListSubItemObj    'swap mode off again
        swapNeeded = False
        Set EventDragSort.SelectedItem = EventDragSort.ListItems.item(SelectedIndex)
    End If
End Sub

Private Sub SingleClick()
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    '@INCLUDE DECLARATION mouse_event
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Public Sub EnableDropFilesFolders(LogFiles As Boolean, LogFolders As Boolean, IncludeSubFolders As Boolean, Optional CommaSeparatedFilters As String = "*")
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    option_DD_LogFiles = LogFiles
    option_DD_LogFolders = LogFolders
    option_DD_IncludeSubfolders = IncludeSubFolders
    option_DD_CommaSeparatedFilters = CommaSeparatedFilters
    oListView.OLEDragMode = ccOLEDragAutomatic    ' ccOLEDragManual
    oListView.OLEDropMode = ccOLEDropManual
    Set EventDropFiles = oListView
End Sub

Private Sub EventDropFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'The Format numbers used in the OLE DragDrop data structure, are:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Text = 1 (vbCFText)
    'Bitmap = 2 (vbCFBitmap)
    'Metafile = 3
    'Emetafile = 14
    'DIB = 8
    'Palette = 9
    'Files = 15 (vbCFFiles)
    'RTF = -16639
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView

    Select Case True
        Case Data.GetFormat(15)    'Format 15 is an array of names from WinExplorer
            If Not option_DD_LogFiles And Not option_DD_LogFiles Then Exit Sub
            Dim FileFullPath As String
            Dim fileItem As Long
            For fileItem = 1 To Data.Files.Count
                FileFullPath = Data.Files(fileItem)
                DD_AddToListview FileFullPath
            Next fileItem
        Case Else
    End Select
End Sub

Private Sub DD_AddToListview(FileFullPath As String)
    '@AssignedModule aListView
    '@INCLUDE PROCEDURE IsFileFolderURL
    '@INCLUDE CLASS aListView
    Dim objFSO      As Scripting.FileSystemObject
    Dim objTopFolder As Scripting.Folder
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim fileName    As String
    Dim var As Variant, element As Variant
    var = Split(option_DD_CommaSeparatedFilters, ",")
    If option_DD_LogFiles = True Then
        If LCase(IsFileFolderURL(FileFullPath)) = "f" Then
            fileName = Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1)
            On Error Resume Next
            If Left(fileName, 1) <> "~" And (var(0) = "*" Or var(0) = "") Then GoTo PASS
            For Each element In var
                If InStr(1, FileFullPath, element, vbTextCompare) > 0 And Left(fileName, 1) <> "~" Then
PASS:

                    'where:
                    'Constant    Value   Description
                    'lvwText     0       (Default) Matches the string with a ListItem object's Text property.
                    'lvwSubitem  1       Matches the string with any string in a ListItem object's SubItems property.
                    'lvwTag      2       Matches the string with any ListItem object's Tag property.

                    'match:
                    'The settings for match are:
                    'Constant        Value   Description
                    'lvwWholeWord    0       (Default) An integer or constant specifying that a match will occur if the item's Text property begins with the whole word being searched. Ignored if the criteria is not text.
                    'lvwPartial      1       An integer or constant specifying that a match will occur if the item's Text property begins with the string being searched. Ignored if the criteria is not text.

                    '                    For full match:
                    '                        Dim itm As MSComctlLib.listItem
                    '                        Set itm = oListView.FindItem(sz:="subitemtext", where:=lvwSubItem, index:=2, fPartial:=lvwPartial)
                    '                    for partial search use this:
                    'Set itm = oListView.FindItem(sz:="partial", where:=lvwSubItem, index:=2, fPartial:lvwPartial)

                    If Not oListView.FindItem(sz:=FileFullPath, where:=0, fPartial:=lvwWhole) Then
                        oListView.ListItems.Add , , FileFullPath
                    End If
                End If
            Next
        Else        'if drag dropped folder
            Set objTopFolder = objFSO.GetFolder(FileFullPath)
            DD_AddToListview objTopFolder.path
        End If
    End If

    If option_DD_LogFolders = True Then
        If UCase(IsFileFolderURL(FileFullPath)) = "D" Then
            Set objTopFolder = objFSO.GetFolder(FileFullPath)
            If Not oListView.FindItem(sz:=objTopFolder.path & "\", where:=0, fPartial:=lvwWhole) Then
                oListView.ListItems.Add , , objTopFolder.path
            End If
            DD_AddToListview objTopFolder.path
        End If
    End If
    Set objFSO = Nothing
    Set objTopFolder = Nothing
End Sub

Sub AutofitColumns()
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    Dim counter     As Long
    For counter = 1 To oListView.ColumnHeaders.Count
        Call SendMessageLong(oListView.hWnd, LVM_SETCOLUMNWIDTH, counter - 1, LVSCW_AUTOSIZE_USEHEADER)
    Next
End Sub

Public Function Init(LV As ListView) As aListView
    'REFERENCE: Microsoft Windows Common Controls 6.0 (SP6)
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView

    Set oListView = LV
    Set Init = Me

End Function
'Public Function Parent()
'
'End Function
Public Sub clear()
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    oListView.ListItems.clear
    oListView.ColumnHeaders.clear
End Sub

Public Sub InitializeFromArray(inputArray As Variant)
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    clear
    Dim vListItem   As ListItem
    Dim vChildItem  As ListSubItem
    Dim vHeader     As Variant
    Dim iRows As Long, iColumns As Long
    For iColumns = LBound(inputArray, 2) To UBound(inputArray, 2)
        Set vHeader = oListView.ColumnHeaders.Add(, , inputArray(LBound(inputArray, 1), iColumns))
    Next
    For iRows = LBound(inputArray, 1) + 1 To UBound(inputArray, 1)
        Set vListItem = oListView.ListItems.Add(, , inputArray(iRows, 1))
        For iColumns = LBound(inputArray, 2) + 1 To UBound(inputArray, 2)
            Set vChildItem = vListItem.ListSubItems.Add(, , inputArray(iRows, iColumns))
        Next
    Next
    oListView.View = lvwReport
End Sub

Public Sub AppendArray(inputArray As Variant)
    '@AssignedModule aListView
    '@INCLUDE PROCEDURE ArrayDimensions
    '@INCLUDE CLASS aListView
    If ArrayDimensions(inputArray) <> 2 Then Exit Sub
    If UBound(inputArray, 2) + IIf(LBound(inputArray, 2) = 0, 1, 0) <> oListView.ColumnHeaders.Count Then Exit Sub
    Dim vListItem   As ListItem
    Dim vChildItem  As ListSubItem
    Dim iRows As Long, iColumns As Long
    For iRows = LBound(inputArray, 1) To UBound(inputArray, 1)
        Set vListItem = oListView.ListItems.Add(, , inputArray(iRows, 1))
        For iColumns = LBound(inputArray, 2) + 1 To UBound(inputArray, 2)
            Set vChildItem = vListItem.ListSubItems.Add(, , inputArray(iRows, iColumns))
        Next
    Next
    '    oListView.View = lvwReport
End Sub

Public Sub EventListener()
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    Set aListViewEvent = oListView
End Sub

'Private Sub aListViewEvent_Click()
'    MsgBox "Clicked the listview"
'End Sub

Private Sub aListViewEvent_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
        ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    '@INCLUDE DECLARATION LOGPIXELSX
    '@INCLUDE DECLARATION LOGPIXELSY
    '@INCLUDE DECLARATION GetDC
    '@INCLUDE DECLARATION GetDeviceCaps
    '@INCLUDE DECLARATION ReleaseDC

    'https://markdagosta.wordpress.com/2010/08/08/listview-hittest-right-click/
    Dim item        As MSComctlLib.ListItem
    Dim lngXPixelsPerInch As Long, lngYPixelsPerInch As Long
    Dim lngDeviceHandle As Long
    'Only capture the standard right-clicks; otherwise get out.
    If (Button <> xlSecondaryButton) Or (Shift <> 0) Then Exit Sub
    'We must determine the Pixels per Inch for the display device.
    lngDeviceHandle = GetDC(0)
    lngXPixelsPerInch = GetDeviceCaps(lngDeviceHandle, LOGPIXELSX)
    lngYPixelsPerInch = GetDeviceCaps(lngDeviceHandle, LOGPIXELSY)
    ReleaseDC 0, lngDeviceHandle
    'Convert the event's x and y arguments from Pixels to Twips
    Set item = oListView.hitTest(x * 1440 / lngXPixelsPerInch, y * 1440 / lngYPixelsPerInch)
    MsgBox "List ID #" & item.index & ": " & item.TEXT & " has been right-clicked!"


End Sub

Public Sub RowsFormatOddEven()
    '@AssignedModule aListView
    '@INCLUDE CLASS aListView
    Dim i As Long, y As Long
    For i = 1 To oListView.ListItems.Count
        If i Mod 2 Then
            oListView.ListItems(i).ForeColor = vbBlue
            For y = 1 To oListView.ColumnHeaders.Count - 1
                oListView.ListItems(i).ListSubItems(y).ForeColor = vbBlue
            Next
        Else
            oListView.ListItems(i).ForeColor = vbRed
            For y = 1 To oListView.ColumnHeaders.Count - 1
                oListView.ListItems(i).ListSubItems(y).ForeColor = vbRed
            Next
        End If
    Next
    oListView.Refresh
End Sub


