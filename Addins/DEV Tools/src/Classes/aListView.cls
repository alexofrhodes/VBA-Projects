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
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare PtrSafe Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    
    Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
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

Public oListView As ListView

Public option_DD_LogFiles As Boolean
Public option_DD_LogFolders As Boolean
Public option_DD_IncludeSubfolders As Boolean
Public option_DD_CommaSeparatedFilters As String


Public LstItmObj As listItem
Public swapNeeded As Boolean 'swap mode


Sub DeselectAll()
    Dim sThisItem As Long, lLvHwnd As Long, lSelectedItems As Long, lItemIndex As Long

    On Error GoTo ErrFailed

    With oListView
        lLvHwnd = .hwnd
        .Visible = False             'For speed. Need to remove the line in VBA
        lSelectedItems = SendMessage(lLvHwnd, LVM_GETSELECTEDCOUNT, 0, ByVal 0&)
        lItemIndex = -1
        For sThisItem = 1 To lSelectedItems
            lItemIndex = SendMessage(lLvHwnd, LVM_GETNEXTITEM, lItemIndex, ByVal LVNI_SELECTED)
            .ListItems(lItemIndex + 1).Selected = False
        Next
        .Visible = True              'For speed. Need to remove the line in VBA
    End With
    Exit Sub

ErrFailed:
    Debug.Print Err.Description
    Debug.Assert False
End Sub

Public Sub EnableDragSort()
    Set EventDragSort = oListView
    swapNeeded = False
End Sub

Private Sub EventDragSort_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    Set LstItmObj = EventDragSort.SelectedItem
End Sub

'when drop occurs we make mouseclick to select next item and then set swap mode on
Private Sub EventDragSort_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'that click will occur only after end of this Sub, that's why we can't make rows swaping here
    Call SingleClick
    swapNeeded = True
End Sub

'this Sub starts after OLEDragDrop ends so new row is already selected and old row is already saved to LstItmObj so here we just need to swap those two rows
Private Sub eventdragsort_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    If LstItmObj Is Nothing Then Exit Sub
    If (swapNeeded) Then
        Sleep 30
        Dim insertedList As listItem
        Dim selectedIndex As Integer
        Dim newListSubItemObj As ListSubItem

        selectedIndex = EventDragSort.SelectedItem.index
        EventDragSort.ListItems.Remove LstItmObj.index

        Set insertedList = EventDragSort.ListItems.Add(selectedIndex, LstItmObj.Key, LstItmObj.TEXT, LstItmObj.Icon, LstItmObj.SmallIcon)
        For Each newListSubItemObj In LstItmObj.ListSubItems
                insertedList.ListSubItems.Add newListSubItemObj.index, newListSubItemObj.Key, newListSubItemObj.TEXT, newListSubItemObj.ReportIcon, newListSubItemObj.TooltipText
        Next newListSubItemObj 'swap mode off again
        swapNeeded = False
        Set EventDragSort.SelectedItem = EventDragSort.ListItems.item(selectedIndex)
    End If
End Sub

Private Sub SingleClick()
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Public Sub EnableDropFilesFolders(LogFiles As Boolean, LogFolders As Boolean, IncludeSubFolders As Boolean, Optional CommaSeparatedFilters As String = "*")
    option_DD_LogFiles = LogFiles
    option_DD_LogFolders = LogFolders
    option_DD_IncludeSubfolders = IncludeSubFolders
    option_DD_CommaSeparatedFilters = CommaSeparatedFilters
    oListView.OLEDragMode = ccOLEDragAutomatic ' ccOLEDragManual
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
    
    Select Case True
    Case Data.GetFormat(15) 'Format 15 is an array of names from WinExplorer
        If Not option_DD_LogFiles And Not option_DD_LogFiles Then Exit Sub
        Dim FileFullPath As String
        Dim fileItem As Long
        For fileItem = 1 To Data.Files.count
            FileFullPath = Data.Files(fileItem)
            DD_AddToListview FileFullPath
        Next fileItem
    Case Else
    End Select
End Sub

Private Sub DD_AddToListview(FileFullPath As String)
    Dim objFSO As scripting.FileSystemObject
    Dim objTopFolder As scripting.Folder
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim FileName As String
    Dim var As Variant, element As Variant
        var = Split(option_DD_CommaSeparatedFilters, ",")
    If option_DD_LogFiles = True Then
        If LCase(IsFileFolderURL(FileFullPath)) = "f" Then
            FileName = Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1)
            On Error Resume Next
            If Left(FileName, 1) <> "~" And (var(0) = "*" Or var(0) = "") Then GoTo PASS
            For Each element In var
                If InStr(1, FileFullPath, element, vbTextCompare) > 0 And Left(FileName, 1) <> "~" Then
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
            Set objTopFolder = objFSO.getFolder(FileFullPath)
            DD_AddToListview objTopFolder.Path
        End If
    End If

    If option_DD_LogFolders = True Then
        If UCase(IsFileFolderURL(FileFullPath)) = "D" Then
            Set objTopFolder = objFSO.getFolder(FileFullPath)
            If Not oListView.FindItem(sz:=objTopFolder.Path & "\", where:=0, fPartial:=lvwWhole) Then
                oListView.ListItems.Add , , objTopFolder.Path
            End If
            DD_AddToListview objTopFolder.Path
        End If
    End If
    Set objFSO = Nothing
    Set objTopFolder = Nothing
End Sub

Sub AutofitColumns()
    Dim counter As Long
    For counter = 1 To oListView.ColumnHeaders.count
        Call SendMessageLong(oListView.hwnd, LVM_SETCOLUMNWIDTH, counter - 1, LVSCW_AUTOSIZE_USEHEADER)
    Next
End Sub

Public Function Init(LV As ListView) As aListView
'REFERENCE: Microsoft Windows Common Controls 6.0 (SP6)

    Set oListView = LV
    Set Init = Me
End Function
Public Function Parent()
    Set Parent = oListView.Parent
End Function
Public Sub Clear()
    oListView.ListItems.Clear
    oListView.ColumnHeaders.Clear
End Sub

Public Sub InitializeFromArray(InputArray As Variant)
    Clear
    Dim vListItem As listItem
    Dim vChildItem As ListSubItem
    Dim vHeader As Variant
    Dim iRows As Long, iColumns As Long
    For iColumns = LBound(InputArray, 2) To UBound(InputArray, 2)
        Set vHeader = oListView.ColumnHeaders.Add(, , InputArray(LBound(InputArray, 1), iColumns))
    Next
    For iRows = LBound(InputArray, 1) + 1 To UBound(InputArray, 1)
        Set vListItem = oListView.ListItems.Add(, , InputArray(iRows, 1))
        For iColumns = LBound(InputArray, 2) + 1 To UBound(InputArray, 2)
            Set vChildItem = vListItem.ListSubItems.Add(, , InputArray(iRows, iColumns))
        Next
    Next
    oListView.View = lvwReport
End Sub

Public Sub AddFromArray(InputArray As Variant)
    If UBound(InputArray, 2) + 1 <> oListView.ColumnHeaders.count Then Exit Sub
    Dim vListItem As listItem
    Dim vChildItem As ListSubItem
    Dim iRows As Long, iColumns As Long
    For iRows = LBound(InputArray, 1) + 1 To UBound(InputArray, 1)
        Set vListItem = oListView.ListItems.Add(, , InputArray(iRows, 1))
        For iColumns = LBound(InputArray, 2) + 1 To UBound(InputArray, 2)
            Set vChildItem = vListItem.ListSubItems.Add(, , InputArray(iRows, iColumns))
        Next
    Next
    oListView.View = lvwReport
End Sub

Public Sub EventListener()
    'add argument by multiple enum to enable desired events?
    Set aListViewEvent = oListView
End Sub

'Private Sub aListViewEvent_Click()
'    MsgBox "Clicked the listview"
'End Sub

Private Sub aListViewEvent_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                    ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)

'https://markdagosta.wordpress.com/2010/08/08/listview-hittest-right-click/

   Dim item As MSComctlLib.listItem
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
  Set item = oListView.HitTest(x * 1440 / lngXPixelsPerInch, y * 1440 / lngYPixelsPerInch)

   Toast "List ID #" & item.index & ": " & item.TEXT & " has been right-clicked!"

End Sub

Sub PopulateListViewFromArray(TargetListView As ListView, InputArray As Variant)
    Dim vListItem As listItem
    Dim vChildItem As ListSubItem
    Dim vHeader As Variant
    Dim iRows As Long, iColumns As Long
    For iColumns = LBound(InputArray, 2) To UBound(InputArray, 2)
        Set vHeader = TargetListView.ColumnHeaders.Add(, , InputArray(LBound(InputArray, 1), iColumns))
    Next
    For iRows = LBound(InputArray, 1) + 1 To UBound(InputArray, 1)
        Set vListItem = TargetListView.ListItems.Add(, , InputArray(iRows, 1))
        For iColumns = LBound(InputArray, 2) + 1 To UBound(InputArray, 2)
            Set vChildItem = vListItem.ListSubItems.Add(, , InputArray(iRows, iColumns))
        Next
    Next
    TargetListView.View = lvwReport
End Sub
