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

Public WithEvents aListViewEvent As ListView
Attribute aListViewEvent.VB_VarHelpID = -1

#If VBA7 Then
Private Declare PtrSafe Function SendMessageLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare PtrSafe Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "USER32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
#Else
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
#End If


'Windows API Constants
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90


Private Const LVM_FIRST = &H1000
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
 
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2
 
Public oListView As ListView

Sub AutoFitColumns()
    Dim counter As Long
    For counter = 1 To oListView.ColumnHeaders.count
        Call SendMessageLong(oListView.hWnd, LVM_SETCOLUMNWIDTH, counter - 1, LVSCW_AUTOSIZE_USEHEADER)
    Next
End Sub

Public Function Init(LV As ListView)
    Set oListView = LV
    Set Init = Me
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
                                    ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)

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
  Set item = oListView.HitTest(X * 1440 / lngXPixelsPerInch, Y * 1440 / lngYPixelsPerInch)

   Toast "List ID #" & item.Index & ": " & item.TEXT & " has been right-clicked!"

End Sub


'Private Sub ListViewControl_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    Dim oLogFiles As Boolean: oLogFiles = True
'
'    Dim dbBefore As Long: dbBefore = oListView.ListItems.Count
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
'
''where:
''Constant    Value   Description
''lvwText     0       (Default) Matches the string with a ListItem object's Text property.
''lvwSubitem  1       Matches the string with any string in a ListItem object's SubItems property.
''lvwTag      2       Matches the string with any ListItem object's Tag property.
'
''match:
''The settings for match are:
''Constant        Value   Description
''lvwWholeWord    0       (Default) An integer or constant specifying that a match will occur if the item's Text property begins with the whole word being searched. Ignored if the criteria is not text.
''lvwPartial      1       An integer or constant specifying that a match will occur if the item's Text property begins with the string being searched. Ignored if the criteria is not text.
'
''                    For full match:
''                        Dim itm As MSComctlLib.listItem
''                        Set itm = oListView.FindItem(sz:="subitemtext", where:=lvwSubItem, index:=2, fPartial:=lvwPartial)
''                    for partial search use this:
'                        'Set itm = oListView.FindItem(sz:="partial", where:=lvwSubItem, index:=2, fPartial:lvwPartial)
'                        If Not oListView.FindItem(sz:=FileFullPath, where:=2, fPartial:=lvwPartial) Then
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

