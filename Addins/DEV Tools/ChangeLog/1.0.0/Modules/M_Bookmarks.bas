Attribute VB_Name = "M_Bookmarks"

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : F_Bookmarks
'* Purpose    : per workbook vbe bookmarks
'* Copyright  :
'*
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 19-07-2023 10:08    Alex                per workbook vbe bookmarks
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Public Sub BmSave1(): SaveBookmark 1: End Sub
Public Sub BmSave2(): SaveBookmark 2: End Sub
Public Sub BmSave3(): SaveBookmark 3: End Sub
Public Sub BmSave4(): SaveBookmark 4: End Sub
Public Sub BmSave5(): SaveBookmark 5: End Sub
Public Sub BmSave6(): SaveBookmark 6: End Sub
Public Sub BmSave7(): SaveBookmark 7: End Sub
Public Sub BmSave8(): SaveBookmark 8: End Sub
Public Sub BmSave9(): SaveBookmark 9: End Sub
Public Sub BmSave10(): SaveBookmark 10: End Sub
Public Sub BmSave11(): SaveBookmark 11: End Sub
Public Sub BmSave12(): SaveBookmark 12: End Sub
Public Sub BmSave13(): SaveBookmark 13: End Sub
Public Sub BmSave14(): SaveBookmark 14: End Sub
Public Sub BmSave15(): SaveBookmark 15: End Sub
Public Sub BmSave16(): SaveBookmark 16: End Sub
Public Sub BmSave17(): SaveBookmark 17: End Sub
Public Sub BmSave18(): SaveBookmark 18: End Sub
Public Sub BmSave19(): SaveBookmark 19: End Sub
Public Sub BmSave20(): SaveBookmark 20: End Sub

Public Sub bmload1(): LoadBookmark 1: End Sub
Public Sub bmload2(): LoadBookmark 2: End Sub
Public Sub bmload3(): LoadBookmark 3: End Sub
Public Sub bmload4(): LoadBookmark 4: End Sub
Public Sub bmload5(): LoadBookmark 5: End Sub
Public Sub bmload6(): LoadBookmark 6: End Sub
Public Sub bmload7(): LoadBookmark 7: End Sub
Public Sub bmload8(): LoadBookmark 8: End Sub
Public Sub bmload9(): LoadBookmark 9: End Sub
Public Sub bmload10(): LoadBookmark 10: End Sub
Public Sub bmload11(): LoadBookmark 11: End Sub
Public Sub bmload12(): LoadBookmark 12: End Sub
Public Sub bmload13(): LoadBookmark 13: End Sub
Public Sub bmload14(): LoadBookmark 14: End Sub
Public Sub bmload15(): LoadBookmark 15: End Sub
Public Sub bmload16(): LoadBookmark 16: End Sub
Public Sub bmload17(): LoadBookmark 17: End Sub
Public Sub bmload18(): LoadBookmark 18: End Sub
Public Sub bmload19(): LoadBookmark 19: End Sub
Public Sub bmload20(): LoadBookmark 20: End Sub


'* Modified   : Date and Time       Author              Description
'* Updated    : 18-08-2023 09:42    Alex                add @Bookmark comment (F_Bookmarks.bas > SaveBookmark)

Sub SaveBookmark(index As Long)
    '@LastModified 2308180942

    Dim TargetWorkbook As Workbook: Set TargetWorkbook = ActiveCodepaneWorkbook

    If IniReadKey(BookmarkFile, ActiveCodepaneWorkbook.Name, "Initialized") = "" Then
        IniWrite BookmarkFile, TargetWorkbook.Name, "Initialized", "True"
        ResetBookmarks
    End If

    Dim delim       As String: delim = " | "
    Dim module      As VBComponent: Set module = ActiveModule
    Dim Procedure   As String: Procedure = ActiveProcedure: If Procedure = "" Then Procedure = "N/A"
    Dim BookmarkLine As String
    Dim lineIndex   As Long: lineIndex = aCodeModule.Init(module).RowFirst
    BookmarkLine = TargetWorkbook.Name & delim & _
            module.Name & delim & _
            Procedure & delim & _
            module.CodeModule.Lines(lineIndex, 1)
    module.CodeModule.InsertLines lineIndex, "'@Bookmark" & index
    IniWrite BookmarkFile, TargetWorkbook.Name, CStr(index), BookmarkLine
End Sub

Function BookmarkFile() As String
    BookmarkFile = ThisWorkbook.path & "\Lib\config\bookmarks.ini"
End Function

Sub ListBookmarks()
    dp IniReadSection(BookmarkFile, ActiveCodepaneWorkbook.Name)
    FindCode "@Bookmark"
    '* @TODO Created: 17-08-2023 08:42 Author: Anastasiou Alex
    '* @TODO show on userform - use uCodeFinder? add '@BOOKMARK comment instead?
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 18-08-2023 09:41    Alex                remove @bookmark comment (F_Bookmarks.bas > RemoveBookmark)

Sub RemoveBookmark(index As Long)
    '@LastModified 2308180941
    Dim BookmarkLine As String: BookmarkLine = IniReadKey(BookmarkFile, ActiveCodepaneWorkbook.Name, CStr(index))
    If BookmarkLine = vbNullString Then Exit Sub
    Dim delim       As String: delim = " | "
    Dim var: var = Split(BookmarkLine, delim)
    Dim targetworkbookname As String: targetworkbookname = var(0)
    Dim ComponentName As String: ComponentName = var(1)
    Dim Procedure   As String: Procedure = var(2)
    BookmarkLine = var(3)

    Dim wb          As Workbook
    Dim module      As VBComponent
    On Error Resume Next
    Set wb = Workbooks(targetworkbookname)
    Set module = wb.VBProject.VBComponents(ComponentName)
    If module Is Nothing Then Set module = ModuleOfProcedure(wb, Procedure)
    On Error GoTo 0
    If module Is Nothing Then Exit Sub
    Dim i           As Long
    Dim ap          As New aProcedure
    ap.Init wb, module, Procedure
    For i = ap.Line_Last To ap.Line_First Step -1
        If InStr(1, ap.Parent.CodeModule.Lines(i, 1), "'@Bookmark" & index, vbTextCompare) > 0 Then
            ap.Parent.CodeModule.DeleteLines i, 1
            Exit For
        End If
    Next

    IniWrite BookmarkFile, ActiveCodepaneWorkbook.Name, CStr(index), ""

End Sub

Sub ResetBookmarks()
    Dim index       As Long
    For index = 1 To 20
        IniWrite BookmarkFile, ActiveCodepaneWorkbook.Name, CStr(index), ""
    Next
    'IniWrite BookmarkFile, ActiveCodepaneWorkbook.Name, vbNullString, vbNullString
End Sub

Sub LoadBookmark(index As Long)
    Dim BookmarkLine As String: BookmarkLine = IniReadKey(BookmarkFile, ActiveCodepaneWorkbook.Name, CStr(index))
    If BookmarkLine = vbNullString Then Exit Sub
retry:
    Dim delim       As String: delim = " | "
    Dim var: var = Split(BookmarkLine, delim)
    Dim targetworkbookname As String: targetworkbookname = var(0)
    Dim ComponentName As String: ComponentName = var(1)
    Dim Procedure   As String: Procedure = var(2)
    BookmarkLine = var(3)

    Dim wb          As Workbook
    Dim module      As VBComponent
    On Error Resume Next
    Set wb = Workbooks(targetworkbookname)
    Set module = wb.VBProject.VBComponents(ComponentName)
    If module Is Nothing Then Set module = ModuleOfProcedure(wb, Procedure)
    On Error GoTo 0
    If module Is Nothing Then Exit Sub
    aModule.Init(module).Activate
    module.CodeModule.CodePane.SetSelection 1, 1, 1, 1
    Dim ProcFirstline As Long
    If Procedure = "N/A" Then Exit Sub
    If Not ProcedureExists(wb, Procedure) Then
        Debug.Print "Procedure " & Procedure & " not found in workbook " & targetworkbookname
        Exit Sub
    End If
    Dim ap          As New aProcedure: ap.Init wb, module, Procedure
    ProcFirstline = ap.Line_First
    module.CodeModule.CodePane.SetSelection ProcFirstline, 1, ProcFirstline, 1
    If BookmarkLine <> "" Then
        If ap.Contains(BookmarkLine, False, False, False) Then
            Dim i   As Long
            For i = ProcFirstline To ap.Line_Last
                If InStr(1, module.CodeModule.Lines(i, 1), BookmarkLine, vbTextCompare) > 0 Then
                    module.CodeModule.CodePane.SetSelection i, 1, i, 10000
                    module.CodeModule.CodePane.TopLine = i
                    Exit Sub
                End If
            Next
        End If
    End If
End Sub

