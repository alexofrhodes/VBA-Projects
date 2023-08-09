Attribute VB_Name = "F_Bookmarks"

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


Sub SaveBookmark(index As Long)
    Dim BookmarkFile As String: BookmarkFile = ThisWorkbook.Path & "config\bookmarks.ini"
    Dim delim As String: delim = " | "
    Dim TargetWorkbook As Workbook: Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim module As VBComponent:  Set module = ActiveModule
    Dim Procedure As String: Procedure = ActiveProcedure: If Procedure = "" Then Procedure = "N/A"
    Dim BookmarkLine As String
    BookmarkLine = TargetWorkbook.Name & delim & _
                   module.Name & delim & _
                   Procedure & delim & _
                   module.CodeModule.Lines(aCodeModule.Init(module).RowFirst, 1)
    IniWrite BookmarkFile, TargetWorkbook.Name, CStr(index), BookmarkLine
End Sub

Sub ListBookmarks()
    Dim BookmarkFile As String: BookmarkFile = ThisWorkbook.Path & "config\bookmarks.ini"
    dp IniReadSection(BookmarkFile, ActiveCodepaneWorkbook.Name)
End Sub
Sub RemoveBookmark(index As Long)
    Dim BookmarkFile As String: BookmarkFile = ThisWorkbook.Path & "config\bookmarks.ini"
    WritePrivateProfileString ActiveCodepaneWorkbook.Name, CStr(index), vbNullString, BookmarkFile
End Sub
Sub ResetBookmarks()
    Dim BookmarkFile As String: BookmarkFile = ThisWorkbook.Path & "config\bookmarks.ini"
    WritePrivateProfileString ActiveCodepaneWorkbook.Name, vbNullString, vbNullString, BookmarkFile
End Sub
Sub LoadBookmark(index As Long)
    Dim BookmarkFile As String: BookmarkFile = ThisWorkbook.Path & "config\bookmarks.ini"
    Dim BookmarkLine As String: BookmarkLine = IniReadKey(BookmarkFile, ActiveCodepaneWorkbook.Name, CStr(index))
    If BookmarkLine = vbNullString Then Exit Sub
Retry:
    Dim delim As String:                delim = " | "
    Dim var:                            var = Split(BookmarkLine, delim)
    Dim targetworkbookname As String:   targetworkbookname = var(0)
    Dim ComponentName As String:        ComponentName = var(1)
    Dim Procedure As String:            Procedure = var(2)
                                        BookmarkLine = var(3)

    Dim WB As Workbook
    Dim module As VBComponent
    On Error Resume Next
        Set WB = Workbooks(targetworkbookname)
        Set module = WB.VBProject.VBComponents(ComponentName)
        If module Is Nothing Then Set module = ModuleOfProcedure(WB, Procedure)
    On Error GoTo 0
    If module Is Nothing Then Exit Sub
    aModule.Init(module).Activate
    module.CodeModule.CodePane.SetSelection 1, 1, 1, 1
    Dim ProcFirstline As Long
    If Procedure = "N/A" Then Exit Sub
    If Not ProcedureExists(WB, Procedure) Then
        Debug.Print "Procedure " & Procedure & " not found in workbook " & targetworkbookname
        Exit Sub
    End If
    Dim AP As New aProcedure: AP.Init WB, module, Procedure
    ProcFirstline = AP.LineIndex(Procedure_First)
    module.CodeModule.CodePane.SetSelection ProcFirstline, 1, ProcFirstline, 1
    If BookmarkLine <> "" Then
        If AP.Contains(BookmarkLine, False, False, False) Then
            Dim i As Long
            For i = ProcFirstline To AP.LineIndex(Procedure_Last)
                If InStr(1, module.CodeModule.Lines(i, 1), BookmarkLine, vbTextCompare) > 0 Then
                    module.CodeModule.CodePane.SetSelection i, 1, i, 10000
                    Exit Sub
                End If
            Next
        End If
    End If
End Sub

