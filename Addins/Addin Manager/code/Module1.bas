Attribute VB_Name = "Module1"
Sub AddinManagerButtonClicked(Control As IRibbonControl)
    uAddinManager.Show
End Sub

Public Function LastModified(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    LastModified = f.DateLastModified
End Function

Sub Admin_AddinsModified()
    '    ThisWorkbook.Sheets("Sheet1").Calculate
    On Error Resume Next
    Application.CalculateFull
    On Error GoTo 0
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim cell As Range
    Set cell = Sheet1.Range("E2")
    Dim ans As Long
    Do While cell <> ""
        If cell.Text = "A" Then
            ans = MsgBox("Replace" & vbTab & Sheet1.Range("A1") & cell.Offset(0, -4) & vbTab & cell.Offset(0, -2) & vbNewLine & " With" & vbTab & _
                         Sheet1.Range("G1") & cell.Offset(0, 2) & vbTab & cell.Offset(0, 4), vbYesNo, "Checking for updates")
        ElseIf cell.Text = "B" Then
            ans = MsgBox("Replace" & vbTab & Sheet1.Range("G1") & cell.Offset(0, 2) & vbTab & cell.Offset(0, 4) & vbNewLine & " With" & vbTab & _
                         Sheet1.Range("A1") & cell.Offset(0, -4) & vbTab & cell.Offset(0, -2), vbYesNo, "Checking for updates")
        End If
        
        If ans = vbYes Then
            If cell.Text = "A" Then
                Workbooks(cell.Offset(0, 2).Text).IsAddin = True
                Workbooks(cell.Offset(0, 2).Text).Close
                fso.CopyFile cell.Offset(0, 3).Text, cell.Offset(0, -3).Text, True
                Workbooks.Open cell.Offset(0, -3)
            ElseIf cell.Text = "B" Then
                fso.CopyFile cell.Offset(0, -3).Text, cell.Offset(0, 3).Text, True
            End If
        End If
        Set cell = cell.Offset(1, 0)
    Loop
    Application.CalculateFull
End Sub

Sub FollowLink(FolderPath As String)
    Dim oShell As Object
    Dim Wnd As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.Name = "File Explorer" Then
            If Wnd.Document.Folder.Self.Path = FolderPath Then Exit Sub
        End If
    Next Wnd
    Application.ThisWorkbook.FollowHyperlink Address:=FolderPath, NewWindow:=True
End Sub

Function WorkbookIsOpen(ByVal sWbkName As String) As Boolean
    WorkbookIsOpen = False
    On Error Resume Next
    WorkbookIsOpen = Len(Workbooks(sWbkName).Name) <> 0
    On Error GoTo 0
End Function


