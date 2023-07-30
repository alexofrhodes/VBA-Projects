Attribute VB_Name = "vbArc"

'***************************************
'* AUTHOR   Anastasiou Alex
'* EMAIL    AnastasiouAlex@gmail.com
'* GITHUB   https://github.com/AlexOfRhodes
'* YOUTUBE  https://bit.ly/3aLZU9M
'* VK       https://vk.com/video/playlist/735281600_1
'***************************************

Public TargetTextbox As MSForms.Textbox

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwNilliseconds As Long)
#Else
    Public Declare  Sub Sleep Lib "kernel32" (ByVal dwNilliseconds As Long)
#End If

'Callback for buttonImmediateWindows onAction
Sub ShowImmediateWindowsUserform(control As IRibbonControl)
    ShowUserformCodeOnTheFly
End Sub


Sub ShowUserformCodeOnTheFly()
    uCodeOnTheFly.Show
End Sub


Sub LoadUserformPosition(Form As Object)
    If GetSetting("My Settings Folder", Form.Name, "Left Position") = "" _
        And GetSetting("My Settings Folder", Form.Name, "Top Position") = "" Then
        Form.StartUpPosition = 1
    Else
        Form.left = GetSetting("My Settings Folder", Form.Name, "Left Position")
        Form.top = GetSetting("My Settings Folder", Form.Name, "Top Position")
    End If
End Sub

Sub SaveUserformPosition(Form As Object)
    SaveSetting "My Settings Folder", Form.Name, "Left Position", Form.left
    SaveSetting "My Settings Folder", Form.Name, "Top Position", Form.top
End Sub

Function ListboxSelectedValues(listboxCollection As Variant) As Collection
    Dim i As Long
    Dim listItem As Long
    Dim selectedCollection As Collection
    Set selectedCollection = New Collection
    Dim listboxCount As Long
    If TypeName(listboxCollection) = "Collection" Then
        For listboxCount = 1 To listboxCollection.count
            If listboxCollection(listboxCount).ListCount > 0 Then
                For listItem = 0 To listboxCollection(listboxCount).ListCount - 1
                    If listboxCollection(listboxCount).Selected(listItem) Then
                        selectedCollection.Add CStr(listboxCollection(listboxCount).List(listItem, listboxCollection(listboxCount).BoundColumn - 1))
                    End If
                Next listItem
            End If
        Next listboxCount
    Else
        If listboxCollection.ListCount > 0 Then
            For i = 0 To listboxCollection.ListCount - 1
                If listboxCollection.Selected(i) Then
                    selectedCollection.Add listboxCollection.List(i, listboxCollection.BoundColumn - 1)
                End If
            Next i
        End If
    End If
    Set ListboxSelectedValues = selectedCollection
End Function

Public Function TextOfControl(c As control) As Variant
    Rem Text of Textbox, Selection of Combobox, Selected items (2d) of Listbox
    '#INCLUDE ListboxSelectedValues
    '#INCLUDE CollectionToArray
    Dim out As New Collection
    If TypeName(c) = "TextBox" Then
        If c.SelLength = 0 Then
            TextOfControl = c.Text
        Else
            TextOfControl = c.SelText
        End If
    ElseIf TypeName(c) = "ComboBox" Then
        If c.Style < 2 Then
            TextOfControl = c.Text
        Else
            TextOfControl = ""
        End If
    ElseIf TypeName(c) = "ListBox" Then
        Set out = ListboxSelectedValues(c)
        If out.count > 0 Then
            TextOfControl = CollectionToArray(out)
        Else
            TextOfControl = ""
        End If
    End If
End Function

Function CollectionToArray(c As Collection) As Variant
    Dim a() As Variant: ReDim a(0 To c.count - 1)
    Dim i As Long
    For i = 1 To c.count
        a(i - 1) = c.Item(i)
    Next
    CollectionToArray = a
End Function

Function ListboxSelectedIndexes(lBox As MSForms.ListBox) As Collection
    Dim i As Long
    Dim SelectedIndexes As Collection
    Set SelectedIndexes = New Collection
    If lBox.ListCount > 0 Then
        For i = 0 To lBox.ListCount - 1
            If lBox.Selected(i) Then SelectedIndexes.Add i
        Next i
    End If
    Set ListboxSelectedIndexes = SelectedIndexes
End Function

Sub SelectListboxItems(lBox As MSForms.ListBox, FindMe As Variant, Optional ByIndex As Boolean)
    Dim i As Long
    Select Case TypeName(FindMe)
    Case Is = "String", "Long", "Integer"
        For i = 0 To lBox.ListCount - 1
            If lBox.List(i) = CStr(FindMe) Then
                lBox.Selected(i) = True
                DoEvents
                If lBox.MultiSelect = fmMultiSelectSingle Then Exit Sub
            End If
        Next
    Case Else
        Dim el As Variant
        If ByIndex Then
            For Each el In FindMe
                lBox.Selected(el) = True
            Next
        Else
            For Each el In FindMe
                For i = 0 To lBox.ListCount - 1
                    If lBox.List(i) = el Then
                        lBox.Selected(i) = True
                        DoEvents
                    End If
                Next
            Next
        End If
    End Select
End Sub

Public Function IsInArray( _
    ByVal value1 As Variant, _
    ByVal array1 As Variant, _
    Optional CaseSensitive As Boolean) _
    As Boolean
    Dim individualElement As Variant
    If CaseSensitive = True Then value1 = UCase(value1)
    For Each individualElement In array1
        If CaseSensitive = True Then individualElement = UCase(individualElement)
        If individualElement = value1 Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Function CreateOrSetSheet(sheetName As String, TargetWorkbook As Workbook) As Worksheet
    '#INCLUDE WorksheetExists
    Dim NewSheet As Worksheet
    If WorksheetExists(sheetName, TargetWorkbook) = True Then
        Set CreateOrSetSheet = TargetWorkbook.Sheets(sheetName)
    Else
        Set CreateOrSetSheet = TargetWorkbook.Sheets.Add
        CreateOrSetSheet.Name = sheetName
    End If
End Function

Function ProceduresOfWorkbook( _
    TargetWorkbook As Workbook, _
    Optional ExcludeDocument As Boolean = True, _
    Optional ExcludeClass As Boolean = True, _
    Optional ExcludeForm As Boolean = True) As Collection
    Dim Module As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim LineNum As Long
    Dim coll As New Collection
    Dim ProcedureName As String
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If ExcludeClass = True Then
            If Module.Type = vbext_ct_ClassModule Then GoTo skip
        End If
        If ExcludeDocument = True Then
            If Module.Type = vbext_ct_Document Then GoTo skip
        End If
        If ExcludeForm = True Then
            If Module.Type = vbext_ct_MSForm Then GoTo skip
        End If
        With Module.CodeModule
            LineNum = .CountOfDeclarationLines + 1
            Do Until LineNum >= .CountOfLines
                ProcedureName = .ProcOfLine(LineNum, ProcKind)
                coll.Add ProcedureName
                LineNum = .ProcStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
            Loop
        End With
skip:
    Next Module
    Set ProceduresOfWorkbook = coll
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 17-10-2022 12:21    Alex                Added option to exclude comments after End Sub/Function (ProcedureEndLine)

Public Function ProcedureEndLine(Module As VBComponent, procName As String, Optional Strict As Boolean) As Long
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim startAt As Long
    Dim endAt As Long
    Dim CountOf As Long
    startAt = Module.CodeModule.ProcStartLine(procName, ProcKind)
    endAt = Module.CodeModule.ProcStartLine(procName, ProcKind) + Module.CodeModule.ProcCountLines(procName, ProcKind) - 1
    CountOf = Module.CodeModule.ProcCountLines(procName, ProcKind)
    
    If Strict Then
        Do While Not Module.CodeModule.Lines(endAt, 1) Like "End *"
            endAt = endAt - 1
        Loop
    End If
    
    ProcedureEndLine = endAt
End Function

Public Sub UpdateProcedureCode( _
    Procedure As Variant, _
    code As String, _
    TargetWorkbook As Workbook, Optional Module As VBComponent)
    '#INCLUDE ModuleOfProcedure
    '#INCLUDE moduleOfProcedure
    Dim StartLine As Integer
    Dim NumLines As Integer
    If Module Is Nothing Then Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
    With Module.CodeModule
        StartLine = .ProcStartLine(Procedure, vbext_pk_Proc)
        NumLines = .ProcCountLines(Procedure, vbext_pk_Proc)
        .DeleteLines StartLine, NumLines
        .InsertLines StartLine, code
    End With
End Sub

Function ModuleOfProcedure(wb As Workbook, ProcedureName As Variant) As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim LineNum As Long, NumProc As Long
    Dim procName As String
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        'If vbComp.name = "A_Test2" Then Stop
        With vbComp.CodeModule
            LineNum = .CountOfDeclarationLines + 1
            Do Until LineNum >= .CountOfLines
                procName = .ProcOfLine(LineNum, ProcKind)
                LineNum = .ProcStartLine(procName, ProcKind) + .ProcCountLines(procName, ProcKind) + 1
                If UCase(procName) = UCase(ProcedureName) Then
                    Set ModuleOfProcedure = vbComp
                    Exit Function
                End If
            Loop
        End With
    Next vbComp
End Function

Public Function ActiveCodepaneWorkbook() As Workbook
    Dim tmpstr As String
    tmpstr = Application.VBE.SelectedVBComponent.Collection.parent.Filename
    tmpstr = Right(tmpstr, Len(tmpstr) - InStrRev(tmpstr, "\"))
    Set ActiveCodepaneWorkbook = Workbooks(tmpstr)
End Function

Function ProcedureExists( _
    ProcedureName As Variant, _
    FromWorkbook As Workbook) _
    As Boolean
    '#INCLUDE ProceduresOfWorkbook
    Dim AllProcedures As Collection: Set AllProcedures = ProceduresOfWorkbook(FromWorkbook)
    Dim Procedure As Variant
    For Each Procedure In AllProcedures
        If UCase(CStr(Procedure)) = UCase(ProcedureName) Then
            ProcedureExists = True
            Exit Function
        End If
    Next
End Function


Public Sub Reframe(Form As Object, control As MSForms.control)
    Dim c As MSForms.control
    For Each c In Form.Controls
        If TypeName(c) = "Frame" Then
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                If c.Name <> control.parent.parent.Name Then c.visible = False
            End If
        End If
    Next
    Form.Controls(control.Caption).visible = True
    For Each c In Form.Controls
        If TypeName(c) = "Label" Then
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                c.BackColor = &H534848
            End If
        End If
    Next
    control.BackColor = &H80B91E
End Sub


Sub appRunOnTime(timeToRun, macroToRun As String, Optional arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5)
    
    If TypeName(arg5) <> "Error" Then
        Application.OnTime timeToRun, "'" & macroToRun & """" & arg1 & """ ,""" & arg2 & """ ,""" & arg3 & """ ,""" & arg4 & """ ,""" & arg5 & " '"
    ElseIf TypeName(arg4) <> "Error" Then
        Application.OnTime timeToRun, "'" & macroToRun & """" & arg1 & """ ,""" & arg2 & """ ,""" & arg3 & """ ,""" & arg4 & " '"
    ElseIf TypeName(arg3) <> "Error" Then
        Application.OnTime timeToRun, "'" & macroToRun & """" & arg1 & """ ,""" & arg2 & """ ,""" & arg3 & " '"
    ElseIf TypeName(arg2) <> "Error" Then
        Application.OnTime timeToRun, "'" & macroToRun & """" & arg1 & """ ,""" & arg2 & " '"
    ElseIf TypeName(arg1) <> "Error" Then
        Application.OnTime timeToRun, "'" & macroToRun & """" & arg1 & """ '"
    Else
        Application.OnTime timeToRun, macroToRun
    End If
End Sub

Sub LoadUserformOptions(Form As Object, Optional ExcludeThese As Variant)
    '#INCLUDE SelectListboxItems
    '#INCLUDE IsInArray
    '#INCLUDE CreateOrSetSheet
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet(Form.Name & "_Settings", ThisWorkbook)
    If ws.Range("A1") = "" Then Exit Sub
    Dim cell As Range
    Set cell = ws.Cells(1, 1)
    Dim c As MSForms.control
    Dim v
    On Error Resume Next
    Do While cell <> ""
        Set c = Form.Controls(cell.Text)
        If Not TypeName(c) = "Nothing " Then
            If Not IsInArray(cell, ExcludeThese) Then
                Select Case TypeName(c)
                Case "TextBox", "CheckBox", "OptionButton", "ToggleButton", "ComboBox"
                    c.Value = cell.Offset(0, 1)
                Case "ListBox"
                    If InStr(1, cell.Offset(0, 1), ",") > 0 Then
                        SelectListboxItems c, Split(cell.Offset(0, 1), ","), True
                    Else
                        c.Selected(CInt(cell.Offset(0, 1))) = True
                    End If
                End Select
            End If
        End If
        Set cell = cell.Offset(1, 0)
    Loop
End Sub


Sub SaveUserformOptions(Form As Object, _
    Optional includeCheckbox As Boolean = True, _
    Optional includeOptionButton As Boolean = True, _
    Optional includeTextBox As Boolean = True, _
    Optional includeListbox As Boolean = True, _
    Optional includeToggleButton As Boolean = True, _
    Optional includeCombobox As Boolean = True)
    '#INCLUDE ListboxSelectedIndexes
    '#INCLUDE CreateOrSetSheet
    '#INCLUDE CollectionToArray
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet(Form.Name & "_Settings", ThisWorkbook)
    ws.Range("A1").CurrentRegion.clear
    Dim coll As New Collection
    Dim cell As Range
    Set cell = ws.Cells(1, 1)
    Dim c As MSForms.control
    For Each c In Form.Controls
        If TypeName(c) Like "CheckBox" Then
            If Not includeCheckbox Then GoTo skip
        ElseIf TypeName(c) Like "OptionButton" Then
            If Not includeOptionButton Then GoTo skip
        ElseIf TypeName(c) Like "TextBox" Then
            If Not includeTextBox Then GoTo skip
        ElseIf TypeName(c) = "ListBox" Then
            If Not includeListbox Then GoTo skip
        ElseIf TypeName(c) Like "ToggleButton" Then
            If Not includeToggleButton Then GoTo skip
        ElseIf TypeName(c) Like "ComboBox" Then
            If Not includeCombobox Then GoTo skip
        Else
            GoTo skip
        End If
        cell = c.Name
        Select Case TypeName(c)
        Case "TextBox", "CheckBox", "OptionButton", "ToggleButton", "ComboBox"
            cell.Offset(0, 1) = c.Value
        Case "ListBox"
            Set coll = ListboxSelectedIndexes(c)
            If coll.count > 0 Then
                cell.Offset(0, 1) = Join(CollectionToArray(coll), ",")
            Else
                cell.Offset(0, 1) = -1
            End If
        End Select
        Set cell = cell.Offset(1, 0)
skip:
    Next
End Sub

