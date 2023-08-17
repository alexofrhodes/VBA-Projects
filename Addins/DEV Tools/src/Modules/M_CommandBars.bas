Attribute VB_Name = "M_CommandBars"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''
''DEVELOPER  Anastasiou Alex
''EMAIL      AnastasiouAlex@gmail.com
''GITHUB     https://github.com/AlexOfRhodes
''YOUTUBE    https://bit.ly/3aLZU9M
''VK         https://vk.com/video/playlist/735281600_1
''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Const rBUILD_ON_OPEN = "I3"
Public Const rC_TAG = "I4"
Public Const rMENU_TYPE = "I5"
Public Const rBAR_LOCATION = "I6"

Public MenuLevel, NextLevel, Caption, Divider, FaceId
Public MenuSheet As Worksheet
Public BarRow As Integer
Public MenuType As Long
Public Action As String

Public Const WorksheetMenu = 1
Public Const VbeMenu = 2
Public Const RightClickMenu = 3

Public BarLocation As String

Public C_TAG As String

Public MenuEvent        As CVBECommandHandler
Public EventHandlers    As New Collection
Public CmdBarItem       As CommandBarControl
Public TargetCommandbar 'As CommandBar
Public TargetControl    As CommandBarControl
Public MainMenu         As CommandBarControl
Public MenuItem         As CommandBarControl
Public Ctrl             As Office.CommandBarControl



Public Sub CreateAllBars()
'@AssignedModule M_CommandBars
'@INCLUDE PROCEDURE CommandBarBuilder
'@INCLUDE CLASS CVBECommandHandler
'@INCLUDE DECLARATION rBUILD_ON_OPEN
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase(Left(ws.Name, 4)) = "BAR_" Then
            If ws.Range(rBUILD_ON_OPEN) = True Then CommandBarBuilder ws
        End If
    Next
End Sub

Public Sub DeleteAllBars()
'@AssignedModule M_CommandBars
'@INCLUDE PROCEDURE DeleteControlsAndHandlers
'@INCLUDE CLASS CVBECommandHandler
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase(Left(ws.Name, 4)) = "BAR_" Then DeleteControlsAndHandlers ws
    Next
End Sub

Public Sub RestoreBars()
'@AssignedModule M_CommandBars
'@INCLUDE PROCEDURE CreateAllBars
'@INCLUDE CLASS CVBECommandHandler
    Application.OnTime Now, "CreateAllBars"
End Sub

Public Sub ListBars()
'@AssignedModule M_CommandBars
'@INCLUDE PROCEDURE ListWorksheetBars
'@INCLUDE PROCEDURE ListVBEBars
'@INCLUDE CLASS CVBECommandHandler
    ListWorksheetBars
    ListVBEBars
End Sub

Public Sub NewBar()
'@AssignedModule M_CommandBars
'@INCLUDE PROCEDURE lastBar
'@INCLUDE CLASS CVBECommandHandler
    Dim wsMain As Worksheet
    Set wsMain = ActiveSheet
    Dim wsCopy As Worksheet
    wsMain.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    Set wsCopy = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    wsCopy.Name = "BAR_" & lastBar + 1
    On Error Resume Next
    wsCopy.Range("A1").ListObject.DataBodyRange.Delete
    On Error GoTo 0
    wsCopy.Range("I5:I6").ClearContents
    wsCopy.Range("I3") = False
End Sub

Private Function lastBar() As Long
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
    Dim counter As Long
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase(Left(ws.Name, 4)) = "BAR_" Then counter = counter + 1
    Next
    lastBar = counter
End Function

'The names for each of the top-level CommandBars in the VBE 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Name                       Description
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Menu Bar                   The normal VBE menu bar
'Standard                   The normal VBE toolbar
'Edit                       The VBE edit toolbar, containing useful code-editing tools
'Debug                      The VBE debug toolbar, containing typical debugging tools
'UserForm                   The VBE UserForm toolbar, containing useful form-editing tools
'MSForms                    The popup menu for a UserForm (shown when you right-click the UserForm background)
'MSForms Control            The popup menu for a normal control on a UserForm
'MSForms Control Group      The popup menu that appears when you right-click a group of controls on a UserForm
'MSForms MPC                The popup menu for the Multi-Page control
'MSForms Palette            The popup menu that appears when you right-click a tool in the Control Toolbox
'MSForms Toolbox            The popup menu that appears when you right-click one of the tabs at the top of the Control Toolbox
'MSForms DragDrop           The popup menu that appears when you use the right mouse button to drag a control between tabs in the Control Toolbox, or onto a UserForm
'Code Window                The popup menu for a code window
'Code Window (Break)        The popup menu for a code window, when in Break (debug) mode
'Watch Window               The popup menu for the Watch window
'Immediate Window           The popup menu for the Immediate window
'Locals Window              The popup menu for the Locals window
'Project Window             The popup menu for the Project Explorer
'Project Window (Break)     The popup menu for the Project Explorer, when in Break mode
'Object Browser             The popup menu for the Object Browser
'Property Browser           The popup menu for the Properties window
'Docked Window              The popup menu that appears when you right-click the title bar of a docked window

Public Function BarExists(findBarName As String) As Boolean
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
    Dim bar As CommandBar
    For Each bar In Application.CommandBars
        If UCase(bar.Name) = UCase(findBarName) Then
            BarExists = True
            Exit Function
        End If
    Next bar
    For Each bar In Application.VBE.CommandBars
        If UCase(bar.Name) = UCase(findBarName) Then
            BarExists = True
            Exit Function
        End If
    Next bar
End Function

Public Sub BuildBarFromShape()
'@AssignedModule M_CommandBars
'@INCLUDE PROCEDURE CommandBarBuilder
'@INCLUDE CLASS CVBECommandHandler
    CommandBarBuilder ActiveSheet
End Sub

Public Sub DeleteBarFromShape()
'@AssignedModule M_CommandBars
'@INCLUDE PROCEDURE DeleteControlsAndHandlers
'@INCLUDE CLASS CVBECommandHandler
    DeleteControlsAndHandlers ActiveSheet
End Sub

Private Sub CommandBarBuilder(ws As Worksheet)
'@AssignedModule M_CommandBars
'@INCLUDE PROCEDURE SetCMDbar
'@INCLUDE PROCEDURE CreateMainMenu
'@INCLUDE PROCEDURE CreatePopup
'@INCLUDE PROCEDURE CreateButton
'@INCLUDE PROCEDURE DirectButton
'@INCLUDE PROCEDURE DeleteControlsAndHandlers
'@INCLUDE PROCEDURE markControlType
'@INCLUDE PROCEDURE SetVariables
'@INCLUDE PROCEDURE ReSetVariables
'@INCLUDE CLASS CVBECommandHandler
'@INCLUDE DECLARATION VbeMenu
'@INCLUDE DECLARATION BarLocation
'@INCLUDE DECLARATION BarRow
'@INCLUDE DECLARATION C_TAG
'@INCLUDE DECLARATION MenuSheet
'@INCLUDE DECLARATION MenuType
'@INCLUDE DECLARATION TargetControl
    DeleteControlsAndHandlers ws
    SetCMDbar ws
    Set MenuSheet = ws
    BarRow = 2
'    If MenuType = VbeMenu Then
'        If BarLocation = "Menu Bar" Then
'            Set TargetControl = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
'        End If
'        If Not TargetControl Is Nothing Then
'            TargetControl.Caption = C_TAG
'            TargetControl.Tag = C_TAG
'        End If
'    End If
    Do Until IsEmpty(MenuSheet.Cells(BarRow, 1))
        With MenuSheet
            SetVariables
        End With
        Select Case MenuLevel
        '@TODO add in the table "Main Menu|Tools" etc
        'and in the following functions create main menu / target... by splitting
        'to get commandbars.controls.add
        'or commandbars.controls(...).add
            Case 1
                If NextLevel > MenuLevel Then
                    CreateMainMenu
                Else
                    DirectButton
                End If
            Case 2
                If NextLevel > MenuLevel Then
                    CreatePopup
                Else
                    DirectButton
                End If
            Case 3
                CreateButton
        End Select
        BarRow = BarRow + 1
        ReSetVariables
    Loop
    markControlType ws
    Debug.Print "Bar created"
End Sub

Private Function SetCMDbar(ws As Worksheet) As Boolean
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
'@INCLUDE DECLARATION rBAR_LOCATION
'@INCLUDE DECLARATION rC_TAG
'@INCLUDE DECLARATION RightClickMenu
'@INCLUDE DECLARATION rMENU_TYPE
'@INCLUDE DECLARATION VbeMenu
'@INCLUDE DECLARATION WorksheetMenu
'@INCLUDE DECLARATION BarLocation
'@INCLUDE DECLARATION C_TAG
'@INCLUDE DECLARATION MenuType
    C_TAG = ws.Range(rC_TAG)
    Select Case LCase(ws.Range(rMENU_TYPE))
        Case Is = LCase("WorksheetMenu")
            MenuType = WorksheetMenu
        Case Is = LCase("vbeMenu")
            MenuType = VbeMenu
        Case Is = LCase("RightClickMenu")
            MenuType = RightClickMenu
        Case Else
    End Select
    If ws.Range(rBAR_LOCATION) <> "" Then
        BarLocation = ws.Range(rBAR_LOCATION)
    Else
        BarLocation = 0
    End If
    If MenuType = VbeMenu Then
        Select Case BarLocation
            'case sensitive
            Case Is = "Menu Bar", "Code Window", "Project Window", "Debug", "Userform", _
                      "MSForms", "MSForms Control", "MSForms Control Group", "MSForms MPC", _
                      "MSForms Palette", "MSForms Toolbox", "MSForms DragDrop", "Code Window (Break)", _
                      "Watch Window", "Immediate Window", "Locals Window", "Project Window (Break)", _
                      "Object Browser", "Property Browser", "Docked Window", _
                      "View", "Insert", "Format", "Add-Ins", "Window"
                Set TargetCommandbar = Application.VBE.CommandBars(BarLocation)
                SetCMDbar = True
            Case Else
                Set TargetCommandbar = Application.VBE.CommandBars.Add(C_TAG, Position:=msoBarTop, Temporary:=True)
                TargetCommandbar.Visible = True
        End Select
    ElseIf MenuType = WorksheetMenu Then
        Select Case ws.Range(rBAR_LOCATION)
            Case Is = "Worksheet Menu Bar", "Cell", "Column", "Row"
                Set TargetCommandbar = Application.CommandBars(BarLocation)
                SetCMDbar = True
            Case Else
        End Select
    Else
    End If
End Function


Private Sub CreateMainMenu()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
'@INCLUDE DECLARATION RightClickMenu
'@INCLUDE DECLARATION VbeMenu
'@INCLUDE DECLARATION WorksheetMenu
'@INCLUDE DECLARATION Action
'@INCLUDE DECLARATION C_TAG
'@INCLUDE DECLARATION MainMenu
'@INCLUDE DECLARATION MenuType
    If MenuType = VbeMenu Then
        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup)
    ElseIf MenuType = WorksheetMenu Then
        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    ElseIf MenuType = RightClickMenu Then
'        Set TargetCommandbar = CommandBars.Add(C_TAG, msoBarPopup, , True)
'        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup)
        On Error Resume Next
        CommandBars.Add C_TAG, msoBarPopup, , True
        On Error GoTo 0
        Set MainMenu = CommandBars(C_TAG).Controls.Add(Type:=msoControlPopup)
        'Exit Sub
    End If
    With MainMenu
        .Caption = Caption
        .BeginGroup = Divider
        If FaceId <> "" And Action <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
End Sub

Private Sub CreatePopup()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
'@INCLUDE DECLARATION RightClickMenu
'@INCLUDE DECLARATION Action
'@INCLUDE DECLARATION C_TAG
'@INCLUDE DECLARATION MainMenu
'@INCLUDE DECLARATION MenuItem
'@INCLUDE DECLARATION MenuType
    If MenuType = RightClickMenu Then
'        Set MenuItem = TargetCommandbar.Controls.Add(Type:=msoControlPopup)
        Set MenuItem = MainMenu.Controls.Add(Type:=msoControlPopup)
        
    Else
        Set MenuItem = MainMenu.Controls.Add(Type:=msoControlPopup)
    End If
    With MenuItem
        .Caption = Caption
        .BeginGroup = Divider
        If FaceId <> "" And Action <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
End Sub

Private Sub CreateButton()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
'@INCLUDE DECLARATION VbeMenu
'@INCLUDE DECLARATION Action
'@INCLUDE DECLARATION C_TAG
'@INCLUDE DECLARATION CmdBarItem
'@INCLUDE DECLARATION EventHandlers
'@INCLUDE DECLARATION MenuEvent
'@INCLUDE DECLARATION MenuItem
'@INCLUDE DECLARATION MenuType
    If MenuType = VbeMenu Then
        Set MenuEvent = New CVBECommandHandler
    End If
    Set CmdBarItem = MenuItem.Controls.Add
    With CmdBarItem
        .Caption = Caption
        .BeginGroup = Divider
        .OnAction = Action
        If FaceId <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
    If MenuType = VbeMenu Then
        Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
        On Error Resume Next
        EventHandlers.Add MenuEvent, CmdBarItem.Caption
        On Error GoTo 0
    End If
End Sub

Private Sub DirectButton()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
'@INCLUDE DECLARATION RightClickMenu
'@INCLUDE DECLARATION VbeMenu
'@INCLUDE DECLARATION Action
'@INCLUDE DECLARATION C_TAG
'@INCLUDE DECLARATION CmdBarItem
'@INCLUDE DECLARATION EventHandlers
'@INCLUDE DECLARATION MainMenu
'@INCLUDE DECLARATION MenuEvent
'@INCLUDE DECLARATION MenuType
    Dim CmdBarItem As CommandBarControl
    If MenuType = VbeMenu Then
        Set MenuEvent = New CVBECommandHandler
    End If
    Select Case MenuLevel
        Case Is = 1
            Set CmdBarItem = TargetCommandbar.Controls.Add
        Case Is = 2
'            If MenuType = RightClickMenu Then
'                Set CmdBarItem = TargetCommandbar.Controls.Add
'            Else
                Set CmdBarItem = MainMenu.Controls.Add
'            End If
    End Select
    With CmdBarItem
        .Style = msoButtonIconAndCaption
        .Caption = Caption
        .BeginGroup = Divider
        .OnAction = Action
        If FaceId <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
    If MenuType = VbeMenu Then
        Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
        EventHandlers.Add MenuEvent
    End If
End Sub

Private Sub DeleteControlsAndHandlers(ws As Worksheet)
'@AssignedModule M_CommandBars
'@INCLUDE PROCEDURE BarExists
'@INCLUDE PROCEDURE DeleteHandlersFor
'@INCLUDE CLASS CVBECommandHandler
'@INCLUDE DECLARATION rC_TAG
'@INCLUDE DECLARATION RightClickMenu
'@INCLUDE DECLARATION rMENU_TYPE
'@INCLUDE DECLARATION VbeMenu
'@INCLUDE DECLARATION WorksheetMenu
'@INCLUDE DECLARATION Ctrl
'@INCLUDE DECLARATION MenuType
    If ws.Range(rC_TAG).text = vbNullString Then Exit Sub
    Select Case LCase(ws.Range(rMENU_TYPE))
        Case "vbemenu"
            MenuType = VbeMenu
        Case "worksheetmenu"
            MenuType = WorksheetMenu
        Case "rightclickmenu"
            MenuType = RightClickMenu
    End Select
    If MenuType = VbeMenu Then
        If BarExists(ws.Range(rC_TAG)) Then
            Application.VBE.CommandBars(ws.Range(rC_TAG).text).Delete
            Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).text)
        End If
    ElseIf MenuType = WorksheetMenu Then
        Set Ctrl = Application.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).text)
    ElseIf MenuType = RightClickMenu Then
        If BarExists(ws.Range(rC_TAG).text) Then
            CommandBars(ws.Range(rC_TAG).text).Delete
        End If
        Exit Sub
    End If
    On Error Resume Next
    Do
        Ctrl.Delete
        If MenuType = VbeMenu Then
            Set Ctrl = Application.VBE.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).text)
        Else
            Set Ctrl = Application.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).text)
        End If
    Loop While Not Ctrl Is Nothing
    On Error GoTo 0
    DeleteHandlersFor ws
End Sub

Private Sub DeleteHandlersFor(ws As Worksheet)
'@AssignedModule M_CommandBars
'@INCLUDE PROCEDURE markControlType
'@INCLUDE CLASS CVBECommandHandler
'@INCLUDE DECLARATION EventHandlers
    On Error Resume Next
    markControlType ws
    Dim cell As Range
    Set cell = ws.Cells(2, 6)
    Do Until IsEmpty(cell)
        If cell.text = "Button" Then
            EventHandlers.Remove cell.offset(0, -3).text
        End If
        Set cell = cell.offset(1)
    Loop
End Sub

Private Sub markControlType(ws As Worksheet)
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
    ws.Columns("F").ClearContents
    Dim idx As Long: idx = 0
    Dim Description() As Variant
    Dim cell As Range
    Set cell = ws.Cells(2, 1)
    Do Until IsEmpty(cell)
        idx = idx + 1
        ReDim Preserve Description(1 To idx)
        Description(idx) = IIf(cell.offset(1) > cell, "PopUp", "Button")
        Set cell = cell.offset(1)
    Loop
    ws.Range("F2").Resize(UBound(Description)) = WorksheetFunction.Transpose(Description)
End Sub

Private Sub SetVariables()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
'@INCLUDE DECLARATION Action
'@INCLUDE DECLARATION BarRow
'@INCLUDE DECLARATION MenuSheet
    With MenuSheet
        MenuLevel = .Cells(BarRow, 1)
        Caption = .Cells(BarRow, 2)
        Action = .Cells(BarRow, 3)
        Divider = .Cells(BarRow, 4)
        FaceId = .Cells(BarRow, 5)
        NextLevel = .Cells(BarRow + 1, 1)
    End With
End Sub

Private Sub ReSetVariables()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
'@INCLUDE DECLARATION Action
    MenuLevel = ""
    Caption = ""
    Action = ""
    Divider = ""
    FaceId = ""
    NextLevel = ""
End Sub


Private Sub ListWorksheetBars()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
    Dim oCB As CommandBar
    Dim oWK As Worksheet
    Set oWK = ThisWorkbook.Sheets("ListSheetBars")
    oWK.Cells.Clear
    Dim arr As Variant
    arr = Array("Type", "Index", "Name", "Built-in", "Visible")
    Dim iCol As Integer
    iCol = UBound(arr) + 1
    oWK.Range("a1").Resize(1, iCol) = arr
    oWK.Range("a1").Resize(1, iCol).Cells.Font.Bold = True
    Dim i As Long
    i = 2
    Dim cbVar(300, 4) As Variant
    For Each oCB In Excel.Application.CommandBars
        cbVar(i - 2, 0) = Choose(oCB.Type + 1, "Toolbar", "Menu", "PopUp")
        cbVar(i - 2, 1) = oCB.index
        cbVar(i - 2, 2) = oCB.Name
        cbVar(i - 2, 3) = oCB.BuiltIn
        cbVar(i - 2, 4) = oCB.Visible
        i = i + 1
    Next
    oWK.Cells(2, 1).Resize(UBound(cbVar, 1) + 1, UBound(cbVar, 2) + 1) = cbVar
    oWK.Columns.AutoFit
End Sub

Private Sub ListVBEBars()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
    Dim oCB As CommandBar
    Dim oWK As Worksheet
    Set oWK = ThisWorkbook.Sheets("ListVBEBars")
    oWK.Cells.Clear
    Dim arr As Variant
    arr = Array("Type", "Index", "Name", "Built-in", "Visible")
    Dim iCol As Integer
    iCol = UBound(arr) + 1
    oWK.Range("a1").Resize(1, iCol) = arr
    oWK.Range("a1").Resize(1, iCol).Cells.Font.Bold = True
    Dim i As Long
    i = 2
    Dim cbVar(300, 4) As Variant
    For Each oCB In Application.VBE.CommandBars
        cbVar(i - 2, 0) = Choose(oCB.Type + 1, "Toolbar", "Menu", "PopUp")
        cbVar(i - 2, 1) = oCB.index
        cbVar(i - 2, 2) = oCB.Name
        cbVar(i - 2, 3) = oCB.BuiltIn
        cbVar(i - 2, 4) = oCB.Visible
        i = i + 1
    Next
    oWK.Cells(2, 1).Resize(UBound(cbVar, 1) + 1, UBound(cbVar, 2) + 1) = cbVar
    oWK.Columns.AutoFit
End Sub

Private Sub exampleOfControls()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
    Dim cbc As CommandBarControl
    Dim cbb As CommandBarButton
    Dim cbcm As CommandBarComboBox
    Dim cbp As CommandBarPopup
    With Application.VBE.CommandBars("CodeArchive")
        Set cbc = .Controls.Add(ID:=3, Temporary:=True)
        Set cbb = .Controls.Add(Temporary:=True)
        cbb.Caption = "A new command"
        cbb.Style = msoButtonCaption
        cbb.OnAction = "NewCommand_OnAction"
        Set cbcm = .Controls.Add(Type:=msoControlComboBox, Temporary:=True)
        cbcm.Caption = "Combo:"
        cbcm.AddItem "list entry 1"
        cbcm.AddItem "list entry 2"
        cbcm.OnAction = "NewCommand_OnAction"
        cbcm.Style = msoComboLabel
        Set cbc = .Controls.Add(Type:=msoControlDropdown, Temporary:=True)
        cbc.Caption = "Dropdown:"
        cbc.AddItem "list entry 1"
        cbc.AddItem "list entry 2"
        cbc.OnAction = "MenuDropdown_OnAction"
        Set cbp = .Controls.Add(Type:=msoControlPopup, Temporary:=True)
        cbp.Caption = "new sub menu"
        Set cbb = cbp.Controls.Add
        cbb.Caption = "sub entry 1"
        Set cbb = cbp.Controls.Add
        cbb.Caption = "sub entry 2"
    End With
End Sub

Private Sub ImageFromEmbedded()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
    Dim p As Excel.Picture
    Dim Btn As Office.CommandBarButton
    Set Btn = Application.CommandBars.FindControl(ID:=30007) _
        .Controls.Add(Type:=msoControlButton, Temporary:=True)
    Set p = Worksheets("Sheet1").Pictures("ThePict")
    p.CopyPicture xlScreen, xlBitmap
    With Btn
        .Caption = "Click Me"
        .Style = msoButtonIconAndCaption
        .PasteFace
    End With
End Sub

Private Sub ImageFromExternalFile()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
    Dim Btn As Office.CommandBarButton
    Set Btn = Application.CommandBars.FindControl(ID:=30007) _
        .Controls.Add(Type:=msoControlButton, Temporary:=True)
    With Btn
        .Caption = "Click Me"
        .Style = msoButtonIconAndCaption
        .Picture = LoadPicture("C:\TestPic.bmp")
    End With
End Sub

Private Sub ResetCBAR()
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
    Excel.Application.CommandBars("Cell").Reset
End Sub

Public Function IsLoaded(formName As String) As Boolean
'@AssignedModule M_CommandBars
'@INCLUDE CLASS CVBECommandHandler
    Dim frm As Object
    For Each frm In VBA.Userforms
        If frm.Name = formName Then
            IsLoaded = True
            Exit Function
        End If
    Next frm
    IsLoaded = False
End Function

Sub openUValiationDropdown()
'@AssignedModule M_CommandBars
    Dim lngValType As Long
    On Error Resume Next
    lngValType = ActiveCell.Validation.Type
    On Error GoTo 0
    Select Case lngValType
        Case Is = 3
            uValidationDropdown.Show
        Case Else
            If IsLoaded("uValidationDropdown") Then
                On Error Resume Next
                Unload uValidationDropdown
                On Error GoTo 0
            End If
    End Select
End Sub

Sub TestProcedure()
    MsgBox "ok"
End Sub

'''''NOTES'''''''
'''''''''''''''''

'-----------
'Use combobox
'-----------
''call a sub through class events handler
''the sub to contain the following
'With Application.VBE.ActiveCodePane
'  Text = Application.VBE.CommandBars(mcToolBar).Controls(mcInsertList).Text
'  .GetSelection StartLine, StartColumn, EndLine, EndColumn
'  .CodeModule.InsertLines StartLine, Text
'  .SetSelection StartLine, 1, StartLine, 1
'End With


