Attribute VB_Name = "z_uRadialMenu"

Option Explicit

Public RadialMenuInstance As Long

'* Modified   : Date and Time       Author              Description
'* Updated    : 21-08-2023 12:18    Alex                (z_uRadialMenu.bas > NewRadialMenu)

Sub NewRadialMenu(Optional instanceCaller As MSForms.Label)
'@LastModified 2308211218
    Dim i As Long
    Dim u As z_RadialMenu
    
        Set u = New z_RadialMenu
        u.Show
        If Not instanceCaller Is Nothing Then
            Dim myParent As Object
            Set myParent = instanceCaller.Parent
            Do Until TypeOf myParent Is MSForms.UserForm And Not LCase(TypeName(myParent)) = "frame"
                Set myParent = myParent.Parent
            Loop
            u.Left = myParent.Left
            u.Top = myParent.Top
            u.Controls("iMain").Left = instanceCaller.Left
            u.Controls("iMain").Top = instanceCaller.Top
        End If
        u.Tag = RadialMenuCount
        u.createRadialMenu _
                        Array(6, 6), _
                        0, _
                        True, _
                        "redCircleGolderBorder"
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 21-08-2023 12:18    Alex                (z_uRadialMenu.bas > RadialMenuCount)

Public Function RadialMenuCount()
'@LastModified 2308211218
    Dim frm As Object
    Dim counter As Long
    counter = 0
    For Each frm In VBA.Userforms
        If frm.Name = "z_RadialMenu" Then
            counter = counter + 1
        End If
    Next
    RadialMenuCount = counter
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 21-08-2023 12:18    Alex                (z_uRadialMenu.bas > RadialMenuUnload)

Public Sub RadialMenuUnload()
'@LastModified 2308211218
    '@INCLUDE USERFORM z_RadialMenu
    '@AssignedModule z_uRadialMenu
'    Unload z_RadialMenu

    Dim frm As Object
    For Each frm In VBA.Userforms
        If frm.Name = "z_RadialMenu" And frm.Tag = RadialMenuInstance Then
            Unload frm
            Exit For
        End If
    Next frm
End Sub

Public Sub RefreshRadialMenuImages()
    '@INCLUDE PROCEDURE FileExists
    '@INCLUDE PROCEDURE IniReadKey
    '@INCLUDE USERFORM z_RadialMenu
    '@AssignedModule z_uRadialMenu
    Dim picForMenuExpanded As String
    picForMenuExpanded = IniReadKey(ThisWorkbook.path & "\Lib\config\radialmenu.ini", "Settings", "picForMenuExpanded")
    If picForMenuExpanded = "" Then picForMenuExpanded = "WhiteCircle"
    picForMenuExpanded = ThisWorkbook.path & "\Lib\img\" & picForMenuExpanded & ".jpg"
    
    Dim picForMenuCollapsed As String
    picForMenuCollapsed = IniReadKey(ThisWorkbook.path & "\Lib\config\radialmenu.ini", "Settings", "picForMenuCollapsed")
    If picForMenuCollapsed = "" Then picForMenuCollapsed = "PurpleCircle"
    picForMenuCollapsed = ThisWorkbook.path & "\Lib\img\" & picForMenuCollapsed & ".jpg"
    
    Dim picForItemsDefault As String
    picForItemsDefault = IniReadKey(ThisWorkbook.path & "\Lib\config\radialmenu.ini", "Settings", "picForItemsDefault")
    If picForItemsDefault = "" Then picForItemsDefault = "BlackCircle"
    picForItemsDefault = ThisWorkbook.path & "\Lib\img\" & picForItemsDefault & ".jpg"
    
    If Not FileExists(picForMenuCollapsed) Or Not FileExists(picForMenuExpanded) Or Not FileExists(picForItemsDefault) Then
        MsgBox "Could not fild image files"
        Stop
    End If
    Dim control     As MSForms.control
    For Each control In z_RadialMenu.Controls
        If control.Name = "iMain" Then
            If control.Tag = "collapsed" Then
                control.Picture = LoadPicture(picForMenuCollapsed)
                If InStr(1, picForMenuCollapsed, "White", vbTextCompare) > 0 Then
                    control.ForeColor = vbBlack
                ElseIf InStr(1, picForMenuCollapsed, "Black", vbTextCompare) > 0 Then
                    control.ForeColor = vbWhite
                End If
            Else
                control.Picture = LoadPicture(picForMenuExpanded)
                If InStr(1, picForMenuExpanded, "White", vbTextCompare) > 0 Then
                    control.ForeColor = vbBlack
                ElseIf InStr(1, picForMenuExpanded, "Black", vbTextCompare) > 0 Then
                    control.ForeColor = vbWhite
                End If
            End If
        ElseIf control.Tag <> "custom" Then
            control.Picture = LoadPicture(picForItemsDefault)
            If InStr(1, picForItemsDefault, "White", vbTextCompare) > 0 Then
                control.ForeColor = vbBlack
            ElseIf InStr(1, picForItemsDefault, "Black", vbTextCompare) > 0 Then
                control.ForeColor = vbWhite
            End If
        End If
        If InStr(1, picForMenuCollapsed, "White", vbTextCompare) Then control.ForeColor = vbWhite
    Next
'    z_RadialMenu.Repaint
End Sub
Sub setRadialMenuTheme(picForMenuCollapsed As String, picForMenuExpanded As String, picForItemsDefault As String)
    '@INCLUDE PROCEDURE IniWrite
    '@INCLUDE PROCEDURE RefreshRadialMenuImages
    '@AssignedModule z_uRadialMenu
    IniWrite ThisWorkbook.path & "\Lib\config\radialmenu.ini", "Settings", "picForMenuExpanded", picForMenuExpanded
    IniWrite ThisWorkbook.path & "\Lib\config\radialmenu.ini", "Settings", "picForMenuCollapsed", picForMenuCollapsed
    IniWrite ThisWorkbook.path & "\Lib\config\radialmenu.ini", "Settings", "picForItemsDefault", picForItemsDefault
    RefreshRadialMenuImages
End Sub
Sub SwitchThemeHexagon()
    '@INCLUDE PROCEDURE setRadialMenuTheme
    '@AssignedModule z_uRadialMenu
    setRadialMenuTheme "HexagonBlack", "HexagonWhite", "HexagonBlack"
End Sub
Sub SwitchThemeCircle()
    '@INCLUDE PROCEDURE setRadialMenuTheme
    '@AssignedModule z_uRadialMenu
    setRadialMenuTheme "PurpleCircle", "WhiteCircle", "BlackCircle"
End Sub

