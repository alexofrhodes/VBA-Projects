Attribute VB_Name = "z_uRadialMenu"
Option Explicit
Sub RadialMenuCommand_Unload()
'@INCLUDE USERFORM z_RadialMenu
'@AssignedModule z_uRadialMenu
    Unload z_RadialMenu
End Sub
Public Sub RefreshRadialMenuImages()
'@INCLUDE PROCEDURE FileExists
'@INCLUDE PROCEDURE IniReadKey
'@INCLUDE USERFORM z_RadialMenu
'@AssignedModule z_uRadialMenu
    Dim picForMenuExpanded As String
    picForMenuExpanded = IniReadKey(ThisWorkbook.Path & "\config\radialmenu.ini", "Settings", "picForMenuExpanded")
    If picForMenuExpanded = "" Then picForMenuExpanded = "WhiteCircle"
    picForMenuExpanded = ThisWorkbook.Path & "\img\" & picForMenuExpanded & ".jpg"
    Dim picForMenuCollapsed As String
    picForMenuCollapsed = IniReadKey(ThisWorkbook.Path & "\config\radialmenu.ini", "Settings", "picForMenuCollapsed")
    If picForMenuCollapsed = "" Then picForMenuCollapsed = "PurpleCircle"
    picForMenuCollapsed = ThisWorkbook.Path & "\img\" & picForMenuCollapsed & ".jpg"
    Dim picForItemsDefault As String
    picForItemsDefault = IniReadKey(ThisWorkbook.Path & "\config\radialmenu.ini", "Settings", "picForItemsDefault")
    If picForItemsDefault = "" Then picForItemsDefault = "BlackCircle"
    picForItemsDefault = ThisWorkbook.Path & "\img\" & picForItemsDefault & ".jpg"
    If Not FileExists(picForMenuCollapsed) Or Not FileExists(picForMenuExpanded) Or Not FileExists(picForItemsDefault) Then
        MsgBox "Could not fild image files"
        Stop
    End If
    Dim control As MSForms.control
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
    z_RadialMenu.Repaint
End Sub
Sub setRadialMenuTheme(picForMenuCollapsed As String, picForMenuExpanded As String, picForItemsDefault As String)
'@INCLUDE PROCEDURE IniWrite
'@INCLUDE PROCEDURE RefreshRadialMenuImages
'@AssignedModule z_uRadialMenu
    IniWrite ThisWorkbook.Path & "\config\radialmenu.ini", "Settings", "picForMenuExpanded", picForMenuExpanded
    IniWrite ThisWorkbook.Path & "\config\radialmenu.ini", "Settings", "picForMenuCollapsed", picForMenuCollapsed
    IniWrite ThisWorkbook.Path & "\config\radialmenu.ini", "Settings", "picForItemsDefault", picForItemsDefault
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

