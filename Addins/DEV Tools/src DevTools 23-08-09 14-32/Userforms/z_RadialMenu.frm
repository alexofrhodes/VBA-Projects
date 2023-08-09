VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} z_RadialMenu 
   Caption         =   "UserForm1"
   ClientHeight    =   2148
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3480
   OleObjectBlob   =   "z_RadialMenu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "z_RadialMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Emitter As EventListenerEmitter
Private Sub Emitter_MouseMove(control As Object, Shift As Integer, x As Single, y As Single)
'@INCLUDE CLASS EventListenerEmitter
'@AssignedModule z_RadialMenu
'@INCLUDE USERFORM z_RadialMenu
    If control.Name = "iMain" Or control.Name = Me.Name Then Exit Sub
    control.ZOrder (0)
End Sub
Private Sub UserForm_Initialize()
'@INCLUDE CLASS aUserform
'@AssignedModule z_RadialMenu
'@INCLUDE USERFORM z_RadialMenu
    Me.BackColor = vbMagenta
    With aUserform.Init(Me)
        .TRANSPARENT (vbMagenta)
        .Borderless
    End With
    Me.Height = 400
    Me.Width = 400
    iMain.Left = Me.Width / 2 - iMain.Width / 2
    iMain.Top = Me.Height / 2 - iMain.Height / 2
    toggleMainImage
End Sub
Sub toggleMainImage()
'@INCLUDE PROCEDURE FileExists
'@INCLUDE PROCEDURE IniReadKey
'@INCLUDE DECLARATION black
'@AssignedModule z_RadialMenu
'@INCLUDE USERFORM z_RadialMenu
    Dim picForMenuExpanded As String
    picForMenuExpanded = IniReadKey(ThisWorkbook.Path & "\config\radialmenu.ini", "Settings", "picForMenuExpanded")
    If picForMenuExpanded = "" Then picForMenuExpanded = "WhiteCircle"
    picForMenuExpanded = ThisWorkbook.Path & "\img\" & picForMenuExpanded & ".jpg"
    Dim picForMenuCollapsed As String
    picForMenuCollapsed = IniReadKey(ThisWorkbook.Path & "\config\radialmenu.ini", "Settings", "picForMenuCollapsed")
    If picForMenuCollapsed = "" Then picForMenuCollapsed = "PurpleCircle"
    picForMenuCollapsed = ThisWorkbook.Path & "\img\" & picForMenuCollapsed & ".jpg"
    If Not FileExists(picForMenuCollapsed) Or Not FileExists(picForMenuExpanded) Then
        MsgBox "Could not fild image files"
        Stop
    End If
    Dim targetFile As String
    With iMain
        If .Picture Is Nothing Then
            .Picture = LoadPicture(picForMenuCollapsed)
            .Tag = "collapsed"
            targetFile = picForMenuCollapsed
        ElseIf .Tag <> "collapsed" Then
            .Picture = LoadPicture(picForMenuCollapsed)
            .Tag = "collapsed"
            targetFile = picForMenuCollapsed
        Else
            .Picture = LoadPicture(picForMenuExpanded)
            .Tag = "expanded"
            targetFile = picForMenuExpanded
        End If
        .BorderStyle = fmBorderStyleNone
        .Width = 48
        .Height = 48
        .BackStyle = fmBackStyleTransparent
    End With
    If InStr(1, Mid(targetFile, InStrRev(targetFile, "\")), "white", vbTextCompare) > 0 Then
        iMain.ForeColor = vbBlack
    ElseIf InStr(1, Mid(targetFile, InStrRev(targetFile, "\")), "black", vbTextCompare) > 0 Then
        iMain.ForeColor = vbWhite
    End If
    Me.Repaint
End Sub
Private Sub iMain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    'if left mouse is pressed
'@INCLUDE PROCEDURE moverForm
'@AssignedModule z_RadialMenu
'@INCLUDE USERFORM z_RadialMenu
    If Button = 1 Then moverForm Me, Me, Button
End Sub
Private Sub iMain_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    'if right click
'@INCLUDE PROCEDURE SwitchThemeHexagon
'@INCLUDE PROCEDURE SwitchThemeCircle
'@AssignedModule z_RadialMenu
'@INCLUDE USERFORM z_RadialMenu
    If Button = 2 Then
        Const popupName = "RM_Menu"
        On Error Resume Next
        CommandBars(popupName).Delete
        On Error GoTo 0
        With CommandBars.Add(popupName, msoBarPopup, , True)
            With .Controls.Add(msoControlButton)
                .OnAction = "RadialMenuCommand_Unload"
                .Caption = "Unload Me"
            End With
            With .Controls.Add(msoControlPopup)
                .Caption = "Themes"
                .BeginGroup = True
                With .Controls.Add(msoControlButton)
                    .Caption = "Circle"
                    .OnAction = "SwitchThemeCircle"
                End With
                With .Controls.Add(msoControlButton)
                    .Caption = "Hexagon"
                    .OnAction = "SwitchThemeHexagon"
                End With
            End With
            .ShowPopup
        End With
    End If
End Sub
Private Sub iMain_Click()
'@AssignedModule z_RadialMenu
'@INCLUDE USERFORM z_RadialMenu
    ToggleMenu
End Sub
Sub ToggleMenu()
'@INCLUDE CLASS EventListenerEmitter
'@AssignedModule z_RadialMenu
'@INCLUDE USERFORM z_RadialMenu
'@INCLUDE DECLARATION Emitter
    toggleMainImage
    Dim Visibility As Boolean
    Dim control As MSForms.control
    If Me.Controls.count = 1 Then
        'createRadialMenu( numLayers As Integer, _
                           numControls As Integer, _
                           startingAngle As Integer, _
                           ClockwisePlacement As Boolean, _
                           ParamArray Pairs() As Variant)
        createRadialMenu 3, _
                         Array(8, 6, 6), _
                         0, _
                         True   ', _
                              "1-1", "caption", _
                              "1-3", "purpleCircle", _
                              "2-6", "hi"
        Set Emitter = New EventListenerEmitter
        Emitter.AddEventListenerAll Me
    Else
        Visibility = Not Controls(1).Visible 'starts at 0
        For Each control In Me.Controls
            If control.Name <> "iMain" Then control.Visible = Visibility
        Next
    End If
    Me.Repaint
End Sub
Sub createRadialMenu(numLayers As Integer, controlsPerLayer As Variant, startingAngle As Integer, ClockwisePlacement As Boolean, ParamArray Pairs() As Variant)
'@INCLUDE PROCEDURE FileExists
'@INCLUDE PROCEDURE IniReadKey
'@AssignedModule z_RadialMenu
'@INCLUDE USERFORM z_RadialMenu
    If UBound(controlsPerLayer) + 1 <> numLayers Then
        Dim dif: dif = UBound(controlsPerLayer) + 1 - numLayers
        dif = IIf(dif > 0, "-", "+") & Abs(dif)
        MsgBox "You are trying to create " & numLayers & " Layers" & vbNewLine & _
               "but passed controlsPerLayer = Array(" & Join(controlsPerLayer, ",") & ")" & vbNewLine & _
               "You should have " & dif & " argument(s) in your array."
        Stop
    End If
    Dim control As MSForms.Label
    Dim centerX As Single
    Dim centerY As Single
    centerX = iMain.Left + iMain.Width / 2 ' X-coordinate of center point
    centerY = iMain.Top + iMain.Height / 2 ' Y-coordinate of center point
    Dim layerIndex As Integer
    For layerIndex = 1 To numLayers
        Dim angleStep As Single
        angleStep = (IIf(ClockwisePlacement, -1, 1) * 360) / controlsPerLayer(layerIndex - 1)
        Dim radius As Single
        radius = CalculateRadius(layerIndex, CInt(controlsPerLayer(layerIndex - 1))) ' Calculate the radius
        Dim angle As Single
        angle = startingAngle  'where to start
        Dim DefaultImage As String
        DefaultImage = IniReadKey(ThisWorkbook.Path & "\config\radialmenu.ini", "Settings", "picForItemsDefault")
        If DefaultImage = "" Then DefaultImage = "BlackCircle"
        Dim targetFile As String
        targetFile = ThisWorkbook.Path & "\img\" & DefaultImage & ".jpg"
        Dim controlIndex As Integer
        For controlIndex = 1 To controlsPerLayer(layerIndex - 1)
            Set control = Me.Controls.Add("Forms.Label.1", "control-" & layerIndex & "-" & controlIndex)
            If FileExists(targetFile) Then
                control.Picture = LoadPicture(targetFile)
            End If
            control.MousePointer = fmMousePointerCustom
            control.MouseIcon = LoadPicture(ThisWorkbook.Path & "\img\Hand Cursor Pointer.ico")
            control.Caption = control.Name
            control.ForeColor = vbWhite
            control.BackStyle = fmBackStyleTransparent
            control.PicturePosition = fmPicturePositionCenter
            control.Height = iMain.Height
            control.Width = iMain.Height
            Dim xPosition As Single
            Dim yPosition As Single
            xPosition = centerX + radius * Cos(DegToRad(angle))
            yPosition = centerY - radius * Sin(DegToRad(angle))
            control.Left = xPosition - control.Width / 2
            control.Top = yPosition - control.Height / 2
            angle = angle + angleStep
        Next controlIndex
    Next layerIndex
    If IsMissing(Pairs) Then Exit Sub
    For Each control In Controls
        If control.Name <> "iMain" Then control.Visible = False
    Next
    Dim i As Long
    For i = LBound(Pairs) To UBound(Pairs) Step 2
        Set control = Controls("control-" & Split(Pairs(i), "-")(0) & "-" & Split(Pairs(i), "-")(1))
        With control
            .Visible = True
            Dim command As String: command = Pairs(i + 1)
            targetFile = ThisWorkbook.Path & "\img\" & command & ".jpg"
            If FileExists(targetFile) Then
                control.Picture = LoadPicture(targetFile)
                control.Tag = "custom"
            Else
                control.Caption = command
            End If
        End With
    Next
    For Each control In Controls
        If control.Name <> "iMain" Then If control.Visible = False Then Controls.Remove control.Name
    Next
End Sub
Function CalculateRadius(layerIndex As Integer, controlsPerLayer As Integer) As Single
    ' Implement your logic to calculate the radius for each layer (distance from the center)
'@AssignedModule z_RadialMenu
'@INCLUDE USERFORM z_RadialMenu
    CalculateRadius = layerIndex * 50 ' adjust as needed
End Function
Function DegToRad(ByVal degrees As Single) As Single
'@AssignedModule z_RadialMenu
'@INCLUDE USERFORM z_RadialMenu
    DegToRad = degrees * WorksheetFunction.Pi / 180
End Function
Private Sub Emitter_Click(control As Object)
'@INCLUDE PROCEDURE ProcedureExists
'@INCLUDE CLASS EventListenerEmitter
'@AssignedModule z_RadialMenu
'@INCLUDE USERFORM z_RadialMenu
    If control.Name = "iMain" Then Exit Sub
    'we can either use the control's name to get its layer & index
    Dim controlIndex As String:     controlIndex = Split(control.Name, "-", 2)(1)
    Select Case controlIndex
    Case "1-1"
        MsgBox "found me"
    Case "2-1"
    '... etc
    Case Else
    End Select
    'or we could use the caption
    Select Case control.Caption
        Case "Sample1"
        Case "Sample2"
        Case Else
            If ProcedureExists(ThisWorkbook, control.Caption) Then
                Application.Run control.Caption
            Else
            End If
    End Select
End Sub
