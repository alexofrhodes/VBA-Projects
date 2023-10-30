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
Attribute Emitter.VB_VarHelpID = -1

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Sub Emitter_MouseMove(control As Object, Shift As Integer, X As Single, Y As Single)
    '@INCLUDE CLASS EventListenerEmitter
    '@AssignedModule z_RadialMenu
    
    Select Case control.Name
    Case "iMain", Me.Name, "background": Exit Sub
    Case Else: control.ZOrder (0)
    End Select
End Sub

Private Sub UserForm_Initialize()
    '@INCLUDE CLASS aUserform
    '@AssignedModule z_RadialMenu
    
    
    Me.Top = 0
    Me.Left = 0
    
    Dim appWidth As Long
    Dim appHeight As Long
    Application.WindowState = xlMaximized
    appWidth = Application.Width
    appHeight = Application.Height
    
    Me.Width = appWidth
    Me.Height = appHeight

    iMain.Left = Me.Width / 2 - iMain.Width / 2
    iMain.Top = Me.Height / 2 - iMain.Height / 2
    toggleMainImage

    Me.BackColor = vbMagenta
    With aUserform.Init(Me)
        .TRANSPARENT (vbMagenta)
        .Borderless
    End With
    
End Sub

Sub toggleMainImage()
    '@INCLUDE PROCEDURE FileExists
    '@INCLUDE PROCEDURE IniReadKey
    '@INCLUDE DECLARATION black
    '@AssignedModule z_RadialMenu
    
    
    Dim picForMenuExpanded As String
    picForMenuExpanded = IniReadKey(ThisWorkbook.path & "\Lib\config\radialmenu.ini", "Settings", "picForMenuExpanded", "WhiteCircle")
    picForMenuExpanded = ThisWorkbook.path & "\Lib\img\" & picForMenuExpanded & ".jpg"
    
    Dim picForMenuCollapsed As String
    picForMenuCollapsed = IniReadKey(ThisWorkbook.path & "\Lib\config\radialmenu.ini", "Settings", "picForMenuCollapsed", "PurpleCircle")
    picForMenuCollapsed = ThisWorkbook.path & "\Lib\img\" & picForMenuCollapsed & ".jpg"
    
    If Not FileExists(picForMenuCollapsed) Or Not FileExists(picForMenuExpanded) Then
        MsgBox "Could not fild image files"
        Stop
    End If
    Dim TargetFile  As String
    With iMain
        If .Picture Is Nothing Then
            .Picture = LoadPicture(picForMenuCollapsed)
            .Tag = "collapsed"
            TargetFile = picForMenuCollapsed
        ElseIf .Tag <> "collapsed" Then
            .Picture = LoadPicture(picForMenuCollapsed)
            .Tag = "collapsed"
            TargetFile = picForMenuCollapsed
        Else
            .Picture = LoadPicture(picForMenuExpanded)
            .Tag = "expanded"
            TargetFile = picForMenuExpanded
        End If
        .BorderStyle = fmBorderStyleNone
        .Width = 48
        .Height = 48
        .BackStyle = fmBackStyleTransparent
    End With
    Dim fileName As String
    fileName = Mid(TargetFile, InStrRev(TargetFile, "\"))
    If InStr(1, fileName, "white", vbTextCompare) > 0 Or InStr(1, fileName, "light", vbTextCompare) > 0 Then
        iMain.ForeColor = vbBlack
    ElseIf InStr(1, fileName, "black", vbTextCompare) > 0 Or InStr(1, fileName, "dark", vbTextCompare) > 0 Then
        iMain.ForeColor = vbWhite
    End If
'    Me.Repaint
End Sub
Private Sub iMain_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'if left mouse is pressed
    '@INCLUDE PROCEDURE moverForm
    '@AssignedModule z_RadialMenu
    
    If Button = 1 Then moverForm Me, Me, Button
End Sub
Private Sub iMain_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'if right click
    '@INCLUDE PROCEDURE SwitchThemeHexagon
    '@INCLUDE PROCEDURE SwitchThemeCircle
    '@AssignedModule z_RadialMenu
    
    If Button = 2 Then
    
        If Me.Tag = "" Then Me.Tag = RadialMenuCount
        RadialMenuInstance = Me.Tag
    
        Const popupName = "RM_Menu"
        On Error Resume Next
        CommandBars(popupName).Delete
        On Error GoTo 0
        With CommandBars.Add(popupName, msoBarPopup, , True)
            With .Controls.Add(msoControlButton)
                If Me.Tag = "" Then Me.Tag = 1
                .OnAction = "RadialMenuUnload"
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
    
    ToggleMenu
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 21-08-2023 13:00    Alex                (z_RadialMenu.frm > ToggleMenu)

Public Sub ToggleMenu()
'@LastModified 2308211300
    '@INCLUDE CLASS EventListenerEmitter
    '@AssignedModule z_RadialMenu
    
    '@INCLUDE DECLARATION Emitter
    toggleMainImage
    Dim Visibility  As Boolean
    Dim control     As MSForms.control
    If Me.Controls.Count = 1 Then
        'createRadialMenu( numControls As Integer, _
                            startingAngle As Integer, _
                            ClockwisePlacement As Boolean, _
                            ParamArray Pairs() As Variant)
                            
                            'Pairs(1) = Target control name (layer-index eg 1-2)
                            'Pairs(2) = Caption
                            'Pairs(3) = Image       pass "" to keep default
        createRadialMenu _
                        Array(8, 6, 6), _
                        0, _
                        True, _
                        "blueCircle", _
                        "1-1", "Create Instance", ""          ', _
                        "1-1", "caption", "", _
                        "1-2", "2nd ctr", "purpleCircle", _
                        "1-3", ""       , "purpleCircle"
        Set Emitter = New EventListenerEmitter
        Emitter.AddEventListenerAll Me
        
    Else
        Visibility = Not Controls(1).Visible    'starts at 0 so iMain = 0th control
        If Me.Tag = "" Then Me.Tag = 1
'        If Me.Tag = 1 Then
            For Each control In Me.Controls
                If control.Name <> "iMain" Then control.Visible = Visibility
            Next
'        Else
'            Unload Me
'        End If
    End If
'    Me.Repaint
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 21-08-2023 13:01    Alex                (z_RadialMenu.frm > createRadialMenu)

Public Sub createRadialMenu(controlsPerLayer As Variant, _
                            startingAngle As Integer, _
                            ClockwisePlacement As Boolean, _
                            backGround As String, _
                            ParamArray Pairs() As Variant)
'@LastModified 2308211301
    '@INCLUDE PROCEDURE FileExists
    '@INCLUDE PROCEDURE IniReadKey
    '@AssignedModule z_RadialMenu
    
    'example call:
'    createRadialMenu _
                        Array(3, 6, _
                        0, _
                        True, _
                        "backgroundImage" _
                        "1-1", "caption", "", _
                        "1-2", "2nd ctr", "purpleCircle", _
                        "1-3", ""       , "purpleCircle"
    Dim control     As MSForms.Label
    Dim centerX     As Single
    Dim centerY     As Single
    centerX = iMain.Left + iMain.Width / 2    ' X-coordinate of center point
    centerY = iMain.Top + iMain.Height / 2    ' Y-coordinate of center point
    
    Dim numLayers As Integer
    numLayers = UBound(controlsPerLayer) + 1
    Dim layerIndex  As Integer
    For layerIndex = 1 To numLayers
        Dim angleStep As Single
        angleStep = (IIf(ClockwisePlacement, -1, 1) * 360) / controlsPerLayer(layerIndex - 1)
        Dim radius  As Single
        radius = CalculateRadius(layerIndex, CInt(controlsPerLayer(layerIndex - 1)))    ' Calculate the radius
        Dim angle   As Single
        angle = startingAngle  'where to start
        Dim DefaultImage As String
        DefaultImage = IniReadKey(ThisWorkbook.path & "\Lib\config\radialmenu.ini", "Settings", "picForItemsDefault")
        If DefaultImage = "" Then DefaultImage = "BlackCircle"
        Dim TargetFile As String
        TargetFile = ThisWorkbook.path & "\Lib\img\" & DefaultImage & ".jpg"
        Dim controlIndex As Integer
        For controlIndex = 1 To controlsPerLayer(layerIndex - 1)
            Set control = Me.Controls.Add("Forms.Label.1", "control-" & layerIndex & "-" & controlIndex)
            If FileExists(TargetFile) Then
                control.Picture = LoadPicture(TargetFile)
            End If
            control.MousePointer = fmMousePointerCustom
            control.MouseIcon = LoadPicture(ThisWorkbook.path & "\Lib\img\Hand Cursor Pointer.ico")
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
    If Not IsMissing(Pairs) Then
        Dim i           As Long
        For i = LBound(Pairs) To UBound(Pairs) Step 3
            Set control = Controls("control-" & Split(Pairs(i), "-")(0) & "-" & Split(Pairs(i), "-")(1))
            With control
                Dim controlCaption As String
                controlCaption = Pairs(i + 1)
                control.Caption = controlCaption
                Dim controlImage As String
                controlImage = Pairs(i + 2)
                TargetFile = ThisWorkbook.path & "\Lib\img\" & controlImage & ".jpg"
                If FileExists(TargetFile) Then
                    control.Picture = LoadPicture(TargetFile)
                    control.Tag = "custom"
                End If
            End With
        Next
    End If
     
    Dim maxDistance As Single
    ' Calculate the maximum distance from the center to any control
    For Each control In Me.Controls
        Dim distance As Single
'        to the center of the outer controls
'        distance = Sqr((control.Left + control.Width / 2 - centerX) ^ 2 + (control.Top + control.Height / 2 - centerY) ^ 2)

'        to the outer edge of the outer controls
        distance = Sqr((control.Left + control.Width - centerX) ^ 2 + (control.Top + control.Height - centerY) ^ 2)

        If distance > maxDistance Then
            maxDistance = distance
        End If
    Next
    
    ' Create the covering control
    Dim coverControl As control
    Set coverControl = Me.Controls.Add("Forms.Label.1", "background")
    
    coverControl.Left = centerX - maxDistance
    coverControl.Top = centerY - maxDistance
    coverControl.Width = 2 * maxDistance
    coverControl.Height = 2 * maxDistance
    coverControl.BackStyle = fmBackStyleTransparent
    
    'if you pass nullstring then it won't add a background
    TargetFile = ThisWorkbook.path & "\Lib\img\" & backGround & ".jpg"
    If FileExists(TargetFile) Then
        coverControl.Picture = LoadPicture(TargetFile)
        coverControl.PicturePosition = fmPicturePositionCenter
    End If
    
    coverControl.ZOrder (1) ' Bring it to the front
End Sub
Function CalculateRadius(layerIndex As Integer, controlsPerLayer As Integer) As Single
    ' Implement your logic to calculate the radius for each layer (distance from the center)
    '@AssignedModule z_RadialMenu
    
    CalculateRadius = layerIndex * 50    ' adjust as needed
End Function
Function DegToRad(ByVal degrees As Single) As Single
    '@AssignedModule z_RadialMenu
    
    DegToRad = degrees * WorksheetFunction.Pi / 180
End Function
Private Sub Emitter_Click(control As Object)
    '@INCLUDE PROCEDURE ProcedureExists
    '@INCLUDE CLASS EventListenerEmitter
    '@AssignedModule z_RadialMenu
    
    Select Case control.Name
    Case "iMain", Me.Name, "background": Exit Sub
    End Select

    'we can either use the control's name to get its layer & index
    Dim controlIndex As String: controlIndex = Split(control.Name, "-", 2)(1)
    Select Case controlIndex
        Case "1-1"
            NewRadialMenu control
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
                '
            End If
    End Select
End Sub
