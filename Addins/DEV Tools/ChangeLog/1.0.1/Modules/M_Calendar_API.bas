Attribute VB_Name = "M_Calendar_API"
'-----------------------------------------------------------------------------------------------------------
#If VBA7 Then

Public Declare PtrSafe Sub ReleaseCapture Lib "user32" ()
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long
'----------------------------------------------------------------------------------------------------
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwNilliseconds As Long)
Public Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare PtrSafe Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare PtrSafe Function MoveJanela Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

#Else
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) As Long
'----------------------------------------------------------------------------------------------------
Public Declare Sub Sleep Lib "kernel32" (ByVal dwNilliseconds As Long)
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function MoveJanela Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

#End If

Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const GWL_STYLE As Long = (-16)
Private Const WS_CAPTION = 55000000
Private Const WS_EX_LAYERED = &H80000
'Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const IDC_HAND = 32649&
Public MeuForm      As Long
Public ESTILO       As Long
Public Const ESTILO_ATUAL As Long = (-16)
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Private lngPixelsX  As Long
Private lngPixelsY  As Long
Private strThunder  As String
Private blnCreate   As Boolean
Private lnghWnd_Form As Long
Private lnghWnd_Sub As Long

Private Const cstMask As Long = &H7FFFFFFF

Public PassiveMonth, PassiveDay, PassiveYear, PassiveDayTag, SelectedDay, SelectedDayTag

Public Function getMonth(iMonth As Integer, Optional language As String)
    '    Select Case languege
    '@AssignedModule M_Calendar_API

    Select Case iMonth Mod 12
        Case Is = 1, "-11"
            getMonth = "JANUARY"
        Case Is = 2, "-10"
            getMonth = "FEBRUARY"
        Case Is = 3, "-9"
            getMonth = "MARCH"
        Case Is = 4, "-8"
            getMonth = "APRIL"
        Case Is = 5, "-7"
            getMonth = "MAY"
        Case Is = 6, "-6"
            getMonth = "JUNE"
        Case Is = 7, "-5"
            getMonth = "JULY"
        Case Is = 8, "-4"
            getMonth = "AUGUST"
        Case Is = 9, "-3"
            getMonth = "SEPTEMBER"
        Case Is = 10, "-2"
            getMonth = "OCTOBER"
        Case Is = 11, "-1"
            getMonth = "NOVEMBER"
        Case Is = 0, 12
            getMonth = "DECEMBER"

    End Select
    '    End Select
End Function

Function HideTitleBarAndBorder(frm As Object)
    '@AssignedModule M_Calendar_API
    '@INCLUDE DECLARATION GWL_EXSTYLE
    '@INCLUDE DECLARATION GWL_STYLE
    '@INCLUDE DECLARATION WS_CAPTION
    '@INCLUDE DECLARATION WS_EX_DLGMODALFRAME
    '@INCLUDE DECLARATION DrawMenuBar
    '@INCLUDE DECLARATION FindWindow
    '@INCLUDE DECLARATION GetWindowLong
    '@INCLUDE DECLARATION SetWindowLong


    Dim lngWindow   As Long
    Dim lFrmHdl     As Long
    lFrmHdl = FindWindow(vbNullString, frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, GWL_EXSTYLE, lngWindow
    DrawMenuBar lFrmHdl

End Function

Function MakeUserformTransparent(frm As Object, colorKey As Variant, Optional color As Variant)
    '@AssignedModule M_Calendar_API
    '@INCLUDE DECLARATION GWL_EXSTYLE
    '@INCLUDE DECLARATION LWA_COLORKEY
    '@INCLUDE DECLARATION WS_EX_LAYERED
    '@INCLUDE DECLARATION FindWindow
    '@INCLUDE DECLARATION GetWindowLong
    '@INCLUDE DECLARATION SetLayeredWindowAttributes
    '@INCLUDE DECLARATION SetWindowLong
    LWA_COLORKEY = colorKey

    Dim formhandle  As Long
    Dim bytOpacity  As Byte

    formhandle = FindWindow(vbNullString, frm.Caption)
    If IsMissing(color) Then color = &H8000&    '//rgbWhite
    bytOpacity = 130

    SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED

    frm.BackColor = color
    SetLayeredWindowAttributes formhandle, color, bytOpacity, LWA_COLORKEY

End Function

Public Function MouseCursor(CursorType As Long)
    '@AssignedModule M_Calendar_API
    '@INCLUDE DECLARATION LoadCursorBynum
    '@INCLUDE DECLARATION SetCursor
    Dim lngRet      As Long
    lngRet = LoadCursorBynum(0&, CursorType)
    lngRet = SetCursor(lngRet)
End Function

Public Function MouseMoveIcon()
    '@AssignedModule M_Calendar_API
    '@INCLUDE PROCEDURE MouseCursor
    '@INCLUDE DECLARATION IDC_HAND
    Call MouseCursor(IDC_HAND)
End Function

Public Sub moverForm(Form As Object, obj As Object, Button As Integer)
    '@AssignedModule M_Calendar_API
    '@INCLUDE DECLARATION HTCAPTION
    '@INCLUDE DECLARATION WM_NCLBUTTONDOWN
    '@INCLUDE DECLARATION FindWindowA
    '@INCLUDE DECLARATION SendMessage
    '@INCLUDE DECLARATION ReleaseCapture
    Dim lngMyHandle As Long, lngCurrentStyle As Long, lngNewStyle As Long
    If val(Application.version) < 9 Then
        lngMyHandle = FindWindowA("ThunderXFrame", Form.Caption)
    Else
        lngMyHandle = FindWindowA("ThunderDFrame", Form.Caption)
    End If

    If Button = 1 Then
        With obj
            Call ReleaseCapture
            Call SendMessage(lngMyHandle, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End With
    End If
End Sub
Public Sub removeTudo(ObjForm As Object)
    '@AssignedModule M_Calendar_API
    '@INCLUDE DECLARATION ESTILO_ATUAL
    '@INCLUDE DECLARATION WS_CAPTION
    '@INCLUDE DECLARATION FindWindowA
    '@INCLUDE DECLARATION MoveJanela
    '@INCLUDE DECLARATION ESTILO
    '@INCLUDE DECLARATION MeuForm
    MeuForm = FindWindowA(vbNullString, ObjForm.Caption)
    ESTILO = ESTILO Or WS_CAPTION
    MoveJanela MeuForm, ESTILO_ATUAL, (ESTILO)
End Sub


