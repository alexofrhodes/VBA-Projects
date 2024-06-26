VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aUserform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

Private oForm As Object

Rem Transparent Declarations
Rem Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
#If VBA7 Then
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#End If

Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME As Long = &H1

Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private m_sngDownX As Single
Private m_sngDownY As Single

Rem Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Rem Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Rem Parent Declarations
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const FORMAT_MESSAGE_TEXT_LEN = 160
Private Const MAX_PATH = 260
Private Const GWL_HWNDPARENT As Long = -8
Private Const GW_OWNER = 4

#If VBA7 Then
    Private Declare PtrSafe Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
#End If

Private VBEditorHWnd As Long
Private ApplicationHWnd As Long
Private ExcelDeskHWnd As Long
Private ActiveWindowHWnd As Long
Private UserFormHWnd As Long
Private WindowsDesktopHWnd As Long
Private Const GA_ROOT As Long = 2
Private Const GA_ROOTOWNER As Long = 3
Private Const GA_PARENT As Long = 1

#If VBA7 Then
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare PtrSafe Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
    Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
    Private Declare PtrSafe Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Private Const C_EXCEL_APP_WINDOWCLASS = "XLMAIN"
Private Const C_EXCEL_DESK_WINDOWCLASS = "XLDESK"
Private Const C_EXCEL_WINDOW_WINDOWCLASS = "EXCEL7"
Private Const USERFORM_WINDOW_CLASS = "ThunderDFrame"
Private Const C_VBA_USERFORM_WINDOWCLASS = "ThunderDFrame"

Rem Window position and more
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

#If VBA7 Then
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hwndInsertAfter As LongPtr, ByVal x As LongPtr, ByVal Y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal uFlags As LongPtr) As Long
#Else
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal uFlags As LongPtr) As Long
#End If

Rem ---
#If VBA7 Then
    Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hwndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long
#Else
    Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
#End If

Rem Closeby
Private Enum CloseBy
    User = 0
    Code = 1
    WindowsOS = 2
    TaskManager = 3
End Enum


Private Const Black As Long = &H80000012
Private Const Red As Long = &HFF&

Rem other
#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Rem userform hwnd
#If Win64 Then
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As Long) As Long
#End If

Rem ___Resizeable___
#If VBA7 Then
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
#Else
    Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#End If
Private mdWidth As Double
Private mdHeight As Double
Private Const WS_THICKFRAME As Long = &H40000
Private Const SW_SHOW As Long = 5



#If VBA7 Then
    Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
    Private Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If


Public Sub Transition(ParamArray Elements() As Variant)
Rem By Robert Todaer
'__usage__
'with aUserform
'    .Transition .Effect(Box, "Top", Me.InsideHeight - Box.Height, 1000) _
                , .Effect(box2, "Top", 0, 100) _
                , .Effect(GoButton, "fontsize", 12, 1000) _
                , .Effect(Me, "Top", 20, 2000)
'End with

    If IsArray(Elements(LBound(Elements, 1))) Then
        Dim Temp As Variant
        Temp = Elements(LBound(Elements, 1))
        Elements = Temp
    End If
    Dim Form As MSForms.UserForm
    Set Form = Elements(LBound(Elements, 1))("form")
    MicroTimer True
    Do
        Dim Index As Integer
        For Index = LBound(Elements, 1) To UBound(Elements, 1)
            IncRementElement Elements(Index), MicroTimer
        Next Index
        Sleep 40
        Form.Repaint
    Loop Until AllTransitionsComplete(CVar(Elements))
End Sub

Public Function Effect(obj As Object, Property As String, Destination As Double, MilSecs As Double) As scripting.Dictionary
    Dim Temp As New scripting.Dictionary
    Set Temp("obj") = obj
    Temp("property") = Property
    Temp("startValue") = CallByName(obj, Property, VbGet)
    Temp("destination") = Destination
    Temp("travel") = Destination - Temp("startValue")
    Temp("milSec") = MilSecs
    Temp("complete") = False
    On Error GoTo Catch:
    Set Temp("form") = obj.Parent
    Set Effect = Temp
    Exit Function
Catch:
    Set Temp("form") = obj
    Resume Next
End Function

Private Function MicroTimer(Optional StartTime As Boolean = False) As Double
    Static dTime As Double
    Dim cyTicks1 As Currency
    Dim cyTicks2 As Currency
    Static cyFrequency As Currency
    MicroTimer = 0
    If cyFrequency = 0 Then getFrequency cyFrequency
    getTickCount cyTicks1
    getTickCount cyTicks2
    If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
    If cyFrequency Then MicroTimer = cyTicks2 / cyFrequency
    If StartTime = True Then
        dTime = MicroTimer
        MicroTimer = 0
    Else
        MicroTimer = (MicroTimer - dTime) * 1000
    End If
End Function

Private Function AllTransitionsComplete(Elements As Variant) As Boolean
    '@INCLUDE TransitionComplete
    Dim el As Object
    Dim Index As Integer
    For Index = LBound(Elements, 1) To UBound(Elements, 1)
        Set el = Elements(Index)
        If Not TransitionComplete(el) Then
            AllTransitionsComplete = False
            Exit Function
        End If
    Next Index
    AllTransitionsComplete = True
End Function

Private Function TransitionComplete(ByVal el As scripting.Dictionary) As Boolean
    If Math.Round(el("destination")) = Math.Round(CallByName(el("obj"), el("property"), VbGet)) Then
        TransitionComplete = True
    End If
End Function

Private Function IncRementElement(ByVal el As scripting.Dictionary, CurrentTime As Double) As Boolean
    '@INCLUDE TransitionComplete
    '@INCLUDE easeInAndOut
    Dim IncRementValue As Double
    Dim CurrentValue As Double
    If TransitionComplete(el) Then
        Exit Function
    End If
    Dim o As Object
    Dim P As String
    Dim Value As Double
    Dim d As Double
    IncRementValue = easeInAndOut(CurrentTime, el("startValue"), el("travel"), el("milSec"))
    If el("travel") < 0 Then
        If Math.Round(IncRementValue, 4) < el("destination") Then
            CallByName el("obj"), el("property"), VbLet, el("destination")
        Else
            CallByName el("obj"), el("property"), VbLet, IncRementValue
        End If
    Else
        If Math.Round(IncRementValue, 4) > el("destination") Then
            CallByName el("obj"), el("property"), VbLet, el("destination")
        Else
            CallByName el("obj"), el("property"), VbLet, IncRementValue
        End If
    End If
End Function

Private Function easeInAndOut(ByVal t As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Double
    d = d / 2
    t = t / d
    If (t < 1) Then
        easeInAndOut = c / 2 * t * t * t + b
    Else
        t = t - 2
        easeInAndOut = c / 2 * (t * t * t + 2) + b
    End If
End Function



Public Sub FormResize()
'
'from STEPHEN BULLEN's USERFORM RESIZER CLASS
'
' _______IN USERFORM_________________
'
' Dim moResizer As New CFormResizer
' Private Sub UserForm_Activate()
'     Set moResizer.Form = Me
' End Sub
' Private Sub UserForm_Resize()
'     moResizer.FormResize
' End Sub
'
' _______Instructions_______
'
' To specify which control(s) to resize (and how), you set the control's .Tag property at design time to
' indicate that the control's top, left, width and height should be adjusted as the form's size changes.
'
' Use the letters t, l, w and h in any order (or not at all) to state that the property should change as the form
' is resized.  Follow the property by a decimal to indicate that the control should change by a percentage of the
' form's change.
'
' For example:
'   hw           Sets the control's height and width to change with the form (e.g. if there's a single list box on the form)
'   tl           Sets the contol's top and left to change in line with the form (e.g. to keep it in the bottom-right corner)
'   w0.5         Sets the control's width to change by 0.5 that of the form's width change
'   w0.5l0.5     Sets the control's width and position to change by 0.5 that of the form's width change

    Dim dWidthAdj As Double, dHeightAdj As Double, sTag As String
    Dim oCtl As MSForms.control
    If mdWidth = 0 Then mdWidth = oForm.Width
    If mdHeight = 0 Then mdHeight = oForm.Height
    dWidthAdj = oForm.Width - mdWidth
    dHeightAdj = oForm.Height - mdHeight
    For Each oCtl In oForm.Controls
        If TypeName(oCtl) = "ListView4" Then Stop
        With oCtl
            sTag = UCase(.Tag)
            If InStr(1, sTag, "L", vbBinaryCompare) Then
                If .Left + dWidthAdj <= 0 Then oForm.Width = mdWidth
            End If
            If InStr(1, sTag, "W", vbBinaryCompare) Then
                If .Width + dWidthAdj <= 0 Then oForm.Width = mdWidth
            End If
            If InStr(1, sTag, "T", vbBinaryCompare) Then
                If .Top + dHeightAdj <= 0 Then oForm.Height = mdHeight
            End If
            If InStr(1, sTag, "H", vbBinaryCompare) Then
                If .Height + dHeightAdj <= 0 Then oForm.Height = mdHeight
            End If
        End With
    Next
    dWidthAdj = oForm.Width - mdWidth
    dHeightAdj = oForm.Height - mdHeight
    For Each oCtl In oForm.Controls
        With oCtl
            sTag = UCase(.Tag)
            If InStr(1, sTag, "L", vbBinaryCompare) Then .Left = .Left + dWidthAdj * ResizeFactor(sTag, "L")
            If InStr(1, sTag, "T", vbBinaryCompare) Then .Top = .Top + dHeightAdj * ResizeFactor(sTag, "T")
            If InStr(1, sTag, "W", vbBinaryCompare) Then .Width = .Width + dWidthAdj * ResizeFactor(sTag, "W")
            If InStr(1, sTag, "H", vbBinaryCompare) Then .Height = .Height + dHeightAdj * ResizeFactor(sTag, "H")
            
        End With
    Next
    mdWidth = oForm.Width
    mdHeight = oForm.Height
End Sub

Private Function ResizeFactor(sTag As String, sChange As String)
    Dim i As Integer, d As Double
    i = InStr(1, sTag, sChange, vbBinaryCompare)
    If i > 0 Then
        d = val(Mid$(sTag, i + 1))
        If d = 0 Then d = 1
    End If
    ResizeFactor = d
End Function

Public Sub Resizable()
    Dim hWndForm As Long, iStyle As Long
    If val(Application.version) < 9 Then
        hWndForm = FindWindow("ThunderXFrame", oForm.Caption)
    Else
        hWndForm = FindWindow("ThunderDFrame", oForm.Caption)
    End If
    iStyle = GetWindowLong(hWndForm, GWL_STYLE)
    iStyle = iStyle Or WS_THICKFRAME
    SetWindowLong hWndForm, GWL_STYLE, iStyle
    ShowWindow hWndForm, SW_SHOW
    DrawMenuBar hWndForm
    SetFocus hWndForm
    mdWidth = oForm.Width
    mdHeight = oForm.Height
End Sub



Public Function Init(Form As Object) As aUserform
    Set oForm = Form
    Set Init = Me
End Function

Private Sub Class_Initialize()
'do something
End Sub



' #If VBA7 Then
'    Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As LongPtr) As Long
'#Else
'    Private Declare Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As Long) As Long
'#End If
'
'
'#If VBA7 Then
'    Function GetUserformHwnd(ByVal ufmTarget As MSForms.UserForm) As LongPtr
'#Else
'    Function GetUserformHwnd(ByVal ufmTarget As MSForms.UserForm) As Long
'#End If
'    WindowFromAccessibleObject ufmTarget, GetUserformHwnd
'End Function

Private Function GetUserformHwnd(ByVal ufmTarget As MSForms.UserForm) As Long
     IUnknown_GetWindow ufmTarget, VarPtr(GetUserformHwnd)
End Function

'Function GetUserformHwnd(Form As Object)
'    GetUserformHwnd = FindWindow(vbNullString, Form.Caption)
'End Function

Public Sub Transparent(Optional Color As Variant)
    '@INCLUDE MakeFormBorderless
    Dim formhandle As Long
    Dim bytOpacity As Byte
    formhandle = CLng(FindWindow(vbNullString, oForm.Caption))
    If IsMissing(Color) Then Color = vbWhite
    bytOpacity = 100
    SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
    oForm.BackColor = Color
    SetLayeredWindowAttributes formhandle, Color, bytOpacity, LWA_COLORKEY
End Sub

Public Sub Borderless()
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = CLng(FindWindow(vbNullString, oForm.Caption))
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, GWL_EXSTYLE, lngWindow
    DrawMenuBar lFrmHdl
End Sub

Public Sub OnTop()
    Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"
    Dim Ret As Long
    Dim formHWnd As Long
    formHWnd = CLng(FindWindow(C_VBA6_USERFORM_CLASSNAME, oForm.Caption))
    If formHWnd = 0 Then
        Debug.Print Err.LastDllError
    End If
    Ret = SetWindowPos(formHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    If Ret = 0 Then
        Debug.Print Err.LastDllError
    End If
End Sub

Sub ParentIsVBE()
    '@INCLUDE DisplayErrorText
    Dim GivenFormCaption As String
        GivenFormCaption = oForm.Caption
    #If VBA7 Then
        Dim VBEWindowPointer As LongPtr
        Dim UserFormWindowPointer As LongPtr
        Dim ReturnOfSetParentAPI As LongPtr
    #Else
        Dim VBEWindowPointer As Long
        Dim UserFormWindowPointer As Long
        Dim ReturnOfSetParentAPI As Long
    #End If
    Dim ErrorNumber As Long
    VBEWindowPointer = Application.VBE.MainWindow.hWnd
    UserFormWindowPointer = FindWindow(lpClassName:=USERFORM_WINDOW_CLASS, lpWindowName:=GivenFormCaption)
    Const ERROR_NUMBER_FOR_SETPARENT_API = 0
    ReturnOfSetParentAPI = SetParent(hwndChild:=UserFormWindowPointer, hWndNewParent:=VBEWindowPointer)
    If ReturnOfSetParentAPI = ERROR_NUMBER_FOR_SETPARENT_API Then
        ErrorNumber = Err.LastDllError
        DisplayErrorText "Error With SetParent", ErrorNumber
    Else
        Debug.Print GivenFormCaption & " is child of VBE Window."
    End If
    SetForegroundWindow UserFormWindowPointer
    Application.VBE.MainWindow.SetFocus
End Sub

Private Sub DisplayErrorText(Context As String, ErrNum As Long)
    Rem  Displays a standard error message box. For this
    Rem  procedure, ErrNum should be the number returned
    Rem  by the GetLastError API function or the value
    Rem  of Err.LastDllError. It is NOT the number
    Rem  returned by Err.Number.
    '@INCLUDE GetSystemErrorMessageText
    Dim ErrText As String
    ErrText = GetSystemErrorMessageText(ErrNum)
    Debug.Print Context & vbCrLf & _
    "Error Number: " & CStr(ErrNum) & vbCrLf & _
    "Error Text:   " & ErrText, vbOKOnly
End Sub

Private Function GetSystemErrorMessageText(ErrorNumber As Long) As String
    Rem  This function gets the system error message text that corresponds to the error code returned by the
    Rem  GetLastError API function or the Err.LastDllError property. It may be used ONLY for these error codes.
    Rem  These are NOT the error numbers returned by Err.Number (for these errors, use Err.Description to get the description of the message).
    Rem  The error number MUST be the value returned by GetLastError or Err.LastDLLError.
    Rem
    Rem  In general, you should use Err.LastDllError rather than GetLastError because under some circumstances the value of
    Rem  GetLastError will be reset to 0 before the value is returned to VB. Err.LastDllError will always reliably return
    Rem  the last error number raised in a DLL.
    Dim ErrorText As String
    Dim ErrorTextLength As Long
    Dim FormatMessageResult As Long
    Dim LanguageID As Long
    LanguageID = 0&
    ErrorText = String$(FORMAT_MESSAGE_TEXT_LEN, " ")
    ErrorTextLength = Len(ErrorText)
    FormatMessageResult = 0&
    #If VBA7 Then
        Dim FormatMessageAPILastArgument As LongPtr
        FormatMessageAPILastArgument = 0
    #Else
        Dim FormatMessageAPILastArgument As Long
        FormatMessageAPILastArgument = 0
    #End If
    FormatMessageResult = FormatMessage( _
    dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
    lpSource:=0&, _
    dwMessageId:=ErrorNumber, _
    dwLanguageId:=0&, _
    lpBuffer:=ErrorText, _
    nSize:=ErrorTextLength, _
    Arguments:=FormatMessageAPILastArgument)
    If FormatMessageResult > 0 Then
        ErrorText = TrimToNull(ErrorText)
        GetSystemErrorMessageText = ErrorText
    Else
        GetSystemErrorMessageText = "NO ERROR DESCRIPTION AVAILABLE"
    End If
End Function

Private Function TrimToNull(TEXT As String) As String
    Rem  Returns all the text in Text to the left of the vbNullChar
    Dim NullCharIndex As Integer
    NullCharIndex = InStr(1, TEXT, vbNullChar, vbTextCompare)
    If NullCharIndex > 0 Then
        TrimToNull = Left(TEXT, NullCharIndex - 1)
    Else
        TrimToNull = TEXT
    End If
End Function

Public Sub MinimizeButton()
    Dim cap     As String:  cap = oForm.Caption
    Dim hWnd    As Long:    hWnd = FindWindowA(vbNullString, cap)
    Dim exLong  As Long:    exLong = GetWindowLongA(hWnd, -16)
    
    If (exLong And &H20000) = 0 Then SetWindowLongA hWnd, -16, exLong Or &H20000
End Sub

Public Sub SaveOptions( _
                Optional includeCheckbox As Boolean = True, _
                Optional includeOptionButton As Boolean = True, _
                Optional includeTextBox As Boolean = True, _
                Optional includeListbox As Boolean = True, _
                Optional includeToggleButton As Boolean = True, _
                Optional includeCombobox As Boolean = True)
    Dim WS As Worksheet
    Set WS = CreateOrSetSheet(oForm.Name & "_Settings", ThisWorkbook)
    WS.Cells.clear
    Dim coll As New Collection
    Dim cell As Range
    Set cell = WS.Cells(1, 1)
    Dim c As MSForms.control
    For Each c In oForm.Controls
        If TypeName(c) Like "CheckBox" Then
            If Not includeCheckbox Then GoTo SKIP
        ElseIf TypeName(c) Like "OptionButton" Then
            If Not includeOptionButton Then GoTo SKIP
        ElseIf TypeName(c) Like "TextBox" Then
            If Not includeTextBox Then GoTo SKIP
        ElseIf TypeName(c) = "ListBox" Then
            If Not includeListbox Then GoTo SKIP
        ElseIf TypeName(c) Like "ToggleButton" Then
            If Not includeToggleButton Then GoTo SKIP
        ElseIf TypeName(c) Like "ComboBox" Then
            If Not includeCombobox Then GoTo SKIP
        Else
            GoTo SKIP
        End If
        cell = c.Name
        Select Case TypeName(c)
        Case "TextBox", "CheckBox", "OptionButton", "ToggleButton", "ComboBox"
            cell.Offset(0, 1) = c.Value
        Case "ListBox"
            Set coll = aListBox.Init(c).selectedIndexes
            If coll.Count > 0 Then
                cell.Offset(0, 1) = aCollection.Init(coll).ToString(",")
            Else
                cell.Offset(0, 1) = -1
            End If
        End Select
        Set cell = cell.Offset(1, 0)
SKIP:
    Next
End Sub

Public Sub LoadOptions(Optional ExcludeThese As Variant)
'ExcludeThese:=array("ListBox1","TextBox1")     '<--control names
    Dim WS As Worksheet
    Set WS = CreateOrSetSheet(oForm.Name & "_Settings", ThisWorkbook)
    If WS.Range("A1") = "" Then Exit Sub
    Dim cell As Range
    Set cell = WS.Cells(1, 1)
    Dim c As MSForms.control
    Dim v
    On Error Resume Next
    Do While cell <> ""
        Set c = oForm.Controls(cell.TEXT)
        If Not TypeName(c) = "Nothing " Then
            If Not ArrayContains(cell, ExcludeThese) Then
                Select Case TypeName(c)
                Case "TextBox", "CheckBox", "OptionButton", "ToggleButton", "ComboBox"
                    c.Value = cell.Offset(0, 1)
                Case "ListBox"
                    If InStr(1, cell.Offset(0, 1), ",") > 0 Then
                        aListBox.Init(c).SelectItems Split(cell.Offset(0, 1), ","), True
                    Else
                        c.Selected(CInt(cell.Offset(0, 1))) = True
                    End If
                End Select
            End If
        End If
        Set cell = cell.Offset(1, 0)
    Loop
End Sub

Public Sub SavePosition()
    SaveSetting "My Settings Folder", oForm.Name, "Left Position", oForm.Left
    SaveSetting "My Settings Folder", oForm.Name, "Top Position", oForm.Top
End Sub

Public Sub LoadPosition()
    If GetSetting("My Settings Folder", oForm.Name, "Left Position") = "" _
        And GetSetting("My Settings Folder", oForm.Name, "Top Position") = "" Then
        oForm.StartUpPosition = 1
    Else
        oForm.Left = GetSetting("My Settings Folder", oForm.Name, "Left Position")
        oForm.Top = GetSetting("My Settings Folder", oForm.Name, "Top Position")
    End If
End Sub

Public Sub ResizeToFitControls()
    oForm.Width = 0
    oForm.Height = 0
    Dim ctr As MSForms.control
    Dim myWidth:    myWidth = oForm.InsideWidth
    Dim myHeight:   myHeight = oForm.InsideHeight
    For Each ctr In oForm.Controls
        If ctr.Visible = True Then
            If ctr.Left + ctr.Width > myWidth Then myWidth = ctr.Left + ctr.Width
            If ctr.Top + ctr.Height > myHeight Then myHeight = ctr.Top + ctr.Height
        End If
    Next
    oForm.Width = myWidth + oForm.Width - oForm.InsideWidth + 50
    oForm.Height = myHeight + oForm.Height - oForm.InsideHeight + 10
End Sub
