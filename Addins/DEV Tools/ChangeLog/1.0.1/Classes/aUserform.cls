VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aUserform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aUserform
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 04-05-2023 12:19    Alex
'* Modified   : 30-06-2023 10:01    Alex                modified save/load options/position to use ini
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Option Compare Text

Private oForm       As Object
Public WithEvents DockResize As UserForm
Attribute DockResize.VB_VarHelpID = -1
Public WithEvents lblResize As MSForms.Label
Attribute lblResize.VB_VarHelpID = -1
Public lblResizer As MSForms.Label

Private Const minHeight = 125
Private Const minWidth = 125
Private resizeEnabled As Boolean
Private mouseX As Double
Private mouseY As Double

Private ActiveWindowHWnd As Long
Private ApplicationHWnd As Long
Private Const C_ALPHA_FULL_OPAQUE As Byte = 255
Private Const C_ALPHA_FULL_TRANSPARENT As Byte = 0
Private Const C_EXCEL_APP_CLASSNAME = "XLMain"
Private Const C_EXCEL_APP_WINDOWCLASS = "XLMAIN"
Private Const C_EXCEL_DESK_CLASSNAME = "XLDesk"
Private Const C_EXCEL_DESK_WINDOWCLASS = "XLDESK"
Private Const C_EXCEL_WINDOW_CLASSNAME = "Excel7"
Private Const C_EXCEL_WINDOW_WINDOWCLASS = "EXCEL7"
Private Const C_USERFORM_CLASSNAME = "ThunderDFrame"
Private Const C_VBA_USERFORM_WINDOWCLASS = "ThunderDFrame"
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const FORMAT_MESSAGE_TEXT_LEN = 160
Private Const GA_PARENT As Long = 1
Private Const GA_ROOT As Long = 2
Private Const GA_ROOTOWNER As Long = 3
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)
Private Const GW_OWNER = 4
Private Const HWND_BOTTOM = 1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const LWA_ALPHA = &H2&
Private Const LWA_COLORKEY = &H1
Private Const MAX_PATH = 260
Private Const MF_BYPOSITION = &H400
Private Const MF_DISABLED = &H2&
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_REMOVE = &H1000
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SW_SHOW As Long = 5
Private Const USERFORM_WINDOW_CLASS = "ThunderDFrame"
Private Const WS_CAPTION = &HC00000
Private Const WS_DLGFRAME = &H400000
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_EX_LAYERED = &H80000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SIZEBOX = &H40000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const black As Long = &H80000012
Private Const red   As Long = &HFF&
Private ExcelDeskHWnd As Long
Private UserFormHWnd As Long
Private VBEditorHWnd As Long
Private WindowsDesktopHWnd As Long
Private m_sngDownX  As Single
Private m_sngDownY  As Single
Private mdHeight    As Double
Private mdWidth     As Double

Public Enum FORM_PARENT_WINDOW_TYPE
    FORM_PARENT_NONE = 0
    FORM_PARENT_APPLICATION = 1
    FORM_PARENT_WINDOW = 2
End Enum

Private Enum CloseBy
    User = 0
    Code = 1
    WindowsOS = 2
    TaskManager = 3
End Enum

#If VBA7 Then
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Private Declare PtrSafe Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare PtrSafe Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal uFlags As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As LongPtr) As Long
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long,ByVal wIDEnableItem As Long,ByVal wEnable As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long,ByVal bRevert As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long,ByVal nPosition As Long,ByVal wFlags As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long,ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal uFlags As LongPtr) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function WindowFromAccessibleObject Lib "oleacc" (ByVal pacc As IAccessible, phwnd As Long) As Long
Private Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If


Public Function ShowMaximizeButton(Optional HideButton As Boolean) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ShowMaximizeButton
    ' Displays (if HideButton is False) or hides (if HideButton is True)
    ' a maximize window button.
    ' NOTE: If EITHER a Minimize or Maximize button is displayed,
    ' BOTH buttons are visible but may be disabled.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION GWL_STYLE
'@INCLUDE DECLARATION WS_MAXIMIZEBOX
'@INCLUDE DECLARATION GetWindowLong
'@INCLUDE DECLARATION SetWindowLong

    Dim UFHWnd As Long
    Dim WinInfo As Long
    Dim R As Long

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        ShowMaximizeButton = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If HideButton = False Then
        WinInfo = WinInfo Or WS_MAXIMIZEBOX
    Else
        WinInfo = WinInfo And (Not WS_MAXIMIZEBOX)
    End If
    R = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)

    ShowMaximizeButton = (R <> 0)

End Function

Public Function ShowMinimizeButton(Optional HideButton As Boolean) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ShowMinimizeButton
    ' Displays (if HideButton is False) or hides (if HideButton is True)
    ' a minimize window button.
    ' NOTE: If EITHER a Minimize or Maximize button is displayed,
    ' BOTH buttons are visible but may be disabled.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION GWL_STYLE
'@INCLUDE DECLARATION WS_MINIMIZEBOX
'@INCLUDE DECLARATION GetWindowLong
'@INCLUDE DECLARATION SetWindowLong

    Dim UFHWnd As Long
    Dim WinInfo As Long
    Dim R As Long

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        ShowMinimizeButton = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If HideButton = False Then
        WinInfo = WinInfo Or WS_MINIMIZEBOX
    Else
        WinInfo = WinInfo And (Not WS_MINIMIZEBOX)
    End If
    R = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)

    ShowMinimizeButton = (R <> 0)

End Function

Private Function HasMinimizeButton() As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' HasMinimizeButton
    ' Returns True if the userform has a minimize button, False
    ' otherwise.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION GWL_STYLE
'@INCLUDE DECLARATION WS_MINIMIZEBOX
'@INCLUDE DECLARATION GetWindowLong

    Dim UFHWnd As Long
    Dim WinInfo As Long
    Dim R As Long

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        HasMinimizeButton = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

    If WinInfo And WS_MINIMIZEBOX Then
        HasMinimizeButton = True
    Else
        HasMinimizeButton = False
    End If

End Function

Private Function HasMaximizeButton() As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' HasMaximizeButton
    ' Returns True if the userform has a maximize button, False
    ' otherwise.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION GWL_STYLE
'@INCLUDE DECLARATION WS_MAXIMIZEBOX
'@INCLUDE DECLARATION GetWindowLong

    Dim UFHWnd As Long
    Dim WinInfo As Long
    Dim R As Long

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        HasMaximizeButton = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

    If WinInfo And WS_MAXIMIZEBOX Then
        HasMaximizeButton = True
    Else
        HasMaximizeButton = False
    End If

End Function


Public Function SetFormParent(Parent As FORM_PARENT_WINDOW_TYPE) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SetFormParent
    ' Set the UserForm UF as a child of (1) the Application, (2) the
    ' Excel ActiveWindow, or (3) no parent. Returns TRUE if successful
    ' or FALSE if unsuccessful.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE PROCEDURE WindowHWnd
'@INCLUDE DECLARATION FORM_PARENT_WINDOW_TYPE
'@INCLUDE DECLARATION SetParent

    Dim UFHWnd As Long
    Dim WindHWnd As Long
    #If VBA7 Then
    Dim R As LongPtr
    #Else
    Dim R As Long
    #End If

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        SetFormParent = False
        Exit Function
    End If

    Select Case Parent
    Case FORM_PARENT_APPLICATION
        R = SetParent(UFHWnd, Application.hwnd)
    Case FORM_PARENT_NONE
        R = SetParent(UFHWnd, 0&)
    Case FORM_PARENT_WINDOW
        If Application.ActiveWindow Is Nothing Then
            SetFormParent = False
            Exit Function
        End If
        WindHWnd = WindowHWnd(Application.ActiveWindow)
        If WindHWnd = 0 Then
            SetFormParent = False
            Exit Function
        End If
        R = SetParent(UFHWnd, WindHWnd)
    Case Else
        SetFormParent = False
        Exit Function
    End Select
    SetFormParent = (R <> 0)

End Function


Private Function IsCloseButtonVisible() As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsCloseButtonVisible
    ' Returns TRUE if UserForm UF has a close button, FALSE if there
    ' is no close button.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION GWL_STYLE
'@INCLUDE DECLARATION WS_SYSMENU
'@INCLUDE DECLARATION GetWindowLong

    Dim UFHWnd As Long
    Dim WinInfo As Long
    Dim R As Long

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        IsCloseButtonVisible = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    IsCloseButtonVisible = (WinInfo And WS_SYSMENU)

End Function

Public Function ShowCloseButton(Optional HideButton As Boolean) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ShowCloseButton
    ' This displays (if HideButton is FALSE) or hides (if HideButton is
    ' TRUE) the Close button on the userform
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION GWL_STYLE
'@INCLUDE DECLARATION WS_SYSMENU
'@INCLUDE DECLARATION GetWindowLong
'@INCLUDE DECLARATION SetWindowLong

    Dim UFHWnd As Long
    Dim WinInfo As Long
    Dim R As Long

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If HideButton = False Then
        ' set the SysMenu bit
        WinInfo = WinInfo Or WS_SYSMENU
    Else
        ' clear the SysMenu bit
        WinInfo = WinInfo And (Not WS_SYSMENU)
    End If

    R = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
    ShowCloseButton = (R <> 0)

End Function


Private Function IsCloseButtonEnabled() As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsCloseButtonEnabled
    ' This returns TRUE if the close button is enabled or FALSE if
    ' the close button is disabled.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE EnableCloseButton
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION MF_BYPOSITION
'@INCLUDE DECLARATION MF_DISABLED
'@INCLUDE DECLARATION MF_ENABLED
'@INCLUDE DECLARATION DrawMenuBar
'@INCLUDE DECLARATION EnableMenuItem
'@INCLUDE DECLARATION GetMenuItemCount
'@INCLUDE DECLARATION GetSystemMenu

    Dim UFHWnd As Long
    Dim hMenu As Long
    Dim ItemCount As Long
    Dim PrevState As Long

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        IsCloseButtonEnabled = False
        Exit Function
    End If
    ' Get the menu handle
    hMenu = GetSystemMenu(UFHWnd, 0&)
    If hMenu = 0 Then
        IsCloseButtonEnabled = False
        Exit Function
    End If

    ItemCount = GetMenuItemCount(hMenu)
    ' Disable the button. This returns MF_DISABLED or MF_ENABLED indicating
    ' the previous state of the item.
    PrevState = EnableMenuItem(hMenu, ItemCount - 1, MF_DISABLED Or MF_BYPOSITION)

    If PrevState = MF_DISABLED Then
        IsCloseButtonEnabled = False
    Else
        IsCloseButtonEnabled = True
    End If
    ' restore the previous state
    EnableCloseButton (PrevState = MF_DISABLED)

    DrawMenuBar UFHWnd

End Function


Public Function EnableCloseButton(Disable As Boolean) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' EnableCloseButton
    ' This function enables (if Disable is False) or disables (if
    ' Disable is True) the "X" button on a UserForm UF.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION MF_BYPOSITION
'@INCLUDE DECLARATION MF_DISABLED
'@INCLUDE DECLARATION MF_ENABLED
'@INCLUDE DECLARATION DrawMenuBar
'@INCLUDE DECLARATION EnableMenuItem
'@INCLUDE DECLARATION GetMenuItemCount
'@INCLUDE DECLARATION GetSystemMenu

    Dim UFHWnd As Long
    Dim hMenu As Long
    Dim ItemCount As Long
    Dim Res As Long

    ' Get the HWnd of the UserForm.
    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        EnableCloseButton = False
        Exit Function
    End If
    ' Get the menu handle
    hMenu = GetSystemMenu(UFHWnd, 0&)
    If hMenu = 0 Then
        EnableCloseButton = False
        Exit Function
    End If

    ItemCount = GetMenuItemCount(hMenu)
    If Disable = True Then
        Res = EnableMenuItem(hMenu, ItemCount - 1, MF_DISABLED Or MF_BYPOSITION)
    Else
        Res = EnableMenuItem(hMenu, ItemCount - 1, MF_ENABLED Or MF_BYPOSITION)
    End If
    If Res = -1 Then
        EnableCloseButton = False
        Exit Function
    End If
    DrawMenuBar UFHWnd

    EnableCloseButton = True


End Function

Public Function ShowTitleBar(HideTitle As Boolean) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ShowTitleBar
    ' Displays (if HideTitle is FALSE) or hides (if HideTitle is TRUE) the
    ' title bar of the userform UF.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION GWL_STYLE
'@INCLUDE DECLARATION WS_CAPTION
'@INCLUDE DECLARATION GetWindowLong
'@INCLUDE DECLARATION SetWindowLong

    Dim UFHWnd As Long
    Dim WinInfo As Long
    Dim R As Long

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        ShowTitleBar = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

    If HideTitle = False Then
        ' turn on the Caption bit
        WinInfo = WinInfo Or WS_CAPTION
    Else
        ' turn off the Caption bit
        WinInfo = WinInfo And (Not WS_CAPTION)
    End If
    R = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
    ShowTitleBar = (R <> 0)
End Function

Function IsTitleBarVisible() As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsTitleBarVisible
    ' Returns TRUE if the title bar of UF is visible or FALSE if the
    ' title bar is not visible.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION GWL_STYLE
'@INCLUDE DECLARATION WS_CAPTION
'@INCLUDE DECLARATION GetWindowLong
    Dim UFHWnd As Long
    Dim WinInfo As Long
    Dim R As Long

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        IsTitleBarVisible = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

    IsTitleBarVisible = (WinInfo And WS_CAPTION)

End Function

Function MakeFormResizable(Sizable As Boolean) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' MakeFormResizable
    ' This makes the userform UF resizable (if Sizable is TRUE) or not
    ' resizable (if Sizalbe is FALSE). Returns TRUE if successful or FALSE
    ' if an error occurred.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION GWL_STYLE
'@INCLUDE DECLARATION WS_SIZEBOX
'@INCLUDE DECLARATION GetWindowLong
'@INCLUDE DECLARATION SetWindowLong
    Dim UFHWnd As Long
    Dim WinInfo As Long
    Dim R As Long

    UFHWnd = CLng(hwnd) 'HWndOfUserForm
    If UFHWnd = 0 Then
        MakeFormResizable = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If Sizable = True Then
        WinInfo = WinInfo Or WS_SIZEBOX
    Else
        WinInfo = WinInfo And (Not WS_SIZEBOX)
    End If

    R = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
    MakeFormResizable = (R <> 0)

    If Sizable Then
        mdWidth = oForm.Width
        mdHeight = oForm.Height
        Set DockResize = oForm
    Else
        Set DockResize = Nothing
    End If
End Function

Function IsFormResizable() As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsFormResizable
    ' Returns TRUE if UF is resizable, FALSE if UF is not resizable.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION GWL_STYLE
'@INCLUDE DECLARATION WS_SIZEBOX
'@INCLUDE DECLARATION GetWindowLong

    Dim UFHWnd As Long
    Dim WinInfo As Long
    Dim R As Long

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        IsFormResizable = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

    IsFormResizable = (WinInfo And WS_SIZEBOX)

End Function


Public Function SetFormOpacity(Opacity As Byte) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SetFormOpacity
    ' This function sets the opacity of the UserForm referenced by the
    ' UF parameter. Opacity specifies the opacity of the form, from
    ' 0 = fully transparent (invisible) to 255 = fully opaque. The function
    ' returns True if successful or False if an error occurred. This
    ' requires Windows 2000 or later -- it will not work in Windows
    ' 95, 98, or ME.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE HWndOfUserForm
'@INCLUDE DECLARATION GWL_EXSTYLE
'@INCLUDE DECLARATION LWA_ALPHA
'@INCLUDE DECLARATION WS_EX_LAYERED
'@INCLUDE DECLARATION GetWindowLong
'@INCLUDE DECLARATION SetLayeredWindowAttributes
'@INCLUDE DECLARATION SetWindowLong
    Dim UFHWnd As Long
    Dim WinL As Long
    Dim Res As Long

    SetFormOpacity = False

    UFHWnd = HWndOfUserForm
    If UFHWnd = 0 Then
        Exit Function
    End If

    WinL = GetWindowLong(UFHWnd, GWL_EXSTYLE)
    If WinL = 0 Then
        Exit Function
    End If

    Res = SetWindowLong(UFHWnd, GWL_EXSTYLE, WinL Or WS_EX_LAYERED)
    If Res = 0 Then
        Exit Function
    End If

    Res = SetLayeredWindowAttributes(UFHWnd, 0, Opacity, LWA_ALPHA)
    If Res = 0 Then
        Exit Function
    End If

    SetFormOpacity = True

End Function


Function HWndOfUserForm() As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' HWndOfUserForm
    ' This returns the window handle (HWnd) of the userform referenced
    ' by UF. It first looks for a top-level window, then a child
    ' of the Application window, then a child of the ActiveWindow.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl
'@INCLUDE PROCEDURE WindowHWnd
'@INCLUDE DECLARATION C_USERFORM_CLASSNAME
'@INCLUDE DECLARATION FindWindow
'@INCLUDE DECLARATION FindWindowEx
    Dim AppHWnd As Long
    Dim DeskHWnd As Long
    Dim WinHWnd As Long
    Dim UFHWnd As Long
    Dim Cap As String
    Dim WindowCap As String

    Cap = oForm.Caption

    ' First, look in top level windows
    UFHWnd = FindWindow(C_USERFORM_CLASSNAME, Cap)
    If UFHWnd <> 0 Then
        HWndOfUserForm = UFHWnd
        Exit Function
    End If
    ' Not a top level window. Search for child of application.
    AppHWnd = Application.hwnd
    UFHWnd = FindWindowEx(AppHWnd, 0&, C_USERFORM_CLASSNAME, Cap)
    If UFHWnd <> 0 Then
        HWndOfUserForm = UFHWnd
        Exit Function
    End If
    ' Not a child of the application.
    ' Search for child of ActiveWindow (Excel's ActiveWindow, not
    ' Window's ActiveWindow).
    If Application.ActiveWindow Is Nothing Then
        HWndOfUserForm = 0
        Exit Function
    End If
    WinHWnd = WindowHWnd(Application.ActiveWindow)
    UFHWnd = FindWindowEx(WinHWnd, 0&, C_USERFORM_CLASSNAME, Cap)
    HWndOfUserForm = UFHWnd

End Function


Function ClearBit(Value As Long, ByVal BitNumber As Long) As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ClearBit
    ' Clears the specified bit in Value and returns the result. Bits are
    ' numbered, right (most significant) 31 to left (least significant) 0.
    ' BitNumber is made positive and then MOD 32 to get a valid bit number.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@AssignedModule modFormControl

    Dim SetMask As Long
    Dim ClearMask As Long

    BitNumber = Abs(BitNumber) Mod 32

    SetMask = Value
    If BitNumber < 30 Then
        ClearMask = Not (2 ^ (BitNumber - 1))
        ClearBit = SetMask And ClearMask
    Else
        ClearBit = Value And &H7FFFFFFF
    End If

End Function



#If VBA7 Then
Function hwnd() As LongPtr
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
#Else
Function hwnd() As Long
#End If
    WindowFromAccessibleObject oForm, hwnd
End Function


Public Sub Transition(ParamArray Elements() As Variant)
    Rem By Robert Todaer
    '__usage__
    'with aUserform
    '    .Transition .Effect(Box, "Top", Me.InsideHeight - Box.Height, 1000) _
         , .Effect(box2, "Top", 0, 100) _
         , .Effect(GoButton, "fontsize", 12, 1000) _
         , .Effect(Me, "Top", 20, 2000)
    'End with
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    '@INCLUDE DECLARATION Sleep

    If IsArray(Elements(LBound(Elements, 1))) Then
        Dim Temp    As Variant
        Temp = Elements(LBound(Elements, 1))
        Elements = Temp
    End If
    Dim Form        As MSForms.UserForm
    Set Form = Elements(LBound(Elements, 1))("form")
    MicroTimer True
    Do
        Dim index   As Integer
        For index = LBound(Elements, 1) To UBound(Elements, 1)
            IncRementElement Elements(index), MicroTimer
        Next index
        Sleep 40
        Form.Repaint
    Loop Until AllTransitionsComplete(CVar(Elements))
End Sub

Public Function Effect(obj As Object, Property As String, Destination As Double, MilSecs As Double) As Scripting.Dictionary
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    Dim Temp        As New Scripting.Dictionary
    Set Temp("obj") = obj
    Temp("property") = Property
    Temp("startValue") = CallByName(obj, Property, VbGet)
    Temp("destination") = Destination
    Temp("travel") = Destination - Temp("startValue")
    Temp("milSec") = MilSecs
    Temp("complete") = False
    On Error GoTo catch:
    Set Temp("form") = obj.Parent
    Set Effect = Temp
    Exit Function
catch:
    Set Temp("form") = obj
    Resume Next
End Function

Private Function MicroTimer(Optional StartTime As Boolean = False) As Double
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    '@INCLUDE DECLARATION getTickCount
    Static dTime    As Double
    Dim cyTicks1    As Currency
    Dim cyTicks2    As Currency
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
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    Dim el          As Object
    Dim index       As Integer
    For index = LBound(Elements, 1) To UBound(Elements, 1)
        Set el = Elements(index)
        If Not TransitionComplete(el) Then
            AllTransitionsComplete = False
            Exit Function
        End If
    Next index
    AllTransitionsComplete = True
End Function

Private Function TransitionComplete(ByVal el As Scripting.Dictionary) As Boolean
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    If Math.Round(el("destination")) = Math.Round(CallByName(el("obj"), el("property"), VbGet)) Then
        TransitionComplete = True
    End If
End Function

Private Function IncRementElement(ByVal el As Scripting.Dictionary, CurrentTime As Double) As Boolean
    '@INCLUDE TransitionComplete
    '@INCLUDE easeInAndOut
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    Dim IncRementValue As Double
    Dim CurrentValue As Double
    If TransitionComplete(el) Then
        Exit Function
    End If
    Dim o           As Object
    Dim p           As String
    Dim Value       As Double
    Dim d           As Double
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
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    d = d / 2
    t = t / d
    If (t < 1) Then
        easeInAndOut = c / 2 * t * t * t + b
    Else
        t = t - 2
        easeInAndOut = c / 2 * (t * t * t + 2) + b
    End If
End Function

Public Sub Resizable(Optional Enable As Boolean = True)
    Dim hWndForm As Long, iStyle As Long
    If val(Application.version) < 9 Then
        hWndForm = FindWindow("ThunderXFrame", oForm.Caption)
    Else
        hWndForm = FindWindow("ThunderDFrame", oForm.Caption)
    End If
    iStyle = GetWindowLong(hWndForm, GWL_STYLE)

    If Enable Then
        iStyle = iStyle Or WS_THICKFRAME
    Else
        iStyle = iStyle And Not WS_THICKFRAME
    End If

    SetWindowLong hWndForm, GWL_STYLE, iStyle
    ShowWindow hWndForm, SW_SHOW
    DrawMenuBar hWndForm
    SetFocus hWndForm

    mdWidth = oForm.Width
    mdHeight = oForm.Height

    Set DockResize = oForm
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 30-10-2023 17:56    Alex                (aUserform.cls > Resizable2 : new method )

Public Sub Resizable2(Optional Enable As Boolean = True)
'@LastModified 2310301756
'use icon at bottom right corner
    If Enable Then
        On Error Resume Next
        Set lblResizer = oForm.Controls.Add("forms.label.1")
        On Error GoTo 0
        With lblResizer
            .MousePointer = fmMousePointerSizeNWSE
            .AutoSize = True
        End With
        IconDesign lblResizer, "F169"
        lblResizer.Left = oForm.InsideWidth - lblResizer.Width
        lblResizer.Top = oForm.InsideHeight - lblResizer.Height
        Set DockResize = oForm
        Set lblResize = lblResizer
    Else
      Set lblResizer = Nothing
    End If
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 30-10-2023 17:56    Alex                (aUserform.cls > lblResize_MouseDown : new method)

Private Sub lblResize_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single)
'@LastModified 2310301756
    'The user clicked on the lblResizer
    resizeEnabled = True
    'Capture the mouse position on click
    mouseX = X
    mouseY = Y
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 30-10-2023 17:56    Alex                (aUserform.cls > lblResize_MouseMove : new method)

Private Sub lblResize_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single)
'@LastModified 2310301756

    'Check if the UserForm is not resized too small
    Dim allowResize As Boolean

    allowResize = True

    If oForm.Width + X - mouseX < minWidth Then allowResize = False
    If oForm.Height + Y - mouseY < minHeight Then allowResize = False

    'Check if the mouse clicked on the lblResizer and above minimum size
    If resizeEnabled = True And allowResize = True Then
        'Resize the UserForm
        oForm.Width = oForm.Width + X - mouseX
        oForm.Height = oForm.Height + Y - mouseY
        'Move the Resizer icon
        lblResizer.Left = oForm.InsideWidth - lblResizer.Width
        lblResizer.Top = oForm.InsideHeight - lblResizer.Height
    End If

End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 30-10-2023 17:56    Alex                (aUserform.cls > lblResize_MouseUp : new method)

Private Sub lblResize_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single)
'@LastModified 2310301756
    'The user un-clicked on the lblResizer
    resizeEnabled = False

End Sub


'* Modified   : Date and Time       Author              Description
'* Updated    : 30-10-2023 17:55    Alex                (aUserform.cls > IconDesign : init )

Public Sub IconDesign(ctrl As control, IconCode As String)
'@LastModified 2310301755
    '@AssignedModule aUserform
    With ctrl
        .Font.Name = "Segoe MDL2 Assets"
        .Caption = ChrW("&H" & IconCode)
        .Font.Size = 12
        .ForeColor = RGB(191, 191, 191)
        .TextAlign = fmTextAlignLeft
        .BorderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleTransparent
    End With
End Sub

Public Sub DockResize_Layout()
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    DockControls
End Sub

Public Sub DockControls()

    '* @TODO Created: 30-10-2023 08:08 Author: Anastasiou Alex
    '* @TODO pass arguments instead of ctrl.tag
    'eg
    'Public TrackHeight  As New Collection
    'Public TrackWidth   As New Collection
    'Public TrackTop     As New Collection
    'Public TrackLeft    As New Collection
    
    'from STEPHEN BULLEN's USERFORM RESIZER CLASS
    '
    ' _______IN USERFORM_________________
    '
    ' Private Sub UserForm_Resize()
    '     auserform.init(me).DockControls
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
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform

    Dim dWidthAdj As Double, dHeightAdj As Double, sTag As String
    Dim oCtl        As MSForms.control
    If mdWidth = 0 Then mdWidth = oForm.Width
    If mdHeight = 0 Then mdHeight = oForm.Height
    dWidthAdj = oForm.Width - mdWidth
    dHeightAdj = oForm.Height - mdHeight
    For Each oCtl In oForm.Controls
        '        If TypeName(oCtl) = "ListView" Then Stop
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
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    Dim i As Integer, d As Double
    i = InStr(1, sTag, sChange, vbBinaryCompare)
    If i > 0 Then
        d = val(Mid$(sTag, i + 1))
        If d = 0 Then d = 1
    End If
    ResizeFactor = d
End Function

Public Function Init(Form As Object) As aUserform
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    Set oForm = Form
    Set Init = Me
End Function

Private Sub Class_Initialize()
    'do something
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
End Sub

'Function GetUserformHwnd(Form As Object)
'    GetUserformHwnd = FindWindow(vbNullString, Form.Caption)
'End Function

Public Sub TRANSPARENT(Optional color As Variant)
    '@INCLUDE MakeFormBorderless
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    '@INCLUDE DECLARATION GWL_EXSTYLE
    '@INCLUDE DECLARATION LWA_COLORKEY
    '@INCLUDE DECLARATION WS_EX_LAYERED
    '@INCLUDE DECLARATION FindWindow
    '@INCLUDE DECLARATION GetWindowLong
    '@INCLUDE DECLARATION SetLayeredWindowAttributes
    '@INCLUDE DECLARATION SetWindowLong
    Dim formhandle  As Long
    Dim bytOpacity  As Byte
    formhandle = CLng(FindWindow(vbNullString, oForm.Caption))
    If IsMissing(color) Then color = vbWhite
    bytOpacity = 100
    SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
    oForm.BackColor = color
    SetLayeredWindowAttributes formhandle, color, bytOpacity, LWA_COLORKEY
End Sub

Public Sub Borderless()
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
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
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    '@INCLUDE DECLARATION HWND_TOPMOST
    '@INCLUDE DECLARATION SWP_NOMOVE
    '@INCLUDE DECLARATION SWP_NOSIZE
    '@INCLUDE DECLARATION FindWindow
    '@INCLUDE DECLARATION SetWindowPos
    Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"
    Dim ret         As Long
    Dim formHWnd    As Long
    formHWnd = CLng(FindWindow(C_VBA6_USERFORM_CLASSNAME, oForm.Caption))
    If formHWnd = 0 Then
        Debug.Print Err.LastDllError
    End If
    ret = SetWindowPos(formHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    If ret = 0 Then
        Debug.Print Err.LastDllError
    End If
End Sub

Sub ParentIsVBE()
    '@INCLUDE DisplayErrorText
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    '@INCLUDE DECLARATION USERFORM_WINDOW_CLASS
    '@INCLUDE DECLARATION FindWindow
    '@INCLUDE DECLARATION SetForegroundWindow
    '@INCLUDE DECLARATION SetParent
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
    VBEWindowPointer = Application.VBE.MainWindow.hwnd
    UserFormWindowPointer = FindWindow(lpClassName:=USERFORM_WINDOW_CLASS, lpWindowName:=GivenFormCaption)
    Const ERROR_NUMBER_FOR_SETPARENT_API = 0
    ReturnOfSetParentAPI = SetParent(hWndChild:=UserFormWindowPointer, hWndNewParent:=VBEWindowPointer)
    If ReturnOfSetParentAPI = ERROR_NUMBER_FOR_SETPARENT_API Then
        ErrorNumber = Err.LastDllError
        DisplayErrorText "Error With SetParent", ErrorNumber
    Else
        '        Debug.Print GivenFormCaption & " is child of VBE Window."
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
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    Dim ErrText     As String
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
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    '@INCLUDE DECLARATION FORMAT_MESSAGE_FROM_SYSTEM
    '@INCLUDE DECLARATION FORMAT_MESSAGE_IGNORE_INSERTS
    '@INCLUDE DECLARATION FORMAT_MESSAGE_TEXT_LEN
    '@INCLUDE DECLARATION FormatMessage
    Dim ErrorText   As String
    Dim ErrorTextLength As Long
    Dim FormatMessageResult As Long
    Dim LanguageID  As Long
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
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    Dim NullCharIndex As Integer
    NullCharIndex = InStr(1, TEXT, vbNullChar, vbTextCompare)
    If NullCharIndex > 0 Then
        TrimToNull = Left(TEXT, NullCharIndex - 1)
    Else
        TrimToNull = TEXT
    End If
End Function
'
'Public Sub MinimizeButton() 'replaced by ShowMinimizedButton
'    '@AssignedModule aUserform
'    '@INCLUDE CLASS aUserform
'    '@INCLUDE DECLARATION FindWindowA
'    '@INCLUDE DECLARATION GetWindowLongA
'    '@INCLUDE DECLARATION SetWindowLongA
'    Dim Cap         As String: Cap = oForm.Caption
'    Dim hwnd        As Long: hwnd = FindWindowA(vbNullString, Cap)
'    Dim exLong      As Long: exLong = GetWindowLongA(hwnd, -16)
'
'    If (exLong And &H20000) = 0 Then SetWindowLongA hwnd, -16, exLong Or &H20000
'End Sub

Public Sub SaveOptions( _
        Optional includeCheckbox As Boolean = True, _
        Optional includeOptionButton As Boolean = True, _
        Optional includeTextBox As Boolean = True, _
        Optional includeListbox As Boolean = True, _
        Optional includeToggleButton As Boolean = True, _
        Optional includeCombobox As Boolean = True, _
        Optional includeMultipage As Boolean = True)
    '@AssignedModule aUserform
    '@INCLUDE PROCEDURE FoldersCreate
    '@INCLUDE PROCEDURE IniWrite
    '@INCLUDE CLASS aCollection
    '@INCLUDE CLASS aListBox
    '@INCLUDE CLASS aUserform

    Dim configFolder As String: configFolder = ThisWorkbook.path & "\Lib\config\"
    FoldersCreate configFolder
    Dim iniFile     As String: iniFile = configFolder & "UserformSettings.ini"
    Dim Key, Value
    Dim c           As MSForms.control
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
        ElseIf TypeName(c) = "MultiPage" Then
            If Not includeMultipage Then GoTo SKIP
        Else
            GoTo SKIP
        End If
        Key = c.Name
        Dim coll    As New Collection
        Select Case TypeName(c)
            Case "CheckBox", "OptionButton", "ToggleButton", "ComboBox", "MultiPage"
                IniWrite iniFile, oForm.Name, Key, c.Value
            Case "TextBox"
                IniWrite iniFile, oForm.Name, Key, VBA.Replace(c.Value, vbNewLine, "||")
            Case "ListBox"
                Set coll = aListBox.Init(c).selectedIndexes
                If coll.Count > 0 Then
                    IniWrite iniFile, oForm.Name, Key, aCollection.Init(coll).ToString(",")
                Else
                    IniWrite iniFile, oForm.Name, Key, -1
                End If
        End Select
SKIP:
    Next

End Sub

Public Sub LoadOptions(ParamArray ExcludeThese() As Variant)
    'ExcludeThese:="ListBox1","TextBox1"     '<--control names
    '@AssignedModule aUserform
    '@INCLUDE PROCEDURE ArrayAllocated
    '@INCLUDE PROCEDURE FileExists
    '@INCLUDE PROCEDURE FoldersCreate
    '@INCLUDE PROCEDURE TxtOverwrite
    '@INCLUDE PROCEDURE ArrayContains
    '@INCLUDE PROCEDURE CreateOrSetSheet
    '@INCLUDE PROCEDURE IniSectionKeys
    '@INCLUDE PROCEDURE IniReadKey
    '@INCLUDE CLASS aListBox
    '@INCLUDE CLASS aUserform
    Dim configFolder As String: configFolder = ThisWorkbook.path & "\Lib\config\"
    FoldersCreate configFolder
    Dim iniFile     As String: iniFile = configFolder & "UserformSettings.ini"
    If Not FileExists(iniFile) Then TxtOverwrite iniFile, ""
    Dim Key, Value
    Dim c           As MSForms.control
    Dim keyArray: keyArray = IniSectionKeys(iniFile, oForm.Name)
    If Not ArrayAllocated(keyArray) Then Exit Sub
    For Each Key In keyArray
        On Error Resume Next
        Set c = oForm.Controls(Key)
        Value = IniReadKey(iniFile, oForm.Name, Key)
        If Not TypeName(c) = "Nothing " Then
            If Not ArrayContains(Key, ExcludeThese) Then
                Select Case TypeName(c)
                    Case "CheckBox", "OptionButton", "ToggleButton", "ComboBox", "MultiPage"
                        c.Value = Value
                    Case "TextBox"
                        c.Value = VBA.Replace(Value, "||", vbNewLine)
                    Case "ListBox"
                        If InStr(1, Value, ",") > 0 Then
                            aListBox.Init(c).SelectItems Split(Value, ","), True
                        Else
                            c.Selected(CInt(Value)) = True
                        End If
                End Select
            End If
        End If
    Next

    '=== OBSOLETE
    '
    '    Dim ws As Worksheet
    '    Set ws = CreateOrSetSheet(oForm.Name & "_Settings", ThisWorkbook)
    '    If ws.Range("A1") = "" Then Exit Sub
    '    Dim cell As Range
    '    Set cell = ws.Cells(1, 1)
    '    Dim c As MSForms.control
    '    Dim v
    '    On Error Resume Next
    '    Do While cell <> ""
    '        Set c = oForm.Controls(cell.TEXT)
    '        If Not TypeName(c) = "Nothing " Then
    '            If Not ArrayContains(cell, ExcludeThese) Then
    '                Select Case TypeName(c)
    '                Case "TextBox", "CheckBox", "OptionButton", "ToggleButton", "ComboBox"
    '                    c.Value = cell.OFFSET(0, 1)
    '                Case "ListBox"
    '                    If InStr(1, cell.OFFSET(0, 1), ",") > 0 Then
    '                        aListBox.Init(c).SelectItems Split(cell.OFFSET(0, 1), ","), True
    '                    Else
    '                        c.Selected(CInt(cell.OFFSET(0, 1))) = True
    '                    End If
    '                End Select
    '            End If
    '        End If
    '        Set cell = cell.OFFSET(1, 0)
    '    Loop
End Sub

Public Sub SavePosition()
    '@AssignedModule aUserform
    '@INCLUDE PROCEDURE FoldersCreate
    '@INCLUDE PROCEDURE IniWrite
    '@INCLUDE CLASS aUserform
    Dim configFolder As String: configFolder = ThisWorkbook.path & "\Lib\config\"
    FoldersCreate configFolder
    Dim iniFile     As String: iniFile = configFolder & "UserformSettings.ini"
    IniWrite iniFile, oForm.Name, "StartupTop", oForm.Top
    IniWrite iniFile, oForm.Name, "StartupLeft", oForm.Left

    '=== OBSOLETE
    '    SaveSetting "My Settings Folder", oForm.Name, "Left Position", oForm.Left
    '    SaveSetting "My Settings Folder", oForm.Name, "Top Position", oForm.Top
End Sub

Public Sub LoadPosition()
    '@AssignedModule aUserform
    '@INCLUDE PROCEDURE FileExists
    '@INCLUDE PROCEDURE FoldersCreate
    '@INCLUDE PROCEDURE TxtOverwrite
    '@INCLUDE PROCEDURE IniReadKey
    '@INCLUDE CLASS aUserform
    Dim configFolder As String: configFolder = ThisWorkbook.path & "\Lib\config\"
    FoldersCreate configFolder
    Dim iniFile     As String: iniFile = configFolder & "UserformSettings.ini"
    If Not FileExists(iniFile) Then TxtOverwrite iniFile, ""
    Dim myLeft, myTop
    myLeft = IniReadKey(iniFile, oForm.Name, "StartupLeft")
    myTop = IniReadKey(iniFile, oForm.Name, "StartupTop")
    If myLeft = "" Or myTop = "" Then
        oForm.startupposition = 1
    Else
        oForm.startupposition = 0
        oForm.Top = myTop
        oForm.Left = myLeft
    End If

    'Setting         Value   Description
    'Manual          0       No initial setting specified.
    'CenterOwner     1       Center on the item to which the UserForm belongs.
    'CenterScreen    2       Center on the whole screen.
    'WindowsDefault  3       Position in upper-left corner of screen.

    '=== OBSOLETE
    '    If GetSetting("My Settings Folder", oForm.Name, "Left Position") = "" _
         '        And GetSetting("My Settings Folder", oForm.Name, "Top Position") = "" Then
    '        oForm.startupposition = 1
    '    Else
    '        oForm.Left = GetSetting("My Settings Folder", oForm.Name, "Left Position")
    '        oForm.Top = GetSetting("My Settings Folder", oForm.Name, "Top Position")
    '    End If
End Sub

Public Sub ResizeToFitControls(Optional marginRight = 0, Optional marginBottom = 0)
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    oForm.Width = 0
    oForm.Height = 0
    Dim ctr         As MSForms.control
    Dim myWidth: myWidth = 0    'oForm.InsideWidth
    Dim myHeight: myHeight = 0    'oForm.InsideHeight
    For Each ctr In oForm.Controls
        If ctr.Parent.Name = oForm.Name Then
            If ctr.Visible = True Then
                If ctr.Left + ctr.Width > myWidth Then myWidth = ctr.Left + ctr.Width
                If ctr.Top + ctr.Height > myHeight Then myHeight = ctr.Top + ctr.Height
            End If
        End If
    Next
    oForm.Width = myWidth + (oForm.Width - oForm.InsideWidth) + marginRight
    oForm.Height = myHeight + (oForm.Height - oForm.InsideHeight) + marginBottom
End Sub


Public Sub MouseOnControl(ctrl As Object)
    '@INCLUDE PointsPerPixelX
    '@INCLUDE PointsPerPixelY
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    '@INCLUDE DECLARATION ClientToScreen
    '@INCLUDE DECLARATION FindWindow
    '@INCLUDE DECLARATION SetCursorPos
    '@INCLUDE DECLARATION Ctrl
    '@INCLUDE DECLARATION tCursor
    Dim p           As tCursor
    Dim lngHwnd     As Long
    lngHwnd = CLng(FindWindow(vbNullString, oForm.Caption))
    p.Left = (ctrl.Left + (ctrl.Width \ 2)) / PointsPerPixelX
    p.Top = (ctrl.Top + (ctrl.Height \ 2)) / PointsPerPixelY
    ClientToScreen lngHwnd, p
    SetCursorPos p.Left, p.Top
End Sub

Private Function PointsPerPixelX() As Double
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    '@INCLUDE DECLARATION LOGPIXELSX
    '@INCLUDE DECLARATION GetDC
    '@INCLUDE DECLARATION GetDeviceCaps
    '@INCLUDE DECLARATION ReleaseDC
    Dim hdc         As Long
    hdc = GetDC(0)
    PointsPerPixelX = 72 / GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc
End Function

Private Function PointsPerPixelY() As Double
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    '@INCLUDE DECLARATION LOGPIXELSY
    '@INCLUDE DECLARATION GetDC
    '@INCLUDE DECLARATION GetDeviceCaps
    '@INCLUDE DECLARATION ReleaseDC
    Dim hdc         As Long
    hdc = GetDC(0)
    PointsPerPixelY = 72 / GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
End Function

Private Function WhereIsTheMouseAt() As tCursor
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    '@INCLUDE DECLARATION GetCursorPos
    '@INCLUDE DECLARATION tCursor
    Dim mPos        As tCursor
    GetCursorPos mPos
    WhereIsTheMouseAt = mPos
End Function

Private Function convertMouseToForm() As tCursor
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    '@INCLUDE DECLARATION tCursor
    Dim mPos        As tCursor
    mPos = WhereIsTheMouseAt
    mPos.Left = PointsPerPixelY * mPos.Left
    mPos.Top = PointsPerPixelX * mPos.Top
    convertMouseToForm = mPos
End Function

Public Sub ShowAtCursor()
    '@AssignedModule aUserform
    '@INCLUDE CLASS aUserform
    oForm.Left = convertMouseToForm.Left
    oForm.Top = convertMouseToForm.Top
End Sub
