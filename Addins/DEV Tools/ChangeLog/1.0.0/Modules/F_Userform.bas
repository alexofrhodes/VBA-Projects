Attribute VB_Name = "F_Userform"

Option Explicit
Option Compare Text

Rem @Subfolder Userforms>Transparent Declarations
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

Private m_sngDownX  As Single
Private m_sngDownY  As Single

Rem Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Rem Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Rem @Subfolder Userforms>Parent Declarations
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Public Const FORMAT_MESSAGE_TEXT_LEN = 160
Public Const MAX_PATH = 260
Public Const GWL_HWNDPARENT As Long = -8
Public Const GW_OWNER = 4

#If VBA7 Then
Private Declare PtrSafe Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
#Else
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
#End If

Public VBEditorHWnd As Long
Public ApplicationHWnd As Long
Public ExcelDeskHWnd As Long
Public ActiveWindowHWnd As Long
Public UserFormHWnd As Long
Public WindowsDesktopHWnd As Long
Public Const GA_ROOT As Long = 2
Public Const GA_ROOTOWNER As Long = 3
Public Const GA_PARENT As Long = 1

#If VBA7 Then
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare PtrSafe Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Public Declare PtrSafe Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Public Const C_EXCEL_APP_WINDOWCLASS = "XLMAIN"
Public Const C_EXCEL_DESK_WINDOWCLASS = "XLDESK"
Public Const C_EXCEL_WINDOW_WINDOWCLASS = "EXCEL7"
Public Const USERFORM_WINDOW_CLASS = "ThunderDFrame"
Public Const C_VBA_USERFORM_WINDOWCLASS = "ThunderDFrame"

Rem Window position and more
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

#If VBA7 Then
Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As LongPtr, ByVal y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal uFlags As LongPtr) As Long
#Else
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal uFlags As LongPtr) As Long
#End If

Rem ---
#If VBA7 Then
Public Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Public Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As LongPtr) As Long
#Else
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
#End If

Rem Closeby
Public Enum CloseBy
    User = 0
    Code = 1
    WindowsOS = 2
    TaskManager = 3
End Enum

Rem FlashControl
#If VBA7 Then
Public Declare PtrSafe Function getTickCount Lib "kernel32" Alias "GetTickCount" () As Long
#Else
Public Declare Function getTickCount Lib "kernel32" Alias "GetTickCount" () As Long
#End If

Public Const black  As Long = &H80000012
Public Const red    As Long = &HFF&
Rem ControlID
Public Const ControlIDCheckBox = "Forms.CheckBox.1"
Public Const ControlIDComboBox = "Forms.ComboBox.1"
Public Const ControlIDCommandButton = "Forms.CommandButton.1"
Public Const ControlIDFrame = "Forms.Frame.1"
Public Const ControlIDImage = "Forms.Image.1"
Public Const ControlIDLabel = "Forms.Label.1"
Public Const ControlIDListBox = "Forms.ListBox.1"
Public Const ControlIDMultiPage = "Forms.MultiPage.1"
Public Const ControlIDOptionButton = "Forms.OptionButton.1"
Public Const ControlIDScrollBar = "Forms.ScrollBar.1"
Public Const ControlIDSpinButton = "Forms.SpinButton.1"
Public Const ControlIDTabStrip = "Forms.TabStrip.1"
Public Const ControlIDTextBox = "Forms.TextBox.1"
Public Const ControlIDToggleButton = "Forms.ToggleButton.1"

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


Public Function IsLoaded(FormName As String) As Boolean
    '@AssignedModule F_Userform
    Dim frm         As Object
    For Each frm In VBA.Userforms
        If frm.Name = FormName Then
            IsLoaded = True
            Exit Function
        End If
    Next frm
    IsLoaded = False
End Function

Sub UserformPersist(FormName As String)
    'In userform put:
    '
    'Private Sub UserForm_Initialize()
    '    setUserformLostStateTrue Me.name
    '    UserformPersist Me.name
    'End Sub
    '
    'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '    setUserformLostStateFalse Me.name
    'End Sub

    '@AssignedModule F_Userform
    '@INCLUDE PROCEDURE UserformLostState
    '@INCLUDE PROCEDURE setUserformLostStateTrue
    '@INCLUDE PROCEDURE setUserformLostStateFalse
    '@INCLUDE PROCEDURE ShowUserform
    '@INCLUDE PROCEDURE appRunOnTime

    appRunOnTime Now() + TimeSerial(0, 0, 5), "UserformPersist", FormName

    If UserformLostState(FormName) Then
        ShowUserform FormName
    Else
        On Error Resume Next
        Application.OnTime Now, "UserformPersist", , False
        On Error GoTo 0
    End If
End Sub

Function UserformLostState(FormName As String) As Boolean
    '@INCLUDE CreateOrSetSheet
    '@AssignedModule F_Userform
    '@INCLUDE PROCEDURE CreateOrSetSheet
    UserformLostState = CBool(CreateOrSetSheet(FormName & "_Settings", ThisWorkbook).Range("Z1").value)
End Function

Sub setUserformLostStateTrue(FormName As String)
    '@INCLUDE CreateOrSetSheet
    '@AssignedModule F_Userform
    '@INCLUDE PROCEDURE CreateOrSetSheet
    CreateOrSetSheet(FormName & "_Settings", ThisWorkbook).Range("Z1").value = True
End Sub

Sub setUserformLostStateFalse(FormName As String)
    '@INCLUDE CreateOrSetSheet
    '@AssignedModule F_Userform
    '@INCLUDE PROCEDURE CreateOrSetSheet
    CreateOrSetSheet(FormName & "_Settings", ThisWorkbook).Range("Z1").value = False
End Sub

Rem var
Public Sub ShowUserform(FormName As String)
    '@INCLUDE IsLoaded
    '@AssignedModule F_Userform
    Dim frm         As Object
    If IsLoaded(FormName) = True Then
        For Each frm In VBA.Userforms
            If frm.Name = FormName Then
                On Error Resume Next
                frm.Show
                Exit Sub
            End If
        Next frm
    Else
        Dim oUserForm As Object
        On Error GoTo Err
        Set oUserForm = Userforms.Add(FormName)
        oUserForm.Show    '(vbModeless)
        Exit Sub
    End If
Err:
    Select Case Err.Number
        Case 424:
            Debug.Print "The Userform with the name " & FormName & " was not found.", vbExclamation, "Load userforn by name"
        Case Else:
            Debug.Print Err.Number & ": " & Err.Description, vbCritical, "Load userforn by name"
    End Select
End Sub


Public Sub Reframe(Form As Object, control As msforms.control)
    'TODO obsolete, replaced with amultipage.init(..).buildmenu
    Dim c           As msforms.control
    For Each c In Form.Controls
        If TypeName(c) = "Frame" Then
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                If c.Name <> control.Parent.Parent.Name Then c.Visible = False
            End If
        End If
    Next
    Form.Controls(control.Caption).Visible = True
    For Each c In Form.Controls
        If TypeName(c) = "Label" Then
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                c.BackColor = &H534848
            End If
        End If
    Next
    control.BackColor = &H80B91E
End Sub
