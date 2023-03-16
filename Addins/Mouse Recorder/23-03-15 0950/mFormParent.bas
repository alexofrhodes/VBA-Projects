Attribute VB_Name = "mFormParent"
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modSupport
' By Chip Pearson, chip@cpearson.com www.cpearson.com
'
' This module contains declarations and code that are used in support of the procedures in frmSetParent but aren't
' directly related to the topic of the workbook.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' used by FormatMessage
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Public Const FORMAT_MESSAGE_TEXT_LEN = 160        ' from ERRORS.H C++ include file.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Various constants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const MAX_PATH = 260        ' Windows mandated value
Public Const GWL_HWNDPARENT As Long = -8
Public Const GW_OWNER = 4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows API Declares
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare PtrSafe Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Public Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare PtrSafe Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' frmSetParent
' This form illustrates the SetParent procedure to make a userform the a child window
' of no window, the Excel Application window, the Excel Desktop Window, and the Active
' Window.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Window Class Names
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const C_EXCEL_APP_WINDOWCLASS = "XLMAIN"
Public Const C_EXCEL_DESK_WINDOWCLASS = "XLDESK"
Public Const C_EXCEL_WINDOW_WINDOWCLASS = "EXCEL7"
Public Const C_VBA_USERFORM_WINDOWCLASS = "ThunderDFrame"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Window HWnds
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public VBEditorHWnd As Long
Public ApplicationHWnd As Long
Public ExcelDeskHWnd As Long
Public ActiveWindowHWnd As Long
Public UserFormHWnd As Long
Public WindowsDesktopHWnd As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Other Consts
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const GA_ROOT As Long = 2
Public Const GA_ROOTOWNER As Long = 3
Public Const GA_PARENT As Long = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Windows API Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare PtrSafe Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Public Declare PtrSafe Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Function SetParent Lib "user32" (ByVal hwndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'''
Rem Form on top
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hwndInsertAfter As LongPtr, ByVal x As LongPtr, ByVal Y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal uFlags As LongPtr) As Long

Public Sub UserformOnTop(Form As Object)
    Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"
    Dim Ret As Long
    Dim formHWnd As Long
    'Get window handle of the userform
    formHWnd = CLng(FindWindow(C_VBA6_USERFORM_CLASSNAME, Form.Caption))
    If formHWnd = 0 Then
        Debug.Print Err.LastDllError
    End If
    'Set userform window to 'always on top'
    Ret = SetWindowPos(formHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    If Ret = 0 Then
        Debug.Print Err.LastDllError
    End If
    'Application.WindowState = xlMinimized
End Sub

'''
Public Sub DisplayLabels(Form As Object)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DipsplayLabelText
    ' This gets the various HWnds if they are not already
    ' set and updates the label captions on the form.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#INCLUDE GetWindowClassName
    '#INCLUDE GetHWndWindowText
    Dim ParentHWnd As Long
    Dim ParentWindowClass As String
    Dim AncestorWindow As Long
    Dim WinLong As Long
    Dim OwnerWindow As Long
    Dim ClassName As String
    Dim s As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get The HWnds
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''
    ' HWnd of the VBEditor
    '''''''''''''''''''''''''''''''
    VBEditorHWnd = Application.VBE.MainWindow.hWnd
    '''''''''''''''''''''''''''''''
    ' HWnd of the Excel Application
    '''''''''''''''''''''''''''''''
    ApplicationHWnd = FindWindow(lpClassName:=C_EXCEL_APP_WINDOWCLASS, lpWindowName:=Application.Caption)
    '''''''''''''''''''''''''''''''
    ' HWnd of the Excel Desktop
    '''''''''''''''''''''''''''''''
    ExcelDeskHWnd = FindWindowEx(hWnd1:=ApplicationHWnd, hWnd2:=0&, lpsz1:=C_EXCEL_DESK_WINDOWCLASS, lpsz2:=vbNullString)
    '''''''''''''''''''''''''''''''
    ' HWnd of the ActiveWindow
    '''''''''''''''''''''''''''''''
    ActiveWindowHWnd = FindWindowEx(hWnd1:=ExcelDeskHWnd, hWnd2:=0&, lpsz1:=C_EXCEL_WINDOW_WINDOWCLASS, lpsz2:=Application.ActiveWindow.Caption)
    '''''''''''''''''''''''''''''''
    ' HWnd of the UserForm
    '''''''''''''''''''''''''''''''
    UserFormHWnd = FindWindow(lpClassName:=C_VBA_USERFORM_WINDOWCLASS, lpWindowName:=Form.Caption)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' update the Option Button Captions with WindowText
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Debug.Print "Active Window (" & GetHWndWindowText(ActiveWindowHWnd) & ")"
    Debug.Print "Application (" & GetHWndWindowText(ApplicationHWnd) & ")"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Update Labels with HWnds and Parent HWnds.
    ' Use GetWindowLong rather than GetParent to
    ' retrieve the Parent windows.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''
    ' VBEditor
    '''''''''''''''''''''
    ClassName = GetWindowClassName(Application.VBE.MainWindow.hWnd)
    Debug.Print "VBEditor -- HWnd: " & CStr(Application.VBE.MainWindow.hWnd) & _
                                                                             "  (Window Class: " & ClassName & ")"
    '''''''''''''''''''''
    ' Windows Desktop
    '''''''''''''''''''''
    ClassName = GetWindowClassName(GetDesktopWindow())
    Debug.Print "Windows Desktop -- HWnd: " & CStr(GetDesktopWindow()) & _
                                                                       "  (Window Class: " & ClassName & ")"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' frmSetParent UserForm.   Class Name "ThunderDFrame".
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ParentHWnd = GetWindowLong(UserFormHWnd, GWL_HWNDPARENT)
    ClassName = GetWindowClassName(UserFormHWnd)
    ParentWindowClass = GetWindowClassName(ParentHWnd)
    Debug.Print "UserForm -- HWnd: " & CStr(UserFormHWnd) & " (Window Class: " & ClassName & _
                                                          ")   Parent HWnd: " & CStr(ParentHWnd) & "  (Window Class: " & ParentWindowClass & ")"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ActiveWindow. Class Name "EXCEL7".
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ParentHWnd = GetWindowLong(ActiveWindowHWnd, GWL_HWNDPARENT)
    ClassName = GetWindowClassName(ActiveWindowHWnd)
    ParentWindowClass = GetWindowClassName(ParentHWnd)
    Debug.Print "ActiveWindow -- HWnd: " & CStr(ActiveWindowHWnd) & " (Window Class: " & ClassName & _
                                                                  ")   Parent HWnd: " & CStr(ParentHWnd) & "  (Window Class: " & ParentWindowClass & ")"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Information About The Excel Desktop. Class Name "XLDESK"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ParentHWnd = GetWindowLong(ApplicationHWnd, GWL_HWNDPARENT)
    ParentHWnd = GetWindowLong(ExcelDeskHWnd, GWL_HWNDPARENT)
    ClassName = GetWindowClassName(ExcelDeskHWnd)
    ParentWindowClass = GetWindowClassName(ParentHWnd)
    Debug.Print "Excel Desktop -- HWnd: " & CStr(ExcelDeskHWnd) & " (Window Class: " & ClassName & _
                                                                ")   Parent HWnd: " & CStr(ParentHWnd) & "  (Window Class: " & ParentWindowClass & ")"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Information About The Application Window. Class Name "XLMAIN"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ParentHWnd = GetWindowLong(ApplicationHWnd, GWL_HWNDPARENT)
    ClassName = GetWindowClassName(ApplicationHWnd)
    ParentWindowClass = GetWindowClassName(ParentHWnd)
    Debug.Print "Excel Application -- HWnd: " & CStr(ApplicationHWnd) & " (Window Class: " & ClassName & _
                                                                      ")   Parent HWnd: " & CStr(ParentHWnd) & "  (Window Class: " & ParentWindowClass & ")"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Display information about the various Ancestor values: GA_ROOT, GA_ROOTOWNER, and GA_PARENT.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    AncestorWindow = GetAncestor(UserFormHWnd, GA_ROOT)
    ClassName = GetWindowClassName(AncestorWindow)
    s = "The Ancestor (GA_ROOT) of this UserForm is " & CStr(AncestorWindow) & "  (Window Class: " & ClassName & ")"
    AncestorWindow = GetAncestor(UserFormHWnd, GA_ROOTOWNER)
    ClassName = GetWindowClassName(AncestorWindow)
    s = s & vbCrLf & "The Ancestor (GA_ROOTOWNER) of this UserForm is " & CStr(AncestorWindow) & "  (Window Class: " & ClassName & ")"
    AncestorWindow = GetAncestor(UserFormHWnd, GA_PARENT)
    ClassName = GetWindowClassName(AncestorWindow)
    Debug.Print s & vbCrLf & "The Ancestor (GA_PARENT) of this UserForm is " & CStr(AncestorWindow) & "  (Window Class: " & ClassName & ")"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Display informationa about this form's owner.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    OwnerWindow = GetWindow(UserFormHWnd, GW_OWNER)
    If OwnerWindow Then
        Debug.Print "The Owner Window of this UserForm is HWnd: " & CStr(OwnerWindow) _
      & "  (Window Class: " & GetWindowClassName(OwnerWindow) & ")"
    Else
        Debug.Print "There is no owner window of this form."
    End If
    Form.Repaint
End Sub

Public Sub MakeFormChildOfVBEditor(Form As Object)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' MakeFormChildOfVBEditor
    ' This makes the form a child window of the VBEditor window (this is what you see if you have
    ' no workbooks open or all workbooks minimized).
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#INCLUDE DisplayErrorText
    Dim Res As Long
    Dim ParentHWnd As Long
    Dim ChildHWnd As Long
    Dim ErrNum As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set Parent HWnd to ActiveWindowHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ParentHWnd = VBEditorHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set Child HWnd to UserFormHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ChildHWnd = UserFormHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Call SetParent to make ChildHWnd a child of
    ' ParentHWnd.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Res = SetParent(hwndChild:=ChildHWnd, hWndNewParent:=ParentHWnd)
    If Res = 0 Then
        '''''''''''''''''''''''''''''
        ' an error occurred
        '''''''''''''''''''''''''''''
        ErrNum = Err.LastDllError
        DisplayErrorText "Error With SetParent", ErrNum
    Else
        Debug.Print "The UserForm is a child of the VBEditor window (Class wndclass_desked_gsk). "
    End If
    SetForegroundWindow UserFormHWnd
    Form.Repaint
End Sub

Public Sub MakeFormChildOfActiveWindow(Form As Object)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' MakeFormChildOfActiveWindow
    ' This makes the form a child window of the ActiveWindow.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#INCLUDE DisplayErrorText
    Dim Res As Long
    Dim ParentHWnd As Long
    Dim ChildHWnd As Long
    Dim ErrNum As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set Parent HWnd to ActiveWindowHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''
    ' Updete the ActiveWindowHWnd
    '''''''''''''''''''''''''''''''
    If Application.ActiveWindow Is Nothing Then
        MsgBox "There is no active window."
        Exit Sub
    End If
    ActiveWindowHWnd = FindWindowEx(hWnd1:=ExcelDeskHWnd, hWnd2:=0&, lpsz1:=C_EXCEL_WINDOW_WINDOWCLASS, _
                                    lpsz2:=Application.ActiveWindow.Caption)
    ParentHWnd = ActiveWindowHWnd
    If ParentHWnd = 0 Then
        MsgBox "ParentHWnd Is 0 In MakeFormChildOfActiveWindow"
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set Child HWnd to UserFormHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ChildHWnd = UserFormHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Call SetParent to make ChildHWnd a child of
    ' ParentHWnd.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Res = SetParent(hwndChild:=ChildHWnd, hWndNewParent:=ParentHWnd)
    If Res = 0 Then
        '''''''''''''''''''''''''''''
        ' an error occurred
        '''''''''''''''''''''''''''''
        ErrNum = Err.LastDllError
        DisplayErrorText "Error With SetParent", ErrNum
    Else
        Debug.Print "The UserForm is a child of the ActiveWindow (Class: EXCEL7). Note that you cannot move the" & _
                    " the form outside of the ActiveWindow, and that the form moves as you move the ActiveWindow. If" & _
                    " you switch to another window such as another workbook, the form is not be visible until you restore" & _
                    " the original window. Note that it is not possible to make a form a child of an individual worksheet tab."
    End If
    SetForegroundWindow UserFormHWnd
    Form.Repaint
End Sub

Public Sub MakeFormChildOfDesktop(Form As Object)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' MakeFormChildOfDesktop
    ' This makes the form a child window of the Excel Desktop (this is what you see if you have
    ' no workbooks open or all workbooks minimized).
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#INCLUDE DisplayErrorText
    Dim Res As Long
    Dim ParentHWnd As Long
    Dim ChildHWnd As Long
    Dim ErrNum As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set Parent HWnd to ActiveWindowHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ParentHWnd = ExcelDeskHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set Child HWnd to UserFormHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ChildHWnd = UserFormHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Call SetParent to make ChildHWnd a child of
    ' ParentHWnd.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Res = SetParent(hwndChild:=ChildHWnd, hWndNewParent:=ParentHWnd)
    If Res = 0 Then
        '''''''''''''''''''''''''''''
        ' an error occurred
        '''''''''''''''''''''''''''''
        ErrNum = Err.LastDllError
        DisplayErrorText "Error With SetParent", ErrNum
    Else
        Debug.Print "The UserForm is a child of the Excel Desktop (Class XLDESK). The window may get lost behind the" & _
                    " worksheet windows. In general, you'll never want to make the form a child of Excel Desktop unless you" & _
                    " don't have any open workbooks, in which case it is better to make a form a child of the Application." & _
                    " If the window gets lost, click on the Show Form button on Sheet1 to restore the form. The form will" & _
                    " still be displayed if you minimize all open windows. Note that you cannot drag the form outside of the " & _
                    " Excel Desktop's window."
    End If
    SetForegroundWindow UserFormHWnd
    Form.Repaint
End Sub

Public Sub MakeFormChildOfApplication(Form As Object)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' MakeFormChildOfApplication
    ' This makes the form a child of the main application window.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#INCLUDE Main
    '#INCLUDE DisplayErrorText
    Dim Res As Long
    Dim ParentHWnd As Long
    Dim ChildHWnd As Long
    Dim ErrNum As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set Parent HWnd to ActiveWindowHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ParentHWnd = ApplicationHWnd
    If ParentHWnd = 0 Then
        MsgBox "ParentHWnd Is 0 In MakeFormChildOfApplication."
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set Child HWnd to UserFormHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ChildHWnd = UserFormHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Call SetParent to make ChildHWnd a child of
    ' ParentHWnd.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Res = SetParent(hwndChild:=ChildHWnd, hWndNewParent:=ParentHWnd)
    If Res = 0 Then
        '''''''''''''''''''''''''''''
        ' an error occurred
        '''''''''''''''''''''''''''''
        ErrNum = Err.LastDllError
        DisplayErrorText "Error With SetParent", ErrNum
    Else
        Debug.Print "The UserForm is a child of the Excel Application (Class XLMAIN). Note that the form will be visible even" & _
                    " as you open and close windows, or minimize windows. If you restore the Excel window and move it around on the" & _
                    " screen, the form will move with the Application window."
    End If
    SetForegroundWindow UserFormHWnd
    Form.Repaint
End Sub

Public Sub MakeFormChildOfNothing(Form As Object)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' MakeFormChildOfNothing
    ' Sets the parent of the form to 0&
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#INCLUDE Main
    '#INCLUDE DisplayErrorText
    Dim Res As Long
    Dim ParentHWnd As Long
    Dim ChildHWnd As Long
    Dim ErrNum As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set Parent HWnd to ActiveWindowHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    'ParentHWnd = 0&
    ParentHWnd = GetDesktopWindow()
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set Child HWnd to UserFormHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ChildHWnd = UserFormHWnd
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Call SetParent to make ChildHWnd a child of
    ' ParentHWnd.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Res = SetParent(hwndChild:=ChildHWnd, hWndNewParent:=ParentHWnd)
    'Res = SetWindowLong(hwnd:=ChildHWnd, nIndex:=GWL_HWNDPARENT, dwNewLong:=0&)
    If Res = 0 Then
        '''''''''''''''''''''''''''''
        ' an error occurred
        '''''''''''''''''''''''''''''
        ErrNum = Err.LastDllError
        DisplayErrorText "Error With SetParent", ErrNum
    Else
        Debug.Print "The UserForm is a child of the Windows Desktop (see the GA_PARENT item above -- it has the same" & _
                    " window handle as the Windows Desktop). The Parent Window shows as XLMAIN because XLMAIN is the " & _
                    " owner of the window (see the GA_ROOTOWNER item above.)  Note that you can move the Excel Application" & _
                    " window and the form will remain at its original location on the screen. You can also move the" & _
                    " form outside of the Application's main window. This is the default behavior of an Excel Userform."
    End If
    SetForegroundWindow UserFormHWnd
    Form.Repaint
End Sub

Public Function GetSystemErrorMessageText(ErrorNumber As Long) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' GetSystemErrorMessageText
    '
    ' This function gets the system error message text that corresponds to the error code returned by the
    ' GetLastError API function or the Err.LastDllError property. It may be used ONLY for these error codes.
    ' These are NOT the error numbers returned by Err.Number (for these errors, use Err.Description to get
    ' the description of the message).
    ' The error number MUST be the value returned by GetLastError or Err.LastDLLError.
    '
    ' In general, you should use Err.LastDllError rather than GetLastError because under some circumstances the value of
    ' GetLastError will be reset to 0 before the value is returned to VB. Err.LastDllError will always reliably return
    ' the last error number raised in a DLL.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#INCLUDE TrimToNull
    Dim ErrorText As String
    Dim TextLen As Long
    Dim FormatMessageResult As Long
    Dim LangID As Long
    LangID = 0&
    ErrorText = String$(FORMAT_MESSAGE_TEXT_LEN, " ")
    TextLen = Len(ErrorText)
    On Error Resume Next
    FormatMessageResult = 0&
    FormatMessageResult = FormatMessage( _
                          dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                          lpSource:=0&, _
                          dwMessageId:=ErrorNumber, _
                          dwLanguageId:=0&, _
                          lpBuffer:=ErrorText, _
                          nSize:=TextLen, _
                          Arguments:=0&)
    On Error GoTo 0
    If FormatMessageResult > 0 Then
        ErrorText = TrimToNull(ErrorText)
        GetSystemErrorMessageText = ErrorText
    Else
        ' Format message didn't return any text. there is no text description for the specified error.
        GetSystemErrorMessageText = "NO ERROR DESCRIPTION AVAILABLE"
    End If
End Function

Public Sub DisplayErrorText(Context As String, ErrNum As Long)
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Displays a standard error message box. For this
    ' procedure, ErrNum should be the number returned
    ' by the GetLastError API function or the value
    ' of Err.LastDllError. It is NOT the number
    ' returned by Err.Number.
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    '#INCLUDE GetSystemErrorMessageText
    Dim ErrText As String
    ErrText = GetSystemErrorMessageText(ErrNum)
    MsgBox Context & vbCrLf & _
           "Error Number: " & CStr(ErrNum) & vbCrLf & _
           "Error Text:   " & ErrText, vbOKOnly
End Sub

Public Function TrimToNull(TEXT As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' TrimToNull
    ' Returns all the text in Text to the left of the vbNullChar
    ' character.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Pos As Integer
    Pos = InStr(1, TEXT, vbNullChar, vbTextCompare)
    If Pos > 0 Then
        TrimToNull = Left(TEXT, Pos - 1)
    Else
        TrimToNull = TEXT
    End If
End Function

Public Function GetWindowClassName(hWnd As Long) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' GetWindowClassName
    ' Returns the window class name of the specified HWnd. Returns
    ' vbNullString if an error occurred.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#INCLUDE DisplayErrorText
    '#INCLUDE TrimToNull
    Dim ClassName As String
    Dim Length As Long
    Dim Res As Long
    If hWnd = 0 Then
        GetWindowClassName = "<none>"
        Exit Function
    End If
    ClassName = String$(MAX_PATH, vbNullChar)
    Length = Len(ClassName)
    Res = GetClassName(hWnd:=hWnd, lpClassName:=ClassName, nMaxCount:=Length)
    If Res = 0 Then
        DisplayErrorText Context:="Error Retrieiving Window Class for HWnd: " & CStr(hWnd), _
        ErrNum:=Err.LastDllError
        GetWindowClassName = vbNullString
    Else
        ClassName = TrimToNull(ClassName)
        GetWindowClassName = ClassName
    End If
End Function

Public Function GetHWndWindowText(hWnd As Long) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' GetHWndWindowText
    ' This returns the WindowText of the HWnd.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#INCLUDE TrimToNull
    Dim txt As String
    Dim Res As Long
    Dim l As Long
    txt = String$(1024, vbNullChar)
    l = Len(txt)
    Res = GetWindowText(hWnd, txt, l)
    If Res Then
        txt = TrimToNull(txt)
        If txt = vbNullString Then
            txt = "<none>"
        End If
    Else
        txt = vbNullString
    End If
    GetHWndWindowText = txt
End Function

Public Function GetParentWindowClass(hWnd As Long) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' GetParentWindowClass
    ' Returns the window class name of the parent window of HWnd.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '#INCLUDE DisplayErrorText
    '#INCLUDE GetWindowClassName
    Dim ParentHWnd As Long
    Dim ClassName As String
    ParentHWnd = GetWindowLong(hWnd:=hWnd, nIndex:=GWL_HWNDPARENT)
    If ParentHWnd = 0 Then
        DisplayErrorText Context:="Error Retrieiving Parent Window for HWnd: " & CStr(hWnd) & _
                                                                                            " Window Class: " & GetWindowClassName(hWnd), ErrNum:=Err.LastDllError
        GetParentWindowClass = vbNullString
        Exit Function
    End If
    ClassName = GetWindowClassName(ParentHWnd)
    GetParentWindowClass = ClassName
End Function


