VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
'
' Application Error Class
' |> Brian J. Satola
' |> https://chejunkie.com
'
'==========================================================================================================================
' This class is an error container.
'==========================================================================================================================
' It can be added to other classes instead of creating separate properties etc. that handle errors in the same way.
'--------------------------------------------------------------------------------------------------------------------------
'
'==========================================================================================================================
' Public Subs, Functions and Properties - oh my!
'==========================================================================================================================
'   |> Get | --- About                  :: Returns description of the class.
'   |> --------- Clear                  :: Clears all error properties.
'   |> --------- Define                 :: Defines an error.
'   |> Get | Set Description            :: Description of an error Number. (NumberDLL description is appended.)
'   |> --------- DisplayMessage         :: Displays error information in a message box.
'   |> --------- HasError
'   |> --------- Initialize
'   |> Get | Set Number
'   |> Get | Set NumberDLL
'   |> Get | --- ParentName
'   |> --------- PrintMessage
'   |> Get | Set Source
'   |> Get | --- Version                :: Returns version string for the class, ##.## (yyyy).
'
'==========================================================================================================================
' Application constants
'==========================================================================================================================
Private Const C_NAME As String = "ApplicationError.cls"
Const C_ERR_OFFSET = 100
'==========================================================================================================================
' Error enum
'==========================================================================================================================
Public Enum ERR_ENUM
    C_ERR_NO_ERROR = 0
    C_ERR_INVALID_FILE = C_ERR_OFFSET + vbObjectError + 1
    C_ERR_INVALID_DIRECTORY = C_ERR_OFFSET + vbObjectError + 2
    C_ERR_LOG_FILE_DNE = C_ERR_OFFSET + vbObjectError + 3
    C_ERR_INVALID_INPUT = C_ERR_OFFSET + vbObjectError + 4
    C_ERR_TYPE_MISMATCH = C_ERR_OFFSET + vbObjectError + 5
End Enum
'==========================================================================================================================
' Private variables
'==========================================================================================================================
Private m_ErrNumber As Long
Private m_ErrNumberDLL As Long
Private m_ErrSource As String
Private m_ErrDescription As String
Private m_ErrDescriptionDLL As String
Private m_ParentName As String
Private m_HasError  As Boolean
'==========================================================================================================================
' API Constants
'==========================================================================================================================
' GetSystemErrorMessageText
' |> www.cpearson.com/Excel/FormatMessage.aspx
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Long = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE As Long = &H800
Private Const FORMAT_MESSAGE_FROM_STRING As Long = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Long = &HFF
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
Private Const FORMAT_MESSAGE_TEXT_LEN As Long = &HA0                  ' from VC++ ERRORS.H file

'==========================================================================================================================
' API Declarations
'==========================================================================================================================
' GetSystemErrorMessageText
' |> www.cpearson.com/Excel/FormatMessage.aspx
Private Declare PtrSafe Function FormatMessage Lib "kernel32" _
        Alias "FormatMessageA" ( _
        ByVal dwFlags As Long, _
        ByVal lpSource As Any, _
        ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long, _
        ByRef Arguments As Long) As Long
' ToggleTrustAccess
' |> https://wellsr.com/vba/2017/excel/macro-to-enable-trust-access-to-the-VBA-project-object-model/
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'''
''' Clears all property settings of the error object.
'''
''' The parent name is not cleared since this is often set just once at the
''' beginning i.e. using the Initialize method.
'''
Public Sub clear()
    m_ErrNumber = C_ERR_NO_ERROR
    m_ErrNumberDLL = C_ERR_NO_ERROR
    m_ErrDescription = ""
    m_ErrSource = ""
    m_HasError = False
End Sub

'''
''' Defines an error.
'''
Public Sub Define(errNumber As Long, errDescription As String, Optional errSource As String)
    clear
    Me.Number = errNumber
    m_ErrDescription = errDescription
    m_ErrSource = errSource
End Sub

'''
''' Returns a descriptive string associated with an error.
''' Instead of having a seperate property for DLL errors, I simply append the result.
'''
Public Property Get Description() As String
    Description = m_ErrDescription
End Property

'''
''' Sets a descriptive string associated with an error.
'''
Property Let Description(str As String)
    m_ErrDescription = str
End Property

'''
''' Displays error details in a message-box: error number, description and source (if available).
''' Option to display custom title using [displayTitle]; else, it is automatically generated.
''' Option to append custom message to error text using [appendMessage].
''' Option to define message box behaviour using [iType] using vbMsgBoxStyle enums.
'''
Public Sub DisplayMessage(Optional displayTitle As String, Optional appendMessage As String, Optional msgBoxStyle As VbMsgBoxStyle = vbExclamation)
    If (displayTitle = "") Then        ' define display title
        displayTitle = "--ERROR!--"
        '--If (m_ErrNumber <> C_ERR_NO_ERROR) And (m_ParentName <> "") Then ' parent error title.
        '--    displayTitle = "--ERROR!-- " & m_ParentName
        '--ElseIf (m_ErrNumber <> C_ERR_NO_ERROR) Then ' error object title
        '--    displayTitle = "--ERROR!-- " & C_NAME
        '--ElseIf (m_ParentName <> "") Then ' parent
        '--    displayTitle = m_ParentName
        '--Else ' error object for everything else
        '--    displayTitle = C_NAME
        '--End If
    End If
    '// Display message to user.
    If (appendMessage <> vbNullString) Then    ' append message to error information
        MsgBox GetDescription & VBA.vbCrLf & VBA.vbCrLf & appendMessage, Title:=displayTitle, Buttons:=msgBoxStyle
    Else
        MsgBox GetDescription, Title:=displayTitle, Buttons:=msgBoxStyle
    End If
End Sub

'
' Returns a detailed error description: source, number and description (including errors from DLL).
'
Private Function GetDescription() As String
    '// Define detailed error description.
    If (m_ParentName <> "") Then
        GetDescription = "|> Parent: " & m_ParentName & VBA.vbCrLf
    End If
    If (m_ErrSource <> "") Then
        GetDescription = GetDescription & _
                "|> Source: " & m_ErrSource & VBA.vbCrLf & _
                "|> Number: " & m_ErrNumber & VBA.vbCrLf & _
                "|> " & m_ErrDescription
    Else
        GetDescription = GetDescription & _
                "|> Number: " & m_ErrNumber & VBA.vbCrLf & _
                "|> " & m_ErrDescription
    End If
    '// Append DLL error information.
    If (m_ErrDescriptionDLL <> "") Then
        GetDescription = GetDescription & VBA.vbCrLf & VBA.vbCrLf & _
                String(50, "-") & VBA.vbCrLf & _
                "|> DLL error number: " & m_ErrNumberDLL & VBA.vbCrLf & _
                "|> " & m_ErrDescription
    End If
End Function

'''
''' Returns a string describing the class.
'''
Public Property Get About() As String
    About = "ChE Junkie VBA Application Error class module, Version " & Me.version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details see:" & VBA.vbCrLf & "https://chejunkie.com/knowledge-base/application-error-class-vba"
End Property

'''
''' Returns the name of the class.
'''
Public Property Get Name() As String
    Name = C_NAME
End Property

'''
''' Returns version information for the class.
'''
Public Property Get version() As String
    ' Contains a list of (historical) changes to the class within the comments of the procedure.
    '--------------------------------------------------------------------------------------------------------------------------
    version = "Version 1.0 (2017)"        ' Initial (official) release.
End Property

'''
''' Will return True if 'Number' and/or 'NumberDLL' error exists.
'''
Property Get HasError() As Boolean
    HasError = m_HasError
End Property

'''
''' The parent name is something that is typically initialized just one time.
'''
Public Sub Initialize(parentName_ As String)
    clear
    m_ParentName = parentName_
End Sub

'''
''' Returns a numeric value specifying an error.
'''
Public Property Get Number() As Long
    Number = m_ErrNumber
End Property

'''
''' Returns a numeric value specifying a DLL error.
'''
Public Property Get NumberDLL() As Long
    NumberDLL = m_ErrNumberDLL
End Property

'''
''' This function gets the system error message text that corresponds to the error code *errNumber*.
''' This value is the value returned by Err.LastDLLError or by GetLastError, or occasionally as the
''' returned result of a Windows API function.
'''
''' These are NOT the error numbers returned by Err.Number (for these errors, use Err.Description
''' to get the description of the error).
'''
''' In general, you should use Err.LastDllError rather than GetLastError because under some
''' circumstances the value of GetLastError will be reset to 0 before the value is returned to VBA.
''' Err.LastDllError will always reliably return the last error number raised in an API function.
'''
''' The function returns vbNullString is an error occurred or if there is no error text for the
''' specified error number.
'''
Private Function GetSystemErrorMessageText(errNumber As Long) As String
    '@INCLUDE Init
    Dim ErrorText   As String
    Dim textLen     As Long
    Dim FormatMessageResult As Long
    Dim langID      As Long
    ''''''''''''''''''''''''''''''''
    ' Initialize the variables
    ''''''''''''''''''''''''''''''''
    langID = 0&        ' Default language
    ErrorText = VBA.String$(FORMAT_MESSAGE_TEXT_LEN, vbNullChar)
    textLen = FORMAT_MESSAGE_TEXT_LEN
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Call FormatMessage to get the text of the error message text
    ' associated with errNumber.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    FormatMessageResult = FormatMessage( _
            dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or _
            FORMAT_MESSAGE_IGNORE_INSERTS, _
            lpSource:=0&, _
            dwMessageId:=errNumber, _
            dwLanguageId:=langID, _
            lpBuffer:=ErrorText, _
            nSize:=textLen, _
            Arguments:=0&)
    If FormatMessageResult = 0& Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' An error occured. Display the error number, but
        ' don't call GetSystemErrorMessageText to get the
        ' text, which would likely cause the error again,
        ' getting us into a loop.
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        MsgBox "An error occurred with the FormatMessage" & _
                " API function call." & vbCrLf & _
                "Error: " & CStr(Err.LastDllError) & _
                " VBA.Hex(" & VBA.Hex(Err.LastDllError) & ")."
        GetSystemErrorMessageText = "An internal system error occurred with the" & vbCrLf & _
                "FormatMessage API function: " & CStr(Err.LastDllError) & ". No futher information" & vbCrLf & _
                "is available."
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' If FormatMessageResult is not zero, it is the number
    ' of characters placed in the ErrorText variable.
    ' Take the left FormatMessageResult characters and
    ' return that text.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ErrorText = VBA.Left$(ErrorText, FormatMessageResult)
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Get rid of the trailing vbCrLf, if present.
    '''''''''''''''''''''''''''''''''''''''''''''
    If VBA.Len(ErrorText) >= 2 Then
        If VBA.Right$(ErrorText, 2) = vbCrLf Then
            ErrorText = VBA.Left$(ErrorText, VBA.Len(ErrorText) - 2)
        End If
    End If
    ''''''''''''''''''''''''''''''''
    ' Return the error text as the
    ' result.
    ''''''''''''''''''''''''''''''''
    GetSystemErrorMessageText = ErrorText
End Function

'''
''' Sets a numeric value specifying an error.
'''
Property Let Number(lng As Long)
    If (lng <> C_ERR_NO_ERROR) Then
        m_ErrNumber = lng
        m_HasError = True
        Description = GetAppErrDescription(lng)
    End If
End Property

'''
''' Sets a numeric value specifying a DLL error.
'''
Property Let NumberDLL(lng As Long)
    If (lng <> C_ERR_NO_ERROR) Then
        m_ErrNumberDLL = lng
        m_ErrDescriptionDLL = GetSystemErrorMessageText(lng)
        m_HasError = True
    End If
End Property

'''
''' Returns the name of the parent that the error object belongs to.
'''
Private Property Get ParentName() As String
    ParentName = m_ParentName
End Property

'
' Prints error details to the immediate window: error number, description and source (if available).
'
Public Sub PrintMessage()
    Debug.Print ""
    Debug.Print String(50, "=")
    Debug.Print " Application Error"
    Debug.Print String(50, "=")
    Debug.Print GetDescription
End Sub

'''
''' Returns the name of the method, object etc. that originally generated the error.
'''
Public Property Get Source() As String
    Source = m_ErrSource
End Property

'''
''' Sets the name of the method, object etc. that originally generated the error.
'''
Property Let Source(str As String)
    m_ErrSource = str
End Property

Private Sub Class_Initialize()
    Debug.Print "|* Initializing Class:= " & C_NAME
End Sub
'''
''' Returns the text description of the application error number.
'''
Public Function GetAppErrDescription(errNumber As Long) As String

    Select Case errNumber
        Case C_ERR_NO_ERROR: GetAppErrDescription = vbNullString
        Case C_ERR_INVALID_FILE: GetAppErrDescription = "Invalid file."
        Case C_ERR_INVALID_DIRECTORY: GetAppErrDescription = "Invalid directory."
        Case C_ERR_LOG_FILE_DNE: GetAppErrDescription = "Logfile DNE."
        Case C_ERR_INVALID_INPUT: GetAppErrDescription = "Invalid input."
        Case C_ERR_TYPE_MISMATCH: GetAppErrDescription = "Type mismatch."
        Case Else
            GetAppErrDescription = "Undefined error."
    End Select
End Function
