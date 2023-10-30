VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem INCLUDE class  ApplicationError
Rem INCLUDE class  RegistryEditor
Rem INCLUDE module mLogFile

Option Explicit
Option Base 1
'
' Log File Class
' |> Brian J. Satola
' |> https://chejunkie.com
' |> brian.satola@gmail.com
'
'==========================================================================================================================
' This class provides several methods for logging intermediate calculation results.
'==========================================================================================================================
' These routines rely on each other, so you should import this entire class module into your project
'   rather than just copying and pasting individual procedures into other modules.
' The Registry Editor class and Application Error class are also required.
'--------------------------------------------------------------------------------------------------------------------------
'
'==========================================================================================================================
' Public Subs, Functions and Properties - oh my!
'==========================================================================================================================
'   |> Get | --- About                  :: Returns description of the class.
'   |> --------- CreateLogFile          :: Creates a logfile for output and stores it in the specified [storageDirectory].
'   |> Get | Let DirectoryPath          :: Returns the directory location of the current logfile.
'   |> Get | --- FilePath               :: Returns the file path of the current logfile.
'   |> --------- PrintBlankLine           :: Outputs a blank line to the logfile.
'   |> Get | --- LogFileName            :: Returns the file name of the current logfile.
'   |> --------- LogDividingLine        :: Outputs a line to the logfile.
'   |> --------- Log                    :: Outputs strOutputMessage to the logfile.
'   |> Get | --- Name                   :: Returns name of the class.
'   |> --------- OpenDirectory          :: Opens the directory that the current logfile is stored in.
'   |> --------- OpenLogFile            :: Opens the current logfile (iff it exists).
'   |> Get | Let RegDirectoryScratch    :: Retrieves or sets the output directory used to save the logfile.
'   |> --------- RemoveRegistrySetting  :: Removes the registry setting for the default save location.
'   |> Get | --- Version                :: Returns version string for the class [e.g. #.# (year)].
'
'==========================================================================================================================
' References
'==========================================================================================================================
Dim m_AppErr        As New ApplicationError

'==========================================================================================================================
' Application constants
'==========================================================================================================================

Private Const C_NAME As String = "LogFile.cls"
Private Const C_DEFAULT_LOG_FILE_NAME As String = "Logger"
Private Const C_MAX_LEN As Integer = 96                 ' size/length of header boarders. (MUST BE AN EVEN NUMBER!)
' 96 seems to be a good size for wide-screen, else 90 or lower.

Private Enum LogOutputType
    LogAppend = 1
    LogOutput = 2
End Enum

'==========================================================================================================================
' Private variables
'==========================================================================================================================
Private m_LogFileExists As Boolean
Private m_FilePath  As String
Private m_DirectoryPath As String
Private m_FileName  As String
Private m_FileExt   As String


Const C_LINE_LENGTH As Integer = 110



'''
''' Version string of the current class.
'''
Public Property Get version() As String
    '==========================================================================================================================
    ' Class changelog here.
    '==========================================================================================================================

    'Version = "Version 1.0 (2015)"  ' Initial (official) release.
    'Version = "Version 2.0 (2015)"  ' Cleaned up code, including procedure descriptions.
    '   |> Added "About" Get property for use in e.g. display messages.
    '   |> Added "Version" Get property to track class changes.
    '   |> Added "Name" Get property (literal defintion).
    '   |> Added 'C_MAX_LEN' and "GetBufferString" to generalise the procedures for
    '       use in defining header titles, line outputs, etc.
    '   |> Updated "Log" procedure to use "GetBufferString" instead of literals.
    '   |> Updated "GetAppErrDisplayMessage" to include titles in dislpay messages.
    '   |> Removed the 'RegistryTRUxl.cls' reference; instead, directly handle read/write
    '       to the srcratch-directory location defined in the registry. This required
    '       adding a local reference to 'Registry.cls', and adding the appropriate
    '       property Get|Let procedures for "RegDirectoryScratch".
    '   |> Added "CreateFolder" procedure.
    '   |> Updated class constructor. If no valid "RegDirectoryScratch" definition exists,
    '       then "CreateFolder" is used to automatically create a scratch folder in the
    '       directory of the workbook to use for outputs. "RegDirectoryScratch" is then
    '       updated to match the new definition. If a valid "RegDirectoryScratch"
    '       defintion exists, then no changes are made.
    'Version = "Version 2.5 (2016)" ' SATOLA, BRIAN J. 26-AUG. 2016:
    '   |> Syntax corrections to module-level variables (prefix:= "m_").
    '   |> Changed enumeration names and values to use Pascal casing.
    '   |> Updated constructor print statement to use new "Name" property.
    '   |> Re-simplified ;) "AppErrDisplayMessage".
    '   |> Updated 'm_AppErr.Source =' definitions (simple formatting changes).
    '   |> FIXED CRITICAL ERROR in class constructor made during version 2 release.
    '       "RegDirectoryScratch" was returning Null to 'm_directoryPath',
    '       a string variable, which was preventing the routine from auto-creating
    '       a directory storage location for the logfile outputs.
    '   |> Fixed formatting problem for titles in the Log method.
    'Version = "3.0 (2017+)"  ' SATOLA, BRIAN J.
    '   |> Updated to use newest registry class (now called RegistryEditor).
    '   |> Removed the Let FileName property, CreateLogFile method already defines the name.
    '   |> Get FileName property is now called FileName_ to avoid naming conflicts.
    '   |> Reformated comments to new convention (for later automation).
    '   |> Replaced application error properties and methods with a new error object/holder.
    '   |> Renamed GetPaddedString to GetBufferString (which is more descriptive).
    '   |> Removed the 'Version' string from the get Version property (more intuitive this way).
    '   |> Added RemoveRegistrySetting() for testing purposes.
    version = "3.1 (2018)"  ' SATOLA, BRIAN J.
    '   |> Automatically define default file name if user does not provide:
    '       C_DEFAULT_LOG_FILE_NAME.

End Property
'''
''' Returns a string describing the class.
'''
Public Property Get About() As String
    About = "ChE Junkie VBA LogFile class module, v" & Me.version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details see https://chejunkie.com/knowledge-base/the-log-file-class-vba/"
End Property
'''
''' Returns the name of the class.
'''
Public Property Get Name() As String
    Name = C_NAME
End Property
'''
''' Returns the application error object of the class.
'''
Public Property Get AppErr() As ApplicationError
    Set AppErr = m_AppErr
End Property

'''
''' Initialize the class
'''
Public Sub Init(Optional storageDirectory As String = vbNullString, Optional ByRef LogFileName As String)
    '//// Class constructor.
    '--------------------------------------------------------------------------------------------------------------------------
    Debug.Print "|> Initializing:= " & C_NAME
    m_AppErr.Initialize C_NAME
    m_DirectoryPath = IniReadKey(ThisWorkbook.path & "\Lib\config\config.ini", "LogFile.cls", "FilePath")
    If Not FileFolderExists(m_DirectoryPath) Then    'key DNE, or invalid path.
        If Not FileFolderExists(ThisWorkbook.path & "\Lib\Log\") Then
            CreateFolder ThisWorkbook.path & "\Lib\Log\"
        End If
        m_DirectoryPath = ThisWorkbook.path & "\Lib\Log\"
    End If
    m_FileName = vbNullString
    m_FilePath = vbNullString
    m_FileExt = ".txt"
    CreateLogFile storageDirectory, LogFileName
End Sub

Private Sub Class_Initialize()
    Init
End Sub

'''
''' Creates a log file for output and stores it in the specified [storageDirectory].
''' If no [storageDirectory] is provided, then the [DirectoryPath] property is used.
'''
Public Function CreateLogFile(Optional storageDirectory As String = vbNullString, Optional ByRef LogFileName As String) As Boolean

    Dim sfx         As String
    Dim i           As Integer

    '//========================================================================================================================
    '// Initialize
    '//========================================================================================================================

    m_LogFileExists = False
    m_AppErr.clear

    If ((storageDirectory = vbNullString) And (m_DirectoryPath <> vbNullString)) Then    ' use existing directory definition.
        storageDirectory = m_DirectoryPath    ' e.g. class initializes to registry pointer.
    End If

    If ((storageDirectory <> vbNullString) And FileFolderExists(storageDirectory)) Then    ' valid input definition.
        m_DirectoryPath = storageDirectory
        If (LogFileName <> "") Then
            m_FileName = LogFileName
        Else
            LogFileName = C_DEFAULT_LOG_FILE_NAME
            m_FileName = LogFileName
        End If
        m_FilePath = m_DirectoryPath & m_FileName & m_FileExt
    Else    ' display application error
        With m_AppErr
            .Number = C_ERR_INVALID_DIRECTORY
            .Source = "CreateLogFile(...)"
            .DisplayMessage
        End With
        Exit Function
    End If

    '==========================================================================================================================
    ' Define LogFileName
    '==========================================================================================================================

    '// Ensure unique name (use integer suffix until found)
'    Do Until (FileFolderExists(m_FilePath) = False)
'        i = i + 1
'        sfx = " (" & i & ")"
'        m_FilePath = m_DirectoryPath & m_FileName & sfx & m_FileExt
'    Loop
    m_FileName = m_FileName & sfx

    m_LogFileExists = True

    '//========================================================================================================================
    '// Print Logfile Header
    '//========================================================================================================================

    '// Function WriteToFile(ByVal strLogMessage As String, ByVal iTypeOutput As EOutputType) As Boolean
    Call WriteToFile("LogFile       :    " & m_FileName, LogAppend)
    Call WriteToFile("Created       :    " & VBA.Now, LogAppend)
    Call WriteToFile(GetBufferString("_", C_MAX_LEN), LogAppend)
    Call WriteToFile("", LogAppend)

    CreateLogFile = True

End Function
'''
''' Creates a folder according to the specified <path> [fullPath].
'''
Private Function CreateFolder(ByVal fullPath As String) As Boolean

    If (fullPath <> vbNullString) Then    'create folder
        '// Ensure correct input format
        If (VBA.Right$(fullPath, 1) = "\") Then
            fullPath = VBA.Left$(fullPath, VBA.Len(fullPath) - 1)
        End If
        '// Create folder
        On Error GoTo ERR_HANDLER
        MkDir fullPath
        CreateFolder = True
        Exit Function
    End If

ERR_HANDLER:
End Function
'''
''' This method is just to change the order of the inputs for the CreateLogFile method.
''' Due to legacy support, the original method could not be directly changed etc.
'''
Public Function CreateLogFileByName(Optional ByRef LogFileName As String, Optional storageDirectory As String) As Boolean
    CreateLogFileByName = CreateLogFile(storageDirectory, LogFileName)
End Function
'''
''' Returns the file path of the current log file.
'''
Public Property Get FilePath() As String
    FilePath = m_FilePath
End Property
'''
''' Returns the file name of the current log file.
'''
Public Property Get fileName() As String
    fileName = m_FileName
End Property
'''
''' Returns the directory location for the current log file.
'''
Public Property Get DirectoryPath() As String
    DirectoryPath = m_DirectoryPath
End Property
'''
''' Sets the directory location in which the log file is saved.
'''
Private Property Let DirectoryPath(path As String)
    m_AppErr.clear
    If (VBA.Right(VBA.Trim(path), 1) <> "\") Then
        path = VBA.Trim(path) & "\"
    End If
    '// Check existence
    If (VBA.Dir(path, vbDirectory) = ".") Then
        m_DirectoryPath = path
    Else    ' display application error
        With m_AppErr
            .Number = C_ERR_INVALID_DIRECTORY
            .Source = "'Let' DirectoryPath"
            .DisplayMessage
        End With
    End If
End Property
'''
''' Function returns TRUE if the specified file or folder exists, FALSE if not.
'''
Private Function FileFolderExists(fullPath As String) As Boolean
    On Error Resume Next
    If Not (Dir(fullPath, vbDirectory) = vbNullString) And (fullPath <> vbNullString) Then
        FileFolderExists = True
    End If
End Function
'''
''' Opens the file/folder by the specified <path>.
'''
Private Sub OpenFileFolder(fullPath As String)
    m_AppErr.clear
    On Error Resume Next
    Excel.ActiveWorkbook.FollowHyperlink Address:=fullPath, NewWindow:=True
    Select Case Err.Number
        Case 0, -2147467260    '(-2147467260) results when someone hits cancel on Excel prompt.
            Exit Sub
        Case Else
            On Error Resume Next
            With m_AppErr
                If (VBA.InStr(VBA.Mid(fullPath, VBA.Len(fullPath) - VBA.Len(m_FileExt) - 1, VBA.Len(m_FileExt)), ".") <> 0) Then
                    .Number = C_ERR_INVALID_FILE
                Else
                    .Number = C_ERR_INVALID_DIRECTORY
                End If
                .Source = "OpenFileFolder(...)"
            End With
    End Select
End Sub
'''
''' Opens the directory location of the current Log File.
'''
Public Sub OpenDirectory()
    OpenFileFolder (DirectoryPath)
End Sub
'''
''' Opens the current log file (if it exists).
'''
Public Sub OpenLogFile()
    OpenFileFolder (m_FilePath)
End Sub









'''
''' Generates a string of singleCharaster(s) of length stringLength.
''' Used to generate outputs for e.g. header boarders.
'''
Private Function GetBufferString(singleCharacter As String, stringLength As Integer) As String
    Dim i           As Long
    m_AppErr.clear
    If (VBA.Len(singleCharacter) = 1) Then
        For i = 1 To stringLength
            GetBufferString = GetBufferString & singleCharacter
        Next i
    Else    'display application error
        With m_AppErr
            .Number = C_ERR_INVALID_INPUT
            .Source = "GetBufferString(...)"
            .DisplayMessage
        End With
    End If
End Function
'''
'''
'''
Public Sub Log(ByVal stringToLog As Variant)
    Dim strg As String
    If IsArray(stringToLog) Then
        strg = Join(stringToLog, vbNewLine)
    Else
        strg = stringToLog
    End If
    Call WriteToFile(CStr(stringToLog), LogAppend)
        
'        Case Else    ' display error message
'            With m_AppErr
'                .Number = C_ERR_INVALID_INPUT
'                .Source = "Log(...)"
'                .Description = GetAppErrDescription(.Number)
'                .DisplayMessage
'            End With
'            Exit Sub
End Sub
'''
'''
'''
Sub PrintLineThrough(ByVal stringToLog As Variant)
    Dim strg As String
    Dim nlen(2) As Integer
    nlen(1) = C_MAX_LEN - VBA.Len(VBA.Now & ": ")
    If (VBA.Len(stringToLog) <= nlen(1)) Then
        strg = VBA.Replace(stringToLog, " ", "-")
        strg = "---" & strg
        nlen(2) = nlen(1) - VBA.Len(stringToLog) - 3
        Dim ilen As Long
        For ilen = 1 To nlen(2)
            strg = strg & "-"
        Next ilen
        stringToLog = VBA.Now & ": " & strg
    End If
    Call WriteToFile(CStr(stringToLog), LogAppend)
End Sub
'''
''' Outputs a blank line to the log file.
'''
Public Sub PrintBlankLine()
    Call WriteToFile("", LogAppend)
End Sub
'''
''' Outputs a line to the Log file: solid (default) or dashed.
'''
Public Sub LogDividingLine(Optional includeBlankLineAfter As Boolean = True, Optional useDashLine As Boolean = False)
    If includeBlankLineAfter Then
        If Not useDashLine Then
            Call WriteToFile(GetBufferString("_", C_MAX_LEN) & VBA.vbCrLf, LogAppend)
        Else
            Call WriteToFile(GetBufferString("-", C_MAX_LEN) & VBA.vbCrLf, LogAppend)
        End If
    ElseIf Not useDashLine Then
        Call WriteToFile(GetBufferString("_", C_MAX_LEN), LogAppend)
    Else
        Call WriteToFile(GetBufferString("-", C_MAX_LEN), LogAppend)
    End If
End Sub
'''
''' Logs the stringToLog to the defined log file.
''' Formats the output according to LogFormatType.
'''
Sub LogTitle(ByVal stringToLog As Variant)
    m_AppErr.clear
    Dim strg As String
    Dim pads(2) As String
    Dim nlen(2) As Integer
    Dim ilen As Long
    Dim i   As Long
    '// Preformatting
    '   [1] Padding for "stringToLog line.
    '   [2] Padding for spacer line.
    Dim ll  As Long
    If IsArray(stringToLog) Then
        Dim element
        For Each element In stringToLog
            If Len(CStr(element)) > ll Then ll = Len(CStr(element))
        Next
    Else
        ll = Len(stringToLog)
    End If
    nlen(1) = C_MAX_LEN - (ll + 12)
    For ilen = 1 To nlen(1)
        pads(1) = pads(1) & "/"
    Next ilen
    pads(1) = "    " & pads(1)
    nlen(2) = ll
    For ilen = 1 To nlen(2)
        pads(2) = pads(2) & " "
    Next ilen
    '// Title definition.
    strg = VBA.vbCrLf & VBA.vbCrLf
    strg = strg & VBA.String(C_MAX_LEN, "/") & VBA.vbCrLf
    strg = strg & "////    " & pads(2) & pads(1) & VBA.vbCrLf
    If IsArray(stringToLog) Then
        For i = LBound(stringToLog) To UBound(stringToLog)
            strg = strg & "////    " & stringToLog(i) & Space(ll - Len(stringToLog(i))) & pads(1) & VBA.vbCrLf
        Next
    Else
        strg = strg & "////    " & stringToLog & pads(1) & VBA.vbCrLf
    End If
    strg = strg & "////    " & pads(2) & pads(1) & VBA.vbCrLf
    strg = strg & VBA.String(C_MAX_LEN, "/") & VBA.vbCrLf
    stringToLog = strg
    Call WriteToFile(CStr(stringToLog), LogAppend)
End Sub
'''
'''
'''
Sub LogHeader1(ByVal stringToLog As Variant)
    m_AppErr.clear
    '==================================================================================================================
    'LOG HEADER NO. 1
    '==================================================================================================================
    Dim strg As String
    '// Preformatting
    strg = VBA.vbCrLf
    '// Header definition
    strg = strg & GetBufferString("=", C_MAX_LEN) & VBA.vbCrLf
    'mod
    If IsArray(stringToLog) Then
        Dim i As Long
        For i = LBound(stringToLog) To UBound(stringToLog)
            strg = strg & stringToLog & VBA.vbCrLf
        Next
    Else
        strg = strg & stringToLog & VBA.vbCrLf
    End If
    '/mod
    strg = strg & GetBufferString("=", C_MAX_LEN) & VBA.vbCrLf
    stringToLog = strg
    Call WriteToFile(CStr(stringToLog), LogAppend)
End Sub
'''
'''
'''
Sub LogHeader2(ByVal stringToLog As Variant)
    '--------------------
    'Header 2
    '--------------------
    m_AppErr.clear
    Dim strg As String
    '// Preformatting
    strg = VBA.vbCrLf
    '// Header definition
    strg = strg & GetBufferString("-", C_MAX_LEN) & VBA.vbCrLf
    'mod
    If IsArray(stringToLog) Then
        Dim i As Long
        For i = LBound(stringToLog) To UBound(stringToLog)
            strg = strg & stringToLog & VBA.vbCrLf
        Next
    Else
        strg = strg & stringToLog & VBA.vbCrLf
    End If
    '/mod
    strg = strg & GetBufferString("-", C_MAX_LEN) & VBA.vbCrLf
    stringToLog = strg
    Call WriteToFile(CStr(stringToLog), LogAppend)
End Sub
'''
'''
'''
Sub LogBullet(lvl As Long, stringToLog As Variant)
    If lvl < 1 Then Exit Sub
    stringToLog = VBA.Now & ": |" & String(2 * (lvl - 1), "-") & "> " & stringToLog
    Call WriteToFile(CStr(stringToLog), LogAppend)
End Sub
'''
'''
'''
Public Sub LogTable(Table As Variant, Optional hasHeader As Boolean = True, Optional padding As Integer = 1)
    Dim nrow        As Integer
    Dim ncol        As Integer
    Dim irow        As Integer
    Dim jcol        As Integer
    Dim vals()      As Variant
    Dim rowStart    As Integer
    Dim widths()    As Integer
    Dim cbuff       As String
    Dim line        As String
    Dim edge        As String

    Dim Output      As String

    nrow = UBound(Table, 1)
    ncol = UBound(Table, 2)
    ReDim vals(1 To ncol * 2)
    ReDim widths(1 To ncol)
    cbuff = String(200, " ")

    ' Define column widths
    For jcol = LBound(Table, 2) To ncol
        For irow = LBound(Table, 1) To nrow
            If (Len(Table(irow, jcol)) > widths(jcol)) Then
                widths(jcol) = Len(Table(irow, jcol))
            End If
        Next irow
    Next jcol

    ' Print top edge and header
    edge = ".-" & Left(String(Len(cbuff), "-"), widths(1) + padding) & "-"
    For jcol = (LBound(Table, 2) + 1) To ncol - 1
        edge = edge & "-" & Left(String(Len(cbuff), "-"), widths(jcol) + padding) & "-"
    Next jcol
    edge = edge & "-" & Left(String(Len(cbuff), "-"), widths(jcol) + padding) & "."
    Output = Output & IIf(Output <> "", vbNewLine, "") & edge
    If hasHeader Then
        line = "| " & Left(Table(LBound(Table, 1), 1) & cbuff, widths(1) + padding) & "|"
        For jcol = (LBound(Table, 2) + 1) To ncol
            line = line & " " & Left(Table(LBound(Table, 1), jcol) & cbuff, widths(jcol) + padding) & "|"
        Next jcol
        Output = Output & IIf(Output <> "", vbNewLine, "") & line
        Output = Output & IIf(Output <> "", vbNewLine, "") & edge

        rowStart = LBound(Table, 1) + 1
    Else
        rowStart = LBound(Table, 1)
    End If

    ' Print data
    For irow = rowStart To nrow
        line = "| " & Left(Table(irow, 1) & cbuff, widths(1) + padding) & "|"
        For jcol = (LBound(Table, 2) + 1) To ncol
            line = line & " " & Left(Table(irow, jcol) & cbuff, widths(jcol) + padding) & "|"
        Next jcol
        Output = Output & IIf(Output <> "", vbNewLine, "") & line
    Next irow

    Output = Output & IIf(Output <> "", vbNewLine, "") & edge
    Call WriteToFile(CStr(Output), LogAppend)
    Exit Sub
ErrHandler:
    Debug.Print VBA.vbCrLf & "** PROBLEM OCCURED ** PRETTY PRINT TABLE **"
    Debug.Print Strings.String(C_LINE_LENGTH, "_")
End Sub
'''
''' Writes "message" to the current log file.
'''
Private Sub WriteToFile(message As String, outputType As LogOutputType)
    Dim lu          As Integer
    On Error Resume Next
    If m_LogFileExists Then
        lu = VBA.FreeFile
        Select Case outputType
            Case LogAppend: Open m_FilePath For Append As #lu
            Case LogOutput: Open m_FilePath For Output As #lu
        End Select
        Print #lu, message
        Close #lu
    Else    ' display application error number
        With m_AppErr
            .Number = C_ERR_LOG_FILE_DNE
            .Source = "WriteToFile(...)"
            .DisplayMessage
        End With
    End If
End Sub





'
''* @TODO Created: 21-08-2023 14:55 Author: Anastasiou Alex
''* @TODO switch to ini method
''''
'''' Returns the default directory save location (scratch folder).
'''' Returns NULL if an error occured.
''''
'Property Get RegDirectoryScratch() As Variant
'
'    Dim var         As Variant
'
'    m_AppErr.clear
'
'    '// Retrieve and check value.
'    var = m_RegEdit.RegistryGetValue(C_BASE_KEY, C_KEY_NAME, C_REG_SCRATCH)
'    If Not VBA.IsNull(var) Then
'        RegDirectoryScratch = VBA.CStr(var)
'    Else
'        RegDirectoryScratch = Null
'        With m_AppErr
'            .Source = "'Get' RegDirecotryScratch"
'            .Number = C_ERR_REGISTRY_CLASS
'            .Description = m_RegEdit.AppErr.Description
'        End With
'    End If
'End Property
'
''''
'''' Sets the directory path for the default save location (scratch folder).
''''
'Public Property Let RegDirectoryScratch(strDirectoryPath As Variant)
'
'    m_AppErr.clear
'
'    '// Force trailing slash
'    If (VBA.Right(VBA.Trim(strDirectoryPath), 1) <> "\") Then
'        strDirectoryPath = VBA.Trim(strDirectoryPath) & "\"
'    End If
'
'    '// Check and set value.
'    If (strDirectoryPath = vbNullString) Then
'        GoTo ErrHandler
'    ElseIf FileFolderExists(VBA.CStr(strDirectoryPath)) Then
'        If m_RegEdit.RegistryUpdateValue(C_BASE_KEY, C_KEY_NAME, C_REG_SCRATCH, VBA.CStr(strDirectoryPath), createKeyIfNotExist:=True) Then
'            Exit Property
'        Else
'            m_AppErr.Number = C_ERR_REGISTRY_CLASS
'        End If
'    Else
'        m_AppErr.Number = C_ERR_INVALID_DIRECTORY
'    End If
'
'ErrHandler:
'    With m_AppErr
'        If (.Number = 0) Then .Number = C_ERR_TYPE_MISMATCH
'        .Source = "'Let' RegDirectoryScratch"
'        .Description = GetAppErrDescription(.Number)
'    End With
'End Property
''''
'''' For testing purposes.
'''' The LogFile class will automatically create a default registry path if none is detected i.e. workbook.path/scr/
'''' This method alows me to remove the defined registry path, to verify that it is correctly created by the class.
''''
'Public Sub RemoveRegistrySetting()
'
'    m_AppErr.clear
'
'    If Not m_RegEdit.RegistryDeleteValue(C_BASE_KEY, C_KEY_NAME, C_REG_SCRATCH) Then    ' display error message
'        With m_AppErr
'            .Number = C_ERR_REGISTRY_CLASS
'            .Source = "RemoveRegistrySetting"
'            .Description = GetAppErrDescription(.Number)
'            .DisplayMessage
'        End With
'    End If
'
'End Sub






