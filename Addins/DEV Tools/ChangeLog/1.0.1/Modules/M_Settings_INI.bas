Attribute VB_Name = "M_Settings_INI"

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : F_Settings_INI
'* Purpose    :
'* Copyright  :
'*
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 30-06-2023 14:11    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'____API METHOD______

#If VBA7 Then
Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
public declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
public declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
public declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#End If


Sub TestINI()
    '@AssignedModule F_Settings_INI
    '@INCLUDE PROCEDURE FollowLink
    '@INCLUDE PROCEDURE dp
    '@INCLUDE PROCEDURE IniSections
    '@INCLUDE PROCEDURE IniReadSection
    '@INCLUDE PROCEDURE IniSectionKeys
    '@INCLUDE PROCEDURE IniReadKey
    '@INCLUDE PROCEDURE IniWrite

    Dim FilePath    As String: FilePath = ThisWorkbook.path & "\test.INI"
    FollowLink ThisWorkbook.path

    IniWrite FilePath, "Settings1", "KeyName1", "Value1"
    IniWrite FilePath, "Settings1", "KeyName2", "2"
    IniWrite FilePath, "Settings1", "KeyName3", "3"     'SEE THE FILE
    Stop
    IniWrite FilePath, "Settings1", "KeyName1", "Updated Value"    'SEE THE FILE
    Stop

    Dim i           As Long
    For i = 1 To 5
        IniWrite FilePath, "Settings" & i, "KeyName" & i, i
    Next
    'SEE THE FILE
    Stop
    dp String(20, "~") & " Printing sections of " & FilePath
    dp IniSections(FilePath)
    Stop
    dp String(20, "~") & " Printing keys of section Settings1"
    dp IniSectionKeys(FilePath, "Settings1")
    Stop
    dp String(20, "~") & " Printing all lines of section Settings1"
    dp IniReadSection(FilePath, "Settings1")
    Stop
    dp String(20, "~") & " Printing value of Section: Settings1, Keyname: Keyname1"
    dp IniReadKey(FilePath, "Settings1", "KeyName1")

End Sub

Public Function IniSections(iniFile As String) As Variant
    '@INCLUDE PROCEDURE TxtRead
    '@AssignedModule F_Settings_INI

    '---sample file content---
    '[settings1]
    '    string1 = aaa
    '    string2 = bbb
    '[settings2]
    '    string1 = ccc
    '    string2 = ddd
    '-------------------------
    IniSections = Split(Replace(Replace(Join(Filter(Split(Replace(TxtRead(iniFile), vbLf, vbNewLine), vbNewLine), "[", True), vbNewLine), "[", ""), "]", ""), vbNewLine)
    '------Result------------------
    'Array("settings1","settings2")
End Function

Public Function IniReadSection(fileName As String, Section As String) As Variant
    '@INCLUDE DECLARATION GetPrivateProfileSection
    '@INCLUDE PROCEDURE ArrayRemoveEmptyElements
    '@AssignedModule F_Settings_INI
    Dim RetVal      As String * 255
    Dim V           As Long: V = GetPrivateProfileSection(Section, RetVal, 255, fileName)
    Dim S           As String: S = Left(RetVal, V + 0)
    Dim VL          As Variant: VL = Split(S, Chr$(0))
    VL = ArrayRemoveEmptyElements(VL)
    IniReadSection = VL
    '-----result for reading "settings1"-----
    'Array("string1=aaa","string2=bbb")
End Function

Public Function IniSectionKeys(fileName As String, Section As String) As Variant
    '@AssignedModule F_Settings_INI
    '@INCLUDE PROCEDURE IniReadSection
    '@INCLUDE PROCEDURE IniSectionExists
    Dim arr()       As Variant
    If Not IniSectionExists(fileName, Section) Then
        IniSectionKeys = arr
        Exit Function
    End If
    arr = IniReadSection(fileName, Section)
    Dim out         As Variant
    ReDim out(UBound(arr))
    Dim i           As Long
    For i = LBound(arr) To UBound(arr)
        out(i) = Trim(Split(arr(i), "=")(0))
    Next i
    IniSectionKeys = out
    '-----result for reading "settings1"-----
    'string1
    'string2
End Function

Public Function IniReadKey(IniFileName As String, ByVal Sect As String, ByVal Keyname As String, Optional ByVal Default As Variant)    'As String
    '@INCLUDE DECLARATION GetPrivateProfileString
    '@AssignedModule F_Settings_INI
    Dim Worked      As Long
    Dim RetStr      As String * 128
    Dim StrSize     As Long
    Dim iNoOfCharInIni As Long: iNoOfCharInIni = 0
    Dim sIniString  As String: sIniString = ""
    If Sect = "" Or Keyname = "" Then
        MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
    Else
        Dim sProfileString As String: sProfileString = ""
        RetStr = Space(128)
        StrSize = Len(RetStr)
        Worked = GetPrivateProfileString(Sect, Keyname, "", RetStr, StrSize, IniFileName)
        If Worked Then
            iNoOfCharInIni = Worked
            sIniString = Left$(RetStr, Worked)
        End If
    End If

    If sIniString <> vbNullString Then
        IniReadKey = sIniString
    ElseIf sIniString = vbNullString And Not IsMissing(Default) Then
        IniReadKey = Default
    End If
    '---- result for reading "settings1", "string1" ----
    'aaa
End Function

Public Sub IniWrite(IniFileName As String, ByVal Sect As String, ByVal Keyname As String, ByVal Wstr As String)
    '@INCLUDE DECLARATION WritePrivateProfileString
    '@AssignedModule F_Settings_INI

    'This macro also creates the file & section & key if they doesn't exist

    Dim Worked      As Long
    Dim iNoOfCharInIni As Long

    iNoOfCharInIni = 0
    Dim sIniString  As String: sIniString = ""
    If Sect = "" Or Keyname = "" Then
        MsgBox "Section Or Key To Write Not Specified !!!", vbExclamation, "INI"
    Else
        Worked = WritePrivateProfileString(Sect, Keyname, Wstr, IniFileName)
        If Worked Then
            iNoOfCharInIni = Worked
            sIniString = Wstr
        End If
    End If

    '---- result for writing "settings1", "string1", "newval" ----
    '[settings1]
    '    string1 = newval
    '    string2 = bbb
    '[settings2]
    '    string1 = ccc
    '    string2 = ddd

    '---- result for writing "settings1", "string3", "newkey" ----
    '[settings1]
    '    string1 = newval
    '    string2 = bbb
    '    string3 = newkey
    '[settings2]
    '    string1 = ccc
    '    string2 = ddd
End Sub






'___NO API METHOD______
'
Public Sub TestReadKey()
    Debug.Print Join( _
                    Array( _
                        "INI File: " & ThisWorkbook.path & "\MyIniFile.ini", _
                        "Section: SETTINGS", _
                        "Section Exist: " & IniSectionExists(ThisWorkbook.path & "\MyIniFile.ini", "SETTINGS"), _
                        "Key: License", _
                        "Key Exist: " & IniKeyExists(ThisWorkbook.path & "\MyIniFile.ini", "SETTINGS", "License"), _
                        "Key Value: " & Ini_ReadKeyVal(ThisWorkbook.path & "\MyIniFile.ini", "SETTINGS", "License")), _
                    vbLf)
    'You can validate the value by checking the bSectionExists and bKeyExists variable to ensure they were actually found in the ini file
    '@AssignedModule F_Settings_INI
    '@INCLUDE PROCEDURE IniSectionExists
    '@INCLUDE PROCEDURE IniKeyExists
End Sub

Public Function IniSectionExists(iniFile As String, Section As String) As Boolean
    'Alex
    '@AssignedModule F_Settings_INI
    '@INCLUDE PROCEDURE TxtRead
    IniSectionExists = InStr(1, TxtRead(iniFile), "[" & Section & "]") > 0
End Function

Public Function IniKeyExists(iniFile As String, Section As String, Key As String) As Boolean
    'Alex
    '@AssignedModule F_Settings_INI
    IniKeyExists = (Ini_ReadKeyVal(iniFile, Section, Key) <> "")
End Function

Public Sub TestWriteKey()
    '@AssignedModule F_Settings_INI
    If Ini_WriteKeyVal(ThisWorkbook.path & "\MyIniFile.ini", "SETTINGS", "License", "JBXR-HHTY-LKIP-HJNB-GGGT") = True Then
        MsgBox "The key was written"
    Else
        MsgBox "An error occurred!"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Ini_ReadKeyVal
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Read an Ini file's Key
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
' Req'd Refs: Uses Late Binding, so none required
'             No APIs either! 100% VBA
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sIniFile  : Full path and filename of the ini file to read
' sSection  : Ini Section to search for the Key to read the Key from
' sKey      : Name of the Key to read the value of
'
' Usage:
' ~~~~~~
' ? Ini_Read(Application.CurrentProject.Path & "\MyIniFile.ini", "LINKED TABLES", "Path")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-08-09              Initial Release
'---------------------------------------------------------------------------------------
Public Function Ini_ReadKeyVal(ByVal sIniFile As String, _
        ByVal sSection As String, _
        ByVal sKey As String) As String
    '@AssignedModule F_Settings_INI
    '@INCLUDE PROCEDURE FileExists
    '@INCLUDE PROCEDURE TxtRead
    '@INCLUDE PROCEDURE ArrayTrim
    On Error GoTo Error_Handler
    Dim bSectionExists As Boolean
    Dim bKeyExists  As Boolean
    Dim sIniFileContent As String
    Dim aIniLines() As String
    Dim sLine       As String
    Dim i           As Long

    sIniFileContent = ""
    bSectionExists = False
    bKeyExists = False

    'Validate that the file actually exists
    If FileExists(sIniFile) = False Then
        MsgBox "The specified ini file: " & vbCrLf & vbCrLf & _
                sIniFile & vbCrLf & vbCrLf & _
                "could not be found.", vbCritical + vbOKOnly, "File not found"
        GoTo Error_Handler_Exit
    End If

    sIniFileContent = TxtRead(sIniFile)    'Read the file into memory
    aIniLines = Split(sIniFileContent, vbLf)
    For i = 0 To UBound(aIniLines)
        sLine = Trim(aIniLines(i))
        sLine = VBA.Replace(sLine, vbTab, vbNullString)
        If InStr(1, sLine, "=") > 0 Then sLine = Join(ArrayTrim(Split(sLine, "=")), "=")    '<- Alex added this
        If bSectionExists = True And Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
            Exit For    'Start of a new section
        End If
        If sLine = "[" & sSection & "]" Then
            bSectionExists = True
        End If
        If bSectionExists = True Then
            If sLine Like sKey & "=*" Then
                bKeyExists = True
                Ini_ReadKeyVal = Mid(sLine, InStr(sLine, "=") + 1)
            End If
        End If
    Next i

Error_Handler_Exit:
    On Error Resume Next
    Exit Function

Error_Handler:
    'Err.Number = 75 'File does not exist, Permission issues to write is denied,
    MsgBox Join( _
                Array( _
                    "The following error has occurred", _
                    "", _
                    "Error Number: " & Err.Number, _
                    "Error Source: Ini_ReadKeyVal", _
                    "Error Description: " & Err.Description & _
                    Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
                    , vbOKOnly + vbCritical, "An Error has Occurred!"), _
                vbLf)
                    
    Resume Error_Handler_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : Ini_WriteKeyVal
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Writes a Key value to the specified Ini file's Section
'               If the file does not exist, it will be created
'               If the Section does not exist, it will be appended to the existing content
'               If the Key does not exist, it will be appended to the existing Section content
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
' Req'd Refs: Uses Late Binding, so none required
'             No APIs either! 100% VBA
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sIniFile  : Full path and filename of the ini file to edit
' sSection  : Ini Section to search for the Key to edit
' sKey      : Name of the Key to edit
' sValue    : Value to associate to the Key
'
' Usage:
' ~~~~~~
' Call Ini_WriteKeyVal(Application.CurrentProject.Path & "\MyIniFile.ini", "LINKED TABLES", "Paths", "D:\")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-08-09              Initial Release
' 2         2020-01-27              Fix to address issue flagged by users
'---------------------------------------------------------------------------------------
Public Function Ini_WriteKeyVal(ByVal sIniFile As String, _
        ByVal sSection As String, _
        ByVal sKey As String, _
        ByVal sValue As String) As Boolean
    '@AssignedModule F_Settings_INI
    '@INCLUDE PROCEDURE FileExists
    '@INCLUDE PROCEDURE TxtOverwrite
    '@INCLUDE PROCEDURE TxtRead
    On Error GoTo Error_Handler
    Dim bSectionExists As Boolean
    Dim bKeyExists  As Boolean
    Dim sIniFileContent As String
    Dim aIniLines() As String
    Dim sLine       As String
    Dim sNewLine    As String
    Dim i           As Long
    Dim bFileExist  As Boolean
    Dim bInSection  As Boolean
    Dim bKeyAdded   As Boolean

    sIniFileContent = ""
    bSectionExists = False
    bKeyExists = False

    'Validate that the file actually exists
    If FileExists(sIniFile) = False Then
        GoTo SectionDoesNotExist
    End If
    bFileExist = True

    sIniFileContent = TxtRead(sIniFile)    'Read the file into memory
    aIniLines = Split(sIniFileContent, vbLf)    'Break the content into individual lines
    sIniFileContent = ""    'Reset it
    For i = 0 To UBound(aIniLines)    'Loop through each line
        sNewLine = ""
        sLine = Trim(aIniLines(i))
        If sLine = "[" & sSection & "]" Then
            bSectionExists = True
            bInSection = True
        End If
        If bInSection = True Then
            If sLine <> "[" & sSection & "]" Then
                If Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
                    'Our section exists, but the key wasn't found, so append it
                    sNewLine = sKey & "=" & sValue
                    i = i - 1
                    bInSection = False    ' we're switching section
                    bKeyAdded = True
                End If
            End If
            If Len(sLine) > Len(sKey) Then
                If Left(sLine, Len(sKey) + 1) = sKey & "=" Then
                    sNewLine = sKey & "=" & sValue
                    bKeyExists = True
                    bKeyAdded = True
                End If
            End If
        End If
        If Len(sIniFileContent) > 0 Then sIniFileContent = sIniFileContent & vbCrLf
        If sNewLine = "" Then
            sIniFileContent = sIniFileContent & sLine
        Else
            sIniFileContent = sIniFileContent & sNewLine
        End If
    Next i

SectionDoesNotExist:
    'if not found, add it to the end
    If bSectionExists = False Then
        If Len(sIniFileContent) > 0 Then sIniFileContent = sIniFileContent & vbCrLf
        sIniFileContent = sIniFileContent & "[" & sSection & "]"
    End If
    If bKeyAdded = False Then
        sIniFileContent = sIniFileContent & vbCrLf & sKey & "=" & sValue
    End If

    'Write to the ini file the new content
    Call TxtOverwrite(sIniFile, sIniFileContent)
    Ini_WriteKeyVal = True

Error_Handler_Exit:
    On Error Resume Next
    Exit Function

Error_Handler:
    MsgBox Join( _
                Array( _
                    "The following error has occurred", _
                    "", _
                    "Error Number: " & Err.Number, _
                    "Error Source: Ini_WriteKeyVal", _
                    "Error Description: " & Err.Description & _
                    Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
                    , vbOKOnly + vbCritical, "An Error has Occurred!"), _
                vbLf)
                    
    Resume Error_Handler_Exit
End Function

'_____________________________________________________________

