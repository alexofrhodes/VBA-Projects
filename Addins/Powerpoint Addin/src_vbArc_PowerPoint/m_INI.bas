Attribute VB_Name = "m_INI"
Option Explicit


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

    Dim FilePath As String: FilePath = ThisProjectPath & "config.INI"
    
    IniWriteKey FilePath, "Settings1", "KeyName1", "Value1"
    IniWriteKey FilePath, "Settings1", "KeyName2", "2"
    IniWriteKey FilePath, "Settings1", "KeyName3", "3"     'SEE THE FILE
    Stop
    IniWriteKey FilePath, "Settings1", "KeyName1", "Updated Value" 'SEE THE FILE
    Stop
    
    Dim i  As Long
    For i = 1 To 5
        IniWriteKey FilePath, "Settings" & i, "KeyName" & i, i
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


Public Function IniReadKey(IniFileName As String, ByVal Sect As String, ByVal Keyname As String) As String
'@INCLUDE DECLARATION GetPrivateProfileString
    Dim Worked As Long
    Dim RetStr As String * 128
    Dim StrSize As Long
    Dim iNoOfCharInIni As Long: iNoOfCharInIni = 0
    Dim sIniString As String: sIniString = ""
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
    IniReadKey = sIniString
'---- result for reading "settings1", "string1" ----
'aaa
End Function

Public Sub IniWriteKey(IniFileName As String, ByVal Sect As String, ByVal Keyname As String, ByVal Wstr As String)
'@INCLUDE DECLARATION WritePrivateProfileString

'This macro also creates the file & section & key if they doesn't exist

    Dim Worked As Long
    Dim iNoOfCharInIni As Long

    iNoOfCharInIni = 0
    Dim sIniString As String: sIniString = ""
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

Public Function IniSections(IniFile As String) As Variant
'@INCLUDE PROCEDURE TxtRead

'---sample file content---
'[settings1]
'    string1 = aaa
'    string2 = bbb
'[settings2]
'    string1 = ccc
'    string2 = ddd
'-------------------------
    IniSections = Split(Replace(Replace(Join(Filter(Split(Replace(TxtRead(IniFile), vbLf, vbNewLine), vbNewLine), "[", True), vbNewLine), "[", ""), "]", ""), vbNewLine)
'------Result------------------
'Array("settings1","settings2")
End Function

Public Function IniReadSection(FileName As String, Section As String) As Variant
'@INCLUDE DECLARATION GetPrivateProfileSection
'@INCLUDE PROCEDURE ArrayRemoveEmptyElements
    Dim RetVal As String * 255
    Dim v As Long:      v = GetPrivateProfileSection(Section, RetVal, 255, FileName)
    Dim s As String:    s = Left(RetVal, v + 0)
    Dim VL As Variant:  VL = Split(s, Chr$(0))
    VL = ArrayRemoveEmptyElements(VL)
    IniReadSection = VL
'-----result for reading "settings1"-----
'Array("string1=aaa","string2=bbb")
End Function

Public Function IniSectionKeys(FileName As String, Section As String) As Variant
    Dim arr As Variant
    arr = IniReadSection(FileName, Section)
    Dim out As Variant
    ReDim out(UBound(arr))
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        out(i) = Trim(Split(arr(i), "=")(0))
    Next i
    IniSectionKeys = out
'-----result for reading "settings1"-----
'string1
'string2
End Function

