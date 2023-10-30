Attribute VB_Name = "M_Notify"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function Shell_NotifyIconW Lib "shell32.dll" (ByVal dwMessage As Long, ByRef nfIconData As NOTIFYICONDATAW) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#Else
Private Declare Function Shell_NotifyIconW Lib "shell32.dll" (ByVal dwMessage As Long, ByRef nfIconData As NOTIFYICONDATAW) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef source As Any, ByVal Length As Long)
#End If

Private Type NOTIFYICONDATAW
    cbSize          As Long

#If Win64 Then
    padding1        As Long
#End If

    hWnd            As LongPtr
    uID             As Long
    uFlags          As Long
    uCallbackMessage As Long

#If Win64 Then
    padding2        As Long
#End If

    hIcon           As LongPtr
    szTip(1 To 128 * 2) As Byte
    dwState         As Long
    dwStateMask     As Long
    szInfo(1 To 256 * 2) As Byte
    uTimeout        As Long
    szInfoTitle(1 To 64 * 2) As Byte
    dwInfoFlags     As Long
End Type

Private Const NIM_ADD As Long = &H0&
Private Const NIM_MODIFY As Long = &H1&
Private Const NIF_INFO As Long = &H10&

Private Function Min(ByVal a As Long, ByVal b As Long) As Long
    '@AssignedModule M_Notify
    If a < b Then Min = a Else Min = b
End Function

Public Sub Toast(Optional ByVal Title As String, Optional ByVal info As String, Optional ByVal flag As Long)
    Rem toast "Hello World", "from Excel",1
    Rem https://github.com/rfl808/Notify
    '@INCLUDE Min
    '@AssignedModule M_Notify
    '@INCLUDE PROCEDURE Min
    '@INCLUDE DECLARATION NIF_INFO
    '@INCLUDE DECLARATION NIM_ADD
    '@INCLUDE DECLARATION NIM_MODIFY
    '@INCLUDE DECLARATION Shell_NotifyIconW
    '@INCLUDE DECLARATION CopyMemory
    '@INCLUDE DECLARATION NOTIFYICONDATAW
    Dim nfIconData  As NOTIFYICONDATAW
    info = info & " "
    Title = Title & " "
    With nfIconData
        .cbSize = Len(nfIconData)
        .uFlags = NIF_INFO
        .dwInfoFlags = flag
        If Len(Title) > 0 Then
            CopyMemory ByVal VarPtr(.szInfoTitle(LBound(.szInfoTitle))), ByVal StrPtr(Title), Min(Len(Title) * 2, UBound(.szInfoTitle) - LBound(.szInfoTitle) + 1 - 2)
        End If
        If Len(info) > 0 Then
            CopyMemory ByVal VarPtr(.szInfo(LBound(.szInfo))), ByVal StrPtr(info), Min(Len(info) * 2, UBound(.szInfo) - LBound(.szInfo) + 1 - 2)
        End If
    End With
    Shell_NotifyIconW NIM_ADD, nfIconData
    Shell_NotifyIconW NIM_MODIFY, nfIconData
End Sub

Rem Flags for the balloon message..
Rem None = 0
Rem Information = 1
Rem Exclamation = 2
Rem Critical = 3
