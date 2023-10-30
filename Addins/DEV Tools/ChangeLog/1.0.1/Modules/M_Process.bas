Attribute VB_Name = "M_Process"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As LongPtr, ByVal bInheritHandle As LongPtr, ByVal dwProcessId As LongPtr) As LongPtr
    Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
#Else
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long,ByVal uExitCode As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
#End If

Public Function KillProcess(ByVal processId As Long) As Boolean
'@INCLUDE DECLARATION CloseHandle
'@INCLUDE DECLARATION OpenProcess
'@INCLUDE DECLARATION TerminateProcess
    Dim hProcess As LongPtr
    hProcess = OpenProcess(&H1F0FFF, 0, processId)
    
    If hProcess <> 0 Then
        Dim success As Long
        success = TerminateProcess(hProcess, 0)
        If success <> 0 Then
            CloseHandle hProcess
            KillProcess = True
        End If
    End If
End Function


