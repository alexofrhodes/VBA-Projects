Attribute VB_Name = "M_Sound"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#Else
    Public Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#End If

Public Const SND_SYNC = &H0         ' (Default) Play the sound synchronously. Code execution pauses until sound is complete.
Public Const SND_ASYNC = &H1        ' Play the sound asynchronously. Code execution does not wait for sound to complete.
Public Const SND_NODEFAULT = &H2    ' If the specified sound is not found, do not play the default sound (no sound is played).
Public Const SND_MEMORY = &H4       ' lpszSoundName is a memory file of the sound. Not used in VBA/VB6.
Public Const SND_LOOP = &H8         ' Continue playing sound in a loop until the next call to sndPlaySound.
Public Const SND_NOSTOP = &H10      ' Do not stop playing the current sound before playing the specified sound.

Sub PlayTheSound(ByVal WhatSound As String, Optional Flags As Long = SND_ASYNC)
'http://www.cpearson.com/excel/PlaySound.aspx
        If Dir(WhatSound, vbNormal) = "" Then
        ' WhatSound is not a file. Get the file named by
        ' WhatSound from the Windows\Media directory.
        WhatSound = Environ("SystemRoot") & "\Media\" & WhatSound
        If InStr(1, WhatSound, ".") = 0 Then
            ' if WhatSound does not have a .wav extension,add one.
            WhatSound = WhatSound & ".wav"
        End If
        If Dir(WhatSound, vbNormal) = vbNullString Then
            ' Can't find the file. Do a simple Beep.
            Beep
            Exit Sub
        End If
    Else
        ' WhatSound is a file. Use it.
    End If
    ' Finally, play the sound.
    sndPlaySound32 WhatSound, Flags
End Sub

Sub ListWavFiles()
    Dim N As Long
    Dim FSO As Object
    Dim FF As Object
    Dim F As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set FF = FSO.GetFolder(Environ("SystemRoot") & "\Media")
    For Each F In FF.Files
        N = N + 1
        Cells(N, 1) = F.Name
        Cells(N, 2) = F.Path
    Next F
    ActiveSheet.Columns(1).AutoFit
End Sub
