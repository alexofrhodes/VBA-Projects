Attribute VB_Name = "F_Clipboard"
#If VBA7 Then
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
#Else
    Public Declare Function CloseClipboard Lib "user32" () As Long
    Public Declare Function EmptyClipboard Lib "user32" () As Long
    Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
#End If

Public Function ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Function

Function IsCopyingCompleted() As Boolean
    'Check if copying data to clipboard is completed
    Dim tempString As String
    Dim myData As DataObject
    'Try to put data from clipboard to string to check if operations on clipboard are completed
    On Error Resume Next
    Set myData = New DataObject
    myData.GetFromClipboard
    tempString = myData.GetText(1)
    If Err.number = 0 Then
        IsCopyingCompleted = True
    Else
        IsCopyingCompleted = False
    End If
    On Error GoTo 0
End Function

Public Function CLIP(Optional StoreText As String) As String
    Dim X As Variant
    X = StoreText
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
            Case Len(StoreText)
                .SetData "text", X
            Case Else
                CLIP = .GetData("text")
            End Select
        End With
    End With
End Function

Sub ClipToTarget(myTarget)
    uCodeOnTheFly.Controls(ThisWorkbook.Sheets("uCodeOnTheFly_Settings").Range("D1").Value) = CLIP
End Sub
