Attribute VB_Name = "m_General"
Option Explicit

Enum MyColors
    FormBackgroundDarkGray = 4208182            ' BACKGROUND DARK GRAY
    FormSidebarMediumGray = 5457992             ' TILE COLORS LIGHTER GRAY
    FormSidebarMouseOverLightGray = &H808080    ' lighter light gray
    FormSelectedGreen = 8435998                 ' green tile
End Enum


Public Sub RotateImage3d()
    Dim img As Shape: Set img = ActiveSheet.Shapes(Application.Caller)
    Dim i As Long
    For i = 0 To 36 Step 1

        img.ThreeD.RotationY = i * 10
        '         img.ThreeD.RotationX = i * 10
         
        ActiveSheet.Shapes("imgGithub").ThreeD.RotationY = i * 10
        ActiveSheet.Shapes("imgBlog").ThreeD.RotationY = i * 10
        ActiveSheet.Shapes("imgYoutube").ThreeD.RotationY = i * 10
        ActiveSheet.Shapes("imgCoffee").ThreeD.RotationY = i * 10
        ActiveSheet.Shapes("imgMail").ThreeD.RotationY = i * 10
    
        DoEvents
    Next i
End Sub

Function FolderContents( _
            FolderOrZipFilePath As Variant, _
            LogFolders As Boolean, _
            LogFiles As Boolean, _
            ScanInSubfolders As Boolean, _
            out As Collection, _
            Optional Filter As String = "*") As Collection
    
    Dim oShell As Object
    Set oShell = CreateObject("Shell.Application")
    Dim oFi As Object
    For Each oFi In oShell.Namespace(FolderOrZipFilePath).Items
        If oFi.IsFolder Then
            If LogFolders Then
                out.Add oFi.Path & "\"
            End If
            If ScanInSubfolders Then FolderContents oFi.Path, LogFolders, LogFiles, ScanInSubfolders, out, Filter
        Else
            If LogFiles Then
                If UCase(oFi.Name) Like UCase(Filter) Then
                    out.Add oFi.Path
                End If
            End If
        End If
    Next
    Set FolderContents = out
    Set oShell = Nothing
End Function

Function getFileName(FilePath As String)
    getFileName = Mid(FilePath, InStrRev(FilePath, "\") + 1)
    getFileName = Left(getFileName, InStr(1, getFileName, ".") - 1)
End Function

Function getFileExtension(FilePath As String)
    Dim pos As Long: pos = InStrRev(FilePath, ".")
    If pos = 0 Then End
    getFileExtension = Mid(FilePath, pos)
End Function

Function getFileFolder(FilePath As String)
    getFileFolder = Left(FilePath, InStrRev(FilePath, "\"))
End Function

Sub FollowLink(FolderPath As String)
    If Right(FolderPath, 1) = "\" Then FolderPath = Left(FolderPath, Len(FolderPath) - 1)
    On Error Resume Next
    Dim oShell As Object
    Dim Wnd As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.Name = "File Explorer" Then
            If Wnd.document.folder.Self.Path = FolderPath Then Exit Sub
        End If
    Next Wnd
    FollowLink FolderPath
End Sub

Function TxtRead(sPath As Variant) As String
    If sPath = "" Or InStr(1, sPath, ":\") = 0 Then Exit Function
    Dim sTXT As String
    If Dir(sPath) = "" Then
        Debug.Print "File was not found."
        Debug.Print sPath
        Exit Function
    End If
    Open sPath For Input As #1
    Do Until EOF(1)
        Line Input #1, sTXT
        TxtRead = TxtRead & sTXT & vbLf
    Loop
    Close
    If Len(TxtRead) = 0 Then
        TxtRead = ""
    Else
        TxtRead = Left(TxtRead, Len(TxtRead) - 1)
    End If
End Function

Sub TxtOverwrite(sFile As String, sText As String)
    On Error GoTo ERR_HANDLER
    Dim FileNumber As Integer
    FileNumber = FreeFile
    Open sFile For Output As #FileNumber
    Print #FileNumber, sText
    Close #FileNumber
Exit_Err_Handler:
    Exit Sub
ERR_HANDLER:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
    "Error Number: " & Err.Number & vbCrLf & _
    "Error Source: TxtOverwrite" & vbCrLf & _
    "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Sub

Public Function ArrayRemoveEmptyElements(varArray As Variant) As Variant
    If UBound(varArray) = -1 Then ArrayRemoveEmptyElements = varArray: Exit Function
    Dim TempArray() As Variant
    Dim OldIndex As Integer
    Dim NewIndex As Integer
    ReDim TempArray(LBound(varArray) To UBound(varArray))
    For OldIndex = LBound(varArray) To UBound(varArray)
        If Not Trim(varArray(OldIndex) & " ") = "" Then
            TempArray(NewIndex) = varArray(OldIndex)
            NewIndex = NewIndex + 1
        End If
    Next OldIndex
    ReDim Preserve TempArray(LBound(varArray) To NewIndex - 1)
    ArrayRemoveEmptyElements = TempArray
    varArray = TempArray
End Function

Function Transpose2DArray(inputArray As Variant) As Variant
    Dim X As Long, yUbound As Long
    Dim Y As Long, xUbound As Long
    Dim TempArray As Variant
    xUbound = UBound(inputArray, 2)
    yUbound = UBound(inputArray, 1)
    ReDim TempArray(1 To xUbound, 1 To yUbound)
    For X = 1 To xUbound
        For Y = 1 To yUbound
            TempArray(X, Y) = inputArray(Y, X)
        Next Y
    Next X
    Transpose2DArray = TempArray
End Function

Sub FoldersCreate(FolderPath As String)
    On Error Resume Next
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim ArrayElement As Variant
    individualFolders = Split(FolderPath, "\")
    For Each ArrayElement In individualFolders
        tempFolderPath = tempFolderPath & ArrayElement & "\"
        If FolderExists(tempFolderPath) = False Then
            MkDir tempFolderPath
        End If
    Next ArrayElement
End Sub

Function FolderExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Sub KillFile(FilePath As String)
    On Error GoTo ErrorHandler
    Dim fso As New FileSystemObject, aFile As file
    If (fso.FileExists(FilePath)) Then
        SetAttr FilePath, vbNormal
        Set aFile = fso.GetFile(FilePath)
        aFile.Delete
    End If
    Exit Sub
ErrorHandler:
    If MsgBox("Could not delete " & vbNewLine & FilePath & vbNewLine & _
              "Cancel operation?", vbYesNo) = vbYes Then
        End
    Else
        Exit Sub
    End If
End Sub

Public Function FileExists(ByVal FileName As String) As Boolean
    If InStr(1, FileName, "\") = 0 Then Exit Function
    If Right(FileName, 1) = "\" Then FileName = Left(FileName, Len(FileName) - 1)
    FileExists = (Dir(FileName, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "")
End Function

