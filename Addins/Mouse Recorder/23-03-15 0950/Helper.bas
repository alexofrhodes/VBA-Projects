Attribute VB_Name = "Helper"
#If VBA7 And Win64 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public MouseFolder As String

Rem mouse
Public MouseArray() As Variant
Rem declaration for keys event reading
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Rem declaration for mouse events
Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_ABSOLUTE = &H8000
Rem declaration for setting mouse position
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Rem declaration for getting mouse position
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
    x As Long
    Y As Long
End Type

Sub FollowLink(FolderPath As String)
    Dim oShell As Object
    Dim Wnd As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.Name = "File Explorer" Then
            If Wnd.Document.Folder.Self.path = FolderPath Then Exit Sub
        End If
    Next Wnd
    Application.ThisWorkbook.FollowHyperlink Address:=FolderPath, NewWindow:=True
End Sub

Function InputboxString(Optional sTitle As String = "Select String", Optional sPrompt As String = "Select String", Optional DefaultValue = "") As String
    Dim stringVariable As String
    stringVariable = Application.InputBox( _
                     title:=sTitle, _
                     Prompt:=sPrompt, _
                     Type:=2, _
                     Default:=DefaultValue)
    InputboxString = CStr(stringVariable)
End Function

Public Function CLIP(Optional StoreText As String) As String
    Dim x As Variant
    x = StoreText
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(StoreText)
                    .SetData "text", x
                Case Else
                    CLIP = .GetData("text")
            End Select
        End With
    End With
End Function

Function IsFileFolderURL(path) As String
    '#INCLUDE FolderExists
    '#INCLUDE FileExists
    '#INCLUDE HttpExists
    Dim retval
    retval = "I"
    If (retval = "I") And FileExists(path) Then retval = "F"
    If (retval = "I") And FolderExists(path) Then retval = "D"
    If (retval = "I") And HttpExists(path) Then retval = "U"
    ' I => Invalid | F => File | D => Directory | U => Valid Url
    CheckPath = retval
End Function

Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)
    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory)        'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If
    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

Function FolderExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Function TrailingSlash(varIn As Variant) As String
    If Len(varIn) > 0 Then
        If Right(varIn, 1) = "\" Then
            TrailingSlash = varIn
        Else
            TrailingSlash = varIn & "\"
        End If
    End If
End Function

Function HttpExists(ByVal sURL As String) As Boolean
    Dim oXHTTP As Object
    Set oXHTTP = CreateObject("MSXML2.XMLHTTP")
    If Not UCase(sURL) Like "HTTP:*" Then
        sURL = "http://" & sURL
    End If
    On Error GoTo haveError
    oXHTTP.Open "HEAD", sURL, False
    oXHTTP.send
    HttpExists = IIf(oXHTTP.Status = 200, True, False)
    Exit Function
haveError:
    Debug.Print Err.Description
    HttpExists = False
End Function

Function ListboxSelectedCount(listboxCollection As Variant) As Long
    Dim i As Long
    Dim listItem As Long
    Dim selectedCollection As Collection
    Set selectedCollection = New Collection
    Dim listboxCount As Long
    'if arguement passed is collection of listboxes
    If TypeName(listboxCollection) = "Collection" Then
        For listboxCount = 1 To listboxCollection.Count
            If listboxCollection(listboxCount).ListCount > 0 Then
                For listItem = 0 To listboxCollection(listboxCount).ListCount - 1
                    If listboxCollection(listboxCount).Selected(listItem) = True Then
                        SelectedCount = SelectedCount + 1
                    End If
                Next listItem
            End If
        Next listboxCount
        'if arguement passed is single Listbox
    Else
        If listboxCollection.ListCount > 0 Then
            For i = 0 To listboxCollection.ListCount - 1
                If listboxCollection.Selected(i) = True Then
                    SelectedCount = SelectedCount + 1
                End If
            Next i
        End If
    End If
    ListboxSelectedCount = SelectedCount
End Function

Function ListboxSelectedIndexes(Lbox As MSForms.ListBox) As Collection
    'listboxes start at 0
    Dim i As Long
    Dim selectedIndexes As Collection
    Set selectedIndexes = New Collection
    If Lbox.ListCount > 0 Then
        For i = 0 To Lbox.ListCount - 1
            If Lbox.Selected(i) Then selectedIndexes.Add i
        Next i
    End If
    Set ListboxSelectedIndexes = selectedIndexes
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


Function WorkbookProjectProtected(ByVal TargetWorkbook As Workbook) As Boolean
        WorkbookProjectProtected = (TargetWorkbook.VBProject.Protection = 1)
End Function


Public Function ArrayDimensions(ByVal vArray As Variant) As Long
    Dim dimnum      As Long
    Dim ErrorCheck As Long
    On Error GoTo FinalDimension
    For dimnum = 1 To 60000
        ErrorCheck = LBound(vArray, dimnum)
    Next
FinalDimension:
    ArrayDimensions = dimnum - 1
End Function

Function GetInputRange(ByRef rInput As Excel.Range, _
                    sPrompt As String, _
                    sTitle As String, _
                    Optional ByVal sDefault As String, _
                    Optional ByVal bActivate As Boolean, _
                    Optional x, _
                    Optional Y) As Boolean

'assigns range to variable passed
'GetInputRange(rng, "Range picker", "Select range to output listbox' list") = False Then Exit Sub
    Dim bGotRng As Boolean
    Dim bEvents As Boolean
    Dim nAttempt As Long
    Dim sAddr As String
    Dim vReturn
    On Error Resume Next
    If Len(sDefault) = 0 Then
        If TypeName(Application.Selection) = "Range" Then
            sDefault = "=" & Application.Selection.Address
            If Len(sDefault) > 240 Then
                sDefault = "=" & Application.ActiveCell.Address
            End If
        ElseIf TypeName(Application.ActiveSheet) = "Chart" Then
            sDefault = " first select a Worksheet"
        Else
            sDefault = " Select Cell(s) or type address"
        End If
    End If
    Set rInput = Nothing
    For nAttempt = 1 To 3
        vReturn = False
        vReturn = Application.InputBox(sPrompt, sTitle, sDefault, x, Y, Type:=0)
        If False = vReturn Or Len(vReturn) = 0 Then
            Exit For
        Else
            sAddr = vReturn
            If Left$(sAddr, 1) = "=" Then sAddr = Mid$(sAddr, 2, 256)
            If Left$(sAddr, 1) = Chr(34) Then sAddr = Mid$(sAddr, 2, 255)
            If Right$(sAddr, 1) = Chr(34) Then sAddr = Left$(sAddr, Len(sAddr) - 1)
            Set rInput = Application.Range(sAddr)
            If rInput Is Nothing Then
                sAddr = Application.ConvertFormula(sAddr, xlR1C1, xlA1)
                Set rInput = Application.Range(sAddr)
                bGotRng = Not rInput Is Nothing
            Else
                bGotRng = True
            End If
        End If
        If bGotRng Then
            If bActivate Then
                On Error GoTo ErrH
                bEvents = Application.EnableEvents
                Application.EnableEvents = False
                If Not Application.ActiveWorkbook Is rInput.Parent.Parent Then
                    rInput.Parent.Parent.Activate
                End If
                If Not Application.ActiveSheet Is rInput.Parent Then
                    rInput.Parent.Activate
                End If
                rInput.Select
            End If
            Exit For
        ElseIf nAttempt < 3 Then
            If MsgBox("Invalid reference, do you want to try again ?", _
                vbOKCancel, sTitle) <> vbOK Then
                Exit For
            End If
        End If
    Next
cleanup:
    On Error Resume Next
    If bEvents Then
        Application.EnableEvents = True
    End If
    GetInputRange = bGotRng
    Exit Function
ErrH:
    Set rInput = Nothing
    bGotRng = False
    Resume cleanup
End Function
''''''''''''''''''''''''''''''
'Contains the following procedures #1
''''''''''''''''''''''''''''''
'IsInArray

Public Function ArrayContains( _
    ByVal value1 As Variant, _
    ByVal array1 As Variant) _
As Boolean

    '@Description: This function checks if a value is in an array
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: value1 is the value that will be checked if its in the array
    '@Param: array1 is the array
    '@Returns: Returns boolean True if the value is in the array, and false otherwise
    '@Example: =IsInArray("hello", {"one", 2, "hello"}) -> True
    '@Example: =IsInArray("hello", {1, "two", "three"}) -> False

    Dim individualElement As Variant
    
    For Each individualElement In array1
        If individualElement = value1 Then
            ArrayContains = True
            Exit Function
        End If
    Next

    ArrayContains = False

End Function

Function WorksheetExists(SheetName As String, TargetWorkbook As Workbook) As Boolean
    Dim TargetWorksheet  As Worksheet
    On Error Resume Next
    Set TargetWorksheet = TargetWorkbook.SHEETS(SheetName)
    On Error GoTo 0
    WorksheetExists = Not TargetWorksheet Is Nothing
End Function

Function CreateOrSetSheet(SheetName As String, TargetWorkbook As Workbook) As Worksheet
    Dim NewSheet As Worksheet
    If WorksheetExists(SheetName, TargetWorkbook) = True Then
        Set CreateOrSetSheet = TargetWorkbook.SHEETS(SheetName)
    Else
        Set CreateOrSetSheet = TargetWorkbook.SHEETS.Add
        CreateOrSetSheet.Name = SheetName
    End If
End Function

