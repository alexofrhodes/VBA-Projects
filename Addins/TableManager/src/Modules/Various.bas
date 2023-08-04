Attribute VB_Name = "Various"
Option Explicit

Public Type tCursor
    Left As Long
    Top As Long
End Type

Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
    Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Public Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As tCursor) As Long
    Public Declare PtrSafe Function GetCursorPos Lib "user32" (p As tCursor) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
    Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As tCursor) As Long
    public Declare Function GetCursorPos Lib "user32" (p As tCursor) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
    Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
#End If

Enum MyColors
    FormBackgroundDarkGray = 4208182            ' BACKGROUND DARK GRAY
    FormSidebarMediumGray = 5457992             ' TILE COLORS LIGHTER GRAY
    FormSidebarMouseOverLightGray = &H808080    ' lighter light gray
    FormSelectedGreen = 8435998                 ' green tile
End Enum

Sub TableManagerButtonClicked(control As IRibbonControl)
    uTableManager.Show
End Sub



Public Function IsValueFormattedCorrectly(ByVal inputRange As Range, ByVal inputValue As Variant) As Boolean
    If IsEmpty(inputValue) Then
        IsValueFormattedCorrectly = True ' Empty values are always valid
        Exit Function
    End If

    On Error Resume Next
    Dim cellFormat As String
    cellFormat = inputRange.NumberFormat
    On Error GoTo 0

    If cellFormat = "General" Then
        IsValueFormattedCorrectly = True ' General format allows any value
        Exit Function
    End If

    On Error Resume Next
    Dim formattedValue As Variant
    formattedValue = Format(inputValue, cellFormat)
    On Error GoTo 0

    If Not IsError(formattedValue) Then
        ' The value can be formatted correctly, now compare the formatted value with the original input value
        IsValueFormattedCorrectly = (formattedValue = inputValue)
    Else
        ' The value cannot be formatted correctly
        IsValueFormattedCorrectly = False
    End If
End Function







Public Function aSwitch(CheckThis, ParamArray OptionPairs() As Variant)
'@LastModified 2307171814
    Dim i As Long
    For i = LBound(OptionPairs) To UBound(OptionPairs) Step 2
        If UCase(CheckThis) = UCase(OptionPairs(i)) Then
            aSwitch = OptionPairs(i + 1)
            Exit Function
        End If
    Next
End Function

Function CountOfCharacters(SearchHere As String, FindThis As String)
    CountOfCharacters = (Len(SearchHere) - Len(Replace(SearchHere, FindThis, ""))) / Len(FindThis)
End Function
Function AvailableFormOrFrameRow(FormOrFrame As Object, Optional AfterWidth As Long = 0, Optional AfterHeight As Long = 0, Optional AddMargin As Long = 0) As Long
    Dim ctr As MSForms.control
    Dim myHeight
    For Each ctr In FormOrFrame.Controls
        If ctr.Visible = True Then
            If ctr.Left >= AfterWidth And ctr.Top >= AfterHeight Then
                If ctr.Top + ctr.Height > myHeight Then myHeight = ctr.Top + ctr.Height
            End If
        End If
    Next
    AvailableFormOrFrameRow = myHeight + AddMargin '6
End Function

Function AvailableFormOrFrameColumn(FormOrFrame As Object, Optional AfterWidth As Long = 0, Optional AfterHeight As Long = 0, Optional AddMargin As Long = 0) As Long
    Dim ctr As MSForms.control
    Dim myWidth
    For Each ctr In FormOrFrame.Controls
        If ctr.Visible = True Then
            If ctr.Left >= AfterWidth And ctr.Top >= AfterHeight Then
                If ctr.Left + ctr.Width > myWidth Then myWidth = ctr.Left + ctr.Width
            End If
        End If
    Next
    AvailableFormOrFrameColumn = myWidth + AddMargin '6
End Function

Function Transpose2DArray(inputArray As Variant) As Variant
    Dim x As Long, yUbound As Long
    Dim y As Long, xUbound As Long
    Dim tempArray As Variant
    xUbound = UBound(inputArray, 2)
    yUbound = UBound(inputArray, 1)
    ReDim tempArray(1 To xUbound, 1 To yUbound)
    For x = 1 To xUbound
        For y = 1 To yUbound
            tempArray(x, y) = inputArray(y, x)
        Next y
    Next x
    Transpose2DArray = tempArray
End Function


Function IsFileFolderURL(Path) As String
    Dim retVal As String
    retVal = "I"
    If (retVal = "I") And FileExists(Path) Then retVal = "F"
    If (retVal = "I") And FolderExists(Path) Then retVal = "D"
    If (retVal = "I") And URLExists(Path) Then retVal = "U"
    ' I => Invalid | F => File | D => Directory | U => Valid Url
    IsFileFolderURL = retVal
End Function

Function URLExists(url) As Boolean
    Dim Request As Object
    Dim FF As Integer
    Dim rc As Variant

    On Error GoTo EndNow
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")

    With Request
      .Open "GET", url, False
      .send
      rc = .statusText
    End With
    Set Request = Nothing
    If rc = "OK" Then URLExists = True

    Exit Function
EndNow:
End Function


Function FolderExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Public Function FileExists(ByVal FileName As String) As Boolean
    If InStr(1, FileName, "\") = 0 Then Exit Function
    If Right(FileName, 1) = "\" Then FileName = Left(FileName, Len(FileName) - 1)
    On Error Resume Next
    FileExists = (Dir(FileName, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "")
End Function


Function WorkbookProjectProtected(ByVal TargetWorkbook As Workbook) As Boolean
        WorkbookProjectProtected = (TargetWorkbook.VBProject.Protection = 1)
End Function



Function TxtRead(sPath As Variant) As String

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

Function GetInputRange(ByRef rInput As Excel.Range, _
                    sPrompt As String, _
                    sTitle As String, _
                    Optional ByVal sDefault As String, _
                    Optional ByVal bActivate As Boolean, _
                    Optional x, _
                    Optional y) As Boolean

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
        vReturn = Application.InputBox(sPrompt, sTitle, sDefault, x, y, Type:=0)
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
Public Function ArrayContains( _
    ByVal value1 As Variant, _
    ByVal array1 As Variant, _
    Optional CaseSensitive As Boolean) _
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
    If CaseSensitive = True Then value1 = UCase(value1)
    For Each individualElement In array1
        If CaseSensitive = True Then individualElement = UCase(individualElement)
        If individualElement = value1 Then
            ArrayContains = True
            Exit Function
        End If
    Next
    ArrayContains = False
End Function

Public Function ArrayAllocated(ByVal arr As Variant) As Boolean
    On Error Resume Next
    ArrayAllocated = IsArray(arr) And (Not IsError(LBound(arr, 1))) And LBound(arr, 1) <= UBound(arr, 1)
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

Sub FollowLink(FolderPath As String)

    If Right(FolderPath, 1) = "\" Then FolderPath = Left(FolderPath, Len(FolderPath) - 1)
    On Error Resume Next
    Dim oShell As Object
    Dim Wnd As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.Name = "File Explorer" Then
            If Wnd.document.Folder.Self.Path = FolderPath Then Exit Sub
        End If
    Next Wnd
    Application.ThisWorkbook.FollowHyperlink Address:=FolderPath, NewWindow:=True
End Sub



Function ArrayTrim(ByVal arr As Variant)
        Dim i As Long
        For i = LBound(arr) To UBound(arr)
            arr(i) = Trim(arr(i))
        Next
        ArrayTrim = arr
End Function

