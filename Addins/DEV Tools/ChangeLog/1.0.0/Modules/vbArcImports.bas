Attribute VB_Name = "vbArcImports"
Option Explicit

Enum MyColors
    FormBackgroundDarkGray = 4208182        ' BACKGROUND DARK GRAY
    FormSidebarMediumGray = 5457992        ' TILE COLORS LIGHTER GRAY
    FormSidebarMouseOverLightGray = &H808080        ' lighter light gray
    FormSelectedGreen = 8435998        ' green tile
End Enum

Public Type tCursor
    Left            As Long
    Top             As Long
End Type

Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
    Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Public Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As tCursor) As Long
    Public Declare PtrSafe Function GetCursorPos Lib "user32" (p As tCursor) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
    Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
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


Rem DspErrMsg Constants and Variables
Global Const success        As Boolean = True
Global Const Failure        As Boolean = False
Global Const NoError        As Long = 0
Global Const LogError       As Long = 997
Global Const RtnError       As Long = 998
Global Const DspError       As Long = 999
Public bLogOnly     As Boolean
Public bDebug       As Boolean

Rem timer constants
Public Const mblncTimer As Boolean = True
Public mvarTimerName
Public mvarTimerStart

Function CanCreateAndEditWorksheet() As Boolean
    Dim ws As Worksheet
    Dim canCreate As Boolean
    Dim canEdit As Boolean
    Application.ScreenUpdating = False
    ' Try to create a new worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets.Add
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ' You can create a worksheet
        canCreate = True
        ' Try to edit the new worksheet
        On Error Resume Next
        ws.Cells(1, 1).value = "Test"
        ws.rows(1).Delete
        If Err.Number = 0 Then
            ' You can edit the worksheet
            canEdit = True
        End If
        On Error GoTo 0
        ' Delete the new worksheet
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Application.ScreenUpdating = False
    ' Check if you can create and edit the worksheet
    CanCreateAndEditWorksheet = canCreate And canEdit
End Function

Sub FolderDelete(targetFolder As String)
'@LastModified 2310251324
'@INCLUDE PROCEDURE FolderExists
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    targetFolder = targetFolder
    If Right(targetFolder, 1) = "\" Then targetFolder = Left(targetFolder, Len(targetFolder) - 1)
    'Delete specified folder
    If FSO.FolderExists(targetFolder) Then FSO.DeleteFolder targetFolder
'    'Delete all subfolders
'    FSO.DeleteFolder targetFolder & "\*"
End Sub

Public Sub ArraySortByLength(list As Variant, first As Long, last As Long)
    '@INCLUDE PROCEDURE ArraySortByLengthHelper
    Dim Pivot As String
    Dim Low As Long
    Dim High As Long
    Low = first
    High = last
    Pivot = list((first + last) \ 2)
    Do While Low <= High
        Do While Low < last And ArraySortByLengthHelper(list(Low), Pivot)
            Low = Low + 1
        Loop
        Do While High > first And ArraySortByLengthHelper(Pivot, list(High))
            High = High - 1
        Loop
        If Low <= High Then
            Dim Swap As String
            Swap = list(Low)
            list(Low) = list(High)
            list(High) = Swap
            Low = Low + 1
            High = High - 1
        End If
    Loop
    If (first < High) Then ArraySortByLength list, first, High
    If (Low < last) Then ArraySortByLength list, Low, last
End Sub

Private Function ArraySortByLengthHelper(one As Variant, two As Variant) As Boolean
    Select Case True
    Case Len(one) < Len(two)
        ArraySortByLengthHelper = True
    Case Len(one) > Len(two)
        ArraySortByLengthHelper = False
    Case Len(one) = Len(two)
        ArraySortByLengthHelper = LCase$(one) < LCase$(two)
    End Select
End Function

Function RandomStringArray(ByVal rowCount As Long, ByVal columnCount As Long, maxStringLength) As Variant
    '@AssignedModule vbArcImports
    Dim resultArray() As Variant
    ReDim resultArray(1 To rowCount, 1 To columnCount)

    Dim i As Long, j As Long
    Dim alphabet    As String
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"

    ' Seed the random number generator
    Randomize

    ' Generate random strings
    For i = 1 To rowCount
        For j = 1 To columnCount
            Dim randomString As String
            randomString = ""

            ' Length of each random string (you can adjust as needed)
            Dim stringLength As Long
            stringLength = WorksheetFunction.RandBetween(1, maxStringLength)

            Dim k   As Long
            For k = 1 To stringLength
                ' Generate a random index to pick a character from the alphabet
                Dim randomIndex As Long
                randomIndex = Int((Len(alphabet) * Rnd) + 1)

                ' Append the random character to the string
                randomString = randomString & Mid(alphabet, randomIndex, 1)
            Next k

            resultArray(i, j) = randomString
        Next j
    Next i

    RandomStringArray = resultArray
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 21-07-2023 11:49    Alex                (ArrayFilter2D)

Function ArrayFilter2D(inputArray As Variant, matchString As String, Optional targetColumn As Long = -1) As Variant
    '@LastModified 2307211149
    '@INCLUDE PROCEDURE ArraySubSet
    '@AssignedModule vbArcImports
    Dim numRows As Long, numCols As Long
    numRows = UBound(inputArray, 1)
    numCols = UBound(inputArray, 2)

    Dim resultArray() As Variant
    Dim resultIndex As Long
    ReDim resultArray(LBound(inputArray, 1) To numRows, LBound(inputArray, 2) To numCols)
    resultIndex = LBound(resultArray, 1) - 1

    Dim i As Long, j As Long
    For i = LBound(resultArray, 1) To numRows
        Dim rowMatches As Boolean
        rowMatches = False

        If targetColumn = -1 Then
            ' Match any cell in the row using Like operator (case-insensitive)
            For j = LBound(resultArray, 2) To numCols
                If LCase(inputArray(i, j)) Like "*" & LCase(matchString) & "*" Then
                    rowMatches = True
                    Exit For
                End If
            Next j
        ElseIf targetColumn >= LBound(resultArray, 2) And targetColumn <= numCols Then
            ' Match the specified column using Like operator (case-insensitive)
            If LCase(inputArray(i, targetColumn)) Like "*" & LCase(matchString) & "*" Then
                rowMatches = True
            End If
        End If

        ' Copy the matching row to the result array
        If rowMatches Then
            resultIndex = resultIndex + 1
            For j = LBound(resultArray, 2) To numCols
                resultArray(resultIndex, j) = inputArray(i, j)
            Next j
        End If
    Next i

    ' Resize the result array to remove any empty rows
    If resultIndex = -1 Then
        ArrayFilter2D = Array()
    Else
        resultArray = ArraySubSet(resultArray, LBound(resultArray, 1), LBound(resultArray, 2), resultIndex, numCols)
        ' Return the filtered array
        ArrayFilter2D = resultArray
    End If
End Function

Sub PrintConditionalFormatting(TargetWorksheet As Worksheet)
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE ArrayCombine
    '@INCLUDE PROCEDURE ArrayToStringTable
    '@INCLUDE PROCEDURE dp
    Dim ws          As Worksheet
    Dim CFrule      As FormatCondition
    Dim Output()
    ReDim Output(1 To 1, 1 To 4)
    Dim i           As Long
    For i = 1 To 4
        Output(1, i) = Choose(i, "Sheet", "Formula", "Range", "Fill")
    Next
    Dim arr(1 To 1, 1 To 4)

    For Each CFrule In TargetWorksheet.Cells.FormatConditions
        For i = 1 To 4

            arr(1, i) = Choose(i, _
                    TargetWorksheet.Name, _
                    "'" & CFrule.Formula1, _
                    CFrule.AppliesTo.Address, _
                    CFrule.Interior.color)



        Next
        ArrayCombine Output, arr, True
    Next CFrule
    dp ArrayToStringTable(Output)
End Sub

Public Function ArraySubSet(vIn As Variant, Optional ByVal iStartRow As Integer = -1, Optional ByVal iStartCol As Integer = -1, Optional ByVal iHeight As Integer = -1, Optional ByVal iWidth As Integer = -1) As Variant
    '@LastModified 2307201044
    '@AssignedModule vbArcImports
    Dim vReturn     As Variant
    Dim iInRowLower As Integer
    Dim iInRowUpper As Integer
    Dim iInColLower As Integer
    Dim iInColUpper As Integer
    Dim iEndRow     As Integer
    Dim iEndCol     As Integer
    Dim irow        As Integer
    Dim iCol        As Integer

    iInRowLower = LBound(vIn, 1)
    iInRowUpper = UBound(vIn, 1)
    iInColLower = LBound(vIn, 2)
    iInColUpper = UBound(vIn, 2)

    If iStartRow = -1 Then
        iStartRow = iInRowLower
    End If
    If iStartCol = -1 Then
        iStartCol = iInColLower
    End If

    If iHeight = -1 Then
        iHeight = iInRowUpper - iStartRow + 1
    End If
    If iWidth = -1 Then
        iWidth = iInColUpper - iStartCol + 1
    End If

    iEndRow = iStartRow + iHeight - IIf(iStartRow = 0, 0, 1)
    iEndCol = iStartCol + iWidth - IIf(iStartCol = 0, 0, 1)

    ReDim vReturn(iStartRow To iEndRow, iStartCol To iEndCol)

    For irow = iStartRow To iEndRow
        For iCol = iStartCol To iEndCol
            vReturn(irow, iCol) = vIn(irow, iCol)
        Next
    Next

    ArraySubSet = vReturn
End Function

Public Function ArrayCombine(ByRef a As Variant, b As Variant, Optional stacked As Boolean = True) As Boolean
    'assumes that A and B are 2-dimensional variant arrays
    'if stacked is true then A is placed on top of B    in this case the number of rows must be the same,
    'otherwise they are placed side by side A|B         in which case the number of columns are the same
    'LBound can be anything but is assumed to be the SAME for A and B (in both dimensions)
    '@AssignedModule vbArcImports

    'False is returned if a clash, so use: If not arraycombe(a,b,true) then goto errorHandler

    Dim LB As Long, m_A As Long, n_A As Long
    Dim m_B As Long, n_B As Long
    Dim M As Long, n As Long
    Dim i As Long, j As Long, k As Long
    Dim c           As Variant

    If TypeName(a) = "Range" Then a = a.value
    If TypeName(b) = "Range" Then b = b.value

    LB = LBound(a, 1)
    m_A = UBound(a, 1)
    n_A = UBound(a, 2)
    m_B = UBound(b, 1)
    n_B = UBound(b, 2)

    If stacked Then
        M = m_A + m_B + 1 - LB
        n = n_A
        If n_B <> n Then
            ArrayCombine = False
            Exit Function
        End If
    Else
        M = m_A
        If m_B <> M Then
            ArrayCombine = False
            Exit Function
        End If
        n = n_A + n_B + 1 - LB
    End If

    ReDim c(LB To M, LB To n)
    For i = LB To M
        For j = LB To n
            If stacked Then
                If i <= m_A Then
                    c(i, j) = a(i, j)
                Else
                    c(i, j) = b(LB + i - m_A - 1, j)
                End If
            Else
                If j <= n_A Then
                    c(i, j) = a(i, j)
                Else
                    c(i, j) = b(i, LB + j - n_A - 1)
                End If
            End If
        Next j
    Next i
    a = c

End Function

Public Function ArrayToStringTable(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional SikiriMoji$ = "|") As String
    'https://gist.github.com/YujiFukami/15c922d41ff06c9b12ad450a14131080#file-
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE ShortenToByteCharacters
    Dim i&, j&, k&, M&, n&
    Dim TateMin&, TateMax&, YokoMin&, YokoMax&
    Dim WithTableHairetu
    Dim NagasaList, MaxNagasaList
    Dim NagasaOnajiList
    Dim OutputList
    '    Const SikiriMoji$ = "|"
    Dim Jigen2%
    On Error Resume Next
    Jigen2 = UBound(Hairetu, 2)
    '    On Error GoTo 0
    If Jigen2 = 0 Then
        Hairetu = Application.Transpose(Hairetu)
    End If
    TateMin = LBound(Hairetu, 1)
    TateMax = UBound(Hairetu, 1)
    YokoMin = LBound(Hairetu, 2)
    YokoMax = UBound(Hairetu, 2)
    ReDim WithTableHairetu(1 To TateMax - TateMin + 1 + 1, 1 To YokoMax - YokoMin + 1 + 1)
    For i = 1 To TateMax - TateMin + 1
        WithTableHairetu(i + 1, 1) = TateMin + i - 1
        For j = 1 To YokoMax - YokoMin + 1
            WithTableHairetu(1, j + 1) = YokoMin + j - 1
            WithTableHairetu(i + 1, j + 1) = Hairetu(i - 1 + TateMin, j - 1 + YokoMin)
        Next j
    Next i
    n = UBound(WithTableHairetu, 1)
    M = UBound(WithTableHairetu, 2)
    ReDim NagasaList(1 To n, 1 To M)
    ReDim MaxNagasaList(1 To M)
    Dim tmpStr$
    For j = 1 To M
        For i = 1 To n
            If j > 1 And HyoujiMaxNagasa <> 0 Then
                tmpStr = WithTableHairetu(i, j)
                WithTableHairetu(i, j) = ShortenToByteCharacters(tmpStr, HyoujiMaxNagasa)
            End If
            NagasaList(i, j) = LenB(StrConv(WithTableHairetu(i, j), vbFromUnicode))
            MaxNagasaList(j) = WorksheetFunction.Max(MaxNagasaList(j), NagasaList(i, j))
        Next i
    Next j
    ReDim NagasaOnajiList(1 To n, 1 To M)
    Dim TmpMaxNagasa&
    For j = 1 To M
        TmpMaxNagasa = MaxNagasaList(j)
        For i = 1 To n
            NagasaOnajiList(i, j) = WithTableHairetu(i, j) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(i, j))
        Next i
    Next j
    ReDim OutputList(1 To n)
    For i = 1 To n
        For j = 1 To M
            If j = 1 Then
                OutputList(i) = NagasaOnajiList(i, j)
            Else
                OutputList(i) = OutputList(i) & SikiriMoji & NagasaOnajiList(i, j)
            End If
        Next j
    Next i
    ArrayToStringTable = Join(OutputList, vbNewLine)
End Function

Public Function isUserform(obj As Object) As Boolean
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE UserformNames
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aModules
    Dim formNames   As New Collection
    Set formNames = aModules.Init(ThisWorkbook).UserformNames
    Dim FormName
    For Each FormName In formNames
        If FormName = obj.Name Then
            isUserform = True
            Exit For
        End If
    Next
End Function
Function IsFileFolderURL(path) As String
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE FileExists
    '@INCLUDE PROCEDURE FolderExists
    '@INCLUDE PROCEDURE URLExists
    Dim RetVal      As String
    RetVal = "I"
    If (RetVal = "I") And FileExists(path) Then RetVal = "F"
    If (RetVal = "I") And FolderExists(path) Then RetVal = "D"
    If (RetVal = "I") And URLExists(path) Then RetVal = "U"
    ' I => Invalid | F => File | D => Directory | U => Valid Url
    IsFileFolderURL = RetVal
End Function

Public Function PadRight(ByVal str As String, ByVal Length As Long, Optional Character As String = " ", Optional RemoveExcess As Boolean)
    '@AssignedModule vbArcImports
    If Len(str) < Length Then
        PadRight = str & String$(Length - Len(str), Character)
    ElseIf RemoveExcess = True Then
        PadRight = Left$(str, Length)
    Else
        PadRight = str
    End If
End Function

Public Function PadLeft(ByVal str As String, ByVal Length As Long, Optional Character As String = " ", Optional RemoveExcess As Boolean)
    '@AssignedModule vbArcImports
    If Len(str) < Length Then
        PadLeft = String$(Length - Len(str), Character) + str
    ElseIf RemoveExcess = True Then
        PadLeft = Right$(str, Length)
    Else
        PadLeft = str
    End If
End Function

Public Function aSwitch(checkThis, ParamArray OptionPairs() As Variant)
    '@LastModified 2307171814
    '@AssignedModule vbArcImports
    Dim i           As Long
    For i = LBound(OptionPairs) To UBound(OptionPairs) Step 2
        If UCase(checkThis) = UCase(OptionPairs(i)) Then
            aSwitch = OptionPairs(i + 1)
            Exit Function
        End If
    Next
End Function

Sub test_ArrayReplace()
    Dim arr
    arr = Selection.value
    ArrayReplace arr, 1, 2
    dp arr
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 25-10-2023 14:33    Alex                (vbArcImports.bas > ArrayReplace) ok for 2d, added start/end Row/Column

Sub ArrayReplace( _
                    ByRef inputArray As Variant, _
                    this As String, _
                    That As String, _
                        Optional startSearchAt As Long = 1, _
                        Optional countChanges As Long = -1, _
                        Optional compareMethod As VbCompareMethod = vbTextCompare, _
                        Optional StartRow As Long, _
                        Optional StartCol As Long, _
                        Optional rowCount As Long = -1, _
                        Optional colCount As Long = -1)
'@LastModified 2310251433
    If StartRow = 0 Then StartRow = LBound(inputArray, 1)
    If StartCol = 0 Then StartCol = LBound(inputArray, 2)
    
    Dim endRow As Long
    If rowCount = -1 Then endRow = UBound(inputArray, 1) - LBound(inputArray, 1) + 1
    Dim endCol As Long
    If colCount = -1 Then endCol = UBound(inputArray, 2) - LBound(inputArray, 2) + 1
    
    Dim x As Long, y As Long
    For x = StartRow To endRow
        For y = StartCol To endCol
            inputArray(x, y) = Replace(inputArray(x, y), this, That, startSearchAt, countChanges, compareMethod)
        Next
    Next
    
End Sub

Public Sub ArraySort(vArray As Variant, Optional inLow As Long = -1, Optional inHi As Long = -1)
    '@BlogPosted
    '@AssignedModule vbArcImports
'    If inLow = -1 Then inLow = LBound(vArray)
'    If inHi = -1 Then inHi = UBound(vArray)
    Dim Pivot       As Variant
    Dim tmpSwap     As Variant
    Dim tmpLow      As Long
    Dim tmpHi       As Long
    tmpLow = inLow
    tmpHi = inHi
    Pivot = vArray((inLow + inHi) \ 2)
    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < Pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        While (Pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    If (inLow < tmpHi) Then ArraySort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then ArraySort vArray, tmpLow, inHi
End Sub
Sub appRunOnTime(timeToRun, macroToRun As String, Optional arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5)
    '@LastModified 2305250729
    '@AssignedModule vbArcImports

    If TypeName(arg5) <> "Error" Then
        Application.OnTime timeToRun, "'" & macroToRun & """" & arg1 & """ ,""" & arg2 & """ ,""" & arg3 & """ ,""" & arg4 & """ ,""" & arg5 & " '"
    ElseIf TypeName(arg4) <> "Error" Then
        Application.OnTime timeToRun, "'" & macroToRun & """" & arg1 & """ ,""" & arg2 & """ ,""" & arg3 & """ ,""" & arg4 & " '"
    ElseIf TypeName(arg3) <> "Error" Then
        Application.OnTime timeToRun, "'" & macroToRun & """" & arg1 & """ ,""" & arg2 & """ ,""" & arg3 & " '"
    ElseIf TypeName(arg2) <> "Error" Then
        Application.OnTime timeToRun, "'" & macroToRun & """" & arg1 & """ ,""" & arg2 & " '"
    ElseIf TypeName(arg1) <> "Error" Then
        Application.OnTime timeToRun, "'" & macroToRun & """" & arg1 & """ '"
    Else
        Application.OnTime timeToRun, macroToRun
    End If
End Sub

Sub appRun(ProcedureName As String, Optional TargetWorkbook As Workbook, Optional arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6, Optional arg7, Optional arg8, Optional arg9, Optional arg10)
    '@LastModified 2305250729
    '@INCLUDE PROCEDURE ActiveCodepaneWorkbook
    '@AssignedModule vbArcImports
    If TypeName(TargetWorkbook) = "Nothing" Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim WorkbookName As String
    WorkbookName = "'" & TargetWorkbook.Name & "'!"

    If TypeName(arg10) <> "Error" Then
        Application.Run WorkbookName & ProcedureName, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9, arg10
    ElseIf TypeName(arg9) <> "Error" Then
        Application.Run WorkbookName & ProcedureName, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9
    ElseIf TypeName(arg8) <> "Error" Then
        Application.Run WorkbookName & ProcedureName, arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8
    ElseIf TypeName(arg7) <> "Error" Then
        Application.Run WorkbookName & ProcedureName, arg1, arg2, arg3, arg4, arg5, arg6, arg7
    ElseIf TypeName(arg6) <> "Error" Then
        Application.Run WorkbookName & ProcedureName, arg1, arg2, arg3, arg4, arg5, arg6
    ElseIf TypeName(arg5) <> "Error" Then
        Application.Run WorkbookName & ProcedureName, arg1, arg2, arg3, arg4, arg5
    ElseIf TypeName(arg4) <> "Error" Then
        Application.Run WorkbookName & ProcedureName, arg1, arg2, arg3, arg4
    ElseIf TypeName(arg3) <> "Error" Then
        Application.Run WorkbookName & ProcedureName, arg1, arg2, arg3
    ElseIf TypeName(arg2) <> "Error" Then
        Application.Run WorkbookName & ProcedureName, arg1, arg2
    ElseIf TypeName(arg1) <> "Error" Then
        Application.Run WorkbookName & ProcedureName, arg1
    Else: Application.Run WorkbookName & ProcedureName
    End If
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 26-10-2023 01:06    Alex                (vbArcImports.bas > InputboxString : )

Function InputboxString(Optional sTitle As String = "Select String", Optional sPrompt As String = "Select String", Optional DefaultValue = "") As String
'@LastModified 2310260106
    '@AssignedModule vbArcImports
    Dim stringVariable As String
    stringVariable = Application.InputBox( _
            Title:=sTitle, _
            Prompt:=sPrompt, _
            Type:=2, _
            Default:=DefaultValue)
    InputboxString = IIf(CStr(stringVariable) <> "False", CStr(stringVariable), "")
End Function

Function LoopThroughFiles(Folder, criteria) As Collection
    '@AssignedModule vbArcImports
    If Right(Folder, 1) <> "\" Then Folder = Folder & "\"
    Dim out         As Collection: Set out = New Collection
    Dim strFile     As String
    strFile = Dir(Folder & criteria)
    Do While Len(strFile) > 0
        out.Add strFile
        strFile = Dir
    Loop
    Set LoopThroughFiles = out
End Function

Function ModuleOfWorksheet(TargetSheet As Worksheet) As VBComponent
    '@LastModified 2305231030
    '@AssignedModule vbArcImports
    Set ModuleOfWorksheet = TargetSheet.Parent.VBProject.VBComponents(TargetSheet.codeName)
End Function

Rem This displays a message box formatted
'based on what the Err object contains and if we want to put our project in debug mode.
'It returns the button the user clicks: vbAbort, vbCancel, vbIgnore, vbRetry

Public Function DspErrMsg(ByVal sRoutineName As String, _
        Optional ByVal sAddText As String = "") As VbMsgBoxResult
    '@AssignedModule vbArcImports
    '@INCLUDE DECLARATION bDebug
    '@INCLUDE DECLARATION bLogOnly
    If bLogOnly Then
        Debug.Print Now(), ThisWorkbook.Name & "!" & sRoutineName, Err.Description, sAddText
    Else
        DspErrMsg = MsgBox( _
                Prompt:="Error#" & Err.Number & vbLf & Err.Description & vbLf & sAddText, _
                Buttons:=IIf(bDebug, vbAbortRetryIgnore, vbCritical) + _
                IIf(Err.Number < 1, 0, vbMsgBoxHelpButton), _
                Title:=sRoutineName, _
                HelpFile:=Err.HelpFile, _
                Context:=Err.HelpContext)
    End If
End Function

Public Function StartTimer(TimerName)
    '@AssignedModule vbArcImports
    '@INCLUDE DECLARATION mblncTimer
    On Error GoTo ERR_HANDLER
    If mblncTimer Then
        mvarTimerName = TimerName
        mvarTimerStart = Timer
    End If
    On Error Resume Next
    Exit Function
ERR_HANDLER:
    MsgBox Err.Number & " " & Err.Description, vbCritical, "StartTimer()"
End Function

Public Function EndTimer()
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE FoldersCreate
    '@INCLUDE PROCEDURE TxtAppend
    '@INCLUDE DECLARATION mblncTimer
    On Error GoTo ERR_HANDLER
    Dim strFile     As String
    Dim strContent  As String
    If mblncTimer Then
        Dim strPath As String
        strPath = Environ$("USERPROFILE") & "\Documents\vbArc\Timers\"
        FoldersCreate strPath
        strFile = strPath & mvarTimerName & ".txt"
        Rem strFile = ThisWorkbook.path & "\" _
            & Left(ThisWorkbook.Name, InStr(1, ThisWorkbook.Name, ".") - 1) _
            & "TimerLog.txt"
        If Len(Dir(strFile)) = 0 Then
            strContent = _
                    "Timestamp" & vbTab & vbTab & vbTab & vbTab & _
                    "ElapsedTime" & vbTab & vbTab & _
                    "TimerName"
            TxtAppend strFile, strContent
        End If
        strContent = Now() & vbTab & vbTab & _
                Format(Timer - mvarTimerStart, "0.00") & vbTab & vbTab & vbTab & _
                mvarTimerName
        TxtAppend strFile, strContent
    End If
    On Error Resume Next
    Exit Function
ERR_HANDLER:
    MsgBox Err.Number & " " & Err.Description, vbCritical, "EndTimer()"
End Function

Function GetInputRange(ByRef rInput As Excel.Range, _
        sPrompt As String, _
        sTitle As String, _
        Optional ByVal sDefault As String, _
        Optional ByVal bActivate As Boolean, _
        Optional x, _
        Optional y) As Boolean
    '@AssignedModule vbArcImports

    'assigns range to variable passed
    'GetInputRange(rng, "Range picker", "Select range to output listbox' list") = False Then Exit Sub
    Dim bGotRng     As Boolean
    Dim bEvents     As Boolean
    Dim nAttempt    As Long
    Dim sAddr       As String
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
CLEANUP:
    On Error Resume Next
    If bEvents Then
        Application.EnableEvents = True
    End If
    GetInputRange = bGotRng
    Exit Function
ErrH:
    Set rInput = Nothing
    bGotRng = False
    Resume CLEANUP
End Function
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
    '@AssignedModule vbArcImports
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
Function OutlookCheck() As Boolean
    '@LastModified 2305220937
    '@AssignedModule vbArcImports
    Dim xOLApp      As Object
    Set xOLApp = CreateObject("Outlook.Application")
    If Not xOLApp Is Nothing Then
        OutlookCheck = True
        Set xOLApp = Nothing
        Exit Function
    End If
    OutlookCheck = False
End Function

Public Function GetInternetConnectedState() As Boolean
    '@LastModified 2305220934
    '@INCLUDE DECLARATION InternetGetConnectedState
    '@AssignedModule vbArcImports
    GetInternetConnectedState = InternetGetConnectedState(0&, 0&)
End Function
Function PickExcelFile()
    '@AssignedModule vbArcImports
    Dim strFile     As String
    Dim fd          As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.clear
        .Filters.Add "Excel Files", "*.xl*", 1
        .Title = "Choose an Excel file"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERprofile") & "\Desktop\"
        If .Show = True Then
            strFile = .SelectedItems(1)
            PickExcelFile = strFile
        End If
    End With
End Function
Function PickFolder(Optional initFolder As String) As String
    '@AssignedModule vbArcImports
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = IIf(initFolder <> "" And FolderExists(initFolder), initFolder, Environ("USERprofile") & "\Desktop\")
        If .Show = -1 Then
            PickFolder = .SelectedItems(1) & "\"
        Else
            Exit Function
        End If
    End With
End Function
Public Function SelectFolder(Optional initFolder As String) As String
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE FolderExists
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Select a folder"
        If FolderExists(initFolder) Then .InitialFileName = initFolder
        .Show
        If .SelectedItems.Count > 0 Then
            SelectFolder = .SelectedItems.item(1)
        Else
        End If
    End With
End Function
Public Function RoundUp(dblNumToRound As Long, lMultiple As Long) As Double
    '@AssignedModule vbArcImports
    Dim asDec       As Variant
    Dim Rounded     As Variant

    asDec = CDec(dblNumToRound) / lMultiple
    Rounded = Int(asDec)

    If Rounded <> asDec Then
        Rounded = Rounded + 1
    End If
    RoundUp = Rounded * lMultiple
End Function
Function StringIndentationNormalize(ByVal txt As String, Optional indentation As Long = 4)
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE RoundUp
    Dim str         As Variant
    str = Split(txt, vbNewLine)
    Dim sLine       As String
    Dim tmpLine     As String
    Dim i           As Long
    Dim iSpaces     As Long
    Dim tmpSpaes    As Long
    For i = LBound(str) To UBound(str)
        sLine = str(i)
        iSpaces = Len(sLine) - Len(LTrim(sLine))
        If iSpaces > 0 Then
            str(i) = Space(RoundUp(iSpaces, indentation)) & Trim(sLine)
        End If
    Next
    If InStr(1, txt, vbNewLine) > 0 Then
        StringIndentationNormalize = Join(str, vbNewLine)
    Else
        StringIndentationNormalize = str
    End If
End Function

Function Parser_Tab(ByVal S As String) As String
    '@LastModified 2305220859
    'https://sites.google.com/site/e90e50/random-topics/tool-for-parsing-formulas-in-excel
    '@AssignedModule vbArcImports
    Dim SS As String, ch As String
    Dim t As Long, z As Long, x As Long

    SS = String(10000, " ")

    t = 1
    z = 1
    For x = 1 To Len(S)
        ch = Mid(S, x, 1)
        If ch = vbCr And x > 1 Then

            If Mid(S, x - 1, 1) = "(" Then
                z = z + 1
            Else
                If Mid(S, x + 1, 1) = ")" Then
                    z = z - 1
                End If
            End If

            Mid(SS, t, z + 1) = ch & Application.Rept(vbTab, z)

            t = t + z
        Else
            Mid(SS, t, 1) = ch
            t = t + 1
        End If
    Next
    S = Left(SS, t - 1)
    Parser_Tab = S

End Function
Function Array_Const_Wrap(ByRef sArraY As String, DelR As String) As String
    '@LastModified 2305220900
    'https://sites.google.com/site/e90e50/random-topics/tool-for-parsing-formulas-in-excel
    '@AssignedModule vbArcImports
    Dim v
    If Len(sArraY) > 1 Then
        v = Split(Mid(sArraY, 2, Len(sArraY) - 2), DelR)
        Array_Const_Wrap = "{" & vbCr & Join(v, DelR & vbCr) & vbCr & "}"
    End If
End Function
Function DataFilePartFolder(fileNameWithExtension, Optional IncludeSlash As Boolean) As String
    '@AssignedModule vbArcImports
    DataFilePartFolder = Left(fileNameWithExtension, InStrRev(fileNameWithExtension, "\") - 1 - IncludeSlash)
End Function

Public Function DataFilePicker(Optional fileType As Variant, Optional multiSelect As Boolean) As Variant
    '@AssignedModule vbArcImports
    Dim blArray     As Boolean
    Dim i           As Long
    Dim strErrMsg As String, strTitle As String
    Dim varItem     As Variant
    If Not IsMissing(fileType) Then
        blArray = IsArray(fileType)
        If Not blArray Then strErrMsg = "Please pass an array in the first parameter of this function!"
    End If
    If strErrMsg = vbNullString Then
        If multiSelect Then strTitle = "Choose one or more files" Else strTitle = "Choose file"
        With Application.FileDialog(msoFileDialogFilePicker)
            .InitialFileName = Environ("USERprofile") & "\Desktop\"
            .AllowMultiSelect = multiSelect
            .Filters.clear
            If blArray Then .Filters.Add "File type", Replace("*." & Join(fileType, ", *."), "..", ".")
            .Title = strTitle
            If .Show <> 0 Then
                ReDim arrResults(1 To .SelectedItems.Count) As Variant
                If blArray Then
                    For Each varItem In .SelectedItems
                        i = i + 1
                        arrResults(i) = varItem
                    Next varItem
                Else
                    arrResults(1) = .SelectedItems(1)
                End If
                DataFilePicker = arrResults
            End If
        End With
    Else
        MsgBox strErrMsg, vbCritical, "Error!"
    End If
End Function

Function DataFilePartExtension(str As String)
    '@AssignedModule vbArcImports
    DataFilePartExtension = Mid(str, InStrRev(str, ".") + 1)
End Function

Function DataFilePartName(fileNameWithExtension As String, Optional IncludeExtension As Boolean) As String
    '@AssignedModule vbArcImports
    If InStr(1, fileNameWithExtension, "\") > 0 Then
        DataFilePartName = Right(fileNameWithExtension, Len(fileNameWithExtension) - InStrRev(fileNameWithExtension, "\"))
    ElseIf InStr(1, fileNameWithExtension, "/") > 0 Then
        DataFilePartName = Right(fileNameWithExtension, Len(fileNameWithExtension) - InStrRev(fileNameWithExtension, "/"))
    Else
        DataFilePartName = fileNameWithExtension
    End If
    If IncludeExtension = False Then DataFilePartName = Left(DataFilePartName, InStr(1, DataFilePartName, ".") - 1)
End Function

Public Function ArrayToString(SourceArray As Variant, Optional Delimiter As String = ",") As String
    '@INCLUDE ArrayDimensionLength
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE ArrayDimensionLength
    Dim Temp        As String
    Select Case ArrayDimensionLength(SourceArray)
        Case 1
            '* @TODO Created: 21-12-2022 19:34 Author: Anastasiou Alex
            '* @TODO find where i put the flattenArray

            Temp = Join(SourceArray, Delimiter)
        Case 2
            Dim rowIndex As Long
            Dim ColIndex As Long
            For rowIndex = LBound(SourceArray, 1) To UBound(SourceArray, 1)
                For ColIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
                    Temp = Temp & SourceArray(rowIndex, ColIndex)
                    If ColIndex <> UBound(SourceArray, 2) Then Temp = Temp & Delimiter
                Next ColIndex
                If rowIndex <> UBound(SourceArray, 1) Then Temp = Temp & vbNewLine
            Next rowIndex
    End Select
    ArrayToString = Temp
End Function
Public Function InputBoxRange(Optional sTitle As String, Optional sPrompt As String) As Range
    '@AssignedModule vbArcImports
    On Error Resume Next
    Set InputBoxRange = Application.InputBox(Title:=sTitle, Prompt:=sPrompt, Type:=8, _
            Default:=IIf(TypeName(Selection) = "Range", Selection.Address, ""))
End Function
Function CreateOrSetSheet(SheetName As String, TargetWorkbook As Workbook) As Worksheet
    '@BlogPosted
    '@INCLUDE PROCEDURE WorksheetExists
    '@AssignedModule vbArcImports
    Dim NewSheet    As Worksheet
    If WorksheetExists(SheetName, TargetWorkbook) = True Then
        Set CreateOrSetSheet = TargetWorkbook.Sheets(SheetName)
    Else
        Set CreateOrSetSheet = TargetWorkbook.Sheets.Add
        CreateOrSetSheet.Name = SheetName
    End If
End Function
Function Parser_Formula( _
        ByVal S As String, _
        sListSeparator As String, _
        sRowSeparator As String) As String
    '@LastModified 2305220859
    'https://sites.google.com/site/e90e50/random-topics/tool-for-parsing-formulas-in-excel
    '@AssignedModule vbArcImports
    Const CW        As String = "[^=\-+*/();:,.$<>^]"
    Dim M As Object, RE As Object, SM As Object, SB As Object
    Dim v As Variant, t As String

    Set RE = CreateObject("vbscript.regexp")
    RE.IgnoreCase = True
    RE.Global = True

    v = Array( _
            "(""[^""]*""|'[^']*')", _
            "(\{[^}]+})", _
            "(\" & sListSeparator & ")", _
            "(" & CW & "+(?:\." & CW & "+)*\()", _
            "(\))", _
            "(^=|\()", _
            "(.)")

    RE.pattern = Join(v, "|")
    If RE.test(S) Then
        Set M = RE.Execute(S)
        S = ""
        For Each SM In M
            Set SB = SM.SubMatches
            If Len(SB(0) & SB(6)) Then
                t = SB(0) & SB(6)
            ElseIf Len(SB(1)) Then
                t = Array_Const_Wrap(SB(1), sRowSeparator) & vbCr
            ElseIf Len(SB(2) & SB(5)) Then
                t = SB(2) & SB(5) & vbCr
            ElseIf Len(SB(3)) Then
                t = vbCr & SB(3) & vbCr
            ElseIf Len(SB(4)) Then
                t = vbCr & SB(4)
            End If
            S = S & t
        Next
    End If

    RE.pattern = "\r{2,}"
    S = RE.Replace(S, vbCr)

    If Left(S, 1) = vbCr Then S = Mid(S, 1 + Len(vbCr))
    If Right(S, 1) = vbCr Then S = Left(S, Len(S) - Len(vbCr))
    Parser_Formula = S
End Function
Function StringFormatFunctionNested( _
        str As String, _
        Optional sListSeparator As String = ",", _
        Optional sRowSeparator As String = ",") As String
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE StringIndentationNormalize
    '@INCLUDE PROCEDURE ArrayRemoveEmptyElements
    Dim txt         As String
    txt = _
            Join( _
            ArrayRemoveEmptyElements( _
            Split( _
            Parser_Tab( _
            Parser_Formula( _
            str, _
            sListSeparator, _
            sRowSeparator _
            ) _
            ), _
            vbCr _
            ) _
            ), _
            " _" & vbNewLine _
            )
    StringFormatFunctionNested = StringIndentationNormalize(txt)
End Function
Function IncreaseAllNumbersInString(str As String)
    '@AssignedModule vbArcImports
    Dim Output      As String
    Dim counter     As Long
    counter = Len(str)
    Dim i           As Long
    For i = 1 To Len(str)
        counter = i
        If IsNumeric(Mid(str, i, 1)) Then
            Do
                Output = Output & Mid(str, counter, 1)
                counter = counter + 1
            Loop While IsNumeric(Mid(str, counter, 1))
            i = counter - 1
            IncreaseAllNumbersInString = IncreaseAllNumbersInString & val(Output + 1)
        Else
            Output = Output & Mid(str, i, 1)
            IncreaseAllNumbersInString = IncreaseAllNumbersInString & Output
        End If
        Output = ""
    Next
End Function
Function ArrayRotate(inputArray, Optional ShiftByNum = 1) As Variant
    'ShiftByNum = Positive number
    Rem @TODO - Rotate right
    Rem rotates array left (first element to end of array)
    '@INCLUDE Len2
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE Len2
    Dim ub          As Long: ub = UBound(inputArray)
    Dim LB          As Long: LB = LBound(inputArray)
    Dim dif         As Long: dif = 1 - LB
    Dim NewArray()  As Variant
    Dim element     As Variant
    Dim counter     As Long
    Dim fromStart   As Long: fromStart = LB
    For counter = LB To ub
        ReDim Preserve NewArray(1 To counter + dif)
        If counter + ShiftByNum <= ub Then
            NewArray(UBound(NewArray)) = inputArray(counter + ShiftByNum)
        ElseIf UBound(NewArray) <= Len2(inputArray) Then
            NewArray(UBound(NewArray)) = inputArray(fromStart)
            fromStart = fromStart + 1
        End If
    Next
    ArrayRotate = NewArray
End Function
Public Function SortSelectionArray(ByVal TempArray As Variant) As Variant
    '@AssignedModule vbArcImports
    Dim MaxVal      As Variant
    Dim MaxIndex    As Integer
    Dim i As Integer, j As Integer
    For i = UBound(TempArray) To 0 Step -1
        MaxVal = TempArray(i)
        MaxIndex = i
        For j = 0 To i
            If TempArray(j) > MaxVal Then
                MaxVal = TempArray(j)
                MaxIndex = j
            End If
        Next j
        If MaxIndex < i Then
            TempArray(MaxIndex) = TempArray(i)
            TempArray(i) = MaxVal
        End If
    Next i
    SortSelectionArray = TempArray
End Function
Public Function RegExpReplace( _
        TEXT As String, _
        pattern As String, _
        text_replace As String, _
        Optional instance_num As Integer = 0, _
        Optional match_case As Boolean = True) As String
    '@AssignedModule vbArcImports
    Dim text_result, text_find As String
    Dim matches_index, pos_start As Integer

    On Error GoTo ErrHandle
    text_result = TEXT
    Dim REGEX       As RegExp
    Set REGEX = CreateObject("VBScript.RegExp")

    REGEX.pattern = pattern
    REGEX.Global = True
    REGEX.MultiLine = True

    If True = match_case Then
        REGEX.IgnoreCase = False
    Else
        REGEX.IgnoreCase = True
    End If
    Dim matches     As Object
    Set matches = REGEX.Execute(TEXT)

    If 0 < matches.Count Then
        If (0 = instance_num) Then
            text_result = REGEX.Replace(TEXT, text_replace)
        Else
            If instance_num <= matches.Count Then
                pos_start = 1
                For matches_index = 0 To instance_num - 2
                    pos_start = InStr(pos_start, TEXT, matches.item(matches_index), vbBinaryCompare) + Len(matches.item(matches_index))
                Next matches_index

                text_find = matches.item(instance_num - 1)
                text_result = Left(TEXT, pos_start - 1) & Replace(TEXT, text_find, text_replace, pos_start, 1, vbBinaryCompare)

            End If
        End If
    End If

    RegExpReplace = text_result
    Exit Function

ErrHandle:
    RegExpReplace = CVErr(xlErrValue)
End Function

Function InStrExact( _
        Start As Long, _
        sourceText As String, _
        WordToFind As String, _
        Optional CaseSensitive As Boolean = False, _
        Optional AllowAccentedCharacters As Boolean = False) As Long
    '@BlogPosted
    '@AssignedModule vbArcImports
    Dim x As Long, Str1 As String, str2 As String, pattern As String
    Const UpperAccentsOnly As String = "ÇÉÑ"
    Const UpperAndLowerAccents As String = "ÇÉÑçéñ"
    If CaseSensitive Then
        Str1 = sourceText
        str2 = WordToFind
        pattern = "[!A-Za-z0-9]"
        If AllowAccentedCharacters Then pattern = Replace(pattern, "!", "!" & UpperAndLowerAccents)
    Else
        Str1 = UCase(sourceText)
        str2 = UCase(WordToFind)
        pattern = "[!A-Z0-9]"
        If AllowAccentedCharacters Then pattern = Replace(pattern, "!", "!" & UpperAccentsOnly)
    End If
    For x = Start To Len(Str1) - Len(str2) + 1
        If Mid(" " & Str1 & " ", x, Len(str2) + 2) Like pattern & str2 & pattern _
                And Not Mid(Str1, x) Like str2 & "'[" & Mid(pattern, 3) & "*" Then
            InStrExact = x
            Exit Function
        End If
    Next
End Function
Function ArrayFilterLike(inputArray As Variant, MatchThis As String, MatchCase As Boolean)
    '@AssignedModule vbArcImports
    Dim OutputArray As Variant
    ReDim OutputArray(1 To 1)
    Dim counter     As Long
    counter = 0
    Dim element
    Dim doesMatch   As Boolean
    For Each element In inputArray
        doesMatch = IIf(MatchCase, _
                element Like MatchThis, _
                UCase(element) Like UCase(MatchThis))
        If doesMatch Then
            counter = counter + 1
            ReDim Preserve OutputArray(1 To counter)
            OutputArray(UBound(OutputArray)) = element
        End If
    Next
    ArrayFilterLike = OutputArray
End Function


Function StringCommentsRemove(ByVal txt As String, RemoveRem As Boolean) As String
    '@BlogPosted
    'modified from Jacob Hilderbrand's code, found at
    'http://www.vbaexpress.com/kb/getarticle.php?kb_id=266
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE ArrayRemoveEmptyElements
    Dim var         As Variant
    ReDim var(0)
    Dim str
    str = Split(txt, vbNewLine)
    '        str = ArrayRemoveEmptyElements(str)
    Dim n           As Long
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim l           As Long
    Dim lineText    As String
    Dim QUOTES      As Long
    Dim q           As Long
    Dim startPos    As Long

    For j = LBound(str) To UBound(str)
        lineText = LTrim(str(j))
        If RemoveRem Then If lineText Like "Rem *" Then GoTo SKIP
        startPos = 1
retry:
        n = InStr(startPos, lineText, "'")
        q = InStr(startPos, lineText, """")
        QUOTES = 0
        If q < n Then
            For l = 1 To n
                If Mid(lineText, l, 1) = """" Then
                    QUOTES = QUOTES + 1
                End If
            Next l
        End If
        If QUOTES = Application.WorksheetFunction.Odd(QUOTES) Then
            startPos = n + 1
            GoTo retry:
        Else
            Select Case n
                Case Is = 0
                    '                    If Len(lineText) > 0 Then
                    var(UBound(var)) = str(j)
                    If j < UBound(str) Then ReDim Preserve var(UBound(var) + 1)
                    '                    End If
                Case Is = 1
                    '
                Case Is > 1
                    var(UBound(var)) = Left(str(j), n - 1)
                    If j < UBound(str) Then ReDim Preserve var(UBound(var) + 1)
            End Select
        End If
SKIP:
    Next j
    '    var = ArrayRemoveEmptyElements(var)
    StringCommentsRemove = Join(var, vbNewLine)
End Function

Public Function IsLineNumberAble( _
        ByVal str As String) As Boolean
    '@AssignedModule vbArcImports
    Dim test        As String
    test = Trim(str)
    If Len(test) = 0 Then Exit Function
    If Right(test, 1) = ":" Then Exit Function
    If IsNumeric(Left(test, 1)) Then Exit Function
    If test Like "'*" Then Exit Function
    If test Like "Rem*" Then Exit Function
    If test Like "Dim*" Then Exit Function
    If test Like "Sub*" Then Exit Function
    If test Like "Public*" Then Exit Function
    If test Like "Private*" Then Exit Function
    If test Like "Function*" Then Exit Function
    If test Like "End Sub*" Then Exit Function
    If test Like "End Function*" Then Exit Function
    If test Like "Debug*" Then Exit Function
    IsLineNumberAble = True
End Function
Public Function NumberOfArrayDimensions(arr As Variant) As Byte
    '@AssignedModule vbArcImports
    Dim Ndx         As Byte
    Dim Res         As Long
    On Error Resume Next
    Do
        Ndx = Ndx + 1
        Res = UBound(arr, Ndx)
    Loop Until Err.Number <> 0
    NumberOfArrayDimensions = Ndx - 1
End Function
Function LargestLength(Optional MyObj As Variant) As Long
    '@LastModified 2305220815
    '@AssignedModule vbArcImports
    LargestLength = 0
    Dim element     As Variant
    If IsMissing(MyObj) Then
        If TypeName(Selection) = "Range" Then
            Set MyObj = Selection
        Else
            Exit Function
        End If
    End If
    Select Case TypeName(MyObj)
        Case Is = "String"
            LargestLength = Len(MyObj)

        Case "Variant", "Variant()", "String()", "Collection"
            For Each element In MyObj
                Select Case TypeName(element)
                    Case Is = "String"
                        If Len(CStr(element)) > LargestLength Then LargestLength = Len(CStr(element))
                    Case "Integer", "Long", "Single", "Date"
                        If element > LargestLength Then LargestLength = element
                    Case Else
                        If element.Width > LargestLength Then LargestLength = element.Width
                End Select
            Next element

        Case Else
    End Select
End Function
Function StringFormatAlignRowsElements( _
                                      txt As String, _
                                      AlignAtString As String, _
                                      SearchFromLeft As Boolean, _
                                      Optional AlignAtColumn As Long)
    '@LastModified 2303171105
    '@AssignedModule vbArcImports
    Dim TextLines: TextLines = Split(txt, vbNewLine)
    Dim elementOriginalColumn As Long
    Dim rightMostColumn As Long
    Dim lineText    As String
    Dim numberOfSpacesToInsert As Long
    Dim i           As Long
    
    For i = LBound(TextLines) To UBound(TextLines)
        TextLines(i) = CleanTrim(TextLines(i), True)
    Next
    If AlignAtColumn = 0 Then
        For i = LBound(TextLines) To UBound(TextLines)
            If SearchFromLeft Then
                elementOriginalColumn = InStr(1, TextLines(i), AlignAtString)
            Else
                elementOriginalColumn = InStrRev(TextLines(i), AlignAtString)
            End If
            If elementOriginalColumn > rightMostColumn Then rightMostColumn = elementOriginalColumn
        Next
        AlignAtColumn = rightMostColumn
    End If

    For i = LBound(TextLines) To UBound(TextLines)
        lineText = TextLines(i)
        If SearchFromLeft Then
            elementOriginalColumn = InStr(lineText, AlignAtString)
        Else
            elementOriginalColumn = InStrRev(lineText, AlignAtString)
        End If

        If elementOriginalColumn > 0 Then
            numberOfSpacesToInsert = AlignAtColumn - elementOriginalColumn + IIf(AlignAtString = ":", 1, 0)
            If numberOfSpacesToInsert > 0 Then
'                If AlignAtString = ":" Then
'                    TextLines(i) = Left(TextLines(i), elementOriginalColumn) & _
'                            Space(numberOfSpacesToInsert) & _
'                            Trim(Mid(TextLines(i), elementOriginalColumn + 1))
'                Else

                    TextLines(i) = RTrim(Left(TextLines(i), elementOriginalColumn - 1)) & _
                            Space(numberOfSpacesToInsert) & _
                            Mid(Trim(TextLines(i)), elementOriginalColumn)
'                End If
            End If
        End If
    Next

    StringFormatAlignRowsElements = Join(TextLines, vbNewLine)

End Function

Public Function Combine2Array(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    '@INCLUDE NumberOfArrayDimensions
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE NumberOfArrayDimensions
    Dim LowRowArr1  As Long
    Dim HighRowArr1 As Long
    Dim LowColumnArr1 As Long
    Dim HighColumnArr1 As Long
    Dim NumOfRowsArr1 As Long
    Dim NumOfColumnsArr1 As Long
    Dim LowRowArr2  As Long
    Dim HighRowArr2 As Long
    Dim LowColumnArr2 As Long
    Dim HighColumnArr2 As Long
    Dim NumOfRowsArr2 As Long
    Dim NumOfColumnsArr2 As Long
    Dim Output      As Variant
    Dim OutputRow   As Long
    Dim OutputColumn As Long
    Dim RowIdx      As Long
    Dim ColIdx      As Long
    If (IsArray(arr1) = False) Or (IsArray(arr2) = False) Then
        Combine2Array = Null
        MsgBox "Both need to be array"
        Exit Function
    End If
    If (NumberOfArrayDimensions(arr1) <> 2) Or (NumberOfArrayDimensions(arr2) <> 2) Then
        Combine2Array = Null
        MsgBox "Both need to be 2D array"
        Exit Function
    End If
    LowRowArr1 = LBound(arr1, 1)
    HighRowArr1 = UBound(arr1, 1)
    LowColumnArr1 = LBound(arr1, 2)
    HighColumnArr1 = UBound(arr1, 2)
    NumOfRowsArr1 = HighRowArr1 - LowRowArr1 + 1
    NumOfColumnsArr1 = HighColumnArr1 - LowColumnArr1 + 1
    LowRowArr2 = LBound(arr2, 1)
    HighRowArr2 = UBound(arr2, 1)
    LowColumnArr2 = LBound(arr2, 2)
    HighColumnArr2 = UBound(arr2, 2)
    NumOfRowsArr2 = HighRowArr2 - LowRowArr2 + 1
    NumOfColumnsArr2 = HighColumnArr2 - LowColumnArr2 + 1
    If NumOfColumnsArr1 <> NumOfColumnsArr2 Then
        Combine2Array = Null
        MsgBox "Both array must have same number of column"
        Exit Function
    End If
    ReDim Output(0 To NumOfRowsArr1 + NumOfRowsArr2 - 1, 0 To NumOfColumnsArr1 - 1)
    For RowIdx = LowRowArr1 To HighRowArr1
        OutputColumn = 0
        For ColIdx = LowColumnArr1 To HighColumnArr1
            Output(OutputRow, OutputColumn) = arr1(RowIdx, ColIdx)
            OutputColumn = OutputColumn + 1
        Next ColIdx
        OutputRow = OutputRow + 1
    Next RowIdx
    For RowIdx = LowRowArr2 To HighRowArr2
        OutputColumn = 0
        For ColIdx = LowColumnArr2 To HighColumnArr2
            Output(OutputRow, OutputColumn) = arr2(RowIdx, ColIdx)
            OutputColumn = OutputColumn + 1
        Next ColIdx
        OutputRow = OutputRow + 1
    Next RowIdx
    Combine2Array = Output
End Function

Public Function IndentationCount(str) As Long
    '@AssignedModule vbArcImports
    IndentationCount = Len(str) - Len(LTrim(str))
End Function

Function WorkbookProjectProtected(ByVal TargetWorkbook As Workbook) As Boolean
    '@BlogPosted
    '@AssignedModule vbArcImports
    WorkbookProjectProtected = (TargetWorkbook.VBProject.Protection = 1)
End Function

Function CountOfCharacters(SearchHere As String, FindThis As String)
    '@AssignedModule vbArcImports
    CountOfCharacters = (Len(SearchHere) - Len(Replace(SearchHere, FindThis, ""))) / Len(FindThis)
End Function

Function IsCommentLine(ByVal str As String) As Boolean
    '@LastModified 2305220757
    '@AssignedModule vbArcImports
    str = Trim(str)
    If str Like "'*" Then IsCommentLine = True
    If str Like "Rem *" Then IsCommentLine = True
End Function

Function CommentsMoveToOwnLine(ByVal txt As String) As String
    '@BlogPosted
    '@INCLUDE PROCEDURE CommentsTrim
    '@AssignedModule vbArcImports

    Dim var         As Variant
    ReDim var(0)
    Dim str         As Variant
    str = Split(txt, vbNewLine)

    Dim n           As Long
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim l           As Long
    Dim lineText    As String
    Dim QUOTES      As Long
    Dim q           As Long
    Dim startPos    As Long

    For j = LBound(str) To UBound(str)
        lineText = Trim(str(j))
        startPos = 1
retry:
        n = InStr(startPos, lineText, "'")
        q = InStr(startPos, lineText, """")
        QUOTES = 0
        If q < n Then
            For l = 1 To n
                If Mid(lineText, l, 1) = """" Then
                    QUOTES = QUOTES + 1
                End If
            Next l
        End If
        If QUOTES = Application.WorksheetFunction.Odd(QUOTES) Then
            startPos = n + 1
            GoTo retry:
        Else
            Select Case n
                Case Is = 0
                    var(UBound(var)) = str(j)
                    ReDim Preserve var(UBound(var) + 1)
                Case Is = 1
                    var(UBound(var)) = CommentsTrim(Array(str(j)))
                    ReDim Preserve var(UBound(var) + 1)
                Case Is > 1
                    var(UBound(var)) = Space(Len(str(j)) - Len(LTrim(str(j)))) & Mid(lineText, n)
                    ReDim Preserve var(UBound(var) + 1)
                    var(UBound(var)) = Space(Len(str(j)) - Len(LTrim(str(j)))) & Left(lineText, n - 1)
                    ReDim Preserve var(UBound(var) + 1)
            End Select
        End If
    Next j

    CommentsMoveToOwnLine = Join(var, vbNewLine)
    CommentsMoveToOwnLine = Left(CommentsMoveToOwnLine, Len(CommentsMoveToOwnLine) - Len(vbNewLine))
End Function
Public Function IsBlockEnd(strline As String) As Boolean
    '@BlogPosted
    '@AssignedModule vbArcImports
    strline = Replace(strline, Chr(13), "")
    Dim bOK         As Boolean
    Dim nPos        As Integer
    Dim strTemp     As String
    nPos = InStr(1, strline, " ") - 1
    If nPos < 0 Then nPos = Len(strline)
    strTemp = Left$(strline, nPos)
    Select Case strTemp
        Case "Next", "Loop", "Wend", "Case", "Else", "#Else", "Else:", "#Else:", "ElseIf", "#ElseIf", "#End"
            bOK = True
        Case "End"
            bOK = (Len(strline) > 3)
    End Select
    IsBlockEnd = bOK
End Function

Function TxtAppend(sFile As String, sText As String)
    '@BlogPosted
    '@AssignedModule vbArcImports
    On Error GoTo ERR_HANDLER
    Dim iFileNumber As Integer
    iFileNumber = FreeFile
    Open sFile For Append As #iFileNumber
    Print #iFileNumber, sText
    Close #iFileNumber
Exit_Err_Handler:
    Exit Function
ERR_HANDLER:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
            "Error Number: " & Err.Number & vbCrLf & _
            "Error Source: Txt_Append" & vbCrLf & _
            "Error Description: " & Err.Description & _
            Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
            , vbOKOnly + vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

Function CommentsTrim(ByVal txt As String) As String
    '@LastModified 2305220838
    '@BlogPosted
    '@INCLUDE PROCEDURE ArrayRemoveEmptyElements
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE tmp
    Dim var         As Variant
    ReDim var(0)
    Dim str         As Variant
    str = Split(txt, vbNewLine)
    Dim j           As Long
    Dim dif         As Long
    Dim lineText    As String
    Dim tmp         As String
    For j = LBound(str) To UBound(str)
        lineText = Trim(str(j))
        If Left(lineText, 2) = "' " Then
            tmp = Mid(lineText, 2)
            dif = Len(tmp) - Len(LTrim(tmp))
            var(UBound(var)) = Space(dif) & "'" & LTrim(tmp)
            ReDim Preserve var(UBound(var) + 1)
        Else
            var(UBound(var)) = str(j)
            ReDim Preserve var(UBound(var) + 1)
        End If
    Next
    var = ArrayRemoveEmptyElements(var)
    CommentsTrim = Join(var, vbNewLine)
End Function
Public Function ArrayRemoveEmptyElements(varArray As Variant) As Variant
    '@LastModified 2305220838
    '@BlogPosted
    '@AssignedModule vbArcImports
    Dim TempArray() As Variant
    Dim OldIndex    As Integer
    Dim NewIndex    As Integer
    If UBound(varArray) = -1 Then Exit Function
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


Function ArrayTrim(ByVal arr As Variant)
    '@BlogPosted
    '@AssignedModule vbArcImports
    Dim i           As Long
    For i = LBound(arr) To UBound(arr)
        arr(i) = Trim(arr(i))
    Next
    ArrayTrim = arr
End Function


Public Function IsBlockStart(strline As String) As Boolean
    '@BlogPosted
    '@AssignedModule vbArcImports
    strline = Replace(strline, Chr(13), "")
    Dim bOK         As Boolean
    Dim nPos        As Integer
    Dim strTemp     As String
    nPos = InStr(1, strline, " ") - 1
    If nPos < 0 Then nPos = Len(strline)
    strTemp = Left$(strline, nPos)
    Select Case strTemp
        Case "With", "For", "Do", "While", "Select", "Case", "Else", "Else:", "#Else", "#Else:", "Sub", "Function", "Property", "Enum", "Type"
            bOK = True
        Case "If", "#If", "ElseIf", "#ElseIf"
            bOK = (Right(strline, 4) = "Then" Or Right(strline, 1) = "_")
        Case "Private", "Public", "Friend"
            nPos = InStr(1, strline, " Static ")
            If nPos Then
                nPos = InStr(nPos + 7, strline, " ")
            Else
                nPos = InStr(Len(strTemp) + 1, strline, " ")
            End If
            On Error GoTo SKIP
            Select Case Mid$(strline, nPos + 1, InStr(nPos + 1, strline, " ") - nPos - 1)
                Case "Sub", "Function", "Property", "Enum", "Type"
                    bOK = True
            End Select
SKIP:
            On Error GoTo 0
    End Select
    IsBlockStart = bOK
End Function


Public Sub dp(var As Variant)
    '@LastModified 2305220815
    '@BlogPosted
    '@INCLUDE DECLARATION i
    '@INCLUDE PROCEDURE PrintXML
    '@INCLUDE PROCEDURE printRange
    '@INCLUDE PROCEDURE printArray
    '@INCLUDE PROCEDURE printCollection
    '@INCLUDE PROCEDURE printDictionary
    '@AssignedModule vbArcImports
    Dim element     As Variant
    Dim i           As Long
    '    Debug.Print TypeName(var)
    Select Case TypeName(var)
        Case Is = "String", "Long", "Integer", "Double", "Boolean"
            Debug.Print var
        Case Is = "Variant()", "String()", "Long()", "Integer()"
            printArray var
        Case Is = "Collection"
            printCollection var
        Case Is = "Dictionary"
            printDictionary var
        Case Is = "Range"
            printRange var
        Case Is = "Date"
            Debug.Print var
        Case Is = "IXMLDOMElement"
            PrintXML var
        Case Else
    End Select
End Sub

Sub PrintXML(var)
    '@BlogPosted
    '@AssignedModule vbArcImports
    Debug.Print var.xml
End Sub
'Sub PrintXML(NodeList)
''   Parse all levels recursively
'    Dim obj
'    On Error Resume Next
'    Set obj = NodeList.ChildNodes
'    If Err.Number = 0 Then
'
'    Else
'        Err.clear
'        Set obj = NodeList.NodeList
'        If Err.Number <> 0 Then: Err.clear: Exit Sub
'    End If
'    On Error GoTo 0
'    Dim child
'    For Each child In obj
'        If Not Left(child.nodename, 1) = "#" Then Debug.Print child.nodename & ":" & child.TEXT
'        If child.ChildNodes.Length > 0 Then PrintXML child.ChildNodes
'    Next
'End Sub

Public Sub PrintLinesContaining(F)
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE ProcedureCode
    '@INCLUDE PROCEDURE ProceduresOfModule
    '@INCLUDE PROCEDURE collectionToString
    '@INCLUDE PROCEDURE WorkbookProjectProtected
    '@INCLUDE PROCEDURE dp
    '@INCLUDE CLASS aWorkbook
    '@INCLUDE CLASS aModule
    '@INCLUDE CLASS aCollection
    Dim i           As Long
    Const ModuleString = vbNewLine & "    Mod|"
    Const Procedurestring = "" & vbTab & "Sub" & "|" & "---" & "| "
    Const FoundString = "" & vbTab & "txt" & "|" & vbTab & " |" & "---" & "| "
    Dim x, y, S, p  As Variant
    Dim module      As VBComponent
    On Error Resume Next
    Dim out         As New Collection
    For Each x In Array(Workbooks, AddIns)
        For Each y In x
            If Not WorkbookProjectProtected(Workbooks(y.Name)) Then
                If Err.Number = 0 Then
                    If UBound(Filter(Split(aWorkbook.Init(Workbooks(y.Name)).Code, vbNewLine), F, True, vbTextCompare)) > -1 Then

                        out.Add ""
                        out.Add String(50, "-")
                        out.Add "| " & y.Name
                        out.Add String(50, "-")

'                        Debug.Print Y.Name

                        For Each module In Workbooks(y.Name).VBProject.VBComponents

'                            Debug.Print vbTab & module.Name

                            If UBound(Filter(Split(aModule.Init(module).Code, vbNewLine), F, True, vbTextCompare)) > -1 Then
                                out.Add ModuleString & module.Name
                                If module.CodeModule.CountOfDeclarationLines > 0 Then
                                    S = Filter(Split(module.CodeModule.Lines(1, module.CodeModule.CountOfDeclarationLines), vbNewLine), F, True, vbTextCompare)
                                    out.Add FoundString & Trim(S(i))
                                End If
                                For Each p In ProceduresOfModule(module)
                                    If UBound(Filter(Split(ProcedureCode(Workbooks(y.Name), module, CStr(p)), vbNewLine), F, True, vbTextCompare)) > -1 Then
                                        out.Add Procedurestring & CStr(p)
                                        S = Filter(Split(ProcedureCode(Workbooks(y.Name), module, CStr(p)), vbNewLine), F, True, vbTextCompare)
                                        For i = 0 To UBound(S)
                                            out.Add FoundString & Trim(S(i))
                                        Next i
                                    End If
                                Next p
                            End If
                        Next module
                    End If
                End If
            End If
            Err.clear
        Next y
    Next x
    dp aCollection.Init(out).ToString(vbNewLine)    'collectionToString(out, vbNewLine)
End Sub

Public Sub printRange(var As Variant)
    '@BlogPosted
    '@INCLUDE PROCEDURE Combine2Array
    '@INCLUDE PROCEDURE dp
    '@AssignedModule vbArcImports
    If var.Areas.Count = 1 Then
        dp var.value
    Else
        Dim out     As Variant
        Dim Temp    As Variant
        Dim i       As Long
        For i = 1 To var.Areas.Count
            Temp = var.Areas(i).value
            If IsEmpty(out) Then
                out = Temp
            Else
                out = Combine2Array(out, Temp)
            End If
        Next
        dp out
    End If
End Sub

Private Sub printArray(var As Variant)
    '@BlogPosted
    '@INCLUDE PROCEDURE DPH
    '@INCLUDE PROCEDURE ArrayDimensions
    '@AssignedModule vbArcImports
    '@INCLUDE PROCEDURE dp
    Dim element
    If ArrayDimensions(var) = 1 Then
        Debug.Print Join(var, vbNewLine)
        '        For Each element In var
        '            dp element
        '        Next
    ElseIf ArrayDimensions(var) > 1 Then
        DPH var
    End If
End Sub

Private Sub printCollection(var As Variant)
    '@BlogPosted
    '@INCLUDE PROCEDURE dp
    '@AssignedModule vbArcImports
    Dim elem        As Variant
    For Each elem In var
        dp elem
    Next elem
End Sub

Private Sub printDictionary(var As Variant)
    '@BlogPosted
    '@INCLUDE PROCEDURE dp
    '@AssignedModule vbArcImports


    '@TODO detect error cause I met when printing a dic from JSON related modules

    Dim i As Long: Dim iCount As Long
    Dim arrKeys
    Dim sKey        As String
    Dim varItem

    Dim Key         As Variant
    For Each Key In var.Keys
        dp var(Key)

    Next Key

    '    Stop

    '    With var
    '        iCount = .Count
    '        arrKeys = .Keys
    '        iCount = UBound(arrKeys, 1)
    '        For i = 0 To iCount
    '            sKey = arrKeys(i)
    '            Debug.Print "Key " & sKey
    '            Debug.Print String(20, "-")
    '            If IsObject(.item(sKey)) Then
    '                dp (.item(sKey))
    '            Else
    '                Debug.Print "Key " & sKey & " : "
    '                dp .item(sKey)
    '            End If
    '        Next i
    '    End With
End Sub

Private Sub DPH(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    '@BlogPosted
    'https://gist.github.com/YujiFukami/15c922d41ff06c9b12ad450a14131080#file-
    '@INCLUDE PROCEDURE DebugPrintHairetu
    '@AssignedModule vbArcImports
    Call DebugPrintHairetu(Hairetu, HyoujiMaxNagasa, HairetuName)
End Sub

Public Function ArrayDimensions(ByVal vArray As Variant) As Long
    '@BlogPosted
    '@AssignedModule vbArcImports
    Dim dimnum      As Long
    Dim ErrorCheck  As Long
    On Error GoTo FinalDimension
    For dimnum = 1 To 60000
        ErrorCheck = LBound(vArray, dimnum)
    Next
FinalDimension:
    ArrayDimensions = dimnum - 1
End Function

Private Sub DebugPrintHairetu(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    '@BlogPosted
    'https://gist.github.com/YujiFukami/15c922d41ff06c9b12ad450a14131080#file-
    '@INCLUDE PROCEDURE ShortenToByteCharacters
    '@AssignedModule vbArcImports


    Dim i&, j&, k&, M&, n&
    Dim TateMin&, TateMax&, YokoMin&, YokoMax&
    Dim WithTableHairetu
    Dim NagasaList, MaxNagasaList
    Dim NagasaOnajiList
    Dim OutputList
    Const SikiriMoji$ = "|"
    Dim Jigen2%
    On Error Resume Next
    Jigen2 = UBound(Hairetu, 2)
    '    On Error GoTo 0

    If Jigen2 = 0 Then
        Hairetu = Application.Transpose(Hairetu)
    End If
    TateMin = LBound(Hairetu, 1)
    TateMax = UBound(Hairetu, 1)
    YokoMin = LBound(Hairetu, 2)
    YokoMax = UBound(Hairetu, 2)
    ReDim WithTableHairetu(1 To TateMax - TateMin + 1 + 1, 1 To YokoMax - YokoMin + 1 + 1)
    For i = 1 To TateMax - TateMin + 1
        WithTableHairetu(i + 1, 1) = TateMin + i - 1
        For j = 1 To YokoMax - YokoMin + 1
            WithTableHairetu(1, j + 1) = YokoMin + j - 1
            WithTableHairetu(i + 1, j + 1) = Hairetu(i - 1 + TateMin, j - 1 + YokoMin)
        Next j
    Next i
    n = UBound(WithTableHairetu, 1)
    M = UBound(WithTableHairetu, 2)
    ReDim NagasaList(1 To n, 1 To M)
    ReDim MaxNagasaList(1 To M)
    Dim tmpStr$
    For j = 1 To M
        For i = 1 To n
            If j > 1 And HyoujiMaxNagasa <> 0 Then
                tmpStr = WithTableHairetu(i, j)
                WithTableHairetu(i, j) = ShortenToByteCharacters(tmpStr, HyoujiMaxNagasa)
            End If
            NagasaList(i, j) = LenB(StrConv(WithTableHairetu(i, j), vbFromUnicode))
            MaxNagasaList(j) = WorksheetFunction.Max(MaxNagasaList(j), NagasaList(i, j))
        Next i
    Next j
    ReDim NagasaOnajiList(1 To n, 1 To M)
    Dim TmpMaxNagasa&
    For j = 1 To M
        TmpMaxNagasa = MaxNagasaList(j)
        For i = 1 To n
            NagasaOnajiList(i, j) = WithTableHairetu(i, j) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(i, j))
        Next i
    Next j
    ReDim OutputList(1 To n)
    For i = 1 To n
        For j = 1 To M
            If j = 1 Then
                OutputList(i) = NagasaOnajiList(i, j)
            Else
                OutputList(i) = OutputList(i) & SikiriMoji & NagasaOnajiList(i, j)
            End If
        Next j
    Next i
    Debug.Print HairetuName
    For i = 1 To n
        Debug.Print OutputList(i)
    Next i
End Sub


Public Function ShortenToByteCharacters(Mojiretu$, ByteNum%)
    '@BlogPosted
    'https://gist.github.com/YujiFukami/15c922d41ff06c9b12ad450a14131080#file-
    '@INCLUDE PROCEDURE CalculateByteCharacters
    '@INCLUDE PROCEDURE TextDecomposition
    '@AssignedModule vbArcImports
    Dim OriginByte%
    Dim Output
    OriginByte = LenB(StrConv(Mojiretu, vbFromUnicode))
    If OriginByte <= ByteNum Then
        Output = Mojiretu
    Else
        Dim RuikeiByteList, BunkaiMojiretu
        RuikeiByteList = CalculateByteCharacters(Mojiretu)
        BunkaiMojiretu = TextDecomposition(Mojiretu)
        Dim AddMoji$
        AddMoji = "."
        Dim i&, n&
        n = Len(Mojiretu)
        For i = 1 To n
            If RuikeiByteList(i) < ByteNum Then
                Output = Output & BunkaiMojiretu(i)
            ElseIf RuikeiByteList(i) = ByteNum Then
                If LenB(StrConv(BunkaiMojiretu(i), vbFromUnicode)) = 1 Then
                    Output = Output & AddMoji
                Else
                    Output = Output & AddMoji & AddMoji
                End If
                Exit For
            ElseIf RuikeiByteList(i) > ByteNum Then
                Output = Output & AddMoji
                Exit For
            End If
        Next i
    End If
    ShortenToByteCharacters = Output
End Function

Private Function CalculateByteCharacters(Mojiretu$)
    '@BlogPosted
    'https://gist.github.com/YujiFukami/15c922d41ff06c9b12ad450a14131080#file-
    '@AssignedModule vbArcImports
    Dim MojiKosu%
    MojiKosu = Len(Mojiretu)
    Dim Output
    ReDim Output(1 To MojiKosu)
    Dim i&
    Dim TmpMoji$
    For i = 1 To MojiKosu
        TmpMoji = Mid(Mojiretu, i, 1)
        If i = 1 Then
            Output(i) = LenB(StrConv(TmpMoji, vbFromUnicode))
        Else
            Output(i) = LenB(StrConv(TmpMoji, vbFromUnicode)) + Output(i - 1)
        End If
    Next i
    CalculateByteCharacters = Output
End Function

Private Function TextDecomposition(Mojiretu$)
    '@BlogPosted
    'https://gist.github.com/YujiFukami/15c922d41ff06c9b12ad450a14131080#file-
    '@AssignedModule vbArcImports
    Dim i&, n&
    Dim Output
    n = Len(Mojiretu)
    ReDim Output(1 To n)
    For i = 1 To n
        Output(i) = Mid(Mojiretu, i, 1)
    Next i
    TextDecomposition = Output
End Function

'* Modified   : Date and Time       Author              Description
'* Updated    : 18-08-2023 12:20    Alex                (vbArcImports.bas > FindCode)

Sub FindCode(S As String)
    '@LastModified 2308181220
    '@INCLUDE USERFORM uCodeFinder
    '@AssignedModule vbArcImports
    Load uCodeFinder
    uCodeFinder.TextBox1.TEXT = S
    uCodeFinder.DoSearch
End Sub

