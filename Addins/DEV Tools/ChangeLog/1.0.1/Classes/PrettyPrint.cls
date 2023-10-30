VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrettyPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
'
' Pretty Print class
' |* Brian J. Satola
' |* https://chejunkie.com
'
'==========================================================================================================================
' Public Subs, Functions and Properties - oh my!
'==========================================================================================================================
' For additional details see About.
'==========================================================================================================================
'     |> Get | --- About                    :: String that describes the current class.
'     |> Get | Let BlankLineBeforeHeaders   :: Switch to insert a blank line before a header is printed.
'     |> Get | Let BlankLineBeforeTitles    :: Switch to insert a blank line before a title is printed.
'     |> Get | --- Name                     :: Returns the name of the class as a string.
'     |> --------- PrintHeader1             :: Print heading style 1 to the Immediate Window.
'     |> --------- PrintHeader2             :: Print heading style 2 to the Immediate Window.
'     |> --------- PrintLineElements        :: Prints a formatted line form single or multiple arguments.
'     |> --------- PrintTable               :: Prints a formatted table to the Immediate Window.
'     |> --------- PrintTitle               :: Prints a formatted title to the Immediate Window.
'     |> Get | --- Version                  :: Returns version string for the class [e.g. #.# (year)].
'
'==========================================================================================================================
' Application Constants
'==========================================================================================================================
Const C_MAX_LEN As Integer = 110
'
'==========================================================================================================================
' Private Variables
'==========================================================================================================================
Dim m_BlankLineBeforeTitles As Boolean
Dim m_BlankLineBeforeHeaders As Boolean
'''
''' Class constructor.
'''
Private Sub Class_Initialize()
    Debug.Print "|> Initializing:= " & Me.Name
    m_BlankLineBeforeTitles = True
    m_BlankLineBeforeHeaders = True
End Sub
'''
''' Returns a string that describes the current class/object.
'''
Public Property Get About() As String
    About = "ChE Junkie VBA Pretty Print class module, " & Me.version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details see https://chejunkie.com/knowledge-base/pretty-print-class-for-vba/"
End Property
'''
''' Switch addds a line before header prints when True.
'''
Public Property Get BlankLineBeforeHeaders() As Boolean
    BlankLineBeforeHeaders = m_BlankLineBeforeHeaders
End Property
'''
''' Switch addds a line before header prints when True.
'''
Public Property Let BlankLineBeforeHeaders(switchValue As Boolean)
    m_BlankLineBeforeHeaders = switchValue
End Property
'''
''' Switch addds a line before title prints when True.
'''
Public Property Get BlankLineBeforeTitles() As Boolean
    BlankLineBeforeTitles = m_BlankLineBeforeTitles
End Property
'''
''' Switch addds a line before title prints when True.
'''
Public Property Let BlankLineBeforeTitles(switchValue As Boolean)
    m_BlankLineBeforeTitles = switchValue
End Property
'''
''' Returns the class name of the object.
'''
Public Property Get Name() As String
    Name = VBA.TypeName(Me) & ".cls"
End Property
'''
''' Print heading style 1.
'''
Public Sub PrintHeader1(ParamArray Args() As Variant)
    If m_BlankLineBeforeHeaders Then
        Debug.Print ""
    End If
    Debug.Print Strings.String(C_MAX_LEN, "=")
    PrintLineElements Args
    Debug.Print Strings.String(C_MAX_LEN, "=")
End Sub
'''
''' Print heading style 2.
'''
Public Sub PrintHeader2(ParamArray Args() As Variant)
    If m_BlankLineBeforeHeaders Then
        Debug.Print ""
    End If
    Debug.Print Strings.String(C_MAX_LEN, "-")
    PrintLineElements Args
    Debug.Print Strings.String(C_MAX_LEN, "-")
End Sub
'''
''' Print bullet
'''
Sub PrintBullet(lvl As Long, stringToLog As Variant)
    If lvl < 1 Then Exit Sub
    stringToLog = VBA.Now & ": |" & String(2 * (lvl - 1), "-") & "> " & stringToLog
    Debug.Print stringToLog
End Sub
'''
''' Print line through
'''
Sub PrintLineThrough(ByVal stringToLog As Variant)
    Dim strg As String
    Dim nlen(2) As Integer
    nlen(1) = C_MAX_LEN - VBA.Len(VBA.Now & ": ")
    If (VBA.Len(stringToLog) <= nlen(1)) Then
        strg = VBA.Replace(stringToLog, " ", "-")
        strg = "---" & strg
        nlen(2) = nlen(1) - VBA.Len(stringToLog) - 3
        Dim ilen As Long
        For ilen = 1 To nlen(2)
            strg = strg & "-"
        Next ilen
        stringToLog = VBA.Now & ": " & strg
    End If
    Debug.Print CStr(stringToLog)
End Sub
'''
''' Print a blank line
'''
Public Sub PrintBlankLine()
    Debug.Print ""
End Sub
'''
''' Print a line : solid (default) or dashed.
'''
Public Sub PrintDividingLine(Optional includeBlankLineAfter As Boolean = True, Optional useDashLine As Boolean = False)
    If includeBlankLineAfter Then
        If Not useDashLine Then
            Debug.Print String("_", C_MAX_LEN) & vbNewLine
        Else
            Debug.Print String("-", C_MAX_LEN) & vbNewLine
        End If
    ElseIf Not useDashLine Then
        Debug.Print String("_", C_MAX_LEN)
    Else
         Debug.Print String("-", C_MAX_LEN)
    End If
End Sub

''' Prints a formatted line depending on the arguments structure:
''' ** Each string will have a length [number] when it exists.
''' |> number, string, string, string... all strings printed to same [number] length.
''' |> string, number, string, number... string-number pairs, where number defines string length.
''' |> string - equivalent to Debug.Print "string".
'''
'* Modified   : Date and Time       Author              Description
'* Updated    : 12-09-2023 14:03    Alex                (PrettyPrint.cls > PrintLineElements) fixed wrong length

Public Sub PrintLineElements(ParamArray Args() As Variant)
'@LastModified 2309121403
    Dim i           As Integer
    Dim line        As String
    Dim cbuff       As String
    On Error GoTo ErrHandler
    cbuff = String(100, " ")
    If (LBound(Args) = UBound(Args)) Then
        If (UBound(Args(0)) > 0) Then    ' array of elements
            If (IsNumeric(Args(0)(0))) Then
                ' First number sets the length of all strings
                For i = 1 To UBound(Args(0))
                    line = line & Left(CStr(Args(0)(i)) & cbuff, CInt(Args(0)(0))) & IIf(i < UBound(Args(0)), Space(1), "")
                Next i
            ElseIf (((UBound(Args(0)) - LBound(Args(0)) + 1) Mod 2 = 0) And IsNumeric(Args(0)(1))) Then
                ' String-length pairs
                For i = 0 To UBound(Args(0))
                    line = line & Left(CStr(Args(0)(i)) & cbuff, Args(0)(i + 1)) & IIf(i < UBound(Args(0)) - 1, Space(1), "")
                    i = i + 1
                Next i
            End If
        Else    ' Debug.Print "string" equivalent
            line = Args(0)(0)
        End If
    ElseIf IsNumeric(Args(0)) Then
        ' First number sets the length of all strings
        For i = 1 To UBound(Args)
            line = line & Left(CStr(Args(i)) & cbuff, CInt(Args(0))) & IIf(i < UBound(Args), Space(1), "")
        Next i
    ElseIf (((UBound(Args) - LBound(Args) + 1) Mod 2 = 0) And IsNumeric(Args(1))) Then
        ' String-length pairs
        For i = 0 To UBound(Args)
            line = line & Left(CStr(Args(i)) & cbuff, Args(i + 1)) & IIf(i < UBound(Args), Space(1), "")
            i = i + 1
        Next i
    Else
        Exit Sub
    End If
    Debug.Print line
    Exit Sub
ErrHandler:
    Debug.Print VBA.vbCrLf & "** PROBLEM OCCURED ** PRETTY PRINT LINE ELEMENTS**"
    Debug.Print Strings.String(C_MAX_LEN, "_")
End Sub
'''
''' Prints formatted table to the immediate widnow.
''' If [hasHeader] then first row entries are taken as column titles.
'''
Public Sub PrintTable(Table As Variant, Optional hasHeader As Boolean = True, Optional padding As Integer = 1)
    On Error GoTo ErrHandler
    Debug.Print ArrayToTable(Table, hasHeader, padding)
    Exit Sub
ErrHandler:
    Debug.Print VBA.vbCrLf & "** PROBLEM OCCURED ** PRETTY PRINT TABLE **"
    Debug.Print Strings.String(C_MAX_LEN, "_")
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 25-10-2023 14:49    Alex                (PrettyPrint.cls > ArrayToTable)

Function ArrayToTable(Table As Variant, Optional hasHeader As Boolean = True, Optional padding As Integer = 1)
'@LastModified 2310251449
    '@SOURCE https://chejunkie.com/knowledge-base/format-debug-print-messages-in-vba/
    Dim nrow        As Integer
    Dim ncol        As Integer
    Dim irow        As Integer
    Dim jcol        As Integer
    Dim vals()      As Variant
    Dim rowStart    As Integer
    Dim widths()    As Integer
    Dim cbuff       As String
    Dim line        As String
    Dim edge        As String
    Dim Output      As String
    
    '!! make sure there are no tabs or linefeeds
    ArrayReplace Table, vbTab, Space(4)
    ArrayReplace Table, vbNewLine, Space(1)
    ArrayReplace Table, vbLf, Space(1)

    nrow = UBound(Table, 1)
    ncol = UBound(Table, 2)
    ReDim vals(1 To ncol * 2)
    ReDim widths(1 To ncol)
    cbuff = String(200, " ")
    ' Define column widths
    For jcol = LBound(Table, 2) To ncol
        For irow = LBound(Table, 1) To nrow
            If (Len(Table(irow, jcol)) > widths(jcol)) Then
                widths(jcol) = Len(Table(irow, jcol))
            End If
        Next irow
    Next jcol
    ' Print top edge and header
    edge = ".-" & Left(String(Len(cbuff), "-"), widths(1) + padding) & "-"
    For jcol = (LBound(Table, 2) + 1) To ncol - 1
        edge = edge & "-" & Left(String(Len(cbuff), "-"), widths(jcol) + padding) & "-"
    Next jcol
    edge = edge & "-" & Left(String(Len(cbuff), "-"), widths(jcol) + padding) & "."
    Output = Output & IIf(Output <> "", vbNewLine, "") & edge
    If hasHeader Then
        line = "| " & Left(Table(LBound(Table, 1), 1) & cbuff, widths(1) + padding) & "|"
        For jcol = (LBound(Table, 2) + 1) To ncol
            line = line & " " & Left(Table(LBound(Table, 1), jcol) & cbuff, widths(jcol) + padding) & "|"
        Next jcol
        Output = Output & IIf(Output <> "", vbNewLine, "") & line
        Output = Output & IIf(Output <> "", vbNewLine, "") & edge
        rowStart = LBound(Table, 1) + 1
    Else
        rowStart = LBound(Table, 1)
    End If
    ' Print data
    For irow = rowStart To nrow
        line = "| " & Left(Table(irow, 1) & cbuff, widths(1) + padding) & "|"
        For jcol = (LBound(Table, 2) + 1) To ncol
            line = line & " " & Left(Table(irow, jcol) & cbuff, widths(jcol) + padding) & "|"
        Next jcol
        Output = Output & IIf(Output <> "", vbNewLine, "") & line
    Next irow
    Output = Output & IIf(Output <> "", vbNewLine, "") & edge
    ArrayToTable = Output
End Function

Sub PrintTitle(ParamArray titleToPrint() As Variant)
    'modified by Anastasiou Alex to print multiple lines
    'eg prettyprint.printtitle array(1, "this","that there")
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '////                  ////////////////////////////////////////////////////////////////////////////////////////
    '////    1             ////////////////////////////////////////////////////////////////////////////////////////
    '////    this          ////////////////////////////////////////////////////////////////////////////////////////
    '////    that there    ////////////////////////////////////////////////////////////////////////////////////////
    '////                  ////////////////////////////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Dim strg        As String
    Dim pads(2)     As String
    Dim nlen(2)     As Integer
    Dim ilen        As Long
    Dim i           As Long
    '// Preformatting
    '   [1] Padding for "stringToLog line.
    '   [2] Padding for spacer line.
    'mod
    Dim ll          As Long
    Dim element
    
        If IsArray(titleToPrint) Then
            For Each element In titleToPrint(0)
                If Len(CStr(element)) > ll Then ll = Len(CStr(element))
            Next
        Else
            For Each element In titleToPrint
                If Len(CStr(element)) > ll Then ll = Len(CStr(element))
            Next
        End If
    
    nlen(1) = C_MAX_LEN - (ll + 12)
    '/mod
    For ilen = 1 To nlen(1)
        pads(1) = pads(1) & "/"
    Next ilen
    pads(1) = "    " & pads(1)
    nlen(2) = ll                                     'mod
    For ilen = 1 To nlen(2)
        pads(2) = pads(2) & " "
    Next ilen
    '// Title definition.
    If m_BlankLineBeforeTitles Then
        strg = VBA.vbCrLf & VBA.vbCrLf
    End If
    strg = strg & VBA.String(C_MAX_LEN, "/") & VBA.vbCrLf
    strg = strg & "////    " & pads(2) & pads(1) & VBA.vbCrLf
    'mod
    If IsArray(titleToPrint) Then
        For i = LBound(titleToPrint(0)) To UBound(titleToPrint(0))
            strg = strg & "////    " & titleToPrint(0)(i) & Space(ll - Len(titleToPrint(0)(i))) & pads(1) & VBA.vbCrLf
        Next
    
    Else
        For i = LBound(titleToPrint) To UBound(titleToPrint)
            strg = strg & "////    " & titleToPrint(i) & Space(ll - Len(titleToPrint(i))) & pads(1) & VBA.vbCrLf
        Next
    End If
    '/mod
    strg = strg & "////    " & pads(2) & pads(1) & VBA.vbCrLf
    strg = strg & VBA.String(C_MAX_LEN, "/") & VBA.vbCrLf
    Debug.Print strg
End Sub
'''
''' Version string of the current class.
'''
Public Property Get version() As String
    '==========================================================================================================================
    ' Class changelog here.
    '==========================================================================================================================
    version = "Version 1.0 (2020)"  ' Initial (official) release.
End Property

