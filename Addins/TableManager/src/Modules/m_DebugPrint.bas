Attribute VB_Name = "m_DebugPrint"
Option Explicit

Public Sub dp(var As Variant)
'@AssignedModule m_DebugPrint
'@INCLUDE PROCEDURE PrintXML
'@INCLUDE PROCEDURE printArray
'@INCLUDE PROCEDURE printCollection
'@INCLUDE PROCEDURE printDictionary
    Dim element     As Variant
    Dim i As Long
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
    Case Is = "Date"
        Debug.Print var
    Case Is = "IXMLDOMElement"
        PrintXML var
    Case Else
    End Select
End Sub

Sub PrintXML(var)
'@AssignedModule m_DebugPrint
    Debug.Print var.XML
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
'        If Not Left(child.nodename, 1) = "#" Then Debug.Print child.nodename & ":" & child.Text
'        If child.ChildNodes.Length > 0 Then PrintXML child.ChildNodes
'    Next
'End Sub

Private Sub printArray(var As Variant)
'@AssignedModule m_DebugPrint
'@INCLUDE PROCEDURE dp
'@INCLUDE PROCEDURE PrintXML
'@INCLUDE PROCEDURE DPH
'@INCLUDE PROCEDURE ArrayDimensions
    Dim element
    If ArrayDimensions(var) = 1 Then
'        Debug.Print Join(var, vbNewLine)
        For Each element In var
            dp element
        Next
    ElseIf ArrayDimensions(var) > 1 Then
        DPH var
    End If
End Sub

Private Sub printCollection(var As Variant)
'@AssignedModule m_DebugPrint
'@INCLUDE PROCEDURE dp
    Dim elem        As Variant
    For Each elem In var
        dp elem
    Next elem
End Sub

Private Sub printDictionary(var As Variant)
'@TODO detect error cause I met when printing a dic from JSON related modules
'@AssignedModule m_DebugPrint
'@INCLUDE PROCEDURE dp
    Dim i As Long: Dim iCount As Long
    Dim arrKeys
    Dim sKey        As String
    Dim varItem
    
    Dim key As Variant
    For Each key In var.Keys
        dp var(key)
        
    Next key
    
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
'https://gist.github.com/YujiFukami/15c922d41ff06c9b12ad450a14131080#file-
'@AssignedModule m_DebugPrint
'@INCLUDE PROCEDURE DebugPrintHairetu
    Call DebugPrintHairetu(Hairetu, HyoujiMaxNagasa, HairetuName)
End Sub

Public Function ArrayDimensions(ByVal vArray As Variant) As Long
'@AssignedModule m_DebugPrint
    Dim dimnum      As Long
    Dim ErrorCheck As Long
    On Error GoTo FinalDimension
    For dimnum = 1 To 60000
        ErrorCheck = LBound(vArray, dimnum)
    Next
FinalDimension:
    ArrayDimensions = dimnum - 1
End Function

Private Sub DebugPrintHairetu(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
'https://gist.github.com/YujiFukami/15c922d41ff06c9b12ad450a14131080#file-
'@AssignedModule m_DebugPrint
'@INCLUDE PROCEDURE Transpose2DArray
'@INCLUDE PROCEDURE ReptText
'@INCLUDE PROCEDURE MaxValue
'@INCLUDE PROCEDURE ShortenToByteCharacters
    Dim i&, j&, k&, M&, N&
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
        Hairetu = Transpose2DArray(Hairetu)
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
    N = UBound(WithTableHairetu, 1)
    M = UBound(WithTableHairetu, 2)
    ReDim NagasaList(1 To N, 1 To M)
    ReDim MaxNagasaList(1 To M)
    Dim tmpStr$
    For j = 1 To M
        For i = 1 To N
            If j > 1 And HyoujiMaxNagasa <> 0 Then
                tmpStr = WithTableHairetu(i, j)
                WithTableHairetu(i, j) = ShortenToByteCharacters(tmpStr, HyoujiMaxNagasa)
            End If
            NagasaList(i, j) = LenB(StrConv(WithTableHairetu(i, j), vbFromUnicode))
            MaxNagasaList(j) = MaxValue(MaxNagasaList(j), NagasaList(i, j))
        Next i
    Next j
    ReDim NagasaOnajiList(1 To N, 1 To M)
    Dim TmpMaxNagasa&
    For j = 1 To M
        TmpMaxNagasa = MaxNagasaList(j)
        For i = 1 To N
            NagasaOnajiList(i, j) = WithTableHairetu(i, j) & ReptText(" ", IIf(WithTableHairetu(i, j) = "", 1, TmpMaxNagasa - NagasaList(i, j)))
        Next i
    Next j
    ReDim OutputList(1 To N)
    For i = 1 To N
        For j = 1 To M
            If j = 1 Then
                OutputList(i) = NagasaOnajiList(i, j)
            Else
                OutputList(i) = OutputList(i) & SikiriMoji & NagasaOnajiList(i, j)
            End If
        Next j
    Next i
    Debug.Print HairetuName
    For i = 1 To N
        Debug.Print OutputList(i)
    Next i
End Sub

Function ReptText(ByVal Text As String, ByVal Count As Integer) As String
'@AssignedModule m_DebugPrint
    Dim result As String
    result = ""
    
    If Count > 0 Then
        Dim i As Integer
        For i = 1 To Count
            result = result & Text
        Next i
    End If
    
    ReptText = result
End Function

Function MaxValue(ParamArray values() As Variant) As Variant
'@AssignedModule m_DebugPrint
'@INCLUDE PROCEDURE GetErrorValue
    If Not IsArray(values) Then
        MaxValue = GetErrorValue
        Exit Function
    End If
    
    Dim i As Long
    Dim Max As Double
    
    If UBound(values) >= LBound(values) Then
        Max = values(LBound(values))
        For i = LBound(values) + 1 To UBound(values)
            If IsNumeric(values(i)) Then
                If values(i) > Max Then
                    Max = values(i)
                End If
            End If
        Next i
    End If
    
    MaxValue = IIf(Max = 0, GetErrorValue, Max)
End Function

Function GetErrorValue() As Variant
'@AssignedModule m_DebugPrint
    GetErrorValue = CVErr(2042) ' 2042 represents the xlErrValue error number in Excel
End Function

Public Function ShortenToByteCharacters(Mojiretu$, ByteNum%)
'https://gist.github.com/YujiFukami/15c922d41ff06c9b12ad450a14131080#file-
'@AssignedModule m_DebugPrint
'@INCLUDE PROCEDURE CalculateByteCharacters
'@INCLUDE PROCEDURE TextDecomposition
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
        Dim i&, N&
        N = Len(Mojiretu)
        For i = 1 To N
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
'https://gist.github.com/YujiFukami/15c922d41ff06c9b12ad450a14131080#file-
'@AssignedModule m_DebugPrint
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
'https://gist.github.com/YujiFukami/15c922d41ff06c9b12ad450a14131080#file-
'@AssignedModule m_DebugPrint
    Dim i&, N&
    Dim Output
    N = Len(Mojiretu)
    ReDim Output(1 To N)
    For i = 1 To N
        Output(i) = Mid(Mojiretu, i, 1)
    Next i
    TextDecomposition = Output
End Function

Function DpHeader( _
                 str As Variant, _
                 Optional lvl As Integer = 1, _
                 Optional Character As String = "'", _
                 Optional Top As Boolean, _
                 Optional Bottom As Boolean) As String
'@AssignedModule m_DebugPrint
'@INCLUDE PROCEDURE LargestLength
    If lvl < 1 Then lvl = 1
    If Character = "" Then Character = "'"
    Dim indentation As Integer
    indentation = (lvl * 4) - 4 + 1
    Dim quote As String: quote = "'"
    Dim s As String
    Dim element As Variant
    If Top = True Then s = vbNewLine & quote & String(indentation + LargestLength(str), Character) & vbNewLine
    If TypeName(str) <> "String" Then
        For Each element In str
            s = s & quote & Character & Space(indentation) & element & vbNewLine
        Next
    Else
        s = s & quote & String(indentation, Character) & str
    End If
    If Bottom = True Then s = s & quote & String(indentation + LargestLength(str), Character)
    DpHeader = s
End Function

Function LargestLength(MyObj) As Long
'@AssignedModule m_DebugPrint
'@INCLUDE PROCEDURE dp
    LargestLength = 0
    Dim element As Variant
    Select Case TypeName(MyObj)
    Case Is = "String"
        LargestLength = Len(MyObj)
    Case "Collection"
        For Each element In MyObj
            If TypeName(element) = "String" Then
                If Len(element) > LargestLength Then LargestLength = Len(element)
            Else
                If element.Width > LargestLength Then LargestLength = element.Width
            End If
        Next element
    Case "Variant", "Variant()", "String()"
       For Each element In MyObj
            If TypeName(element) = "String" Then
                If Len(element) > LargestLength Then LargestLength = Len(element)
'                dp element & vbTab & Len(element)
            End If
        Next
    Case Else
    End Select
End Function




