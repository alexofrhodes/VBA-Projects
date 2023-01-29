Attribute VB_Name = "vbArcImports"

Rem --------------------

Rem Debug Print
Sub dp(Var As Variant)
    '#INCLUDE DPH
    '#INCLUDE ArrayDimensions
    '#IMPORTS DPH
    '#IMPORTS ArrayDimensions
    Dim element As Variant
    Select Case TypeName(Var)
        Case Is = "String", "Long", "Integer", "Boolean"
            Debug.Print Var
            Rem todo How to handle multidimensional array?
        Case Is = "Variant()", "String()", "Long()", "Integer()"
            If ArrayDimensions(Var) = 1 Then
                Dim i As Long
                For i = LBound(Var) To UBound(Var)
                    Debug.Print Var(i)
                Next i
            ElseIf ArrayDimensions(Var) > 1 Then
                DPH Var
            End If
        Case Is = "Collection"
            For Each element In Var
                dp element
            Next element
        Case Else
    End Select
End Sub

Function TextDecomposition(Mojiretu$)
    Dim i&, n&
    Dim output
    n = Len(Mojiretu)
    ReDim output(1 To n)
    For i = 1 To n
        output(i) = Mid(Mojiretu, i, 1)
    Next i
    TextDecomposition = output
End Function

Function CalculateByteCharacters(Mojiretu$)
    Dim MojiKosu%
    MojiKosu = Len(Mojiretu)
    Dim output
    ReDim output(1 To MojiKosu)
    Dim i&
    Dim TmpMoji$
    For i = 1 To MojiKosu
        TmpMoji = Mid(Mojiretu, i, 1)
        If i = 1 Then
            output(i) = LenB(StrConv(TmpMoji, vbFromUnicode))
        Else
            output(i) = LenB(StrConv(TmpMoji, vbFromUnicode)) + output(i - 1)
        End If
    Next i
    CalculateByteCharacters = output
End Function

Function ShortenToByteCharacters(Mojiretu$, ByteNum%)
    '#INCLUDE CalculateByteCharacters
    '#INCLUDE TextDecomposition
    '#IMPORTS CalculateByteCharacters
    '#IMPORTS TextDecomposition
    Dim OriginByte%
    Dim output
    OriginByte = LenB(StrConv(Mojiretu, vbFromUnicode))
    If OriginByte <= ByteNum Then
        output = Mojiretu
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
                output = output & BunkaiMojiretu(i)
            ElseIf RuikeiByteList(i) = ByteNum Then
                If LenB(StrConv(BunkaiMojiretu(i), vbFromUnicode)) = 1 Then
                    output = output & AddMoji
                Else
                    output = output & AddMoji & AddMoji
                End If
                Exit For
            ElseIf RuikeiByteList(i) > ByteNum Then
                output = output & AddMoji
                Exit For
            End If
        Next i
    End If
    ShortenToByteCharacters = output
End Function

Sub DebugPrintHairetu(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    '#INCLUDE ShortenToByteCharacters
    '#IMPORTS ShortenToByteCharacters
    Dim i&, j&, k&, m&, n&
    Dim TateMin&, TateMax&, YokoMin&, YokoMax&
    Dim WithTableHairetu
    Dim NagasaList, MaxNagasaList
    Dim NagasaOnajiList
    Dim OutputList
    Const SikiriMoji$ = "|"
    Dim Jigen2%
    On Error Resume Next
    Jigen2 = UBound(Hairetu, 2)
    On Error GoTo 0
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
                m = UBound(WithTableHairetu, 2)
                ReDim NagasaList(1 To n, 1 To m)
                ReDim MaxNagasaList(1 To m)
                Dim TmpStr$
                For j = 1 To m
                    For i = 1 To n
                        If j > 1 And HyoujiMaxNagasa <> 0 Then
                            TmpStr = WithTableHairetu(i, j)
                            WithTableHairetu(i, j) = ShortenToByteCharacters(TmpStr, HyoujiMaxNagasa)
                            End If
                            NagasaList(i, j) = LenB(StrConv(WithTableHairetu(i, j), vbFromUnicode))
                            MaxNagasaList(j) = WorksheetFunction.Max(MaxNagasaList(j), NagasaList(i, j))
                        Next i
                    Next j
                    ReDim NagasaOnajiList(1 To n, 1 To m)
                    Dim TmpMaxNagasa&
                    For j = 1 To m
                        TmpMaxNagasa = MaxNagasaList(j)
                        For i = 1 To n
                            NagasaOnajiList(i, j) = WithTableHairetu(i, j) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(i, j))
                        Next i
                    Next j
                    ReDim OutputList(1 To n)
                    For i = 1 To n
                        For j = 1 To m
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

Function ArrayDimensions(ByVal vArray As Variant) As Long
    Dim dimnum As Long
    On Error GoTo FinalDimension
    For dimnum = 1 To 60000
        ErrorCheck = LBound(vArray, dimnum)
    Next
FinalDimension:
    ArrayDimensions = dimnum - 1
End Function

Rem ------------------------


Sub DPH(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    '#INCLUDE DebugPrintHairetu
    '#IMPORTS DebugPrintHairetu
    Call DebugPrintHairetu(Hairetu, HyoujiMaxNagasa, HairetuName)
End Sub

