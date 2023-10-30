Attribute VB_Name = "vbArcImports"

Sub SaveUserformOptions(FORM As Object, _
                Optional includeCheckbox As Boolean = True, _
                Optional includeOptionButton As Boolean = True, _
                Optional includeTextBox As Boolean = True, _
                Optional includeListbox As Boolean = True)
'#INCLUDE CreateOrSetSheet
'#INCLUDE ListboxSelectedIndexes
'#INCLUDE CollectionToArray

    Dim ws As Worksheet
    Set ws = CreateOrSetSheet(FORM.Name & "_Settings", ThisWorkbook) ' change to activeworkbook for testing
    ws.Cells.clear
    Dim coll As New Collection
    Dim cell As Range
    Set cell = ws.Cells(1, 1)
    Dim c As MSForms.Control
    For Each c In FORM.Controls
        If TypeName(c) Like "CheckBox" Then
            If Not includeCheckbox Then GoTo SKIP
        ElseIf TypeName(c) Like "OptionButton" Then
            If Not includeOptionButton Then GoTo SKIP
        ElseIf TypeName(c) Like "TextBox" Then
            If Not includeTextBox Then GoTo SKIP
        ElseIf TypeName(c) = "ListBox" Then
            If Not includeListbox Then GoTo SKIP
        Else
            GoTo SKIP
        End If
        cell = c.Name
        Select Case TypeName(c)
        Case "TextBox", "CheckBox", "OptionButton"
            cell.Offset(0, 1) = c.Value
        Case "ListBox"
            Set coll = ListboxSelectedIndexes(c)
            If coll.Count > 0 Then
                cell.Offset(0, 1) = Join(CollectionToArray(coll), ",")
            Else
                cell.Offset(0, 1) = -1
            End If
        End Select
        Set cell = cell.Offset(1, 0)
SKIP:
    Next
End Sub

Sub LoadUserformOptions(FORM As Object, Optional ExcludeThese As Variant)
'#INCLUDE CreateOrSetSheet
'#INCLUDE IsInArray
'#INCLUDE SelectListboxItems
'#IMPORTS CreateOrSetSheet
'#IMPORTS IsInArray
'#IMPORTS SelectListboxItems
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet(FORM.Name & "_Settings", ThisWorkbook) ' change to activeworkbook for testing
    If ws.Range("A1") = "" Then Exit Sub
    Dim cell As Range
    Set cell = ws.Cells(1, 1)
    Dim c As MSForms.Control
    Dim v
    On Error Resume Next
    Do While cell <> ""
        Set c = FORM.Controls(cell.Text)
        If Not TypeName(c) = "Nothing " Then
            If Not IsInArray(cell, ExcludeThese) Then
                Select Case TypeName(c)
                Case "TextBox", "CheckBox", "OptionButton"
                    c.Value = cell.Offset(0, 1)
                Case "ListBox"
                    If InStr(1, cell.Offset(0, 1), ",") > 0 Then
                        SelectListboxItems c, Split(cell.Offset(0, 1), ","), True
                    Else
                        c.Selected(CInt(cell.Offset(0, 1))) = True
                    End If
                End Select
            End If
        End If
        Set cell = cell.Offset(1, 0)
    Loop
End Sub
Function CreateOrSetSheet(SheetName As String, TargetWorkbook As Workbook) As Worksheet
    '#INCLUDE WorksheetExists
    Dim NewSheet As Worksheet
    If WorksheetExists(SheetName, TargetWorkbook) = True Then
        Set CreateOrSetSheet = TargetWorkbook.Sheets(SheetName)
    Else
        Set CreateOrSetSheet = TargetWorkbook.Sheets.Add
        CreateOrSetSheet.Name = SheetName
    End If
End Function
Function ListboxSelectedIndexes(LBox As MSForms.ListBox) As Collection
    Dim i As Long
    Dim SelectedIndexes As Collection
    Set SelectedIndexes = New Collection
    If LBox.ListCount > 0 Then
        For i = 0 To LBox.ListCount - 1
            If LBox.Selected(i) Then SelectedIndexes.Add i
        Next i
    End If
    Set ListboxSelectedIndexes = SelectedIndexes
End Function

Function SelectListboxItems(LBox As MSForms.ListBox, FindMe As Variant, Optional ByIndex As Boolean)
    Dim i As Long
    Select Case TypeName(FindMe)
    Case Is = "String", "Long", "Integer"
        For i = 0 To LBox.ListCount - 1
            If LBox.list(i) = CStr(FindMe) Then
                LBox.Selected(i) = True
                DoEvents
                If LBox.MultiSelect = fmMultiSelectSingle Then Exit Function
            End If
        Next
    Case Else
        Dim el As Variant
        If ByIndex Then
            For Each el In FindMe
                LBox.Selected(el) = True
            Next
        Else
            For Each el In FindMe
                For i = 0 To LBox.ListCount - 1
                    If LBox.list(i) = el Then
                        LBox.Selected(i) = True
                        DoEvents
                    End If
                Next
            Next
        End If
    End Select
End Function

Public Function IsInArray( _
       ByVal value1 As Variant, _
       ByVal array1 As Variant, _
       Optional CaseSensitive As Boolean) _
        As Boolean
    Dim individualElement As Variant
    If CaseSensitive = True Then value1 = UCase(value1)
    For Each individualElement In array1
        If CaseSensitive = True Then individualElement = UCase(individualElement)
        If individualElement = value1 Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function
Function CollectionToArray(c As Collection) As Variant

    Dim A() As Variant: ReDim A(0 To c.Count - 1)
    Dim i As Long
    For i = 1 To c.Count
        A(i - 1) = c.Item(i)
    Next
    CollectionToArray = A
End Function
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
'#INCLUDE sheetExists
    Dim sht As Worksheet
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Rem --------------------

Rem Debug Print
Sub dp(var As Variant)
    '#INCLUDE DPH
    '#INCLUDE ArrayDimensions
    '#IMPORTS DPH
    '#IMPORTS ArrayDimensions
    Dim element As Variant
    Select Case TypeName(var)
        Case Is = "String", "Long", "Integer", "Boolean"
            Debug.Print var
            Rem todo How to handle multidimensional array?
        Case Is = "Variant()", "String()", "Long()", "Integer()"
            If ArrayDimensions(var) = 1 Then
                Dim i As Long
                For i = LBound(var) To UBound(var)
                    Debug.Print var(i)
                Next i
            ElseIf ArrayDimensions(var) > 1 Then
                DPH var
            End If
        Case Is = "Collection"
            For Each element In var
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
