Attribute VB_Name = "mFindRange"
Option Compare Text

Public Enum OPERATOR
    IS_LIKE
    IS_EQUAL
    NOT_EQUAL
    IS_CONTAINS
    NOT_CONTAINS
    STARTS_WITH
    ENDS_WITH
    GREATER_THAN
    GREATER_OR_EQUAL
    LESS_THAN
    LESS_OR_EQUAL
    IS_BETWEEN
    NOT_BETWEEN
End Enum

Function FindDateRange(rng As Range) As Range
    Dim CopyRange As Range
    Dim cell As Range
    For Each cell In rng
        If IsDate(cell) Then
            If CopyRange Is Nothing Then
                Set CopyRange = cell
            Else
                Set CopyRange = Union(CopyRange, cell)
            End If
        End If
    Next cell
    Set FindDateRange = CopyRange
End Function

Function FindNumericRange(rng As Range) As Range
Dim startTime
startTime = Now
    Dim CopyRange As Range
    Dim cell As Range
    For Each cell In rng
        If IsNumeric(cell) And Not IsDate(cell) Then
            If CopyRange Is Nothing Then
                Set CopyRange = cell
            Else
                Set CopyRange = Union(CopyRange, cell)
            End If
        End If
'        If Now() - startTime > TimeSerial(0, 0, 10) Then Stop
    Next cell
    Set FindNumericRange = CopyRange
End Function

Function FindStringRange(rng As Range) As Range
    Dim CopyRange As Range
    Dim cell As Range
    For Each cell In rng
        If IsDate(cell) Then
            If CopyRange Is Nothing Then
                Set CopyRange = cell
            Else
                Set CopyRange = Union(CopyRange, cell)
            End If
        End If
    Next cell
    Set FindStringRange = CopyRange
End Function

Function whichOption(Frame As Variant, controlType As String) As Variant
    Dim out As New Collection
    For Each Control In Frame.Controls
        If UCase(TypeName(Control)) = UCase(controlType) Then
            If Control.Value = True Then
                out.Add Control
            End If
        End If
    Next
    If out.Count = 1 Then
        Set whichOption = out(1)
    ElseIf out.Count > 1 Then
        Set whichOption = out
    End If
End Function

Function CellRow(cell As Range) As Range
    Dim ws As Worksheet
    Set ws = cell.Parent
    Dim R: R = cell.Row
    Dim c As Long: c = cell.CurrentRegion.Column
    If uFinder.oCurrentRegion.Value = True Then Set CellRow = ws.Range(ws.Cells(R, c), ws.Cells(R, c + cell.CurrentRegion.Columns.Count - 1))
    If uFinder.oUsedRange.Value = True Then Set CellRow = ws.Range(ws.UsedRange(R, 1), ws.UsedRange(R, ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1))
End Function

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Function Listbox_Selected(LBox As MSForms.ListBox, Count_Indexes_Values As Integer)
    Dim SelectedIndexes As String
    Dim SelectedValues As String
    Dim SelectedCount As Integer
    Dim i As Long
    With LBox
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                SelectedCount = SelectedCount + 1
                SelectedIndexes = SelectedIndexes & i & ","
                SelectedValues = SelectedValues & .list(i) & ","
            End If
        Next i
    End With
    If SelectedCount = 0 Then
        Listbox_Selected = 0
        Exit Function
    End If
    SelectedIndexes = left(SelectedIndexes, Len(SelectedIndexes) - 1)
    SelectedValues = left(SelectedValues, Len(SelectedValues) - 1)
    Select Case Count_Indexes_Values
        Case Is = 1
            Listbox_Selected = SelectedCount
        Case Is = 2
            Listbox_Selected = SelectedIndexes
        Case Is = 3
            Listbox_Selected = SelectedValues
    End Select
End Function

Function SheetAdd(SheetName As String, TargetWorkbook As Workbook) As Worksheet
    Dim NewSheet As Worksheet
    If WorksheetExists(SheetName, TargetWorkbook) = True Then
        Set SheetAdd = TargetWorkbook.Sheets(SheetName)
    Else
        Set SheetAdd = TargetWorkbook.Sheets.Add
        SheetAdd.Name = SheetName
    End If
End Function
Sub ListboxToRange(LBox As MSForms.ListBox, rng As Range)
rng.Resize(LBox.ListCount, LBox.ColumnCount) = LBox.list
End Sub
Function ArrayColumn(arr As Variant, col As Long) As Variant
    ArrayColumn = WorksheetFunction.index(arr, 0, col)
End Function
Sub ResizeControlColumns(ctr As MSForms.Control, Optional ResizeListbox As Boolean)
    If ctr.ListCount = 0 Then Exit Sub
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") 'SheetAdd("ListboxColumnwidth", ThisWorkbook)
    ws.Cells.clear
    
    On Error Resume Next
    Dim v As Variant: v = ctr.list
    Dim i As Long, Y As Long
    For i = LBound(v) To UBound(v)
        For Y = 0 To ctr.ColumnCount - 1
            If left(v(i, Y), 1) = "=" Then v(i, Y) = "'" & v(i, Y)
        Next
    Next
    
    '---Listbox to range-----
    Dim rng As Range
    Set rng = ws.Range("A1")
'    ListboxToRange ctr, rng
    Set rng = rng.Resize(UBound(ctr.list) + 1, ctr.ColumnCount)
    rng = v 'ctr.List
    rng.Font.Size = ctr.Font.Size + 2
    
    '---Get ColumnWidths------
    rng.Columns.AutoFit
    Dim sWidth As String
    Dim vR() As Variant
    Dim n As Integer
    Dim cell As Range
    For Each cell In rng.Resize(1)
        n = n + 1
        ReDim Preserve vR(1 To n)
        vR(n) = cell.EntireColumn.Width
    Next cell
    sWidth = Join(vR, ";")
    'Debug.Print sWidth

    '---assign ColumnWidths----
    With ctr
        .ColumnWidths = sWidth
        '.RowSource = "A1:A3"
        .BorderStyle = fmBorderStyleSingle
    End With
        
    'remove worksheet
'    Application.DisplayAlerts = False
    'ws.Delete
'    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    
    '----Resize Listbox--------
    If ResizeListbox = False Then Exit Sub
    Dim w As Long
    For i = LBound(vR) To UBound(vR)
        w = w + vR(i)
    Next
    DoEvents
    ctr.Width = w + 10
End Sub
Private Function SortCompare(one As Variant, two As Variant) As Boolean
    Select Case True
        Case Len(one) < Len(two)
            SortCompare = True
        Case Len(one) > Len(two)
            SortCompare = False
        Case Len(one) = Len(two)
            SortCompare = LCase$(one) < LCase$(two)
    End Select
End Function

Public Sub CustomQuickSort(list As Variant, first As Long, last As Long)
    Dim pivot As String
    Dim low As Long
    Dim high As Long

    low = first
    high = last
    pivot = list((first + last) \ 2)

    Do While low <= high
        Do While low < last And SortCompare(list(low), pivot)
            low = low + 1
        Loop
        Do While high > first And SortCompare(pivot, list(high))
            high = high - 1
        Loop
        If low <= high Then
            Dim swap As String
            swap = list(low)
            list(low) = list(high)
            list(high) = swap
            low = low + 1
            high = high - 1
        End If
    Loop

    If (first < high) Then CustomQuickSort list, first, high
    If (low < last) Then CustomQuickSort list, low, last
End Sub

Function Transpose2DArray(inputArray As Variant) As Variant

    Dim x As Long, yUbound As Long
    Dim Y As Long, xUbound As Long
    Dim tempArray As Variant

    xUbound = UBound(inputArray, 2)
    yUbound = UBound(inputArray, 1)
    
    ReDim tempArray(1 To xUbound, 1 To yUbound)
    
    For x = 1 To xUbound
        For Y = 1 To yUbound
            tempArray(x, Y) = inputArray(Y, x)
        Next Y
    Next x
    
    Transpose2DArray = tempArray
    
End Function

Function RangeToString(ByVal myRange As Range, Optional delim As String = ",") As String
    RangeToString = ""
    If Not myRange Is Nothing Then
        Dim myCell As Range
        For Each myCell In myRange
            RangeToString = RangeToString & delim & myCell.Value
        Next myCell
        'Remove extra comma
        RangeToString = right(RangeToString, Len(RangeToString) - Len(delim))
    End If
End Function

Function FindIfGetRow(FirstValue, operation As OPERATOR, SecondValue, offsetRow As Integer, offsetColumn As Integer, _
                      Optional wb As Workbook, Optional ws As Worksheet, Optional delim As String = ",") As Collection
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim output As New Collection
    Dim arr, element
    Dim str As String
    Dim c As Range
    Dim firstAddress As String
  
    If TypeName(ws) = "Nothing" Then
        For Each ws In wb.Worksheets
            With ws.Cells
                Set c = .Find(FirstValue, LookIn:=xlValues)
                If Not c Is Nothing Then
                    firstAddress = c.Address
                    Do
                        If compare(c.Offset(offsetRow, offsetColumn).Value, operation, SecondValue) = True Then
                            arr = ws.Range(ws.Cells(c.Row, c.CurrentRegion.Column), ws.Cells(c.Row, c.CurrentRegion.Column + c.CurrentRegion.Columns.Count - 1)).Value
                            str = ArrayToString(arr, delim)
                            output.Add str
                            Debug.Print str
                        End If
                        Set c = .FindNext(c)
                    Loop While Not c Is Nothing And c.Address <> firstAddress
                End If
            End With
        Next
    Else
        With ws.Cells
            Set c = .Find(FirstValue, LookIn:=xlValues)
            If Not c Is Nothing Then
                firstAddress = c.Address
                Do
                    If compare(c.Offset(offsetRow, offsetColumn).Value, operation, SecondValue) = True Then
                        arr = ws.Range(ws.Cells(c.Row, c.CurrentRegion.Column), ws.Cells(c.Row, c.CurrentRegion.Column + c.CurrentRegion.Columns.Count - 1)).Value
                        str = ArrayToString(arr, delim)
                        output.Add str
                        Debug.Print str
                    End If
                    Set c = .FindNext(c)
                Loop While Not c Is Nothing And c.Address <> firstAddress
            End If
        End With
    End If

    Set FindIfGetRow = output
End Function

'RETURNS A STRING FROM A 2 DIM ARRAY, SPERATED BY OPTIONAL DELIMITER AND VBNEWLINE FOR EACH ROW
'
'@AUTHOR ROBERT TODAR
Public Function ArrayToString(SourceArray As Variant, Optional Delimiter As String = ",") As String
    
    Dim Temp As String
    
    Select Case ArrayDimensionLength(SourceArray)
        'SINGLE DIMENTIONAL ARRAY
    Case 1
        Temp = Join(SourceArray, Delimiter)
        
        '2 DIMENSIONAL ARRAY
    Case 2
        Dim RowIndex As Long
        Dim ColIndex As Long
            
        'LOOP EACH ROW IN MULTI ARRAY
        For RowIndex = LBound(SourceArray, 1) To UBound(SourceArray, 1)
                
            'LOOP EACH COLUMN ADDING VALUE TO STRING
            For ColIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
                Temp = Temp & SourceArray(RowIndex, ColIndex)
                If ColIndex <> UBound(SourceArray, 2) Then Temp = Temp & Delimiter
            Next ColIndex
                
            'ADD NEWLINE FOR THE NEXT ROW (MINUS LAST ROW)
            If RowIndex <> UBound(SourceArray, 1) Then Temp = Temp & vbNewLine
        
        Next RowIndex
    End Select
    
    ArrayToString = Temp
    
End Function

'RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
Public Function ArrayDimensionLength(SourceArray As Variant) As Integer
    
    Dim arrayDimention As Integer
    Dim Test As Long

    On Error GoTo catch
    Do
        arrayDimention = arrayDimention + 1
        Test = UBound(SourceArray, arrayDimention)
    Loop
    
catch:
    ArrayDimensionLength = arrayDimention - 1

End Function

Function FindAllGetRow(FirstValue, Optional wb As Workbook, Optional ws As Worksheet) As Collection
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim output As New Collection
    Dim arr, element
    Dim str As String
    Dim c As Range
    Dim firstAddress As String
  
    If TypeName(ws) = "Nothing" Then
        For Each ws In wb.Worksheets
            With ws.Cells
                Set c = .Find(FirstValue, LookIn:=xlValues)
                If Not c Is Nothing Then
                    firstAddress = c.Address
                    Rem dp vbNewLine & ws.Name
                    Do
                        Rem dp c.Address
                        arr = ws.Range(ws.Cells(c.Row, c.CurrentRegion.Column), ws.Cells(c.Row, c.CurrentRegion.Column + c.CurrentRegion.Columns.Count - 1)).Value
                        str = ArrayToString(arr)
                        Do While InStr(1, str, ",,") > 0
                            str = Replace(str, ",,", ",")
                        Loop
                        str = str
                        output.Add str
                        Debug.Print str
                        Set c = .FindNext(c)
                    Loop While Not c Is Nothing And c.Address <> firstAddress
                End If
            End With
        Next
    Else
        With ws.Cells
            Set c = .Find(FirstValue, LookIn:=xlValues)
            If Not c Is Nothing Then
                firstAddress = c.Address
                Do
                    arr = ws.Range(ws.Cells(c.Row, c.CurrentRegion.Column), ws.Cells(c.Row, c.CurrentRegion.Column + c.CurrentRegion.Columns.Count - 1)).Value
                    str = ArrayToString(arr)
                    output.Add str
                    Debug.Print str
                    Set c = .FindNext(c)
                Loop While Not c Is Nothing And c.Address <> firstAddress
            End If
        End With
    End If

    Set FindAllGetRow = output
End Function

Function InStrExact(Start As Long, SourceText As String, WordToFind As String, _
                    Optional CaseSensitive As Boolean = False, _
                    Optional AllowAccentedCharacters As Boolean = False) As Long
    Dim x As Long, Str1 As String, Str2 As String, Pattern As String
    Const UpperAccentsOnly As String = "ÇÉÑ"
    Const UpperAndLowerAccents As String = "ÇÉÑçéñ"
    If CaseSensitive Then
        Str1 = SourceText
        Str2 = WordToFind
        Pattern = "[!A-Za-z0-9]"
        If AllowAccentedCharacters Then Pattern = Replace(Pattern, "!", "!" & UpperAndLowerAccents)
    Else
        Str1 = UCase(SourceText)
        Str2 = UCase(WordToFind)
        Pattern = "[!A-Z0-9]"
        If AllowAccentedCharacters Then Pattern = Replace(Pattern, "!", "!" & UpperAccentsOnly)
    End If
    For x = Start To Len(Str1) - Len(Str2) + 1
        If Mid(" " & Str1 & " ", x, Len(Str2) + 2) Like Pattern & Str2 & Pattern _
                                                   And Not Mid(Str1, x) Like Str2 & "'[" & Mid(Pattern, 3) & "*" Then
            InStrExact = x
            Exit Function
        End If
    Next
End Function

Function compare(inputValue, operation As OPERATOR, FirstComparison, Optional SecondComparison, Optional CaseSensitive As Boolean) As Boolean

    If TypeName(inputValue) = "Range" Then inputValue = inputValue.Value
    
    Select Case TypeName(inputValue)
        Case "String()", "Variant", "Variant()", "Collection"
            MsgBox "Not able to proccess this case at the moment"
            Stop
    End Select
    
    If TypeName(inputValue) = "String" Then
        If CaseSensitive = True Then
            inputValue = UCase(inputValue)
            FirstComparison = UCase(FirstComparison)
            If Not IsMissing(SecondComparison) Then SecondComparison = UCase(SecondComparison)
        End If
    ElseIf IsDate(inputValue) Then
        inputValue = CDate(inputValue)
    ElseIf IsNumeric(inputValue) Then
        inputValue = CDbl(inputValue)
    End If
        
    If IsDate(FirstComparison) Then
        FirstComparison = CDate(FirstComparison)
        If Not IsMissing(SecondComparison) Then
            If IsDate(SecondComparison) Then SecondComparison = CDate(SecondComparison)
        End If
    ElseIf IsNumeric(FirstComparison) Then
        FirstComparison = CDbl(FirstComparison)
        If Not IsMissing(SecondComparison) Then
            If IsNumeric(SecondComparison) Then SecondComparison = CDbl(SecondComparison)
        End If
    End If
    
    If operation = OPERATOR.IS_LIKE Then
        'If TypeName(inputValue) = TypeName(FirstComparison) Then
        If inputValue Like FirstComparison Then
            compare = True
        End If
        'End If
    ElseIf operation = OPERATOR.IS_CONTAINS Then
        If InStrExact(1, CStr(inputValue), CStr(FirstComparison)) > 0 Then compare = True
    ElseIf operation = OPERATOR.NOT_CONTAINS Then
        If InStrExact(1, CStr(inputValue), CStr(FirstComparison)) = 0 Then compare = True
    ElseIf operation = OPERATOR.NOT_EQUAL Then
        If inputValue <> FirstComparison Then compare = True
    ElseIf operation = OPERATOR.STARTS_WITH Then
        If inputValue Like FirstComparison & "*" Then compare = True
    ElseIf operation = OPERATOR.ENDS_WITH Then
        If inputValue Like "*" & FirstComparison Then compare = True
    ElseIf operation = OPERATOR.IS_EQUAL Then
        If inputValue = FirstComparison Then compare = True
    ElseIf operation = OPERATOR.GREATER_THAN Then
        If inputValue > FirstComparison Then compare = True
    ElseIf operation = OPERATOR.GREATER_OR_EQUAL Then
        If inputValue >= FirstComparison Then compare = True
    ElseIf operation = OPERATOR.IS_BETWEEN Then
        If FirstComparison < inputValue And inputValue < SecondComparison Then compare = True
    ElseIf operation = OPERATOR.NOT_BETWEEN Then
        If Not (FirstComparison < inputValue And inputValue < SecondComparison) Then compare = True
    ElseIf operation = OPERATOR.LESS_THAN Then
        If inputValue < FirstComparison Then compare = True
    ElseIf operation = OPERATOR.LESS_OR_EQUAL Then
        If inputValue <= FirstComparison Then compare = True
    End If
End Function

'Returns a range containing only cells that match the given value
Public Function RangeFindAll(ByRef SearchRange As Range, ByVal Value As Variant, Optional ByVal LookIn As XlFindLookIn = xlValues, Optional LookAt As XlLookAt = xlPart) As Range
    Dim FoundValues As Range, Found As Range, Prev As Range, Looper As Boolean: Looper = True
    Do While Looper
        'If we've found something before, then search from after that point
        If Not Prev Is Nothing Then Set Found = SearchRange.Find(Value, Prev, LookIn, LookAt)
        'If we haven't searched for anything before, then do an initial search
        If Found Is Nothing Then Set Found = SearchRange.Find(Value, LookIn:=LookIn, LookAt:=LookAt)
        If Not Found Is Nothing Then
            'If our search found something
            If FoundValues Is Nothing Then
                'If our found value repository is empty, then set it to what we just found
                Set FoundValues = Found
            Else
                If Not Intersect(Found, FoundValues) Is Nothing Then Looper = False
                'If the found value intersects with what we've already found, then we've looped through the SearchRange
                'Note: This check is performed BEFORE we insert the newly found data into our repository
                Set FoundValues = Union(FoundValues, Found)
                'If our found value repository contains data, then add what we just found to it
            End If
            Set Prev = Found
        End If
        If Found Is Nothing And Prev Is Nothing Then Exit Function
    Loop
    Set RangeFindAll = FoundValues
    Set FoundValues = Nothing
    Set Found = Nothing
    Set Prev = Nothing
End Function


