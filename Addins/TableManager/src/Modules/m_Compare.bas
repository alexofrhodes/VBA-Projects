Attribute VB_Name = "m_Compare"
Option Explicit

Public Enum operator
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
End Enum

Function CompareValues(ByVal value1 As Variant, ByVal value2 As Variant, ByVal operator As operator) As Boolean
'@AssignedModule m_Compare
'@INCLUDE DECLARATION operator
    Select Case operator
        Case IS_LIKE
            CompareValues = (UCase(CStr(value1)) Like "*" & UCase(CStr(value2)) & "*")
        Case IS_EQUAL
            CompareValues = (value1 = value2)
        Case NOT_EQUAL
            CompareValues = (value1 <> value2)
        Case IS_CONTAINS
            CompareValues = (InStr(1, CStr(value1), CStr(value2), vbTextCompare) > 0)
        Case NOT_CONTAINS
            CompareValues = (InStr(1, CStr(value1), CStr(value2), vbTextCompare) = 0)
        Case STARTS_WITH
            CompareValues = (Left(CStr(value1), Len(CStr(value2))) = CStr(value2))
        Case ENDS_WITH
            CompareValues = (Right(CStr(value1), Len(CStr(value2))) = CStr(value2))
        Case GREATER_THAN
            CompareValues = (value1 > value2)
        Case GREATER_OR_EQUAL
            CompareValues = (value1 >= value2)
        Case LESS_THAN
            CompareValues = (value1 < value2)
        Case LESS_OR_EQUAL
            CompareValues = (value1 <= value2)
    End Select
End Function


Function FilterArray2d(inputArray As Variant, HasHeader As Boolean, MatchString As String, _
        ComparisonOperator As operator, Optional ColumnIndex As Long = -1) As Variant
'@AssignedModule m_Compare
'@INCLUDE PROCEDURE CompareValues
'@INCLUDE PROCEDURE ArrayCombine
'@INCLUDE DECLARATION operator
    Dim numRows As Long
    Dim numCols As Long
    Dim resultArray() As Variant
    Dim resultIndex As Long
    Dim i As Long
    Dim firstRow As Long: firstRow = LBound(inputArray, 1)
    Dim firstColumn As Long: firstColumn = LBound(inputArray, 2)
    numRows = UBound(inputArray, 1) - firstRow + 1
    numCols = UBound(inputArray, 2) - firstColumn + 1
    
    ' Define the result array to have the same number of columns as the input array
    ReDim resultArray(firstRow To 1, firstColumn To numCols)
    resultIndex = 0
    
    ' Check if the array has a header
    Dim startRow As Long
    If HasHeader Then
        startRow = LBound(inputArray, 1)
        ' Copy the header row to the result array
        For i = 1 To numCols
            resultArray(1, i) = inputArray(1, i)
        Next i
        startRow = startRow + 1
        resultIndex = resultIndex + 1
    Else
        startRow = LBound(inputArray, 1) + 1
    End If
    
    ' Loop through the rows of the input array
    For i = startRow To numRows
        Dim rowMatches As Boolean
        rowMatches = False
        
        ' Check if the row matches the specified criteria
        If ColumnIndex < 0 Or ColumnIndex > numCols Then
            ' Match any cell in the row
            Dim j As Long
            For j = firstColumn To numCols
                If CompareValues(inputArray(i, j), MatchString, ComparisonOperator) Then
                    rowMatches = True
                    Exit For
                End If
            Next j
        Else
            ' Match the specified column
            If CompareValues(inputArray(i, ColumnIndex), MatchString, ComparisonOperator) Then
                rowMatches = True
            End If
        End If
        
        ' Copy the matching row to the result array
        If rowMatches Then
            Dim tmp()
            ReDim tmp(1 To 1, firstColumn To numCols)
            For j = firstColumn To numCols
                tmp(1, j) = inputArray(i, j)
            Next
            resultArray = ArrayCombine(resultArray, tmp, True)
        End If
    Next i
    
    ' Return the filtered array
    If resultIndex = 0 Then
        FilterArray2d = Array()
    Else
        FilterArray2d = resultArray
    End If
End Function

Private Function ArrayCombine(a As Variant, b As Variant, Optional stacked As Boolean = True) As Variant
    'assumes that A and B are 2-dimensional variant arrays
    'if stacked is true then A is placed on top of B
    'in this case the number of rows must be the same,
    'otherwise they are placed side by side A|B
    'in which case the number of columns are the same
    'LBound can be anything but is assumed to be
    'the same for A and B (in both dimensions)
    'False is returned if a clash
'@AssignedModule m_Compare

    Dim LB As Long, m_A As Long, n_A As Long
    Dim m_B As Long, n_B As Long
    Dim M As Long, N As Long
    Dim i As Long, j As Long, k As Long
    Dim c As Variant

    If TypeName(a) = "Range" Then a = a.Value
    If TypeName(b) = "Range" Then b = b.Value

    LB = LBound(a, 1)
    m_A = UBound(a, 1)
    n_A = UBound(a, 2)
    m_B = UBound(b, 1)
    n_B = UBound(b, 2)

    If stacked Then
        M = m_A + m_B + 1 - LB
        N = n_A
        If n_B <> N Then
            ArrayCombine = False
            Exit Function
        End If
    Else
        M = m_A
        If m_B <> M Then
            ArrayCombine = False
            Exit Function
        End If
        N = n_A + n_B + 1 - LB
    End If
    ReDim c(LB To M, LB To N)
    For i = LB To M
        For j = LB To N
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
    ArrayCombine = c
End Function

Public Function NumberOfArrayDimensions(arr As Variant) As Byte
'@AssignedModule m_Compare
    Dim Ndx As Byte
    Dim Res As Long
    On Error Resume Next
    Do
        Ndx = Ndx + 1
        Res = UBound(arr, Ndx)
    Loop Until Err.Number <> 0
    NumberOfArrayDimensions = Ndx - 1
End Function

Function isUserform(thing) As Boolean
'@AssignedModule m_Compare
    On Error Resume Next
    Dim Module As VBComponent
    Set Module = ThisWorkbook.VBProject.VBComponents(thing.Name)
    isUserform = Module.Type = vbext_ct_MSForm
End Function

