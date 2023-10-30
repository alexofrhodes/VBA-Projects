VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aCollection
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 04-05-2023 12:17    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Public RestrictionTypename As String
Private myCollection As Collection

Public Property Get items() As Collection
    Set items = myCollection
End Property

Public Function Init(coll As Collection) As aCollection
    Select Case RestrictionTypename
        Case ""
            Set myCollection = coll

        Case Else
            Set myCollection = New Collection
            Dim element
            For Each element In coll
                Add element
            Next
    End Select
    Set Init = Me
End Function

Public Property Get item(NameOrNumber As Variant) As Variant
    Set item = myCollection(IndexOf(NameOrNumber))
End Property

Public Function IndexOf(ByVal item As Variant, _
        Optional ByVal StartIndex As Long = 1) As Long
    Dim collindex   As Long
    Dim collitemtype As Integer
    Dim itemtype    As Integer

    itemtype = VarType(item)
    For collindex = StartIndex To myCollection.Count
        collitemtype = VarType(myCollection(collindex))
        If collitemtype = itemtype Then
            Select Case collitemtype
                Case 0 To 1: IndexOf = collindex: Exit Function
                Case 2 To 8, 11, 14, 17: If myCollection(collindex) = item Then IndexOf = collindex: Exit Function
                Case 9: If myCollection(collindex) Is item Then IndexOf = collindex: Exit Function
                Case Else
                    Debug.Print "Unsupported type for CollectionIndexOf."
                    Debug.Assert False
            End Select
        End If
    Next
    IndexOf = 0
End Function

Public Sub RestrictionSet(ByVal Value As String)
    RestrictionTypename = Value
End Sub

Public Sub RestrictionRemove()
    RestrictionTypename = ""
End Sub

Public Sub Add(NewItem As Variant, Optional Description As String)
    Err.clear
    Select Case RestrictionTypename
        Case ""
            myCollection.Add NewItem
        Case Else
            If TypeName(NewItem) <> RestrictionTypename Then
                MsgBox "TypeName does not match restriction: " & RestrictionTypename
            Else
                On Error Resume Next
                myCollection.Add NewItem, Description
                On Error GoTo 0
                If Err.Number <> 0 Then Debug.Print "Item " & Description & " already contained"
            End If
    End Select
End Sub

Public Property Get Count() As Long
    Count = myCollection.Count
End Property

Public Sub Remove(NameOrNumber As Variant)
    myCollection.Remove NameOrNumber
End Sub

Public Function ToString(delim As String) As String
    Dim element
    Dim out         As String
    For Each element In myCollection
        out = IIf(out = "", element, out & delim & element)
    Next
    ToString = out
End Function

Public Function Distinct() As aCollection
    Dim result      As Collection: Set result = New Collection
    Dim eachItem    As Variant
    On Error Resume Next
    For Each eachItem In myCollection
        result.Add eachItem, CStr(eachItem)
    Next
    On Error GoTo 0
    Set myCollection = result
    Set Distinct = Me
End Function

Public Function Unique() As aCollection
    '-----------------------------------------------------------------------------------------------------------
    ' CollectionUnique     - Returns a collection of unique values from a full collection
    '                                   - In : myCollection As Collection
    '                                   - Out: Result of values, or error
    '                                   - Last Updated: 8/7/11 by AJS
    '-----------------------------------------------------------------------------------------------------------
    Dim result      As New Collection
    Dim eachItem As Variant, eachUnique As Variant
    Dim MatchFound  As Boolean
    '    On Error GoTo IsError:
    For Each eachItem In myCollection
        MatchFound = False
        For Each eachUnique In result
            If eachItem = eachUnique Then
                MatchFound = True
                Exit For
            End If
        Next
        If MatchFound = False Then result.Add eachItem
    Next
    Set myCollection = result
    Set Unique = Me
    '    Exit Function
    'IsError:
    '    Unique = CVErr(xlErrNA)
    '    Debug.Print "Error in aCollection.Unique: " & Err.Number & ": " & Err.Description
End Function

Public Function Sort() As aCollection
    Dim iCounter    As Integer
    Dim iCounter2   As Integer
    Dim Temp        As Variant
    For iCounter = 1 To myCollection.Count - 1
        For iCounter2 = iCounter + 1 To myCollection.Count
            If myCollection(iCounter) > myCollection(iCounter2) Then
                Temp = myCollection(iCounter2)
                myCollection.Remove iCounter2
                myCollection.Add Temp, , iCounter
            End If
        Next iCounter2
    Next iCounter
    Set Sort = Me
End Function

Public Function Contains( _
        Optional Key As Variant, _
        Optional item As Variant) As Boolean
    Dim strKey      As String
    Dim var         As Variant
    If Not IsMissing(Key) Then
        strKey = CStr(Key)
        On Error Resume Next
        Contains = True
        var = myCollection(strKey)
        If Err.Number = 91 Then GoTo CheckForObject
        If Err.Number = 5 Then GoTo NotFound
        On Error GoTo 0
        Exit Function
CheckForObject:
        If IsObject(myCollection(strKey)) Then
            Contains = True
            On Error GoTo 0
            Exit Function
        End If
NotFound:
        Contains = False
        On Error GoTo 0
        Exit Function
    ElseIf Not IsMissing(item) Then
        Contains = False
        For Each var In myCollection
            If var = item Then
                Contains = True
                Exit Function
            End If
        Next var
    Else
        Contains = False
    End If
End Function

Function ToArray() As Variant
    Dim a()         As Variant: ReDim a(0 To myCollection.Count - 1)
    Dim i           As Long
    For i = 1 To myCollection.Count
        a(i - 1) = myCollection.item(i)
    Next
    ToArray = a
End Function

Public Function Reverse() As aCollection
    Dim Output      As New Collection
    Dim i           As Long
    For i = myCollection.Count To 1 Step -1
        Output.Add myCollection.item(i)
    Next i
    Set myCollection = Output
    Set Reverse = Me
End Function

'------------------------------------------------------------------------'
'Function Mid : returns subset of a collection                       '
'  Similar to Mid$() on strings.                                         '
'See also: FromToColl, LeftColl, RightColl, ButLastColl                  '
'------------------------------------------------------------------------'
Public Function Mid(ByVal Start As Long, Optional ByVal Length As Variant) As aCollection
    Dim result      As New Collection
    Dim Count       As Long
    Dim c           As Long

    Count = myCollection.Count
    If Start < 1 Then Error 5
    If Start > Count Then
        Set result = result
    Else
        If IsMissing(Length) Then
            For c = Start To Count
                result.Add myCollection(c)
            Next
        ElseIf (Length >= (Count + 1 - Start)) Then
            For c = Start To Count
                result.Add myCollection(c)
            Next
        ElseIf Length < 0 Then
            Error 5
        Else
            For c = Start To Start + Count - 1
                result.Add myCollection(c)
            Next
        End If
    End If
    Set myCollection = result
    Set Mid = Me
End Function

'------------------------------------------------------------------------'
'Function FromTo : returns subset of a collection                    '
'See also: MidColl, LeftColl, RightColl                                  '
'------------------------------------------------------------------------'
Public Function FromTo(ByVal FromOffset As Long, ByVal ToOffset As Long) As aCollection
    Dim result      As New Collection
    Dim c           As Long

    If FromOffset > myCollection.Count Then
        Set FromTo = result
        Exit Function
    ElseIf FromOffset < 1 Then
        Error 5
    End If

    If ToOffset > myCollection.Count Then
        ToOffset = myCollection.Count
    ElseIf ToOffset < 1 Then
        Error 5
    End If

    If ToOffset < FromOffset Then
        Set FromTo = result
        Exit Function
    Else
        For c = FromOffset To ToOffset
            result.Add myCollection(c)
        Next
    End If

    Set myCollection = result
    Set FromTo = Me
End Function

'------------------------------------------------------------------------'
'Function LeftColl : returns left-most elements of a collection          '
'  Similar to Left$() on strings.                                        '
'See also: MidColl, RightColl, ButLastColl                               '
'------------------------------------------------------------------------'
Public Function Left(ByVal Length As Long) As aCollection
    Dim result      As New Collection
    Dim c           As Long

    If Length > myCollection.Count Then Length = myCollection.Count
    For c = 1 To Length
        result.Add myCollection(c)
    Next
    Set myCollection = result
    Set Left = Me
End Function

'------------------------------------------------------------------------'
'Function RightColl : returns right-most elements of a collection        '
'  Similar to Right$() on strings.                                       '
'See also: MidColl, LeftColl                                             '
'------------------------------------------------------------------------'
Public Function Right(ByVal Length As Long) As aCollection
    Dim result      As New Collection
    Dim c           As Long

    If Length > myCollection.Count Then Length = myCollection.Count
    For c = myCollection.Count - Length + 1 To myCollection.Count
        result.Add myCollection(c)
    Next
    Set myCollection = result
    Set Right = Me

End Function

Function Filter( _
        Match As String, _
        IgnoreCase As Boolean, _
        Include As Boolean) As aCollection
    Dim result      As New Collection
    If IgnoreCase Then Match = UCase(Match)
    Dim element
    For Each element In myCollection
        If IgnoreCase Then
            If Include Then
                If UCase(element) Like Match Then result.Add element
            Else
                If Not UCase(element) Like Match Then result.Add element
            End If
        Else
            If Include Then
                If UCase(element) Like Match Then result.Add element
            Else
                If UCase(element) Like Match Then result.Add element
            End If
        End If
    Next
    Set myCollection = result
    Set Filter = Me
End Function


Public Function CollectionsToArray2D(collections As Collection) As Variant
    If collections.Count = 0 Then Exit Function
    Dim columnCount As Long: columnCount = collections.Count
    Dim rowCount    As Long: rowCount = collections.item(1).Count
    Dim var         As Variant
    ReDim var(1 To rowCount, 1 To columnCount)
    Dim cols        As Long
    Dim rows        As Long
    For rows = 1 To rowCount
        For cols = 1 To collections.Count
            var(rows, cols) = collections(cols).item(rows)
        Next cols
    Next rows
    CollectionsToArray2D = var
End Function
