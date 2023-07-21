Attribute VB_Name = "z_M"
Option Explicit

Function RandomStringArray(ByVal rowCount As Long, ByVal columnCount As Long, maxStringLength) As Variant
    Dim resultArray() As Variant
    ReDim resultArray(1 To rowCount, 1 To columnCount)
    
    Dim i As Long, j As Long
    Dim alphabet As String
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
            
            Dim k As Long
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

