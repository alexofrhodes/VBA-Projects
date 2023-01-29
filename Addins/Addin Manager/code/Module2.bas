Attribute VB_Name = "Module2"
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Sub DownloadFile(FileUrl As String, SaveAs As String)
    URLDownloadToFile 0, FileUrl, SaveAs, 0, 0
End Sub

Function vbArcAddins() As Variant
'#IMPORTS DownloadTextFile
    Dim v
    v = Filter(Split(DownloadTextFile("https://github.com/alexofrhodes/vbArc-addins/raw/main/ListOfAddins.txt"), "  " & vbLf), "xlam", True)
    Dim i As Long
'    For i = LBound(v) To UBound(v)
'        v(i) = Mid(v(i), InStr(1, v(i), ">") + 1)
'    Next
    vbArcAddins = v
End Function

Public Function FileExists(ByVal FileName As String) As Boolean
    If Right(FileName, 1) = "\" Then FileName = Left(FileName, Len(FileName) - 1)
    FileExists = (Dir(FileName, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "")
    
End Function

Sub DownloadFileFromURL()        'FileUrl As String, SaveAs As String
'#IMPORTS DownloadFile

    Dim FileUrl As String
    Dim objXmlHttpReq As Object
    Dim objStream As Object
    'example
    FileUrl = "https://github.com/alexofrhodes/vbArc-addins/raw/main/Code Printer/CodePrinter.xlam"

    Set objXmlHttpReq = CreateObject("Microsoft.XMLHTTP")
    objXmlHttpReq.Open "GET", FileUrl, False, "username", "password"
    objXmlHttpReq.send

    If objXmlHttpReq.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Open
        objStream.Type = 1
        objStream.Write objXmlHttpReq.responseBody
        objStream.SaveToFile "C:\Users\acer\Downloads\CodePrinter.xlam", 2
        objStream.Close
    End If

End Sub

Function DownloadTextFile(URL As String) As String
    On Error GoTo Err_GetFromWebpage
    Dim objWeb As Object
    Dim strXML As String
    Set objWeb = CreateObject("Msxml2.ServerXMLHTTP")
    objWeb.Open "GET", URL, False
    objWeb.setRequestHeader "Content-Type", "text/xml"
    objWeb.setRequestHeader "Cache-Control", "no-cache"
    objWeb.setRequestHeader "Pragma", "no-cache"
    objWeb.send
    strXML = objWeb.responseText
    DownloadTextFile = strXML
End_GetFromWebpage:
    Set objWeb = Nothing
    Exit Function
Err_GetFromWebpage:
    MsgBox Err.Description & " (" & Err.Number & ")"
    Resume End_GetFromWebpage
End Function


