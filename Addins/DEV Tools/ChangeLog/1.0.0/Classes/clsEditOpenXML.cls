VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEditOpenXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mbCreateBackup As Boolean
Private mvSourceFile As Variant
Private msSheet2Change As String
Private msSheetId   As String
Private msSheetFileName As String
Private mbAddedZip  As Boolean
Private mvXMLFolderRoot As Variant
Private mvxmlfolder As Variant
Public Enum XMLFolder
    XMLFolder_root = 1
    XMLFolder_rels = 2
    XMLFolder_xl = 3
    XMLFolder_customUI = 4
    XMLFolder_docProps = 5
End Enum

Public Sub ExtractRibbonX(sFullFile As String, sSaveFile As String)
    '@AssignedModule clsEditOpenXML
    '@INCLUDE CLASS clsEditOpenXML
    Dim cEditOpenXML As clsEditOpenXML
    Dim sXML        As String
    Dim oxmldoc     As MSXML2.DOMDocument
    Set cEditOpenXML = New clsEditOpenXML
    With cEditOpenXML
        .CreateBackup = False
        .SourceFile = sFullFile
        .UnzipFile
        sXML = .GetXMLFromFile("customUI.xml", XMLFolder_customUI)
        If Len(sXML) > 0 Then
            Set oxmldoc = New DOMDocument
            oxmldoc.LoadXML sXML
            oxmldoc.Save sSaveFile
        End If
        sXML = .GetXMLFromFile("customUI14.xml", XMLFolder_customUI)
        If Len(sXML) > 0 Then
            Set oxmldoc = New DOMDocument
            oxmldoc.LoadXML sXML
            oxmldoc.Save Replace(sSaveFile, ".xml", "14.xml")
        End If
    End With
    Set cEditOpenXML = Nothing
End Sub

Public Sub WriteRibbonXML2File(sFile As String, sXML As String)
    '@AssignedModule clsEditOpenXML
    '@INCLUDE CLASS clsEditOpenXML
    Dim cEditOpenXML As clsEditOpenXML

    Set cEditOpenXML = New clsEditOpenXML

    With cEditOpenXML
        .SourceFile = sFile

        .UnzipFile

        .WriteXML2File sXML, "customUI.xml", XMLFolder_customUI

        .ZipAllFilesInFolder
    End With

    Set cEditOpenXML = Nothing
End Sub

Public Function GetXMLFromFile(sFileName As String, sXMLFolder As XMLFolder) As String
    '@AssignedModule clsEditOpenXML
    '@INCLUDE CLASS clsEditOpenXML
    Dim oxmldoc     As MSXML2.DOMDocument
    If Len(XMLFolder(sXMLFolder)) = 0 Then
        GetXMLFromFile = ""
    Else
        Set oxmldoc = New MSXML2.DOMDocument
        oxmldoc.Load XMLFolder(sXMLFolder) & sFileName
        GetXMLFromFile = oxmldoc.xml
        Set oxmldoc = Nothing
    End If
End Function

Public Sub WriteXML2File(sXML As String, sFileName As String, sXMLFolder As XMLFolder)
    '@AssignedModule clsEditOpenXML
    '@INCLUDE PROCEDURE FolderExists
    '@INCLUDE CLASS clsEditOpenXML
    Dim oxmldoc     As MSXML2.DOMDocument
    Set oxmldoc = New MSXML2.DOMDocument
    If sXMLFolder = XMLFolder_customUI Then
        If Not FolderExists(XMLFolder(XMLFolder_customUI)) Then
            MkDir XMLFolder(XMLFolder_customUI)
            oxmldoc.LoadXML sXML
            oxmldoc.Save XMLFolder(sXMLFolder) & sFileName
            AddCustomUIToRels
        End If
    End If
    oxmldoc.LoadXML sXML
    oxmldoc.Save XMLFolder(sXMLFolder) & sFileName
End Sub

Public Sub AddCustomUIToRels()
    '@AssignedModule clsEditOpenXML
    '@INCLUDE CLASS clsEditOpenXML
    Dim oxmldoc     As MSXML2.DOMDocument
    Dim oXMLElement As MSXML2.IXMLDOMNode
    Dim oXMLAttrib  As MSXML2.IXMLDOMAttribute
    Dim oNamedNodeMap As MSXML2.IXMLDOMNamedNodeMap
    Dim oXMLRelsList As MSXML2.IXMLDOMNodeList
    Set oxmldoc = New MSXML2.DOMDocument
    oxmldoc.Load XMLFolder(XMLFolder_rels) & ".rels"
    Set oXMLElement = oxmldoc.createNode(1, "Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")
    Set oNamedNodeMap = oXMLElement.Attributes
    Set oXMLAttrib = oxmldoc.createAttribute("Id")
    oXMLAttrib.NodeValue = "cuID"
    oNamedNodeMap.setNamedItem oXMLAttrib
    Set oXMLAttrib = oxmldoc.createAttribute("Type")
    oXMLAttrib.NodeValue = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
    oNamedNodeMap.setNamedItem oXMLAttrib
    Set oXMLAttrib = oxmldoc.createAttribute("Target")
    oXMLAttrib.NodeValue = "customUI/customUI.xml"
    oNamedNodeMap.setNamedItem oXMLAttrib
    Set oXMLRelsList = oxmldoc.SelectNodes("/Relationships")
    oXMLRelsList.item(0).appendChild oXMLElement
    oxmldoc.Save XMLFolder(XMLFolder_rels) & ".rels"
    Set oXMLAttrib = Nothing
    Set oXMLElement = Nothing
    Set oxmldoc = Nothing
End Sub

Private Function GetSheetIdFromSheetName(sSheetName) As String
    '@AssignedModule clsEditOpenXML
    '@INCLUDE CLASS clsEditOpenXML
    Dim oxmldoc     As MSXML2.DOMDocument
    Dim oxmlNode    As MSXML2.IXMLDOMNode
    Dim oXMLChildNode As MSXML2.IXMLDOMNode
    Dim oXMLTemp    As MSXML2.IXMLDOMNode
    If XMLFolder(XMLFolder_xl) <> "" And Sheet2Change <> "" Then
        Set oxmldoc = New MSXML2.DOMDocument
        oxmldoc.Load XMLFolder(XMLFolder_xl) & "workbook.xml"
        For Each oxmlNode In oxmldoc.ChildNodes
            For Each oXMLChildNode In oxmlNode.ChildNodes
                If oXMLChildNode.BaseName = "sheets" Then
                    For Each oXMLTemp In oXMLChildNode.ChildNodes
                        If oXMLTemp.Attributes.getNamedItem("name").NodeValue = sSheetName Then
                            GetSheetIdFromSheetName = oXMLTemp.Attributes.getNamedItem("r:id").NodeValue
                            Exit Function
                        End If
                    Next
                End If
            Next
        Next
    End If
End Function

Public Function GetSheetFileNameFromId(sSheetId As String) As String
    '@AssignedModule clsEditOpenXML
    '@INCLUDE CLASS clsEditOpenXML
    Dim oxmldoc     As MSXML2.DOMDocument
    Dim oxmlNode    As MSXML2.IXMLDOMNode
    Dim oXMLChildNode As MSXML2.IXMLDOMNode
    If XMLFolder(XMLFolder_xl) <> "" And Sheet2Change <> "" Then
        Set oxmldoc = New MSXML2.DOMDocument
        oxmldoc.Load XMLFolder(XMLFolder_xl) & "_rels\workbook.xml.rels"
        For Each oxmlNode In oxmldoc.ChildNodes
            For Each oXMLChildNode In oxmlNode.ChildNodes
                If oXMLChildNode.Attributes.getNamedItem("Id").NodeValue = sSheetId Then
                    GetSheetFileNameFromId = oXMLChildNode.Attributes.getNamedItem("Target").NodeValue
                    Exit Function
                End If
            Next
        Next
    End If
End Function

Private Function GetSheetNameFromId(sId As String) As String
    '@AssignedModule clsEditOpenXML
    '@INCLUDE CLASS clsEditOpenXML
    Dim oxmldoc     As MSXML2.DOMDocument
    Dim oxmlNode    As MSXML2.IXMLDOMNode
    Dim oXMLChildNode As MSXML2.IXMLDOMNode
    Dim oXMLChildChildNode As MSXML2.IXMLDOMNode
    If mvxmlfolder(XMLFolder_xl) <> "" Then
        Set oxmldoc = New MSXML2.DOMDocument
        oxmldoc.Load XMLFolder(XMLFolder_xl) & "workbook.xml"
        For Each oxmlNode In oxmldoc.ChildNodes
            For Each oXMLChildNode In oxmlNode.ChildNodes
                If oXMLChildNode.nodename = "sheets" Then
                    For Each oXMLChildChildNode In oXMLChildNode.ChildNodes
                        If oXMLChildChildNode.Attributes.getNamedItem("r:id").NodeValue = "rId" & val(sId) + 1 Then
                            GetSheetNameFromId = oXMLChildChildNode.Attributes.getNamedItem("name").NodeValue
                            Exit Function
                        End If
                    Next
                    Exit Function
                End If
            Next
        Next
    End If
End Function

Public Sub ZipAllFilesInFolder()
    '@AssignedModule clsEditOpenXML
    '@INCLUDE CLASS clsEditOpenXML
    Dim oShellApp   As Object
    Dim sDate       As String
    Dim sDefPath    As String
    Dim vFileNameZip As Variant
    Dim FSO         As Object
    Dim lFileCt     As Long
    Set FSO = CreateObject("scripting.filesystemobject")
    sDate = Format(Now, " dd-mmm-yy h-mm-ss")
    vFileNameZip = SourceFile & sDate & ".zip"
    NewZip vFileNameZip
    Set oShellApp = CreateObject("Shell.Application")
    lFileCt = oShellApp.Namespace(FolderName & "Unzipped " & fileName & Application.PathSeparator).items.Count
    oShellApp.Namespace(vFileNameZip).CopyHere oShellApp.Namespace(FolderName & "Unzipped " & fileName & Application.PathSeparator).items
    On Error Resume Next
    Do Until oShellApp.Namespace(vFileNameZip).items.Count = lFileCt
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    DoEvents
    Kill SourceFile
    Name vFileNameZip As SourceFile
    On Error Resume Next
    FSO.DeleteFolder FolderName & "Unzipped " & fileName, True
    On Error GoTo 0
    Set oShellApp = Nothing
End Sub

Public Sub UnzipFile()
    '@AssignedModule clsEditOpenXML
    '@INCLUDE PROCEDURE FolderExists
    '@INCLUDE CLASS clsEditOpenXML
    Dim FSO         As Object
    Dim oShellApp   As Object
    Set FSO = CreateObject("scripting.filesystemobject")
    XMLFolderRoot = FolderName
    If Right(XMLFolderRoot, 1) <> Application.PathSeparator Then
        XMLFolderRoot = XMLFolderRoot & "\UnZipped " & fileName & Application.PathSeparator
    Else
        XMLFolderRoot = XMLFolderRoot & "UnZipped " & fileName & Application.PathSeparator
    End If
    On Error Resume Next
    FSO.DeleteFolder XMLFolderRoot & "*", True
    Kill XMLFolderRoot & "*.*"
    If FolderExists(XMLFolderRoot) = False Then
        MkDir XMLFolderRoot
    End If
    Set oShellApp = CreateObject("Shell.Application")
    oShellApp.Namespace(XMLFolderRoot).CopyHere oShellApp.Namespace(SourceFile).items
    On Error Resume Next
    FSO.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
    Set oShellApp = Nothing
    Set FSO = Nothing
    Exit Sub
End Sub

Sub NewZip(sPath)
    '@AssignedModule clsEditOpenXML
    '@INCLUDE CLASS clsEditOpenXML
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Public Property Get CreateBackup() As Boolean
    CreateBackup = mbCreateBackup
End Property

Public Property Let CreateBackup(ByVal bCreateBackup As Boolean)
    mbCreateBackup = bCreateBackup
End Property

Private Sub Class_Initialize()
    CreateBackup = True
End Sub

Public Property Get SourceFile() As Variant
    SourceFile = mvSourceFile
End Property

Public Property Let SourceFile(ByVal vSourceFile As Variant)
    mvSourceFile = vSourceFile
    If CreateBackup Then
        If Len(Dir(vSourceFile & "(backup)")) > 0 Then
            Kill vSourceFile & "(backup)"
        End If
        FileCopy vSourceFile, vSourceFile & "(backup)"
    End If
    If Not vSourceFile Like "*.zip" Then
        Name vSourceFile As vSourceFile & ".zip"
        mvSourceFile = mvSourceFile & ".zip"
        AddedZip = True
    End If
End Property

Public Property Get FolderName() As Variant
    FolderName = Mid(SourceFile, 1, InStrRev(SourceFile, Application.PathSeparator))
End Property

Public Property Get fileName() As Variant
    If SourceFile <> "" Then
        fileName = Mid(SourceFile, InStrRev(SourceFile, Application.PathSeparator) + 1, Len(SourceFile))
    End If
End Property

Public Property Get XMLFolder(sXMLFolder As XMLFolder) As String
    Select Case sXMLFolder
        Case Is = XMLFolder_root
            XMLFolder = mvXMLFolderRoot
        Case Is = XMLFolder_customUI
            XMLFolder = mvXMLFolderRoot & "customUI" & Application.PathSeparator
        Case Is = XMLFolder_docProps
            XMLFolder = mvXMLFolderRoot & "docProps" & Application.PathSeparator
        Case Is = XMLFolder_rels
            XMLFolder = mvXMLFolderRoot & "_rels" & Application.PathSeparator
        Case Is = XMLFolder_xl
            XMLFolder = mvXMLFolderRoot & "xl" & Application.PathSeparator
    End Select
End Property

Public Property Get Sheet2Change() As String
    Sheet2Change = msSheet2Change
End Property

Public Property Let Sheet2Change(ByVal sSheet2Change As String)
    msSheet2Change = sSheet2Change
    SheetId = GetSheetIdFromSheetName(sSheet2Change)
    If SheetId <> "" Then
        SheetFileName = GetSheetFileNameFromId(SheetId)
    End If
End Property

Public Property Get SheetId() As String
    SheetId = msSheetId
End Property

Public Property Let SheetId(ByVal sSheetId As String)
    msSheetId = sSheetId
End Property

Public Property Get SheetFileName() As String
    SheetFileName = msSheetFileName
End Property

Public Property Let SheetFileName(ByVal sSheetFileName As String)
    msSheetFileName = sSheetFileName
End Property

Private Property Get AddedZip() As Boolean
    AddedZip = mbAddedZip
End Property

Private Property Let AddedZip(ByVal bAddedZip As Boolean)
    mbAddedZip = bAddedZip
End Property

Private Sub Class_Terminate()
    Dim FSO         As Object
    If AddedZip Then
        Name SourceFile As Left(SourceFile, Len(SourceFile) - 4)
    End If
    On Error Resume Next
    FSO.DeleteFolder XMLFolderRoot, True
End Sub

Private Property Get XMLFolderRoot() As Variant
    XMLFolderRoot = mvXMLFolderRoot
End Property

Private Property Let XMLFolderRoot(ByVal vXMLFolderRoot As Variant)
    mvXMLFolderRoot = vXMLFolderRoot
End Property

