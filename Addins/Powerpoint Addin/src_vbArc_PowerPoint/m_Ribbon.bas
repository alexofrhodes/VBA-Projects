Attribute VB_Name = "m_Ribbon"
Option Explicit

'Rem https://www.thespreadsheetguru.com/vba/refresh-ribbon-ui-custom-excel-add-in
'
Public myRibbon As IRibbonUI

#If VBA7 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As LongPtr)
#Else
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
#End If


Sub Ribbon_OnLoad(ribbon As IRibbonUI)

    'PURPOSE: Run code when Ribbon loads the UI to store Ribbon Object's Pointer ID code
    'Handle variable declaration if 32-bit or 64-bit Excel
    #If VBA7 Then
        Dim StoreRibbonPointer As LongPtr
    #Else
        Dim StoreRibbonPointer As Long
    #End If
    Set myRibbon = ribbon                   'Store Ribbon Object to Public variable
    StoreRibbonPointer = ObjPtr(ribbon)     'Store pointer to IRibbonUI
    IniWriteKey RibbonIni, "Ribbon", "Pointer", StoreRibbonPointer
    If IniReadKey(RibbonIni, "T1", "Label") = "" Then
        createRibbonIni
    End If
    
    FoldersCreate IMAGE_FOLDER
    
End Sub

Sub Ribbon_RefreshRibbon()
    If IniReadKey(RibbonIni, "Ribbon", "Pointer") = "" Then GoTo ErrorHandler
    On Error GoTo ErrorHandler
    If myRibbon Is Nothing Then
        Set myRibbon = GetRibbon(IniReadKey(RibbonIni, "Ribbon", "Pointer"))
    End If
    myRibbon.Invalidate
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    MsgBox "Ribbon UI Refresh Failed"
End Sub

#If VBA7 Then
Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
Function GetRibbon(ByVal lRibbonPointer As Long) As Object
#End If
    Dim objRibbon As Object
    CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
    Set GetRibbon = objRibbon
    Set objRibbon = Nothing
End Function

Sub createRibbonIni()
    MsgBox "The ribbon ini was not found. It will be created now." & vbNewLine & _
              "It will take only a moment, do not interrupt!"

    Dim IniFile As String
        IniFile = RibbonIni
        
    IniWriteKey IniFile, "T1", "Label", "vbArc"
    IniWriteKey IniFile, "T1", "Visible", True
    
    'the following loop can faster
    'change to append a string variable and write it to the file
    Dim g As Long, b As Long, s As Long, m As Long
    For g = 1 To 5
        IniWriteKey IniFile, "T1_G" & g, "Label", "Group" & g
        IniWriteKey IniFile, "T1_G" & g, "Visible", IIf(g = 1, True, False)
        For b = 1 To 10
            IniWriteKey IniFile, "T1_G" & g & "_B" & b, "Size", True
            IniWriteKey IniFile, "T1_G" & g & "_B" & b, "Label", "T1_G" & g & "_B" & b
            IniWriteKey IniFile, "T1_G" & g & "_B" & b, "ShowLabel", True
            IniWriteKey IniFile, "T1_G" & g & "_B" & b, "Visible", IIf(g = 1, True, False)
            IniWriteKey IniFile, "T1_G" & g & "_B" & b, "Enabled", True
            IniWriteKey IniFile, "T1_G" & g & "_B" & b, "Description", ""
            IniWriteKey IniFile, "T1_G" & g & "_B" & b, "ScreenTip", ""
            IniWriteKey IniFile, "T1_G" & g & "_B" & b, "SuperTip", ""
            IniWriteKey IniFile, "T1_G" & g & "_B" & b, "Image", "FileSaveAsPowerPointPptx"
            IniWriteKey IniFile, "T1_G" & g & "_B" & b, "ShowImage", False
        Next
        
        For s = 1 To 9
            IniWriteKey IniFile, "T1_G" & g & "_S" & s, "Size", False
            IniWriteKey IniFile, "T1_G" & g & "_S" & s, "Visible", IIf(g = 1, True, False)
            
            For b = 1 To 10
                IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_B", "Label", "T1_G" & g & "_S" & s & "_B"
                IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_B", "ShowLabel", True
                IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_B", "Enabled", True
                IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_B", "Description", ""
                IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_B", "ScreenTip", ""
                IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_B", "SuperTip", ""
                IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_B", "Image", "FileSaveAsPowerPointPptx"
                IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_B", "ShowImage", False

                For m = 1 To 10
                    IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_M" & m, "Label", "T1_G" & g & "_S" & s & "_M" & m
                    IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_M" & m, "ShowLabel", True
                    IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_M" & m, "Visible", True
                    IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_M" & m, "Enabled", True
                    IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_M" & m, "Description", ""
                    IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_M" & m, "ScreenTip", ""
                    IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_M" & m, "SuperTip", ""
                    IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_M" & m, "Image", "FileSaveAsPowerPointPptx"
                    IniWriteKey IniFile, "T1_G" & g & "_S" & s & "_M" & m, "ShowImage", False
                Next
            Next
        Next
    Next
    
    Dim content As Variant
        content = Split(TxtRead(IniFile), vbLf)
    Dim output As String
    Dim element As String
    Dim i As Long
    For i = LBound(content) To UBound(content)
        element = content(i)
        If Left(element, 1) = "[" Then
            output = output & element
        Else
            output = output & vbTab & element
        End If
        If i < UBound(content) Then output = output & vbLf
    Next
    TxtOverwrite IniFile, output
    Debug.Print "Ini created"
End Sub

Function RibbonIni() As String
    RibbonIni = ThisProjectPath & getFileName(ThisProject.FileName) & "_ribbon.ini"
End Function

Function Ribbon_SetValue(id, TargetProperty, NewValue)
    IniWriteKey RibbonIni, id, TargetProperty, NewValue
End Function

Function Ribbon_GetValue(id, TargetProperty)
    Ribbon_GetValue = IniReadKey(RibbonIni, id, TargetProperty)
End Function



Sub Ribbon_ButtonClick(control As IRibbonControl)
    Select Case Ribbon_GetValue(control.id, "Label")
    Case "Reload Ribbon"
        Ribbon_RefreshRibbon
    Case "Dev"
        uDEV.Show
    Case "IO"
        uVBProjects.Show
    Case "NewName"
        MsgBox "hello world"
    End Select
End Sub

Sub Ribbon_CheckboxClick(control As IRibbonControl, pressed As Boolean)
    Ribbon_SetValue control, "pressed", pressed
    Select Case Ribbon_GetValue(control.id, "Label")
    Case "Button1"
        '<CODE>
    End Select
End Sub

Sub Ribbon_TogglebuttonClick(control As IRibbonControl, pressed As Boolean)
    Ribbon_SetValue control, "pressed", pressed
    Select Case Ribbon_GetValue(control.id, "Label")
    Case "Button1"
        '<CODE>
    End Select
End Sub

Sub Ribbon_EditboxChanged(control As IRibbonControl, Text As String)
    Ribbon_SetValue control, "Text", Text
    Select Case Ribbon_GetValue(control.id, "Label")
    Case Is = "EditBox1"
        Select Case Text
            '<CODE>
        End Select
    End Select
End Sub

Sub Ribbon_ComboboxChanged(control As IRibbonControl, Text As String)
    Ribbon_SetValue control, "Text", Text
    Select Case Ribbon_GetValue(control.id, "Label")
    Case "Combo 1"
        Select Case Text
        Case "Item 1"
            '<CODE>
        End Select
    End Select
End Sub

'Sub Ribbon_GalleryClick(control As IRibbonControl, id As String, index As Integer)
'    Ribbon_SetValue control, "Text", LabelOf(id)
'    Select Case LabelOf(control.id)
'    Case "Gallery 1"
'        Select Case LabelOf(id)
'        Case "G1 Item 1"
'        '<CODE>
'        End Select
'    End Select
'End Sub

'Sub Ribbon_DropdownClick(control As IRibbonControl, id As String, index As Integer)
'    Ribbon_SetValue control, "Text", LabelOf(id)
'    Select Case LabelOf(control.id)
'    Case "DD1"
'        Select Case LabelOf(id)
'        Case "Item1"
'            '<CODE>
'        End Select
'    End Select
'End Sub

Function IMAGE_FOLDER() As String
    IMAGE_FOLDER = ThisProjectPath & "RibbonImages\"
End Function

Sub Ribbon_getImage(control As IRibbonControl, ByRef returnedVal)
    Dim image
    Dim ImageName As String
    ImageName = Ribbon_GetValue(control.id, "Image")
    If InStr(1, ImageName, ".") > 0 Then
        On Error GoTo ErrorHandler
        Dim strPath As String
        strPath = IMAGE_FOLDER
        If FileExists(strPath & ImageName) Then
            Set returnedVal = LoadPictureGDI(strPath & ImageName)
        Else
            returnedVal = "FileSaveAsPowerPointPptx"
        End If
    Else
        returnedVal = ImageName
    End If
ErrorHandler:
End Sub

Sub Ribbon_getVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_GetValue(control.id, "Visible")
End Sub

Sub Ribbon_getSize(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_GetValue(control.id, "Size")
End Sub

Sub Ribbon_getLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_GetValue(control.id, "Label")
End Sub

Sub Ribbon_getScreenTip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_GetValue(control.id, "ScreenTip")
End Sub

Sub Ribbon_getSuperTip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_GetValue(control.id, "superTip")
End Sub

Sub Ribbon_getDescription(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_GetValue(control.id, "Description")
End Sub

Sub Ribbon_getShowLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_GetValue(control.id, "ShowLabel")
End Sub

Sub Ribbon_getShowImage(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_GetValue(control.id, "ShowImage")
End Sub

Sub Ribbon_getEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_GetValue(control.id, "Enabled")
End Sub

Sub Ribbon_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_GetValue(control.id, "Pressed")
End Sub

Sub Ribbon_getEditboxText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Ribbon_GetValue(control.id, "Text")
End Sub
