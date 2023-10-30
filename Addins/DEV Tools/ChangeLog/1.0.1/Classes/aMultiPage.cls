VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aMultiPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TODO
'test buildmenu targeting a multipage inside a multipage

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aMultiPage
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 29-06-2023 13:41    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Const AUTHOR_EMAIL = "anastasioualex@gmail.com"

Private sbLeftWidth As Long
Private sbRightWidth As Long
Private sbBottomHeight As Long

#If Win64 Then
Public hwnd         As LongPtr    'LongLong
#Else
Public hwnd         As Long
#End If

Private oForm       As Object

Public oMultipage   As MultiPage


''''''''''''''''''''''''''
'Multipage BackColor start
''''''''''''''''''''''''''
Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

Private Type uPicDesc
    Size            As Long
Type            As Long
#If VBA7 Then
    hPic            As LongPtr
    hPal            As LongPtr
#Else
    hPic            As Long
    hPal            As Long
#End If
End Type

Private Type GUID
    Data1           As Long
    Data2           As Integer
    Data3           As Integer
    Data4(0 To 7)   As Byte
End Type


#If VBA7 Then
Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function OleCreatePictureIndirectAut Lib "oleAut32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Private Declare PtrSafe Function OleCreatePictureIndirectPro Lib "olepro32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private hdc As LongPtr, hMemDc As LongPtr, hMemBmp As LongPtr, hBrush As LongPtr, hCopy As LongPtr, ar() As LongPtr

#Else
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function OleCreatePictureIndirectAut Lib "oleAut32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Private Declare Function OleCreatePictureIndirectPro Lib "olepro32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private hdc As Long, hMemDc As Long, hMemBmp As Long, hBrush As Long, hCopy As Long, ar() As Long

#End If


Private Const IMAGE_BITMAP = 0
Private Const PICTYPE_BITMAP = 1
Private Const LR_COPYRETURNORG = &H4
Private Const S_OK = 0
''''''''''''''''''''''''''
'Multipage BackColor end
''''''''''''''''''''''''''


'Public WithEvents oMultiPageEvent As MultiPage

'Private MenuLabels() As New aMultiPage

'Private MenuLabelCounter As Long
'Public WithEvents MenuLabelEvent As MSForms.Label
'
'Private Sub MenuLabelEvent_Click()
'    MsgBox MenuLabelEvent.Caption
'End Sub

Public Function Init(MP As control) As aMultiPage
    '@AssignedModule aMultiPage
    '@INCLUDE CLASS aMultiPage
    If TypeName(MP) <> "MultiPage" Then Stop
    Set oForm = MP.Parent
    Set oMultipage = MP
    '    Set oMultiPageEvent = oMultiPage
    hwnd = MP.[_GethWnd]
    Set Init = Me
End Function

Public Sub BuildMenu(createSidebarMinimizers As Boolean)
    '@AssignedModule aMultiPage
    '@INCLUDE PROCEDURE AvailableFormOrFrameRow
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aUserform
    '@INCLUDE CLASS aMultiPage
    '@INCLUDE DECLARATION MyColors
    '@INCLUDE DECLARATION Ctrl
    '@INCLUDE DECLARATION Emitter
    Dim MP          As MultiPage: Set MP = oMultipage
    oForm.Height = MP.Height + oForm.Height - oForm.InsideHeight
    oForm.BackColor = MyColors.FormBackgroundDarkGray


    Dim SidebarLeft As Frame
    Set SidebarLeft = oForm.Controls.Add("forms.frame.1")
    SidebarLeft.Name = "SidebarLeft"
    With SidebarLeft
        .Caption = ""
        .SpecialEffect = fmSpecialEffectFlat
        '        .BorderStyle = fmBorderStyleSingle
        .Left = 0
        .Top = 0
        .Height = MP.Height
        .Width = 60
        .BackColor = MyColors.FormSidebarMediumGray

    End With

    With MP
        .Left = SidebarLeft.Left + SidebarLeft.Width
        .Top = 0
        '        .Width = oForm.InsideWidth - SidebarLeft.Width
        .Height = oForm.InsideHeight
        .Style = fmTabStyleNone
        Dim i       As Long
        For i = 0 To MP.Pages.Count - 1
            SetBackColor i, BackColor:=MyColors.FormBackgroundDarkGray
        Next
    End With

    Dim lbl         As MSForms.Label

    Dim page
    For Each page In MP.Pages
        Set lbl = SidebarLeft.Controls.Add("forms.label.1")
        lbl.Name = "sidebarLabel" & page.Name
        With lbl
            If page.index = 0 Then .BackColor = MyColors.FormSelectedGreen
            .BorderStyle = fmBorderStyleSingle
            '            .BorderColor = MyColors.FormSelectedGreen
            .WordWrap = False
            .AutoSize = True
            .Font.Size = 12
            .Top = AvailableFormOrFrameRow(SidebarLeft)
            .Left = 0
            .ForeColor = vbWhite
            .Caption = Space(4) & page.Caption
            .AutoSize = False
            .Height = 24
            .Tag = "page-" & page.index
            .ControlTipText = WorksheetFunction.Trim(.Caption)
        End With
        '        MenuLabelCounter = MenuLabelCounter + 1
        '        ReDim Preserve MenuLabels(1 To MenuLabelCounter)
        '        Set MenuLabels(MenuLabelCounter).MenuLabelEvent = lbl
    Next

    Dim ctrl        As control
    Dim maxWidth    As Long: maxWidth = 0
    For Each ctrl In SidebarLeft.Controls
        If maxWidth < ctrl.Width Then maxWidth = ctrl.Width
    Next
    For Each ctrl In SidebarLeft.Controls
        ctrl.Width = maxWidth + 12
    Next
    SidebarLeft.Width = maxWidth + 12

    MP.Left = SidebarLeft.Left + SidebarLeft.Width

    Dim SidebarRight As MSForms.Frame
    Dim SidebarBottom As MSForms.Frame
    On Error Resume Next
    Set SidebarRight = oForm.Controls("SidebarRight")
    Set SidebarBottom = oForm.Controls("SidebarBottom")
    If Not SidebarRight.Visible Then Set SidebarRight = Nothing
    If Not SidebarBottom.Visible Then Set SidebarBottom = Nothing
    On Error GoTo 0

    If Not SidebarRight Is Nothing Then
        If SidebarRight.Visible Then
            With SidebarRight
                .Caption = ""
                .SpecialEffect = fmSpecialEffectFlat
                .BorderStyle = fmBorderStyleNone
                .BackColor = MyColors.FormSidebarMediumGray    'vbred
                .Top = MP.Top
                .Left = MP.Left + MP.Width
                If SidebarBottom Is Nothing Then
                    .Height = MP.Height
                Else
                    If SidebarBottom.Visible = True Then
                        .Height = MP.Height + SidebarBottom.Height
                    Else
                        .Height = MP.Height
                    End If
                End If
            End With
        End If
    End If

    If Not SidebarBottom Is Nothing Then
        If SidebarBottom.Visible Then
            SidebarLeft.Height = SidebarLeft.Height + SidebarBottom.Height
            With SidebarBottom
                .Caption = ""
                .SpecialEffect = fmSpecialEffectFlat
                .BorderStyle = fmBorderStyleNone
                .BackColor = MyColors.FormSidebarMediumGray    'vbred
                .Top = MP.Top & MP.Height
                .Left = SidebarLeft.Left + SidebarLeft.Width - 1
                '            If SidebarRight Is Nothing Then
                .Width = MP.Width + 2
                '            Else
                '                .Width = MP.Width + SidebarRight.Width
                '            End If

            End With
        End If
    End If

    aUserform.Init(oForm).ResizeToFitControls

    sbLeftWidth = SidebarLeft.Width
    If Not SidebarRight Is Nothing Then sbRightWidth = SidebarRight.Width
    If Not SidebarBottom Is Nothing Then sbBottomHeight = SidebarBottom.Height

    If createSidebarMinimizers Then
        Dim item
        For Each item In Array("Left", "Right")
            If Not (item = "Right" And SidebarRight Is Nothing) Then
                Set lbl = oForm.Controls("Sidebar" & item).Controls.Add("forms.label.1", "ResizeSidebar" & item)
                With lbl
                    .Font.Name = "Consolas"
                    .AutoSize = False
                    .Width = oForm.Controls("Sidebar" & item).Width
                    .Height = 16
                    .Left = 0
                    .Top = oForm.Controls("Sidebar" & item).Top + oForm.Controls("Sidebar" & item).Height - lbl.Height    '- 6
                    .ForeColor = vbWhite
                    .Font.Bold = True
                    .Font.Size = 14
                    .Caption = "<"
                End With
            End If
        Next
    End If

    For Each ctrl In oForm.Controls
        If TypeName(ctrl) = "Label" Then
            ctrl.MouseIcon = LoadPicture(ThisWorkbook.path & "\Lib\img\Hand Cursor Pointer.ico")
            ctrl.MousePointer = fmMousePointerCustom
        End If
    Next

    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll oForm

End Sub


Private Sub Emitter_LabelMouseOut(Label As MSForms.Label)
    '@AssignedModule aMultiPage
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aMultiPage
    If InStr(1, Label.Tag, "page-") > 0 Or InStr(1, Label.Name, "ResizeSidebar") > 0 Then
        If Label.BackColor <> &H80B91E Then Label.BackColor = &H534848
    End If
End Sub

Private Sub Emitter_LabelMouseOver(Label As MSForms.Label)
    '@AssignedModule aMultiPage
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aMultiPage
    If InStr(1, Label.Tag, "page-") > 0 Or InStr(1, Label.Name, "ResizeSidebar") > 0 Then
        If Label.BackColor <> &H80B91E Then Label.BackColor = &H808080
    End If
End Sub

Private Sub Emitter_LabelClick(ByRef Label As MSForms.Label)
    '@AssignedModule aMultiPage
    '@INCLUDE PROCEDURE FollowLink
    '@INCLUDE PROCEDURE CLIP
    '@INCLUDE PROCEDURE OutlookCheck
    '@INCLUDE PROCEDURE GetInternetConnectedState
    '@INCLUDE CLASS EventListenerEmitter
    '@INCLUDE CLASS aUserform
    '@INCLUDE CLASS aMultiPage
    '@INCLUDE DECLARATION AUTHOR_EMAIL
    '@INCLUDE DECLARATION Ctrl
    If InStr(1, Label.Tag, "page-") > 0 Then
        Dim ctrl    As control
        For Each ctrl In oForm.Controls
            If InStr(1, ctrl.Tag, "page-") > 0 Then ctrl.BackColor = &H534848
        Next
        Label.BackColor = &H80B91E
        If InStr(1, Label.Tag, "page-") > 0 Then
            oMultipage.Value = Split(Label.Tag, "-")(1)
        End If
    End If

    If InStr(1, Label.Caption, ">") > 0 Or InStr(1, Label.Caption, "<") > 0 Then
        Dim SidebarLeft As MSForms.control
        Dim SidebarRight As MSForms.control
        Dim SidebarBottom As MSForms.control
        Dim MP      As MSForms.control
        On Error Resume Next
        Set SidebarLeft = oForm.Controls("SidebarLeft")
        Set SidebarRight = oForm.Controls("SidebarRight")
        Set SidebarBottom = oForm.Controls("SidebarBottom")
        Set MP = oForm.Controls("MultiPage1")
        On Error GoTo 0

        If InStr(1, Label.Name, "Sidebar") > 0 Then
            If InStr(1, Label.Name, "Right") > 0 Then
                If SidebarRight.Width <> sbRightWidth Then
                    SidebarRight.Width = sbRightWidth
                    oForm.Controls("ResizeSidebarRight").Caption = "<"
                Else
                    SidebarRight.Width = 12
                    oForm.Controls("ResizeSidebarRight").Caption = ">"
                End If

            ElseIf InStr(1, Label.Name, "Bottom") > 0 Then
                '            oForm.Controls("SidebarBottom").Height = sbBottomHeight
            ElseIf InStr(1, Label.Name, "Left") > 0 Then
                If SidebarLeft.Width <> sbLeftWidth Then
                    SidebarLeft.Width = sbLeftWidth
                    oForm.Controls("ResizeSidebarLeft").Caption = "<"
                Else
                    SidebarLeft.Width = 12
                    oForm.Controls("ResizeSidebarLeft").Caption = ">"
                End If
                MP.Left = SidebarLeft.Left + SidebarLeft.Width
                If Not SidebarRight Is Nothing Then SidebarRight.Left = MP.Left + MP.Width
                If Not SidebarBottom Is Nothing Then SidebarBottom.Left = SidebarLeft.Left + SidebarLeft.Width
            End If

            '       aUserform.Init(oForm).ResizeToFitControls
            If SidebarRight Is Nothing Then
                oForm.Width = MP.Left + MP.Width
            Else
                If SidebarRight.Visible Then
                    oForm.Width = SidebarRight.Left + SidebarRight.Width + 12
                Else
                    oForm.Width = MP.Left + MP.Width
                End If
            End If
        End If
    End If

    If InStr(1, Label.Tag, "followlink-") > 0 Then
        Select Case Split(Label.Tag, "-")(1)
            Case "youtube": FollowLink "https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg"
            Case "github": FollowLink "https://github.com/alexofrhodes/"
            Case "blog": FollowLink "https://alexofrhodes.github.io/"
            Case "buymeacoffee": FollowLink "http://paypal.me/alexofrhodes"
            Case "email"
                If GetInternetConnectedState = False Then
                    MsgBox "Seems Internet is not available"
                    Exit Sub
                End If
                If OutlookCheck = True Then
                    MailDev
                Else
                    Dim out As String
                    out = AUTHOR_EMAIL
                    CLIP out
                    MsgBox ("Seems Outlook is not available" & Chr(10) & _
                            "DEV's email address " & vbNewLine & out & vbNewLine & "copied to clipboard")
                End If
        End Select
    End If
End Sub


Public Sub AddThemeControlsSidbarRight()
    '@AssignedModule aMultiPage
    '@INCLUDE CLASS aMultiPage
    '@INCLUDE CLASS aFrame
    aFrame.Init(oForm.Controls("SidebarRight")).AddThemeControls
End Sub

Public Sub AddContactsToSidebarBottom()
    '@AssignedModule aMultiPage
    '@INCLUDE PROCEDURE FileExists
    '@INCLUDE PROCEDURE AvailableFormOrFrameColumn
    '@INCLUDE CLASS aMultiPage
    Dim SidebarBottom As MSForms.Frame
    On Error Resume Next
    Set SidebarBottom = oForm.Controls("SidebarBottom")
    On Error GoTo 0
    If SidebarBottom Is Nothing Then Exit Sub
    Dim lbl         As MSForms.Label
    Dim newPicture  As IPictureDisp
    Dim FilePath    As String
    Dim myLinks(1 To 5)
    myLinks(1) = "github"
    myLinks(2) = "blog"
    myLinks(3) = "youtube"
    myLinks(4) = "buymeacoffee"
    myLinks(5) = "email"

    Dim item
    For Each item In myLinks
        FilePath = ThisWorkbook.path & "\Lib\img\" & item & ".bmp"
        If FileExists(FilePath) Then
            Set newPicture = LoadPicture(FilePath)
            Set lbl = SidebarBottom.Controls.Add("Forms.Label.1")
            lbl.Visible = False
            With lbl
                .Top = 6
                .Width = 24
                .Height = 24
                .Left = AvailableFormOrFrameColumn(SidebarBottom, 0, , 12)
                Set .Picture = Nothing
                .Picture = newPicture
                .Tag = "followlink-" & item
                .ControlTipText = item
            End With
            lbl.Visible = True
        End If
    Next
End Sub


Public Function ActivePage() As MSForms.page
    '@AssignedModule aMultiPage
    '@INCLUDE CLASS aMultiPage
    Set ActivePage = oMultipage.Pages(oMultipage.Value)
End Function

Public Sub SetBackColor(PageIndex, BackColor As Long)
    'eg
    '    Call SetBackColor(Page:=MultiPage1.Pages(0), BackColor:=vbRed)
    '    Call SetBackColor(Page:=MultiPage1.Pages(1), BackColor:=RGB(20, 200, 100))
    '@AssignedModule aMultiPage
    '@INCLUDE CLASS aMultiPage
    '@INCLUDE DECLARATION GUID
    '@INCLUDE DECLARATION GetDC
    '@INCLUDE DECLARATION ReleaseDC

    'by Jaafar
    'https://www.mrexcel.com/board/threads/can-a-userform-multipage-backcolor-be-changed.79069/page-4

    Dim page        As MSForms.page
    Set page = oMultipage.Pages(PageIndex)
#If VBA7 Then
    Dim hLib        As LongPtr
#Else
    Dim hLib        As Long
#End If

    Dim R           As RECT
    Dim IID_IDispatch As GUID
    Dim uPicinfo    As uPicDesc
    Dim iPic        As IPicture
    Dim lRet        As Long
    Static i        As Integer

    page.PictureSizeMode = fmPictureSizeModeStretch

    hdc = GetDC(0)
    SetRect R, 0, 0, 1, 1

    With R
        hMemBmp = CreateCompatibleBitmap(hdc, .Right - .Left, .Bottom - .Top)
    End With

    hMemDc = CreateCompatibleDC(hdc)
    DeleteObject SelectObject(hMemDc, hMemBmp)
    hBrush = CreateSolidBrush(BackColor)
    FillRect hMemDc, R, hBrush
    hCopy = CopyImage(hMemBmp, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)

    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With uPicinfo
        .Size = Len(uPicinfo)
        .Type = PICTYPE_BITMAP
        .hPic = hCopy
        .hPal = 0
    End With

    hLib = LoadLibrary("oleAut32.dll")
    If hLib Then
        lRet = OleCreatePictureIndirectAut(uPicinfo, IID_IDispatch, True, iPic)
    Else
        hLib = LoadLibrary("olepro32.dll")
        lRet = OleCreatePictureIndirectPro(uPicinfo, IID_IDispatch, True, iPic)
    End If
    FreeLibrary hLib

    If lRet = S_OK Then
        Set page.Picture = iPic
    Else
        MsgBox "Unable to create Picture", vbCritical, "Error."
    End If

    DeleteObject hMemBmp
    DeleteObject hMemDc
    DeleteObject hBrush
    ReleaseDC 0, hdc

    ReDim Preserve ar(i)
    ar(i) = hCopy
    i = i + 1

End Sub

Private Sub DeleteResources()
    'private sub userform_Terminate()
    'Call DeleteResources
    'end sub
    '@AssignedModule aMultiPage
    '@INCLUDE CLASS aMultiPage
    Dim element     As Variant

    For Each element In ar
        DeleteObject element
    Next

End Sub


Function OutlookCheck() As Boolean
    '@AssignedModule aMultiPage
    '@INCLUDE CLASS aMultiPage
    Dim xOLApp      As Object
    Set xOLApp = CreateObject("Outlook.Application")
    If Not xOLApp Is Nothing Then
        OutlookCheck = True
        Set xOLApp = Nothing
        Exit Function
    End If
    OutlookCheck = False
End Function


Sub MailDev()
    '@AssignedModule aMultiPage
    '@INCLUDE CLASS aMultiPage
    '@INCLUDE DECLARATION AUTHOR_EMAIL
    Dim OutApp      As Object
    Dim OutMail     As Object
    Dim strBody     As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    With OutMail
        .To = AUTHOR_EMAIL
        .cc = vbNullString
        .BCC = vbNullString
        .Subject = "Feedback or request - " & ThisWorkbook.Name
        .body = strBody
        .display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

