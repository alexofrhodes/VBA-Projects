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

#If Win64 Then
Public hWnd As LongPtr 'LongLong
#Else
Public hWnd As Long
#End If

Private oForm As Object
Private aForm As aUserform

Public oMultiPage As MultiPage
Private Initialized As Boolean

Public WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

''''''''''''''''''''''''''
'Multipage BackColor start
''''''''''''''''''''''''''
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type uPicDesc
    Size As Long
    Type As Long
    #If VBA7 Then
        hPic As LongPtr
        hPal As LongPtr
    #Else
       hPic As Long
       hPal As Long
    #End If
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


#If VBA7 Then
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirectAut Lib "oleAut32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirectPro Lib "olepro32.dll" Alias "OleCreatePictureIndirect" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
    Private Declare PtrSafe Function CopyImage Lib "USER32" (ByVal handle As LongPtr, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
    Private Declare PtrSafe Function GetDC Lib "USER32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "USER32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
    Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
    Private Declare PtrSafe Function FillRect Lib "USER32" (ByVal hdc As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
    Private Declare PtrSafe Function SetRect Lib "USER32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    
    Private hdc As LongPtr, hMemDc As LongPtr, hMemBmp As LongPtr, hBrush As LongPtr, hCopy As LongPtr, ar() As LongPtr
    
#Else
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
    If TypeName(MP) <> "MultiPage" Then Stop
    Set oForm = MP.Parent
    Set oMultiPage = MP
'    Set oMultiPageEvent = oMultiPage
    hWnd = MP.[_GethWnd]
    Set Init = Me
    Initialized = True
    Set aForm = aUserform.Init(oForm)
End Function

Public Sub Test()
    If Not Initialized Then End
End Sub

Public Sub BuildMenu()
'    MenuLabelCounter = 0
    'For MultiPage, TabStrip:
    '   SetObject = Object.Add([ Name [, Caption [, index ]]])
    '
    'For other controls:
    '   SetControl = object. Add(ProgID [, Name [, Visible ]] )
    Dim MP As MultiPage: Set MP = oMultiPage
    oForm.Height = MP.Height + oForm.Height - oForm.InsideHeight
    oForm.BackColor = MyColors.FormBackgroundDarkGray

    Dim sidebar As Frame
    Set sidebar = oForm.Controls.Add("forms.frame.1")
    sidebar.Name = "Sidebar"
    With sidebar
        .Caption = ""
        .BorderStyle = fmBorderStyleSingle
        .Left = 0
        .Top = 0
        .Height = MP.Height
        .Width = 60
        .BackColor = MyColors.FormSidebarMediumGray

    End With
    
    With MP
        .Left = sidebar.Left + sidebar.Width
        .Top = 0
        .Width = oForm.InsideWidth - sidebar.Width
        .Height = oForm.InsideHeight
        .style = fmTabStyleNone
        Dim i As Long
        For i = 0 To MP.Pages.count - 1
            SetBackColor i, BackColor:=MyColors.FormBackgroundDarkGray
        Next
    End With
    
    Dim lbl As MSForms.Label
    
    Dim Page
    For Each Page In MP.Pages
        Set lbl = sidebar.Controls.Add("forms.label.1")
        With lbl
            .WordWrap = False
            .AutoSize = True
            .Font.Size = 12
            .Top = AvailableFormOrFrameRow(sidebar)
            .Left = 0
            .ForeColor = vbWhite
            .Caption = "  " & Page.Caption
            .AutoSize = False
            .Height = 24
            .Tag = "page-" & Page.Index
        End With
'        MenuLabelCounter = MenuLabelCounter + 1
'        ReDim Preserve MenuLabels(1 To MenuLabelCounter)
'        Set MenuLabels(MenuLabelCounter).MenuLabelEvent = lbl
    Next
    
    Dim ctrl As control
    Dim maxWidth As Long: maxWidth = 0
    For Each ctrl In sidebar.Controls
        If maxWidth < ctrl.Width Then maxWidth = ctrl.Width
    Next
    For Each ctrl In sidebar.Controls
        ctrl.Width = maxWidth + 10
    Next
    sidebar.Width = maxWidth + 10
    
    MP.Left = sidebar.Left + sidebar.Width
    
    oForm.Width = sidebar.Width + MP.Width + oForm.Width - oForm.InsideWidth
    
    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll oForm
    
End Sub

Public Sub Emitter_LabelMouseOut(Label As MSForms.Label)
    If InStr(1, Label.Tag, "page-") > 0 Then
        If Label.BackColor <> &H80B91E Then Label.BackColor = &H534848
    End If
End Sub

Public Sub Emitter_LabelMouseOver(Label As MSForms.Label)
    If InStr(1, Label.Tag, "page-") > 0 Then
        If Label.BackColor <> &H80B91E Then Label.BackColor = &H808080
    End If
End Sub

Public Sub Emitter_LabelClick(ByRef Label As MSForms.Label)
    Dim ctrl As control
    If InStr(1, Label.Tag, "page-") > 0 Then
        For Each ctrl In oForm.Controls
            If InStr(1, ctrl.Tag, "page-") > 0 Then ctrl.BackColor = &H534848
        Next
        Label.BackColor = &H80B91E
        If InStr(1, Label.Tag, "page-") > 0 Then
            oMultiPage.Value = Split(Label.Tag, "-")(1)
        End If
    End If
End Sub

Public Function ActivePage() As MSForms.Page
    Set ActivePage = oMultiPage.Pages(oMultiPage.Value)
End Function

Public Sub SetBackColor(PageIndex, BackColor As Long)
'eg
'    Call SetBackColor(Page:=MultiPage1.Pages(0), BackColor:=vbRed)
'    Call SetBackColor(Page:=MultiPage1.Pages(1), BackColor:=RGB(20, 200, 100))

'by Jaafar
'https://www.mrexcel.com/board/threads/can-a-userform-multipage-backcolor-be-changed.79069/page-4

    Dim Page As MSForms.Page
    Set Page = oMultiPage.Pages(PageIndex)
    #If VBA7 Then
        Dim hLib As LongPtr
    #Else
        Dim hLib As Long
    #End If

    Dim R As RECT
    Dim IID_IDispatch As GUID
    Dim uPicinfo As uPicDesc
    Dim iPic As IPicture
    Dim lRet As Long
    Static i As Integer
    
    Page.PictureSizeMode = fmPictureSizeModeStretch
    
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
        Set Page.Picture = iPic
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
    Dim element As Variant
    
    For Each element In ar
        DeleteObject element
    Next

End Sub
