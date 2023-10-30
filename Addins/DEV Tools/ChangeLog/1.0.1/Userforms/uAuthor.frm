VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uAuthor 
   Caption         =   "UserForm1"
   ClientHeight    =   10068
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   19380
   OleObjectBlob   =   "uAuthor.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "uAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Userform   : uAuthor
'* Created    : 06-10-2022 10:34
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* DONATE     : http://paypal.me/alexofrhodes
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Const AUTHOR_EMAIL = "anastasioualex@gmail.com"

#If VBA7 Then
Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#Else
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
#End If


Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private m_sngDownX  As Single
Private m_sngDownY  As Single

Public Function GetInternetConnectedState() As Boolean
    '@AssignedModule uAuthor
    '@INCLUDE USERFORM uAuthor
    '@INCLUDE DECLARATION InternetGetConnectedState
    GetInternetConnectedState = InternetGetConnectedState(0&, 0&)
End Function

Private Sub MakeFormTransparent(frm As Object, Optional color As Variant)
    '@AssignedModule uAuthor
    '@INCLUDE USERFORM uAuthor
    '@INCLUDE DECLARATION GWL_EXSTYLE
    '@INCLUDE DECLARATION LWA_COLORKEY
    '@INCLUDE DECLARATION WS_EX_LAYERED
    '@INCLUDE DECLARATION FindWindow
    '@INCLUDE DECLARATION GetWindowLong
    '@INCLUDE DECLARATION SetLayeredWindowAttributes
    '@INCLUDE DECLARATION SetWindowLong
    Dim formhandle  As Long
    Dim bytOpacity  As Byte
    formhandle = CLng(FindWindow(vbNullString, frm.Caption))
    If IsMissing(color) Then color = vbWhite
    bytOpacity = 100
    SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
    frm.BackColor = color
    SetLayeredWindowAttributes formhandle, color, bytOpacity, LWA_COLORKEY
End Sub

Private Sub MakeFormBorderless(frm As Object)
    '@AssignedModule uAuthor
    '@INCLUDE USERFORM uAuthor
    '@INCLUDE DECLARATION GWL_EXSTYLE
    '@INCLUDE DECLARATION GWL_STYLE
    '@INCLUDE DECLARATION WS_CAPTION
    '@INCLUDE DECLARATION WS_EX_DLGMODALFRAME
    '@INCLUDE DECLARATION DrawMenuBar
    '@INCLUDE DECLARATION FindWindow
    '@INCLUDE DECLARATION GetWindowLong
    '@INCLUDE DECLARATION SetWindowLong
    Dim lngWindow   As Long
    Dim lFrmHdl     As Long
    lFrmHdl = CLng(FindWindow(vbNullString, frm.Caption))
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, GWL_EXSTYLE, lngWindow
    DrawMenuBar lFrmHdl
End Sub

Private Sub LVK_Click()
    '@AssignedModule uAuthor
    '@INCLUDE PROCEDURE FollowLink
    '@INCLUDE USERFORM uAuthor
    FollowLink ("https://vk.com/video/playlist/735281600_1")
End Sub

Private Sub LExit_Click()
    '@AssignedModule uAuthor
    '@INCLUDE USERFORM uAuthor
    Unload Me
End Sub

Private Sub LBLOG_Click()
    '@AssignedModule uAuthor
    '@INCLUDE PROCEDURE FollowLink
    '@INCLUDE USERFORM uAuthor
    FollowLink ("https://alexofrhodes.github.io")
End Sub

Private Sub LGitHub_Click()
    '@AssignedModule uAuthor
    '@INCLUDE PROCEDURE FollowLink
    '@INCLUDE USERFORM uAuthor
    FollowLink ("https://github.com/alexofrhodes")
End Sub

Private Sub LYouTube_Click()
    '@AssignedModule uAuthor
    '@INCLUDE PROCEDURE FollowLink
    '@INCLUDE USERFORM uAuthor
    FollowLink ("https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg")
End Sub

Private Sub LBuyMeACoffee_Click()
    '@AssignedModule uAuthor
    '@INCLUDE PROCEDURE FollowLink
    '@INCLUDE USERFORM uAuthor
    FollowLink "http://paypal.me/alexofrhodes"
End Sub

Private Function CLIP(Optional StoreText As String) As String
    '@AssignedModule uAuthor
    '@INCLUDE USERFORM uAuthor
    Dim X           As Variant
    X = StoreText
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(StoreText)
                    .SetData "text", X
                Case Else
                    CLIP = .GetData("text")
            End Select
        End With
    End With
End Function

Private Sub LEmail_Click()
    '@AssignedModule uAuthor
    '@INCLUDE PROCEDURE CLIP
    '@INCLUDE PROCEDURE OutlookCheck
    '@INCLUDE PROCEDURE GetInternetConnectedState
    '@INCLUDE USERFORM uAuthor
    '@INCLUDE DECLARATION AUTHOR_EMAIL
    If GetInternetConnectedState = False Then
        MsgBox "Seems Internet is not available"
        Exit Sub
    End If
    If OutlookCheck = True Then
        MailDev
    Else
        Dim out     As String
        out = AUTHOR_EMAIL
        CLIP out
        MsgBox ("Seems Outlook is not available" & Chr(10) & _
                "DEV's email address " & vbNewLine & out & vbNewLine & "copied to clipboard")
    End If
End Sub

Function OutlookCheck() As Boolean
    '@AssignedModule uAuthor
    '@INCLUDE USERFORM uAuthor
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
    '@AssignedModule uAuthor
    '@INCLUDE USERFORM uAuthor
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

Private Sub UserForm_Initialize()
    '@AssignedModule uAuthor
    '@INCLUDE USERFORM uAuthor
    '@INCLUDE DECLARATION GetSystemMetrics

    Dim screenWidth As Long
    Dim screenHeight As Long

    screenWidth = GetSystemMetrics(0)  ' SM_CXSCREEN
    screenHeight = GetSystemMetrics(1)  ' SM_CYSCREEN

    Me.Width = screenWidth
    Me.Height = screenHeight

    Frame1.Left = Me.Width / 2 - Image1.Width
    Frame1.Top = Me.Height / 2 - Image1.Height
    MakeFormBorderless Me
    '    MakeFormTransparent Me, vbYellow
End Sub



