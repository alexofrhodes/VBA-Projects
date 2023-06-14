VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uDEV 
   Caption         =   "vbArc ~ Anastasiou Alex"
   ClientHeight    =   7308
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10356
   OleObjectBlob   =   "uDEV.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uDEV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Const AUTHOR_EMAIL = "anastasioualex@gmail.com"

#If VBA7 Then
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
#Else
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
#End If

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uDEV
'* Created    : 06-10-2022 10:34
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Rem MakeFormTransparent me
Rem MakeFormBorderless Me
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private m_sngDownX As Single
Private m_sngDownY As Single

Public Function GetInternetConnectedState() As Boolean
    GetInternetConnectedState = InternetGetConnectedState(0&, 0&)
End Function

Private Sub MakeFormTransparent(Frm As Object, Optional Color As Variant)
    Dim formhandle As Long
    Dim bytOpacity As Byte
    formhandle = CLng(FindWindow(vbNullString, Frm.Caption))
    If IsMissing(Color) Then Color = vbWhite
    bytOpacity = 100
    SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
    Frm.BackColor = Color
    SetLayeredWindowAttributes formhandle, Color, bytOpacity, LWA_COLORKEY
End Sub

Private Sub MakeFormBorderless(Frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = CLng(FindWindow(vbNullString, Frm.Caption))
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, GWL_EXSTYLE, lngWindow
    DrawMenuBar lFrmHdl
End Sub

Private Sub LVK_Click()
    FollowLink ("https://vk.com/video/playlist/735281600_1")
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub LFaceBook_Click()
    MsgBox "No longer exists."
End Sub

Private Sub LGitHub_Click()
    FollowLink ("https://github.com/alexofrhodes")
End Sub

Private Sub LYouTube_Click()
    FollowLink ("https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg")
End Sub

Private Sub LBuyMeACoffee_Click()
    FollowLink ("http://paypal.me/alexofrhodes")
End Sub

Private Function CLIP(Optional StoreText As String) As String
    Dim x As Variant
    x = StoreText
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
            Case Len(StoreText)
                .SetData "text", x
            Case Else
                CLIP = .GetData("text")
            End Select
        End With
    End With
End Function

Private Sub LEmail_Click()
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
End Sub

Function OutlookCheck() As Boolean
    Dim xOLApp As Object
    Set xOLApp = CreateObject("Outlook.Application")
    If Not xOLApp Is Nothing Then
        OutlookCheck = True
        Set xOLApp = Nothing
        Exit Function
    End If
    OutlookCheck = False
End Function


Sub MailDev()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strBody As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    With OutMail
        .To = AUTHOR_EMAIL
        .cc = vbNullString
        .BCC = vbNullString
        .Subject = "DEV REQUEST OR FEEDBACK FOR -CODE ARCHIVE-"
        .body = strBody
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub



Private Sub UserForm_Initialize()
    MakeFormBorderless Me
    MakeFormTransparent Me, vbBlack
End Sub
