VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class      : aComboBox
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 27-06-2023 12:35    Alex                figured how to get the hwnd
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Public oComboBox    As ComboBox

Rem resize combobox List Portion
#If VBA7 Then
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hdc As LongPtr, ByVal lpStr As LongPtr, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare PtrSafe Function SetMapMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nMapMode As Long) As Long
Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hdc As LongPtr, ByVal lpsz As LongPtr, ByVal cbString As Long, lpSize As ComboboxListSize) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As LongPtr, ByVal lpStr As LongPtr, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal nMapMode As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As LongPtr, ByVal lpsz As LongPtr, ByVal cbString As Long, lpSize As ComboboxListSize) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

#If Win64 Then
Private Const NULL_PTR = 0^
#Else
Private Const NULL_PTR = 0&
#End If

Private Type ComboboxListSize
    cx              As Long
    cy              As Long
End Type

Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

Public Function Init(cbx As ComboBox) As aComboBox
    '@AssignedModule aComboBox
    '@INCLUDE CLASS aComboBox
    Set oComboBox = cbx
    Set Init = Me
End Function

Public Function AutoSizeDropDown(Optional ByVal JumpToWidestEntry As Boolean = False) As Long
    '@AssignedModule aComboBox
    '@INCLUDE CLASS aComboBox
    '@INCLUDE DECLARATION GetSystemMetrics

    Const MM_TEXT = 1&
    Const SM_CXVSCROLL = 2&

    Static bDropDownState As Boolean
    Dim hMemDc      As LongPtr
    Dim hPrevFont   As LongPtr
    Dim IFont       As stdole.IFont
    Dim lPrevMPMode As Long
    Dim lIndex      As Long
    Dim lMaxWidth   As Long
    Dim lMaxIndex   As Long
    Dim sItemText   As String
    Dim sWidestText As String
    Dim tSize       As ComboboxListSize

    hMemDc = CreateCompatibleDC(NULL_PTR)
    If hMemDc Then
        lPrevMPMode = SetMapMode(hMemDc, MM_TEXT)
        Set IFont = oComboBox.Font
        hPrevFont = SelectObject(hMemDc, IFont.hFont)
        With oComboBox
            For lIndex = 0& To .ListCount - 1&
                sItemText = .list(lIndex)
                Call GetTextExtentPoint32(hMemDc, StrPtr(sItemText), Len(sItemText), tSize)
                If lMaxWidth < tSize.cx Then
                    lMaxWidth = tSize.cx
                    sWidestText = sItemText
                    lMaxIndex = lIndex
                End If
            Next lIndex
            'Debug.Print lMaxIndex, sWidestText
            AutoSizeDropDown = lMaxIndex
            lMaxWidth = PXtoPT(hMemDc, lMaxWidth + 2.5 * GetSystemMetrics(SM_CXVSCROLL), False)
            Call SetMapMode(hMemDc, lPrevMPMode)
            Call SelectObject(hMemDc, hPrevFont)
            Call DeleteDC(hMemDc)
            .ListWidth = lMaxWidth
            If JumpToWidestEntry Then
                If bDropDownState = False Then
                    .ListIndex = lMaxIndex
                End If
            End If
            bDropDownState = Not bDropDownState
        End With
    End If

End Function

Private Function ScreenDPI(ByVal bVert As Boolean, ByVal hdc As LongPtr) As Long
    '@AssignedModule aComboBox
    '@INCLUDE CLASS aComboBox
    '@INCLUDE DECLARATION LOGPIXELSX
    '@INCLUDE DECLARATION LOGPIXELSY
    '@INCLUDE DECLARATION GetDeviceCaps
    Const LOGPIXELSX As Long = 88&
    Const LOGPIXELSY As Long = 90&
    Static lDPI(1&) As Long

    If lDPI(0&) = 0& Then
        lDPI(0&) = GetDeviceCaps(hdc, LOGPIXELSX)
        lDPI(1&) = GetDeviceCaps(hdc, LOGPIXELSY)
    End If
    ScreenDPI = lDPI(Abs(bVert))
End Function

Private Function PXtoPT(ByVal hdc As LongPtr, ByVal Pixels As Single, ByVal bVert As Boolean) As Single
    '@AssignedModule aComboBox
    '@INCLUDE CLASS aComboBox
    Const POINTSPERINCH As Long = 72&
    PXtoPT = (Pixels / (ScreenDPI(bVert, hdc) / POINTSPERINCH))
End Function



Public Sub LoadVBProjects()
    '@AssignedModule aComboBox
    '@INCLUDE PROCEDURE WorkbookProjectProtected
    '@INCLUDE CLASS aComboBox
    Dim coll        As New Collection
    Dim wb          As Workbook
    For Each wb In Workbooks
        If Len(wb.path) > 0 Then
            If WorkbookProjectProtected(wb) = False Then
                On Error Resume Next
                coll.Add wb.Name, wb.Name
                On Error GoTo 0
            End If
        End If
    Next
    Rem list addins
    Dim vbProj      As VBProject
    Dim wbPath      As String
    For Each vbProj In Application.VBE.VBProjects
        On Error GoTo ErrorHandler
        wbPath = vbProj.fileName
        If Right(wbPath, 4) = "xlam" Or Right(wbPath, 3) = "xla" Then
            Dim wbName As String
            wbName = Mid(wbPath, InStrRev(wbPath, "\") + 1)
            If WorkbookProjectProtected(Workbooks(wbName)) = False Then
                On Error Resume Next
                coll.Add wbName, wbName
                On Error GoTo 0
            End If
        End If
SKIP:
    Next vbProj

    Dim el          As Variant
    For Each el In coll
        oComboBox.AddItem el
    Next

    Exit Sub
ErrorHandler:
    If Err.Number = 76 Then GoTo SKIP
End Sub

