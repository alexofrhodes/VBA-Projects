VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem ***************************************************************************
Rem *
Rem * MODULE NAME:     USERFORM RESIZER CLASS
Rem * AUTHOR & DATE:   STEPHEN BULLEN, Office Automation Ltd
Rem *
Rem * CONTACT:         Stephen@oaltd.co.uk
Rem * WEB SITE:        http://www.oaltd.co.uk
Rem *
Rem * DESCRIPTION:     Handles the resizing and repositioning of controls on a userform
Rem *
Rem ***************************************************************************
Rem IN USERFORM
Rem
Rem Dim moResizer As New CFormResizer
Rem Private Sub UserForm_Activate()
Rem     Set moResizer.Form = Me
Rem End Sub
Rem Private Sub UserForm_Resize()
Rem     moResizer.FormResize
Rem End Sub
Rem This class makes a userform resizable and handles the resizing of all the controls on the userform,
Rem such that their physical dimensions (e.g. size and position) change as the form is resized.
Rem To specify which control(s) to resize (and how), you set the control's .Tag property at design time to
Rem indicate that the control's top, left, width and height should be adjusted as the form's size changes.
Rem
Rem Use the letters t, l, w and h in any order (or not at all) to state that the property should change as the form
Rem is resized.  Follow the property by a decimal to indicate that the control should change by a percentage of the
Rem form's change.
Rem
Rem For example:
Rem   hw           Sets the control's height and width to change with the form (e.g. if there's a single list box on the form)
Rem   tl           Sets the contol's top and left to change in line with the form (e.g. to keep it in the bottom-right corner)
Rem   w0.5         Sets the control's width to change by 0.5 that of the form's width change
Rem   w0.5l0.5     Sets the control's width and position to change by 0.5 that of the form's width change
Rem Windows API calls to do all the dirty work!

Option Explicit
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Const GWL_STYLE As Long = (-16)
Private Const WS_THICKFRAME As Long = &H40000
Private Const SW_SHOW As Long = 5
Dim moForm As Object
Dim mdWidth As Double
Dim mdHeight As Double

Public Property Set Form(oNew As Object)
    Dim hWndForm As Long, iStyle As Long
    Set moForm = oNew
    If val(Application.Version) < 9 Then
        hWndForm = FindWindow("ThunderXFrame", moForm.Caption)
    Else
        hWndForm = FindWindow("ThunderDFrame", moForm.Caption)
    End If
    iStyle = GetWindowLong(hWndForm, GWL_STYLE)
    iStyle = iStyle Or WS_THICKFRAME
    SetWindowLong hWndForm, GWL_STYLE, iStyle
    ShowWindow hWndForm, SW_SHOW
    DrawMenuBar hWndForm
    SetFocus hWndForm
    mdWidth = moForm.Width
    mdHeight = moForm.Height
End Property

Public Sub FormResize()
    Dim dWidthAdj As Double, dHeightAdj As Double, sTag As String
    Dim oCtl As MSForms.control
    If mdWidth = 0 Then mdWidth = moForm.Width
    If mdHeight = 0 Then mdHeight = moForm.Height
    dWidthAdj = moForm.Width - mdWidth
    dHeightAdj = moForm.Height - mdHeight
    For Each oCtl In moForm.Controls
        With oCtl
            sTag = UCase(.Tag)
            If InStr(1, sTag, "L", vbBinaryCompare) Then
                If .left + dWidthAdj <= 0 Then moForm.Width = mdWidth
            End If
            If InStr(1, sTag, "W", vbBinaryCompare) Then
                If .Width + dWidthAdj <= 0 Then moForm.Width = mdWidth
            End If
            If InStr(1, sTag, "T", vbBinaryCompare) Then
                If .top + dHeightAdj <= 0 Then moForm.Height = mdHeight
            End If
            If InStr(1, sTag, "H", vbBinaryCompare) Then
                If .Height + dHeightAdj <= 0 Then moForm.Height = mdHeight
            End If
        End With
    Next
    dWidthAdj = moForm.Width - mdWidth
    dHeightAdj = moForm.Height - mdHeight
    For Each oCtl In moForm.Controls
        With oCtl
            sTag = UCase(.Tag)
            If InStr(1, sTag, "L", vbBinaryCompare) Then .left = .left + dWidthAdj * ResizeFactor(sTag, "L")
            If InStr(1, sTag, "T", vbBinaryCompare) Then .top = .top + dHeightAdj * ResizeFactor(sTag, "T")
            If InStr(1, sTag, "W", vbBinaryCompare) Then .Width = .Width + dWidthAdj * ResizeFactor(sTag, "W")
            If InStr(1, sTag, "H", vbBinaryCompare) Then .Height = .Height + dHeightAdj * ResizeFactor(sTag, "H")
        End With
    Next
    mdWidth = moForm.Width
    mdHeight = moForm.Height
End Sub

Private Function ResizeFactor(sTag As String, sChange As String)
    Dim i As Integer, d As Double
    i = InStr(1, sTag, sChange, vbBinaryCompare)
    If i > 0 Then
        d = val(Mid$(sTag, i + 1))
        If d = 0 Then d = 1
    End If
    ResizeFactor = d
End Function

