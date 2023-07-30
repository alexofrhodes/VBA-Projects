VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFormResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'*
'* MODULE NAME:     USERFORM RESIZER CLASS
'* AUTHOR & DATE:   STEPHEN BULLEN, Office Automation Ltd
'*
'* CONTACT:         Stephen@oaltd.co.uk
'* WEB SITE:        http://www.oaltd.co.uk
'*
'* DESCRIPTION:     Handles the resizing and repositioning of controls on a userform
'*
'***************************************************************************
'IN USERFORM
'
'Dim moResizer As New CFormResizer
'Private Sub UserForm_Activate()
'    Set moResizer.Form = Me
'End Sub
'Private Sub UserForm_Resize()
'    moResizer.FormResize
'End Sub
'This class makes a userform resizable and handles the resizing of all the controls on the userform,
'such that their physical dimensions (e.g. size and position) change as the form is resized.
'To specify which control(s) to resize (and how), you set the control's .Tag property at design time to
'indicate that the control's top, left, width and height should be adjusted as the form's size changes.
'
'Use the letters t, l, w and h in any order (or not at all) to state that the property should change as the form
'is resized.  Follow the property by a decimal to indicate that the control should change by a percentage of the
'form's change.
'
'For example:
'  hw           Sets the control's height and width to change with the form (e.g. if there's a single list box on the form)
'  tl           Sets the contol's top and left to change in line with the form (e.g. to keep it in the bottom-right corner)
'  w0.5         Sets the control's width to change by 0.5 that of the form's width change
'  w0.5l0.5     Sets the control's width and position to change by 0.5 that of the form's width change
Option Explicit
'Windows API calls to do all the dirty work!
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Const GWL_STYLE As Long = (-16)        'The offset of a window's style
Private Const WS_THICKFRAME As Long = &H40000        'Style to add a sizable frame
Private Const SW_SHOW As Long = 5
Dim moForm As Object
Dim mdWidth As Double
Dim mdHeight As Double

'Property to set the userform to be resizable
Public Property Set FORM(oNew As Object)
    Dim hWndForm As Long, iStyle As Long
    'Remember the form for later
    Set moForm = oNew
    'Get the userform's window handle
    If val(Application.Version) < 9 Then
        hWndForm = FindWindow("ThunderXFrame", moForm.Caption)        'XL97
    Else
        hWndForm = FindWindow("ThunderDFrame", moForm.Caption)        'XL2000
    End If
    'Make the form resizable
    iStyle = GetWindowLong(hWndForm, GWL_STYLE)
    iStyle = iStyle Or WS_THICKFRAME
    SetWindowLong hWndForm, GWL_STYLE, iStyle
    'Show the window with the changes
    ShowWindow hWndForm, SW_SHOW
    DrawMenuBar hWndForm
    SetFocus hWndForm
    'Remember the current size for later
    mdWidth = moForm.Width
    mdHeight = moForm.Height
End Property

'Handle the form's resize event, by resizing and repositioning controls
Public Sub FormResize()
    Dim dWidthAdj As Double, dHeightAdj As Double, sTag As String
    Dim oCtl As MSForms.Control
    'If not set before, remember the old width and height
    If mdWidth = 0 Then mdWidth = moForm.Width
    If mdHeight = 0 Then mdHeight = moForm.Height
    'How much are we changing by?
    dWidthAdj = moForm.Width - mdWidth
    dHeightAdj = moForm.Height - mdHeight
    'Check if we can perform the adjustment (i.e. widths and heights can't be -ve)
    For Each oCtl In moForm.Controls
        With oCtl
            sTag = UCase(.Tag)
            'Check if the left would become -ve
            If InStr(1, sTag, "L", vbBinaryCompare) Then
                If .left + dWidthAdj <= 0 Then moForm.Width = mdWidth
            End If
            'Check if the width would become -ve
            If InStr(1, sTag, "W", vbBinaryCompare) Then
                If .Width + dWidthAdj <= 0 Then moForm.Width = mdWidth
            End If
            'Check if the top would become -ve
            If InStr(1, sTag, "T", vbBinaryCompare) Then
                If .Top + dHeightAdj <= 0 Then moForm.Height = mdHeight
            End If
            'Check if the height would become -ve
            If InStr(1, sTag, "H", vbBinaryCompare) Then
                If .Height + dHeightAdj <= 0 Then moForm.Height = mdHeight
            End If
        End With
    Next
    'OK to do it, so perform the resize
    dWidthAdj = moForm.Width - mdWidth
    dHeightAdj = moForm.Height - mdHeight
    'Loop through the controls on the form, changing their size and/or position
    For Each oCtl In moForm.Controls
        With oCtl
            sTag = UCase(.Tag)
            If InStr(1, sTag, "L", vbBinaryCompare) Then .left = .left + dWidthAdj * ResizeFactor(sTag, "L")
            If InStr(1, sTag, "T", vbBinaryCompare) Then .Top = .Top + dHeightAdj * ResizeFactor(sTag, "T")
            If InStr(1, sTag, "W", vbBinaryCompare) Then .Width = .Width + dWidthAdj * ResizeFactor(sTag, "W")
            If InStr(1, sTag, "H", vbBinaryCompare) Then .Height = .Height + dHeightAdj * ResizeFactor(sTag, "H")
        End With
    Next
    'Remember the new dimensions for next time
    mdWidth = moForm.Width
    mdHeight = moForm.Height
End Sub

'Get the resize factor from the control's Tag property
Private Function ResizeFactor(sTag As String, sChange As String)
    Dim i As Integer, D As Double
    'Find the position of the required change designator (L, T, W or H)
    i = InStr(1, sTag, sChange, vbBinaryCompare)
    If i > 0 Then
        'Get the value of any numbers following the designator
        D = val(Mid$(sTag, i + 1))
        'If none there, change by 100%
        If D = 0 Then D = 1
    End If
    ResizeFactor = D
End Function

