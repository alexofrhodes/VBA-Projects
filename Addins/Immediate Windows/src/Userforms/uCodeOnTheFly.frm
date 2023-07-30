VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uCodeOnTheFly 
   Caption         =   "Immediate Windows"
   ClientHeight    =   7332
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12900
   OleObjectBlob   =   "uCodeOnTheFly.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uCodeOnTheFly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1
Dim moResizer As New CFormResizer

Private Sub UserForm_Initialize()
    LoadUserformPosition Me
    startFrameForm Me
    For i = 1 To 5
        Me.Controls("Textbox" & i).Width = 800
        Me.Controls("Textbox" & i).Height = 500
        Me.Controls("Window" & i).Width = 500
    Next
    LoadUserformOptions Me
    Me.Height = 200
    Me.Width = 430
    
    Dim windowName As String
    windowName = ThisWorkbook.Sheets("uCodeOnTheFly_Settings").Range("D2").Value
    If windowName = "" Then windowName = "Window1"
    Me.Controls(windowName).visible = True
    Set TargetTextbox = Me.Controls("Textbox" & Right(windowName, 1))
    ThisWorkbook.Sheets("uCodeOnTheFly_Settings").Range("D1").Value = TargetTextbox.Name
    Me.Controls("Label" & Right(windowName, 1)).BackColor = 8435998
End Sub

Private Sub UserForm_Activate()
    Set moResizer.Form = Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveUserformOptions Me
    SaveUserformPosition Me
End Sub

Sub startFrameForm(Form As Object)
    Dim anc As MSForms.control
    Dim c As MSForms.control
    For Each c In Form.Controls
        If TypeName(c) = "Frame" Then
            'c.Caption = ""
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                c.visible = False
                If InStr(1, c.Tag, "anchor") > 0 Then
                    On Error Resume Next
                     Set anc = Me.Controls(Split(c.Tag)(0))
                     If anc Is Nothing Then Stop
                    On Error GoTo 0
                    c.top = anc.top 'Anchor01.Top
                    c.left = anc.left ' Anchor01.Left
                    Set anc = Nothing
                End If
            End If
        End If
    Next
    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll Me
End Sub

Private Sub CommandButton1_Click()
    SaveUserformOptions Me

    Dim code As String
    code = TextOfControl(TargetTextbox)
    If left(code, 1) <> "?" Then
        'if pointer width=0 then code=all text, else code=selection
        appRunOnTime Now, "ShowUserformCodeOnTheFly"
        RunCodeOnTheFly code
    Else
        Dim qAsked As Long
        qAsked = Len(code) - Len(replace(code, "?", ""))
        If qAsked > 1 Then
            MsgBox qAsked & " questions detected. I can only answer one."
            Exit Sub
        End If
        appRunOnTime Now, "ShowUserformCodeOnTheFly"
        appRunOnTime Now, _
                    "RunCodeOnTheFly", _
                    "targettextbox.text=targettextbox.text & vbNewLine & "" & ""  " & Mid(code, 2)
'        EvaluateQuestion Mid(code, 2)
    End If
End Sub

Private Sub info_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub


Private Sub UserForm_Resize()
    On Error Resume Next
    moResizer.FormResize
End Sub


Private Sub Emitter_LabelMouseOut(Label As MSForms.Label)
    If InStr(1, Label.Tag, "reframe", vbTextCompare) > 0 Then
        If Label.BackColor <> &H80B91E Then Label.BackColor = &H534848
    End If
End Sub

Private Sub Emitter_LabelMouseOver(Label As MSForms.Label)
    If InStr(1, Label.Tag, "reframe", vbTextCompare) > 0 Then
        If Label.BackColor <> &H80B91E Then Label.BackColor = &H808080
    End If
End Sub

Sub Emitter_LabelClick(ByRef Label As MSForms.Label)
    If InStr(1, Label.Tag, "reframe", vbTextCompare) > 0 Then
        Reframe Me, Label
        Set TargetTextbox = Me.Controls("Textbox" & Right(Label.Caption, 1))
        ThisWorkbook.Sheets("uCodeOnTheFly_Settings").Range("D1").Value = TargetTextbox.Name
        ThisWorkbook.Sheets("uCodeOnTheFly_Settings").Range("D2").Value = Label.Caption
    End If
End Sub

