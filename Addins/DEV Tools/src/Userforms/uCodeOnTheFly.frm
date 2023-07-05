VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uCodeOnTheFly 
   Caption         =   "Immediate Windows"
   ClientHeight    =   7332
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12900
   OleObjectBlob   =   "uCodeOnTheFly.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uCodeOnTheFly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm   : uCodeOnTheFly
'* Purpose    :
'* Copyright  :
'*
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 30-06-2023 11:53    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



Option Explicit

Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Private Sub GetInfo_Click()
    uAuthor.Show
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    startFrameForm Me
    For i = 1 To 5
        Me.Controls("Textbox" & i).Width = 800
        Me.Controls("Textbox" & i).Height = 1500
        Me.Controls("Window" & i).Width = 1500
        Me.Controls("Window" & i).Height = 1500
    Next
    
    aUserform.Init(Me).LoadOptions
    Me.Height = 200
    Me.Width = 430

    Dim configFolder As String: configFolder = ThisWorkbook.Path & "\configurations\"
    FoldersCreate configFolder
    Dim iniFile As String: iniFile = configFolder & "UserformSettings.ini"
    If Not FileExists(iniFile) Then TxtOverwrite iniFile, ""
    
    Dim windowName As String
'    windowName = ThisWorkbook.Sheets("uCodeOnTheFly_Settings").Range("D2").Value
    windowName = IniReadKey(iniFile, Me.Name, "Window")
    If windowName = "" Then windowName = "Window1"
    
    Me.Controls(windowName).Visible = True
    Set CodeOnTheFlyTextbox = Me.Controls("Textbox" & Right(windowName, 1))
'    ThisWorkbook.Sheets("uCodeOnTheFly_Settings").Range("D1").Value = CodeOnTheFlyTextbox.Name
        IniWrite iniFile, Me.Name, "Textbox", CodeOnTheFlyTextbox.Name
        
    Me.Controls("Label" & Right(windowName, 1)).BackColor = 8435998
End Sub

Private Sub UserForm_Activate()
    Dim myForm As New aUserform
    myForm.Init(Me).Resizable
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    aUserform.Init(Me).SaveOptions
End Sub

Private Sub startFrameForm(Form As Object)
    Dim anc As MSForms.control
    Dim c As MSForms.control
    For Each c In Form.Controls
        If TypeName(c) = "Frame" Then
            'c.Caption = ""
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                c.Visible = False
                If InStr(1, c.Tag, "anchor") > 0 Then
                    On Error Resume Next
                     Set anc = Me.Controls(Split(c.Tag)(0))
                     If anc Is Nothing Then Stop
                    On Error GoTo 0
                    c.Top = anc.Top 'Anchor01.Top
                    c.Left = anc.Left ' Anchor01.Left
                    Set anc = Nothing
                End If
            End If
        End If
    Next
    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll Me
End Sub

Private Sub CommandButton1_Click()
    aUserform.Init(Me).SaveOptions

    Dim Code As String
    Code = TextOfControl(CodeOnTheFlyTextbox)
    If Left(Code, 1) <> "?" Then
        'if pointer width=0 then code=all text, else code=selection
        appRunOnTime Now, "ShowUserformCodeOnTheFly"
        RunCodeOnTheFly Code
    Else
        Dim qAsked As Long
        qAsked = Len(Code) - Len(Replace(Code, "?", ""))
        If qAsked > 1 Then
            MsgBox qAsked & " questions detected. I can only answer one."
            Exit Sub
        End If
        appRunOnTime Now, "ShowUserformCodeOnTheFly"
        appRunOnTime Now, _
                    "RunCodeOnTheFly", _
                    "CodeOnTheFlyTextbox.text=CodeOnTheFlyTextbox.text & vbNewLine & "" & ""  " & Mid(Code, 2)
'        EvaluateQuestion Mid(code, 2)
    End If
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

Private Sub Emitter_LabelClick(ByRef Label As MSForms.Label)
    If InStr(1, Label.Tag, "reframe", vbTextCompare) > 0 Then
        Reframe Me, Label
        Set CodeOnTheFlyTextbox = Me.Controls("Textbox" & Right(Label.Caption, 1))
        
        Dim configFolder As String: configFolder = ThisWorkbook.Path & "\configurations\"
        FoldersCreate configFolder
        Dim iniFile As String: iniFile = configFolder & "UserformSettings.ini"
        If Not FileExists(iniFile) Then TxtOverwrite iniFile, ""
        IniWrite iniFile, Me.Name, "Window", Label.Caption
        IniWrite iniFile, Me.Name, "Textbox", CodeOnTheFlyTextbox.Name
        
'        ThisWorkbook.Sheets("uCodeOnTheFly_Settings").Range("D1").Value = CodeOnTheFlyTextbox.Name
'        ThisWorkbook.Sheets("uCodeOnTheFly_Settings").Range("D2").Value = Label.Caption
    End If
End Sub

