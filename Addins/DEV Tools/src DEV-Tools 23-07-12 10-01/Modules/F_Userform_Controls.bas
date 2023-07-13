Attribute VB_Name = "F_Userform_Controls"
Option Explicit

Function CreateOrSetFrame(Form As Object, Optional FrameName As String, Optional LTWH As Variant) As MSForms.Frame
    Dim cFrame As MSForms.Frame
    On Error Resume Next
    Set cFrame = Form.Controls(FrameName)
    On Error GoTo 0
    If cFrame Is Nothing Then
        If TypeName(Form) = "VBComponent" Then
            Set cFrame = Form.Designer.Controls.Add("Forms.Frame.1")
        Else
            Set cFrame = Form.Controls.Add("Forms.Frame.1")
        End If
    End If
    If Not IsMissing(FrameName) Then cFrame.Name = FrameName
    If Not IsMissing(LTWH) Then
        cFrame.Left = LTWH(0)
        cFrame.Top = LTWH(1)
        cFrame.Width = LTWH(2)
        cFrame.Height = LTWH(3)
    End If
    Set CreateOrSetFrame = cFrame
End Function

Function AvailableFormOrFrameRow(FormOrFrame As Object, Optional AfterWidth As Long = 0, Optional AfterHeight As Long = 0, Optional AddMargin As Long = 0) As Long
    Dim ctr As MSForms.control
    Dim myHeight
    For Each ctr In FormOrFrame.Controls
        If ctr.Visible = True Then
            If ctr.Left >= AfterWidth And ctr.Top >= AfterHeight Then
                If ctr.Top + ctr.Height > myHeight Then myHeight = ctr.Top + ctr.Height
            End If
        End If
    Next
    AvailableFormOrFrameRow = myHeight + AddMargin '6
End Function

Function AvailableFormOrFrameColumn(FormOrFrame As Object, Optional AfterWidth As Long = 0, Optional AfterHeight As Long = 0, Optional AddMargin As Long = 0) As Long
    Dim ctr As MSForms.control
    Dim myWidth
    For Each ctr In FormOrFrame.Controls
        If ctr.Visible = True Then
            If ctr.Left >= AfterWidth And ctr.Top >= AfterHeight Then
                If ctr.Left + ctr.Width > myWidth Then myWidth = ctr.Left + ctr.Width
            End If
        End If
    Next
    AvailableFormOrFrameColumn = myWidth + AddMargin '6
End Function

Sub AddFormControls(controlID As String, _
                    CountOrArrayOfNames As Variant, _
                    Optional Captions As Variant = 0, _
                    Optional Vertical As Boolean = True, _
                    Optional OFFSET As Long = 0, _
                    Optional Form As Object)
    If IsNumeric(CountOrArrayOfNames) And IsArray(Captions) Then
        If UBound(Captions) + 1 <> CLng(CountOrArrayOfNames) Then Exit Sub
    End If
    Dim Module As VBComponent
    If Form Is Nothing Then
        Set Module = ActiveModule
        If Module.Type <> vbext_ct_MSForm Then Exit Sub
    End If
    Dim c As MSForms.control
    Dim i As Long
    If IsNumeric(CountOrArrayOfNames) Then
        For i = 1 To CLng(CountOrArrayOfNames)
            If Form Is Nothing Then
                Set c = Module.Designer.Controls.Add(controlID)
            Else
                Set c = Form.Controls.Add(controlID)
            End If
            If Vertical Then
                c.Top = i * c.Height + i * 5 - c.Height
                c.Left = OFFSET
            Else
                c.Top = OFFSET
                c.Left = i * c.Width + i * 5 - c.Width
            End If
            If IsArray(Captions) Then
                c.Caption = Captions(i - 1)
            Else
                On Error Resume Next
                c.Caption = CountOrArrayOfNames(i - 1)
                If c.Caption = "" Then c.Caption = c.Name
                On Error GoTo 0
            End If
        Next
    Else
        For i = 1 To UBound(CountOrArrayOfNames) + 1
            If Form Is Nothing Then
                Set c = Module.Designer.Controls.Add(controlID)
            Else
                Set c = Form.Controls.Add(controlID)
            End If
            If Vertical Then
                c.Top = i * c.Height + i * 5 - c.Height
                c.Left = OFFSET
            Else
                c.Top = OFFSET
                c.Left = i * c.Width + i * 5 - c.Width
            End If
            c.Name = CountOrArrayOfNames(i - 1)
            If IsArray(Captions) Then
                c.Caption = Captions(i - 1)
            Else
                On Error Resume Next
                c.Caption = CountOrArrayOfNames(i - 1)
                If c.Caption = "" Then c.Caption = c.Name
                On Error GoTo 0
            End If
        Next
    End If
End Sub

Sub AddMultipleControls(ControlTypes As Variant, count As Long, Optional Vertical As Boolean = True, Optional Form As Object = Nothing)
    Dim i As Long
    For i = 1 To UBound(ControlTypes) + 1
        If Vertical Then
            AddFormControls CStr(ControlTypes(i - 1)), count, , Vertical, i * 60 - 50, Form
        Else
            AddFormControls CStr(ControlTypes(i - 1)), count, , Vertical, i * 20 - 20, Form
        End If
    Next
    Dim c As MSForms.control
    On Error Resume Next
    If Form Is Nothing Then
        For Each c In ActiveModule.Designer.Controls
            If Not TypeName(c) Like "TextBox" Then c.AutoSize = True
        Next
    Else
        For Each c In Form.Controls
            If Not TypeName(c) Like "TextBox" Then c.AutoSize = True
        Next
    End If
End Sub


Public Sub flashControl(ctr As MSForms.control, blinkCount As Integer)
    Rem if blinkCount = odd then the control will become hidden
    Dim lngTime As Long
    Dim i As Integer
    If blinkCount Mod 2 <> 0 Then blinkCount = blinkCount + 1
    For i = 1 To blinkCount * 2
        lngTime = getTickCount
        If ctr.Visible = True Then
            ctr.Visible = False
        Else
            ctr.Visible = True
        End If
        DoEvents
        Do While getTickCount - lngTime < 200
        Loop
    Next
End Sub

Public Function TextOfControl(c As control) As Variant
    Rem Text of Textbox, Selection of Combobox, Selected items (2d) of Listbox
    Dim out As New Collection
    If TypeName(c) = "TextBox" Then
        If c.SelLength = 0 Then
            TextOfControl = c.TEXT
        Else
            TextOfControl = c.SelText
        End If
    ElseIf TypeName(c) = "ComboBox" Then
        If c.Style < 2 Then
            TextOfControl = c.TEXT
        Else
            TextOfControl = ""
        End If
    ElseIf TypeName(c) = "ListBox" Then
        Set out = aListBox.Init(c).SelectedValues
        If out.count > 0 Then
            TextOfControl = aCollection.Init(out).ToArray
        Else
            TextOfControl = ""
        End If
    End If
End Function
