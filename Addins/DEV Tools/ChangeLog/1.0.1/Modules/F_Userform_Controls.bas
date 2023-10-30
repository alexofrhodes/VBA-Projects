Attribute VB_Name = "F_Userform_Controls"
Option Explicit

Function CreateOrSetFrame( _
                         Form As Object, _
                         Optional FrameName As String, _
                         Optional LTWH As Variant) As MSForms.Frame
    '@AssignedModule F_Userform_Controls
    Dim cFrame      As MSForms.Frame
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

Function AvailableFormOrFrameRow( _
                                FormOrFrame As Object, _
                                Optional AfterWidth As Long = 0, _
                                Optional AfterHeight As Long = 0, _
                                Optional AddMargin As Long = 0) As Long
    '@LastModified 2307171805
    '@AssignedModule F_Userform_Controls
    Dim ctr         As MSForms.control
    Dim myHeight
    For Each ctr In FormOrFrame.Controls
        If ctr.Visible = True Then
            If ctr.Left >= AfterWidth And ctr.Top >= AfterHeight Then
                If ctr.Top + ctr.Height > myHeight Then myHeight = ctr.Top + ctr.Height
            End If
        End If
    Next
    AvailableFormOrFrameRow = myHeight + AddMargin    '6
End Function

Function AvailableFormOrFrameColumn( _
                                   FormOrFrame As Object, _
                                   Optional AfterWidth As Long = 0, _
                                   Optional AfterHeight As Long = 0, _
                                   Optional AddMargin As Long = 0) As Long
    '@LastModified 2307171805
    '@AssignedModule F_Userform_Controls
    Dim ctr         As MSForms.control
    Dim myWidth
    For Each ctr In FormOrFrame.Controls
        If ctr.Visible = True Then
            If ctr.Left >= AfterWidth And ctr.Top >= AfterHeight Then
                If ctr.Left + ctr.Width > myWidth Then myWidth = ctr.Left + ctr.Width
            End If
        End If
    Next
    AvailableFormOrFrameColumn = myWidth + AddMargin    '6
End Function

Sub AddFormControls( _
                   controlID As String, _
                   CountOrArrayOfNames As Variant, _
                   Optional Captions As Variant = 0, _
                   Optional Vertical As Boolean = True, _
                   Optional offset As Long = 0, _
                   Optional Form As Object)
    '@AssignedModule F_Userform_Controls
    '@INCLUDE PROCEDURE ActiveModule
    If IsNumeric(CountOrArrayOfNames) And IsArray(Captions) Then
        If UBound(Captions) + 1 <> CLng(CountOrArrayOfNames) Then Exit Sub
    End If
    Dim module      As VBComponent
    If Form Is Nothing Then
        Set module = ActiveModule
        If module.Type <> vbext_ct_MSForm Then Exit Sub
    End If
    Dim c           As MSForms.control
    Dim i           As Long
    If IsNumeric(CountOrArrayOfNames) Then
        For i = 1 To CLng(CountOrArrayOfNames)
            If Form Is Nothing Then
                Set c = module.Designer.Controls.Add(controlID)
            Else
                Set c = Form.Controls.Add(controlID)
            End If
            If Vertical Then
                c.Top = i * c.Height + i * 5 - c.Height
                c.Left = offset
            Else
                c.Top = offset
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
                Set c = module.Designer.Controls.Add(controlID)
            Else
                Set c = Form.Controls.Add(controlID)
            End If
            If Vertical Then
                c.Top = i * c.Height + i * 5 - c.Height
                c.Left = offset
            Else
                c.Top = offset
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

Sub AddMultipleControls( _
                       ControlTypes As Variant, _
                       Count As Long, _
                       Optional Vertical As Boolean = True, _
                       Optional Form As Object = Nothing)
    '@AssignedModule F_Userform_Controls
    '@INCLUDE PROCEDURE ActiveModule
    '@INCLUDE PROCEDURE AddFormControls
    Dim i           As Long
    For i = 1 To UBound(ControlTypes) + 1
        If Vertical Then
            AddFormControls CStr(ControlTypes(i - 1)), Count, , Vertical, i * 60 - 50, Form
        Else
            AddFormControls CStr(ControlTypes(i - 1)), Count, , Vertical, i * 20 - 20, Form
        End If
    Next
    Dim c           As MSForms.control
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
    '@AssignedModule F_Userform_Controls
    '@INCLUDE DECLARATION getTickCount
    Dim lngTime     As Long
    Dim i           As Integer
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
    '@AssignedModule F_Userform_Controls
    '@INCLUDE CLASS aCollection
    '@INCLUDE CLASS aListBox
    Dim out         As New Collection
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
        If out.Count > 0 Then
            TextOfControl = aCollection.Init(out).ToArray
        Else
            TextOfControl = ""
        End If
    End If
End Function
