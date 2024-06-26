VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventListenerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem Author Robert Todar

Option Explicit
Option Compare Text

'SET FROM AddEventListener - NEEDED TO EMIT EVENT BACK TO IT.
Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

'CONTROLS THAT HAVE THE EVENTS
Private WithEvents Form As MSForms.UserForm
Attribute Form.VB_VarHelpID = -1
Private WithEvents Textbox As MSForms.Textbox
Attribute Textbox.VB_VarHelpID = -1
Private WithEvents Label As MSForms.Label
Attribute Label.VB_VarHelpID = -1
Private WithEvents CommandButton As MSForms.CommandButton
Attribute CommandButton.VB_VarHelpID = -1
Private WithEvents ComboBox As MSForms.ComboBox
Attribute ComboBox.VB_VarHelpID = -1
Private WithEvents Frame As MSForms.Frame
Attribute Frame.VB_VarHelpID = -1

Private Type state
    control As Object
    IsHoveredControl As Boolean
    IsFocusedControl As Boolean
End Type

Private This As state

'TODO
' - ADD EVENTS FOR ALL USERFORM CONTROLS
' - ADD MORE EVENTS THAN THE LIST BELOW
' - DOUBLE CHECK THAT EMITTS ARE ALL CURRENT

'CURRENT EVENTS (WILL ADD MORE AS I HAVE NEED)
' - MouseOver
' - MouseOut
' - MouseMove
' - Click
' - DblClick
' - Focus
' - Blur
' - KeyUp
' - KeyDown

' The Only public method. This will be called from EventListeneRemitter class module
Public Sub AddEventListener(ByRef ControlOrForm As Object, ByRef EmitterReference As EventListenerEmitter)
    ' Capture the emitter class. This will be used to emit events from each control.
    Set Emitter = EmitterReference

    ' This is used to compare and check to see if this is the control triggering the event.
    Set This.control = ControlOrForm

    ' Set control (or form) based on its type
    Select Case TypeName(ControlOrForm)
    Case "CommandButton"
        Set CommandButton = ControlOrForm

    Case "ComboBox"
        Set ComboBox = ControlOrForm

    Case "Frame"
        Set Frame = ControlOrForm

    Case "Label"
        Set Label = ControlOrForm

    Case "TextBox"
        Set Textbox = ControlOrForm

    Case Else

        If TypeOf ControlOrForm Is MSForms.UserForm Then
            Set Form = ControlOrForm
        Else
            Rem todo
            Rem err.Raise 5, TypeName(Me), "Invalid control: Currently unable to listen to events on " & TypeName(ControlOrForm)
        End If

    End Select
End Sub

'***********************************************************************************
' Helper functions
'***********************************************************************************

' Called on mousemove event. This is a way of creating a mouseover and mouseout event.
Private Sub CheckIfHoveredControl()
    If Not This.IsHoveredControl Then
        This.IsHoveredControl = True
        Emitter.EmitEvent This.control, MouseOver
    End If
End Sub

' Called on mousemove event. This is a way of creating a mouseover and mouseout event.
Private Sub CheckIfFocusedControl()
    If Not This.IsFocusedControl Then
        If TypeName(This.control) = "Frame" Then
            Emitter.SetFocusToControl This.control.ActiveControl
        Else
            This.IsFocusedControl = True
            Emitter.EmitEvent This.control, Focus
        End If
    End If
End Sub

' Simple Collection factory � for ease of use.
Private Function ToCollection(ParamArray keyValuePairs() As Variant) As Collection
    ' Check to see if there is even number of parameters
    Dim ArrayLenght As Long
    ArrayLenght = UBound(keyValuePairs) - LBound(keyValuePairs) + 1

    If ArrayLenght Mod 2 <> 0 Then
        Err.Raise 5, TypeName(Me), "Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
    End If

    Set ToCollection = New Collection
    Dim index As Long
    For index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        ToCollection.Add keyValuePairs(index + 1), keyValuePairs(index)
    Next index
End Function

'***********************************************************************************
' EVENTS
'***********************************************************************************

' ONCE AN EVENT HAS EMMITED, EACH EVENTListenerITEM WILL LISTEN FOR THAT EVENT
Private Sub Emitter_EmittedEvent(ByRef control As Object, ByVal EventName As EmittedEvent, ByRef EventParameters As Collection)
    ' CREATE A MOUSEOVER MOUSEOUT EVENTS
    On Error GoTo EH '<-- TODO proper error handling
    Select Case EventName
    Case MouseOver
        If This.control.Name <> control.Name And This.IsHoveredControl Then 'todo catastrophic error occurs with dynamic controls
            This.IsHoveredControl = False
            Emitter.EmitEvent This.control, MouseOut
        End If

    Case Focus
        If This.control.Name <> control.Name And This.IsFocusedControl Then
            This.IsFocusedControl = False
            Emitter.EmitEvent This.control, Blur
        ElseIf This.control.Name = control.Name And This.IsFocusedControl = False Then
            This.IsFocusedControl = True
        End If

    End Select
EH:
End Sub

'------------------------------------------------------------------------
' USERFORM
'------------------------------------------------------------------------
Private Sub Form_Click()
    Emitter.EmitEvent This.control, Click
End Sub

Private Sub Form_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Emitter.EmitEvent This.control, DoubleClick, ToCollection("Cancel", Cancel)
End Sub

Private Sub Form_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CheckIfHoveredControl
    Emitter.EmitEvent This.control, MouseMove, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub Form_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Emitter.EmitEvent This.control, MouseDown, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub Form_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Emitter.EmitEvent This.control, MouseUp, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

'------------------------------------------------------------------------
' COMMAND BUTTON
'------------------------------------------------------------------------
Private Sub CommandButton_Click()
    Emitter.EmitEvent This.control, Click
End Sub

Private Sub CommandButton_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Emitter.EmitEvent This.control, DoubleClick, ToCollection("Cancel", Cancel)
End Sub

Private Sub CommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CheckIfHoveredControl
    Emitter.EmitEvent This.control, MouseMove, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub CommandButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CheckIfFocusedControl
    Emitter.EmitEvent This.control, MouseUp, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub CommandButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Emitter.EmitEvent This.control, MouseDown, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub CommandButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckIfFocusedControl
    Emitter.EmitEvent This.control, KeyUp, ToCollection("KeyCode", KeyCode, "Shift", Shift)
End Sub

Private Sub CommandButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Emitter.EmitEvent This.control, KeyDown, ToCollection("KeyCode", KeyCode, "Shift", Shift)
End Sub

'------------------------------------------------------------------------
' LABEL
'------------------------------------------------------------------------
Private Sub Label_Click()
    Emitter.EmitEvent This.control, Click
End Sub

Private Sub Label_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Emitter.EmitEvent This.control, DoubleClick, ToCollection("Cancel", Cancel)
End Sub

Private Sub Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CheckIfHoveredControl
    Emitter.EmitEvent This.control, MouseMove, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub Label_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Emitter.EmitEvent This.control, MouseDown, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub Label_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Emitter.EmitEvent This.control, MouseUp, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

'------------------------------------------------------------------------
' Frame
'------------------------------------------------------------------------
Private Sub Frame_Click()
    Emitter.EmitEvent This.control, Click
End Sub

Private Sub Frame_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Emitter.EmitEvent This.control, DoubleClick, ToCollection("Cancel", Cancel)
End Sub

Private Sub Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CheckIfHoveredControl
    Emitter.EmitEvent This.control, MouseMove, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub Frame_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CheckIfFocusedControl        'FRAME DOESN'T TAKE FOCUS BUT ACTIVE CONTROL IN FRAME DOES
    Emitter.EmitEvent This.control, MouseDown
End Sub

Private Sub Frame_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Emitter.EmitEvent This.control, MouseUp, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

'------------------------------------------------------------------------
' Textbox
'------------------------------------------------------------------------
Private Sub Textbox_Click()
    Emitter.EmitEvent This.control, Click
End Sub

Private Sub Textbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Emitter.EmitEvent This.control, DoubleClick, ToCollection("Cancel", Cancel)
End Sub

Private Sub Textbox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CheckIfHoveredControl
    Emitter.EmitEvent This.control, MouseMove, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub Textbox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CheckIfFocusedControl
    Emitter.EmitEvent This.control, MouseUp, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub Textbox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Emitter.EmitEvent This.control, MouseDown, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub Textbox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    CheckIfFocusedControl
    Emitter.EmitEvent This.control, KeyUp, ToCollection("KeyCode", KeyCode, "Shift", Shift)
End Sub

Private Sub Textbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Emitter.EmitEvent This.control, KeyDown, ToCollection("KeyCode", KeyCode, "Shift", Shift)
End Sub

Private Sub TextBox_Change() 'added
    Emitter.EmitEvent This.control, Change
End Sub

'------------------------------------------------------------------------
' Combobox
'------------------------------------------------------------------------
Private Sub ComboBox_Click()
    Emitter.EmitEvent This.control, Click
End Sub

Private Sub ComboBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Emitter.EmitEvent This.control, DoubleClick, ToCollection("Cancel", Cancel)
End Sub

Private Sub ComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    CheckIfHoveredControl
    Emitter.EmitEvent This.control, MouseMove, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub

Private Sub ComboBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 'added
    Emitter.EmitEvent This.control, KeyDown, ToCollection("KeyCode", KeyCode, "Shift", Shift)
End Sub

Private Sub ComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 'added
    CheckIfFocusedControl
    Emitter.EmitEvent This.control, KeyUp, ToCollection("KeyCode", KeyCode, "Shift", Shift)
End Sub

Private Sub ComboBox_Change() 'added
    Emitter.EmitEvent This.control, Change
End Sub

Private Sub ComboBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single) 'added
    CheckIfFocusedControl
    Emitter.EmitEvent This.control, MouseUp, ToCollection("Button", Button, "Shift", Shift, "X", x, "Y", y)
End Sub
