VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVBECommandHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents EvtHandler As VBIDE.CommandBarEvents
Attribute EvtHandler.VB_VarHelpID = -1

Public Sub EvtHandler_Click( _
                           ByVal CommandBarControl As Object, _
                           handled As Boolean, _
                           CancelDefault As Boolean)
    '@AssignedModule CVBECommandHandler
    '@INCLUDE CLASS CVBECommandHandler
    On Error Resume Next
    Application.Run CommandBarControl.OnAction
    handled = True
    CancelDefault = True
End Sub


