VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallStackFrameInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'https://codereview.stackexchange.com/questions/135926/managing-a-programmatically-accessible-stack-trace

Option Explicit

Public Property Get ModuleName() As String
Attribute ModuleName.VB_Description = "Gets the name of the module for this instance."
End Property

Public Property Get MemberName() As String
Attribute MemberName.VB_Description = "Gets the name of the member for this instance."
End Property

Public Property Get ParameterValue(ByVal index As Integer) As Variant
Attribute ParameterValue.VB_Description = "Gets the value of the parameter at the specified index."
End Property

Public Function ToString() As String
Attribute ToString.VB_Description = "Returns a string representation of the member and its arguments."
End Function
