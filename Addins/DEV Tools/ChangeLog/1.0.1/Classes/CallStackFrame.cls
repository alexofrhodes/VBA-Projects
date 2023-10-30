VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallStackFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'https://codereview.stackexchange.com/questions/135926/managing-a-programmatically-accessible-stack-trace

Option Explicit
Implements CallStackFrameInterface

Private Type ThisStackFrame
    ModuleName As String
    MemberName As String
    values As Collection
End Type

Private this As ThisStackFrame

Public Function Create(ByVal module As String, ByVal member As String, ByRef parameterValues() As Variant) As CallStackFrameInterface
Attribute Create.VB_Description = "Creates a new instance of an object representing a stack frame, i.e. a procedure call and its arguments."
    With New CallStackFrame
        .ModuleName = module
        .MemberName = member

        Dim index As Integer
        For index = LBound(parameterValues) To UBound(parameterValues)
            .AddParameterValue parameterValues(index)
        Next

        Set Create = .Self
    End With
End Function

Public Property Get Self() As CallStackFrameInterface
Attribute Self.VB_Description = "Gets a reference to this instance."
    Set Self = Me
End Property

Public Property Get ModuleName() As String
Attribute ModuleName.VB_Description = "Gets/sets the name of the module for this instance."
    ModuleName = this.ModuleName
End Property

Public Property Let ModuleName(ByVal Value As String)
    this.ModuleName = Value
End Property

Public Property Get MemberName() As String
Attribute MemberName.VB_Description = "Gets/sets the name of the member for this instance."
    MemberName = this.MemberName
End Property

Public Property Let MemberName(ByVal Value As String)
    this.MemberName = Value
End Property

Public Property Get ParameterValue(ByVal index As Integer) As Variant
Attribute ParameterValue.VB_Description = "Gets the value of the parameter at the specified index."
    ParameterValue = this.values(index)
End Property

Public Sub AddParameterValue(ByRef Value As Variant)
Attribute AddParameterValue.VB_Description = "Adds the specified parameter value to this instance."
    this.values.Add Value
End Sub

Private Sub Class_Initialize()
    Set this.values = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.values = Nothing
End Sub

Private Property Get CallStackFrameInterface_MemberName() As String
    CallStackFrameInterface_MemberName = this.MemberName
End Property

Private Property Get CallStackFrameInterface_ModuleName() As String
    CallStackFrameInterface_ModuleName = this.ModuleName
End Property

Private Property Get CallStackFrameInterface_ParameterValue(ByVal index As Integer) As Variant
    CallStackFrameInterface_ParameterValue = this.values(index)
End Property

Private Function CallStackFrameInterface_ToString() As String

    Dim result As String
    result = this.ModuleName & "." & this.MemberName & "("

    Dim index As Integer
    Dim Value As Variant
    For Each Value In this.values

        index = index + 1

        result = result & "{" & TypeName(Value) & ":"
        If IsObject(Value) Then
            result = result & ObjPtr(Value)
        ElseIf IsArray(Value) Then
            result = result & "[" & LBound(Value) & "-" & UBound(Value) & "]"
        ElseIf VarType(Value) = vbString Then
            result = result & Chr$(34) & Value & Chr$(34)
        Else
            result = result & CStr(Value)
        End If
        result = result & "}" & IIf(index = this.values.Count, vbNullString, ",")

    Next

    result = result & ")"
    CallStackFrameInterface_ToString = result

End Function
