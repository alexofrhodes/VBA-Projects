Attribute VB_Name = "M_RunCodeOnTheFly"
Option Explicit
Function NamelessCodeOnTheFly()
uCodeOnTheFly.TextBox1.TEXT = uCodeOnTheFly.TextBox1.TEXT & vbNewLine & ThisWorkbook.path

End Function

Public Sub ShowUserformCodeOnTheFly()
    '@AssignedModule M_RunCodeOnTheFly
    '@INCLUDE USERFORM uCodeOnTheFly
    uCodeOnTheFly.Show
End Sub
Function EvaluateQuestion(str As String)
    'use in immediate window:
    '
    '   ?EvaluateQuestion("now")
    '@AssignedModule M_RunCodeOnTheFly
    '@INCLUDE PROCEDURE ClearClipboard
    '@INCLUDE PROCEDURE NamelessCodeOnTheFly
    '@INCLUDE PROCEDURE RunCodeOnTheFly
    '@INCLUDE USERFORM uCodeOnTheFly

    Dim var
    Dim Code        As String
    '    code = "on error resume next" & vbnewlilne
    Code = Code & "ClearClipboard" & vbNewLine
    Code = Code & "dim var" & vbNewLine
    Code = Code & "var=" & str & vbNewLine
    Code = Code & "clip cstr(var)" & vbNewLine
    Code = Code & "namelesscodeonthefly=cstr(var)" & vbNewLine

    '    code = code & "uCodeOnTheFly.Controls(ThisWorkbook.Sheets(""uCodeOnTheFly_Settings"").Range(""D1"").Value).text= _" & vbNewLine
    '    code = code & "ThisWorkbook.Sheets(""uCodeOnTheFly_Settings"").columns(1).find( _" & vbNewLine
    '    code = code & "ThisWorkbook.Sheets(""uCodeOnTheFly_Settings"").Range(""D1"").Value).offset(0,1).value & vbNewLine & cstr(var)"
    '
    RunCodeOnTheFly Code

    EvaluateQuestion = NamelessCodeOnTheFly
End Function

Function EvaluateString(strTextString As String)
    '@AssignedModule M_RunCodeOnTheFly
    Application.Volatile
    EvaluateString = Application.Caller.Parent.Evaluate(strTextString)
    Debug.Print strTextString & vbTab & ":" & vbTab & EvaluateString
End Function

Sub RunCodeFromRange()
    '@INCLUDE RunCodeOnTheFly
    '@AssignedModule M_RunCodeOnTheFly
    '@INCLUDE PROCEDURE RunCodeOnTheFly
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Columns.Count <> 1 Then Exit Sub
    Dim Code        As String
    If Selection.Cells.Count = 1 Then
        Code = Selection.Value
    Else
        Dim var
        var = WorksheetFunction.Transpose(Selection.Value)
        Code = Join(var, vbNewLine)
    End If
    RunCodeOnTheFly Code
End Sub

Sub RunMacroFromSelection()
    '@AssignedModule M_RunCodeOnTheFly
    '@INCLUDE PROCEDURE ActiveModule
    '@INCLUDE PROCEDURE ActiveCodepaneWorkbook
    '@INCLUDE PROCEDURE ProcedureExists
    '@INCLUDE PROCEDURE RunCodeOnTheFly
    '@INCLUDE CLASS aCodeModule
    Dim Code        As String
    Code = aCodeModule.Init(ActiveModule).CodemoduleSelection
    If ProcedureExists(ActiveCodepaneWorkbook, Code) Then
        Application.Run Code
    Else
        RunCodeOnTheFly Code
    End If
End Sub

Sub RunMacroFromClipboard()
    '@AssignedModule M_RunCodeOnTheFly
    '@INCLUDE PROCEDURE ActiveCodepaneWorkbook
    '@INCLUDE PROCEDURE ProcedureExists
    '@INCLUDE PROCEDURE CLIP
    '@INCLUDE PROCEDURE RunCodeOnTheFly
    Dim Code        As String
    Code = CLIP
    If ProcedureExists(ActiveCodepaneWorkbook, Code) Then
        Application.Run Code
    Else
        RunCodeOnTheFly Code
    End If
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 18-11-2022 18:22    Alex                (RunCodeOnTheFly) Initial Release

Public Sub RunCodeOnTheFly(CodeOnTheFly As String)
    Rem Do not move this procedure !!!
    Rem All lines after this procedure will be deleted and replaced.
    '@AssignedModule M_RunCodeOnTheFly
    '@INCLUDE PROCEDURE ModuleOfProcedure
    '@INCLUDE PROCEDURE ProcedureLinesLast
    '@INCLUDE PROCEDURE ProcedureReplace
    '@INCLUDE PROCEDURE NamelessCodeOnTheFly
    '@INCLUDE PROCEDURE appRunOnTime
    '@INCLUDE CLASS aProcedure

    'The following are considered true
    '1. If the CodeOnTheFly you pass as an argument contains multiple macros,
    '   then the first macro is the main one, which calls the subsequent ones
    '2. No declarations (@TODO use a helper module to overcome this) or missing references are needed
    '3. Make sure your manually typed code is able to run, it's up to you

    On Error GoTo ErrorHandler
    CodeOnTheFly = Replace(CodeOnTheFly, "Public", "Private")
    Dim module      As VBComponent
    Set module = ModuleOfProcedure(ThisWorkbook, "RunCodeOnTheFly")

    Dim subName     As String
    Dim SubStart    As Long
    SubStart = InStr(1, CodeOnTheFly, "Sub ", vbTextCompare)
    Dim FunctionStart As Long
    FunctionStart = InStr(1, CodeOnTheFly, "Function ", vbTextCompare)
    If SubStart > 0 Or FunctionStart > 0 Then
        If (SubStart > 0 And SubStart < FunctionStart) Or _
                (SubStart > 0 And FunctionStart = 0) Then
            subName = Mid(CodeOnTheFly, SubStart)
            subName = Split(subName, "Sub ", , vbTextCompare)(1)
            subName = Split(subName, "(")(0)
        ElseIf FunctionStart > 0 And FunctionStart < SubStart Or _
                (SubStart = 0 And FunctionStart > 0) Then
            subName = Mid(CodeOnTheFly, FunctionStart)
            subName = Split(subName, "Function ", , vbTextCompare)(1)
            subName = Split(subName, "(")(0)
        Else
            Stop
        End If
    Else
        subName = "NamelessCodeOnTheFly"
        ProcedureReplace module, "NamelessCodeOnTheFly", _
                "Function NamelessCodeOnTheFly()" & vbLf & _
                CodeOnTheFly & vbLf & _
                "End Function"
    End If

    If subName <> "NamelessCodeOnTheFly" Then
        Dim procEndLine As Long
        procEndLine = aProcedure.Init(ThisWorkbook, module, "RunCodeOnTheFly").Line_Last
        With module.CodeModule
            .DeleteLines procEndLine + 1, .CountOfLines - procEndLine
            .InsertLines .CountOfLines + 1, vbNewLine & CodeOnTheFly
        End With
    End If
    appRunOnTime Now + TimeSerial(0, 0, 1), subName
    Exit Sub
ErrorHandler:
    MsgBox "An error occured"
End Sub

