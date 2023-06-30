Attribute VB_Name = "M_RunCodeOnTheFly"
Option Explicit

Public CodeOnTheFlyTextbox As MSForms.Textbox
Function NamelessCodeOnTheFly()

End Function

Function EvaluateQuestion(str As String)
'use in immediate window:
'
'   ?EvaluateQuestion("now")

    Dim var
    Dim Code As String
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
    Application.Volatile
    EvaluateString = Application.Caller.Parent.Evaluate(strTextString)
    Debug.Print strTextString & vbTab & ":" & vbTab & EvaluateString
End Function

Sub RunCodeFromRange()
'@INCLUDE RunCodeOnTheFly
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Columns.count <> 1 Then Exit Sub
    Dim Code As String
    If Selection.Cells.count = 1 Then
        Code = Selection.Value
    Else
        Dim var
        var = WorksheetFunction.Transpose(Selection.Value)
        Code = Join(var, vbNewLine)
    End If
    RunCodeOnTheFly Code
End Sub

Sub RunMacroFromSelection()
    Dim Code As String
    Code = aCodeModule.Init(ActiveModule).Selection
    If ProcedureExists(ActiveCodepaneWorkbook, Code) Then
        Application.Run Code
    Else
        RunCodeOnTheFly Code
    End If
End Sub

Sub RunMacroFromClipboard()
    Dim Code As String
    Code = CLIP
    If ProcedureExists(ActiveCodepaneWorkbook, Code) Then
        Application.Run Code
    Else
        RunCodeOnTheFly Code
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''AUTHOR     Anastasiou Alex
''EMAIL      AnastasiouAlex@gmail.com
''GITHUB     https://github.com/AlexOfRhodes
''YOUTUBE    https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
''VK         https://vk.com/video/playlist/735281600_1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'* Modified   : Date and Time       Author              Description
'* Updated    : 18-11-2022 18:22    Alex                (RunCodeOnTheFly) Initial Release

Sub RunCodeOnTheFly(CodeOnTheFly As String)
    Rem Do not move this procedure !!!
    Rem All lines after this procedure will be deleted and replaced.
'@INCLUDE ProcedureReplace
'@INCLUDE ProcedureLinesLast
'@INCLUDE ModuleOfProcedure
'@INCLUDE appRunOnTime
'@INCLUDE NamelessCodeOnTheFly

    'The following are considered true
    '1. If the CodeOnTheFly you pass as an argument contains multiple macros,
    '   then the first macro is the main one, which calls the subsequent ones
    '2. No declarations (@TODO use a helper module to overcome this) or missing references are needed
    '3. Make sure your manually typed code is able to run, it's up to you

    On Error GoTo ErrorHandler
    CodeOnTheFly = Replace(CodeOnTheFly, "Public", "Private")
    Dim Module As VBComponent
    Set Module = ModuleOfProcedure(ThisWorkbook, "RunCodeOnTheFly")

    Dim subName As String
    Dim SubStart As Long
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
        ProcedureReplace Module, "NamelessCodeOnTheFly", _
                            "Function NamelessCodeOnTheFly()" & vbLf & _
                            CodeOnTheFly & vbLf & _
                            "End Function"
    End If

    If subName <> "NamelessCodeOnTheFly" Then
        Dim procEndLine As Long
        procEndLine = aProcedure.Init(ThisWorkbook, Module, "RunCodeOnTheFly").LineIndex(Procedure_Last)
        With Module.CodeModule
            .DeleteLines procEndLine + 1, .CountOfLines - procEndLine
            .InsertLines .CountOfLines + 1, vbNewLine & CodeOnTheFly
        End With
    End If
    appRunOnTime Now + TimeSerial(0, 0, 1), subName
    Exit Sub
ErrorHandler:
    MsgBox "An error occured"
End Sub

