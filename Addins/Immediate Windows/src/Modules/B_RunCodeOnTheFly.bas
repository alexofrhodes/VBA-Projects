Attribute VB_Name = "B_RunCodeOnTheFly"
Function NamelessCodeOnTheFly()
ActiveCell.Value = 5
ActiveCell.Offset(1).Value = 100
End Function


Function EvaluateQuestion(str As String)
'use in immediate window:
'
'   ?EvaluateQuestion("now")

    Dim var
    Dim code As String
'    code = "on error resume next" & vbnewlilne
    code = code & "ClearClipboard" & vbNewLine
    code = code & "dim var" & vbNewLine
    code = code & "var=" & str & vbNewLine
    code = code & "clip cstr(var)" & vbNewLine
    code = code & "namelesscodeonthefly=cstr(var)" & vbNewLine
    

'    code = code & "uCodeOnTheFly.Controls(ThisWorkbook.Sheets(""uCodeOnTheFly_Settings"").Range(""D1"").Value).text= _" & vbNewLine
'    code = code & "ThisWorkbook.Sheets(""uCodeOnTheFly_Settings"").columns(1).find( _" & vbNewLine
'    code = code & "ThisWorkbook.Sheets(""uCodeOnTheFly_Settings"").Range(""D1"").Value).offset(0,1).value & vbNewLine & cstr(var)"
'
    RunCodeOnTheFly code
    
    EvaluateQuestion = NamelessCodeOnTheFly
End Function

Function EvaluateString(strTextString As String)
    Application.Volatile
    EvaluateString = Application.Caller.parent.Evaluate(strTextString)
    Debug.Print strTextString & vbTab & ":" & vbTab & EvaluateString
End Function

Sub RunCodeFromRange()
'#INCLUDE RunCodeOnTheFly
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Columns.count <> 1 Then Exit Sub
    Dim code As String
    If Selection.Cells.count = 1 Then
        code = Selection.Value
    Else
        Dim var
        var = WorksheetFunction.Transpose(Selection.Value)
        code = Join(var, vbNewLine)
    End If
    RunCodeOnTheFly code
End Sub

Sub RunMacroFromSelection()
    Dim code As String
    code = CodepaneSelection
    If ProcedureExists(code, ActiveCodepaneWorkbook) Then
        Application.Run code
    Else
        RunCodeOnTheFly code
    End If
End Sub

Sub RunMacroFromClipboard()
    Dim code As String
    code = CLIP
    If ProcedureExists(code, ActiveCodepaneWorkbook) Then
        Application.Run code
    Else
        RunCodeOnTheFly code
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
'#INCLUDE UpdateProcedureCode
'#INCLUDE ProcedureEndLine
'#INCLUDE ModuleOfProcedure
'#INCLUDE appRunOnTime
'#INCLUDE NamelessCodeOnTheFly
    
    'The following are considered true
    '1. If the CodeOnTheFly you pass as an argument contains multiple macros,
    '   then the first macro is the main one, which calls the subsequent ones
    '2. No declarations (@TODO use a helper module to overcome this) or missing references are needed
    '3. Make sure your manually typed code is able to run, it's up to you
    
    On Error GoTo ErrorHandler
    CodeOnTheFly = replace(CodeOnTheFly, "Public", "Private")
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
        UpdateProcedureCode "NamelessCodeOnTheFly", _
                            "Function NamelessCodeOnTheFly()" & vbLf & _
                            CodeOnTheFly & vbLf & _
                            "End Function", _
                            ThisWorkbook
    End If
    
    If subName <> "NamelessCodeOnTheFly" Then
        Dim ProcEndLine As Long
        ProcEndLine = ProcedureEndLine(Module, "RunCodeOnTheFly", True)
        With Module.CodeModule
            .DeleteLines ProcEndLine + 1, .CountOfLines - ProcEndLine
            .InsertLines .CountOfLines + 1, vbNewLine & CodeOnTheFly
        End With
    End If
    appRunOnTime Now + TimeSerial(0, 0, 1), subName
    Exit Sub
ErrorHandler:
    MsgBox "An error occured"
End Sub

Sub Macro1()
MsgBox Macro2
End Sub

Function Macro2()
Macro2 = Now
End Function
