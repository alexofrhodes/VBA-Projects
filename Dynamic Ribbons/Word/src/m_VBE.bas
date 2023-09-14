Attribute VB_Name = "m_VBE"
Option Explicit

Function ProceduresOfProject(TargetProject As vbProject) As Collection
    Dim Module As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim lineNum As Long
    Dim coll As New Collection
    Dim ProcedureName As String
    For Each Module In TargetProject.VBComponents
        If Module.Type = vbext_ct_StdModule Then
            With Module.CodeModule
                lineNum = .CountOfDeclarationLines + 1
                Do Until lineNum >= .CountOfLines
                    ProcedureName = .ProcOfLine(lineNum, ProcKind)
                    ' _ is used in events. Events may have the same name in different components
                    If InStr(1, ProcedureName, "_") = 0 Then
                        coll.Add ProcedureName
                    End If
                    lineNum = .ProcStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
                Loop
            End With
        End If
    Next Module
    Set ProceduresOfProject = coll
End Function

Function ProcedureExists( _
                        TargetProject As vbProject, _
                        ProcedureName As Variant) As Boolean
    Dim Procedures As Collection
    Set Procedures = ProceduresOfProject(TargetProject)
    Dim Procedure As Variant
    For Each Procedure In Procedures
        If UCase(CStr(Procedure)) = UCase(ProcedureName) Then
            ProcedureExists = True
            Exit Function
        End If
    Next
End Function


Function ListContains(ctrl As control, str As String, _
                        Optional ColumnIndexZeroBased As Long = -1, _
                        Optional CaseSensitive As Boolean = False) As Boolean
    Dim i      As Long
    Dim N      As Long
    Dim sTemp  As String
        If ColumnIndexZeroBased > ctrl.columnCount - 1 Or ColumnIndexZeroBased < 0 Then
            ColumnIndexZeroBased = -1
        End If
        N = ctrl.ListCount
        If ColumnIndexZeroBased <> -1 Then
            For i = N - 1 To 0 Step -1
                If CaseSensitive = True Then
                    sTemp = ctrl.List(i, ColumnIndexZeroBased)
                Else
                    str = LCase(str)
                    sTemp = LCase(ctrl.List(i, ColumnIndexZeroBased))
                End If
                If InStr(1, sTemp, str) > 0 Then
                    ListContains = True
                    Exit Function
                End If
            Next i
        Else
            Dim columnCount As Long
            N = ctrl.ListCount
            For i = N - 1 To 0 Step -1
                For columnCount = 0 To ctrl.columnCount - 1
                    If CaseSensitive = True Then
                        sTemp = ctrl.List(i, columnCount)
                    Else
                        str = LCase(str)
                        sTemp = LCase(ctrl.List(i, columnCount))
                    End If
                    If InStr(1, sTemp, str) > 0 Then
                        ListContains = True
                        Exit Function
                    End If
                Next columnCount
            Next i
        End If
End Function


