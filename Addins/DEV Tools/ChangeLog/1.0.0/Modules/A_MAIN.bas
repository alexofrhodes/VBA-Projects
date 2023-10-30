Attribute VB_Name = "A_MAIN"

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : A_MAIN
'* Purpose    : These macros use the new classes and target ACTIVE procedure/module/designer
'*              There are plenty more procedures in the classes, have a look. Give feedback
'* Copyright  :
'*
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* BLOG       : https://alexofrhodes.github.io/
'* GITHUB     : https://github.com/alexofrhodes/
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'*
'* Modified   : Date and Time       Author              Description
'* Created    : 22-08-2023 14:04    Alex
'* Modified   : 22-08-2023 14:04    Alex
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

'* Modified   : Date and Time       Author              Description
'* Updated    : 25-08-2023 12:23    Alex                (A_MAIN.bas > RunVbaGui) store process ID

Public Sub RunVbaGui()
'@LastModified 2308251223
'@INCLUDE DECLARATION LongPtr
'@INCLUDE PROCEDURE FileExists
'@INCLUDE PROCEDURE IniWrite
    Dim strProgramName As String
    strProgramName = ThisWorkbook.path & "\AHK\vbaGUI.exe"
    If Not FileExists(strProgramName) Then Exit Sub
    Dim hProcess As LongPtr
    hProcess = Shell("""" & strProgramName & """, vbNormalFocus)")
    IniWrite ThisWorkbook.path & "\AHK\process.ini", "ProgID", "vbaGUI", CStr(hProcess) ' Convert to string
End Sub

'* Modified   : Date and Time       Author              Description
'* Updated    : 25-08-2023 12:57    Alex                (A_MAIN.bas > KillVbaGui)

Public Sub KillVbaGui()
'@LastModified 2308251257
'@INCLUDE PROCEDURE IniReadKey
'@INCLUDE PROCEDURE KillProcess
    Dim strProcessId As String
    strProcessId = IniReadKey(ThisWorkbook.path & "\AHK\process.ini", "ProgID", "vbaGUI")
    Dim processId As Long
    If IsNumeric(strProcessId) Then
        processId = CLng(strProcessId)
        If KillProcess(processId) Then
'            MsgBox "Process terminated successfully."
        Else
'            MsgBox "Failed to terminate the process."
        End If
    End If
End Sub

'____ Codemodule Formatting _____

Public Sub cm_ActivateProcedure():      aCodeModule.Active.ActivateProcedure:              End Sub
Public Sub cm_AlignAs():                aCodeModule.Active.AlignAs:                        End Sub
Public Sub cm_AlignColumn():            aCodeModule.Active.AlignColumn:                    End Sub
Public Sub cm_AlignComments():          aCodeModule.Active.AlignComments:                  End Sub
Public Sub cm_AssignEnumValues():       aCodeModule.Active.AssignEnumValues:               End Sub
Public Sub cm_BeautifyFunction():       aCodeModule.Active.BeautifyFunction:               End Sub
Public Sub cm_BringProcedureHere():     aCodeModule.Active.BringProcedureHere:             End Sub
Public Sub cm_CaseLower():              aCodeModule.Active.CaseLower:                      End Sub
Public Sub cm_CaseProper():             aCodeModule.Active.CaseProper:                     End Sub
Public Sub cm_CaseUpper():              aCodeModule.Active.CaseUpper:                      End Sub
Public Sub cm_Comment():                aCodeModule.Active.Comment:                        End Sub
Public Sub cm_Copy():                   aCodeModule.Active.Copy:                           End Sub
Public Sub cm_Cut():                    aCodeModule.Active.Cut:                            End Sub
Public Sub cm_DimMerge():               aCodeModule.Active.DimMerge:                       End Sub
Public Sub cm_DimSeparate():            aCodeModule.Active.DimSeparate:                    End Sub
Public Sub cm_Duplicate():              aCodeModule.Active.Duplicate:                      End Sub
Public Sub cm_EncapsulateParenthesis(): aCodeModule.Active.Encapsulate_Parenthesis:        End Sub
Public Sub cm_EncapsulateQuotes():      aCodeModule.Active.Encapsulate_Quotes:             End Sub
Public Sub cm_EnumToCase():             aCodeModule.Active.EnumToCase:                     End Sub
Public Sub cm_FindCode():               FindCode aCodeModule.Active.CodemoduleSelection:   End Sub
Public Sub cm_FoldLine():               aCodeModule.Active.FoldLine:                       End Sub
Public Sub cm_FormatVBA7():             aCodeModule.Active.Format_VBA7:                    End Sub
Public Sub cm_ImportProcedure():        aCodeModule.Active.ImportProcedure:                End Sub
Public Sub cm_Increment():              aCodeModule.Active.Increment:                      End Sub
Public Sub cm_InjectArgumentStyle():    aCodeModule.Active.InjectArgumentStyleFolded:      End Sub
Public Sub cm_injectDivider():          aCodeModule.Active.injectDivider:                  End Sub
Public Sub cm_MoveDown():               aCodeModule.Active.Move_Down:                      End Sub
Public Sub cm_MoveUp():                 aCodeModule.Active.Move_Up:                        End Sub
Public Sub cm_PrintLinesLike():         PrintLinesContaining aCodeModule.Active.CodemoduleSelection: End Sub
Public Sub cm_RemAdd():                 aCodeModule.Active.RemAdd:                         End Sub
Public Sub cm_RemRemove():              aCodeModule.Active.RemRemove:                      End Sub
Public Sub cm_RotateCommas():           aCodeModule.Active.RotateCommas:                   End Sub
Public Sub cm_SortComma():              aCodeModule.Active.Sort_Comma:                     End Sub
Public Sub cm_SortLines():              aCodeModule.Active.Sort_Lines:                     End Sub
Public Sub cm_ToDo():                   aCodeModule.Active.Todo:                           End Sub
Public Sub cm_ToggleComments():         aCodeModule.Active.ToggleComments:                 End Sub
Public Sub cm_UnFoldLine():             aCodeModule.Active.UnFoldLine:                     End Sub
Public Sub cm_Uncomment():              aCodeModule.Active.UnComment:                      End Sub

'____ PROCEDURE Ops _____

Public Sub ap_AddToLinkedTable():            aProcedure.Active.AddToLinkedTable:                                 End Sub
Public Sub ap_BringLinkedProceduresHere():   aProcedure.Active.BringLinkedProceduresHere:                        End Sub
Public Sub ap_CommentsRemove():              aProcedure.Active.CommentsRemove False, Body_Code:                  End Sub
Public Sub ap_CommentsToOwnLine():           aProcedure.Active.CommentsToOwnLine:                                End Sub
Public Sub ap_ConvertBlankLinesToDividers(): aProcedure.Active.ConvertBlankLinesToDividers:                      End Sub
Public Sub ap_CreateCaller():                On Error Resume Next: aProcedure.Active.CreateCaller InputBoxRange: End Sub
Public Sub ap_Export():                      aProcedure.Active.Export:                                           End Sub
Public Sub ap_ExportLinkedCode():            aProcedure.Active.ExportLinkedCode:                                 End Sub
Public Sub ap_FoldDeclaration():             aProcedure.Active.FoldDeclaration:                                  End Sub
Public Sub ap_ImportDependencies():          aProcedure.Active.ImportDependencies:                               End Sub
Public Sub ap_Indent():                      aProcedure.Active.Indent:                                           End Sub
Public Sub ap_InjectDescription():           aProcedure.Active.InjectDescription:                                End Sub
Public Sub ap_InjectLinkedLists():           aProcedure.Active.InjectLinkedLists:                                End Sub
Public Sub ap_InjectModification():          aProcedure.Active.InjectModification:                               End Sub
Public Sub ap_InjectObjectsRelease():        aProcedure.Active.InjectObjectsReleaseHere:                         End Sub
Public Sub ap_InjectTemplate():              aProcedure.Active.InjectTemplate:                                   End Sub
Public Sub ap_InjectTemplateObject():        aProcedure.Active.InjectTemplateObject:                             End Sub
Public Sub ap_InjectTimer():                 aProcedure.Active.InjectTimer:                                      End Sub
Public Sub ap_MoveDown():                    aProcedure.Active.MoveDown:                                         End Sub
Public Sub ap_MoveToAssignedModule():        aProcedure.Active.MoveToAssignedModule:                             End Sub
Public Sub ap_MoveToBottom():                aProcedure.Active.MoveToBottom:                                     End Sub
Public Sub ap_MoveToTop():                   aProcedure.Active.MoveToTop:                                        End Sub
Public Sub ap_MoveUp():                      aProcedure.Active.MoveUp:                                           End Sub
Public Sub ap_NumbersAdd():                  aProcedure.Active.NumbersAdd:                                       End Sub
Public Sub ap_NumbersRemove():               aProcedure.Active.NumbersRemove:                                    End Sub
Public Sub ap_PrintDims():                   aProcedure.Active.PrintDims:                                        End Sub
Public Sub ap_RemoveEmptyLines():            aProcedure.Active.RemoveEmptyLines:                                 End Sub
Public Sub ap_RemoveIncludeLines():          aProcedure.Active.RemoveIncludeLines:                               End Sub
Public Sub ap_TestCreate():                  aProcedure.Active.TestCreate:                                       End Sub
Public Sub ap_UnfoldDeclaration():           aProcedure.Active.UnfoldDeclaration:                                End Sub
Public Sub ap_Update():                      aProcedure.Active.Update:                                           End Sub

'____ MODULE    Ops _____
Public Sub am_CodeRemove():                  aModule.Active.CodeRemove:                  End Sub
Public Sub am_CommentsRemove():              aModule.Active.CommentsRemove:              End Sub
Public Sub am_CommentsToOwnLine():           aModule.Active.CommentsToOwnLine:           End Sub
Public Sub am_EnableDebugPrint():            aModule.Active.EnableDebugPrint:            End Sub
Public Sub am_DisableDebugPrint():           aModule.Active.DisableDebugPrint:           End Sub
Public Sub am_EnableStop():                  aModule.Active.EnableStop:                  End Sub
Public Sub am_DisableStop():                 aModule.Active.DisableStop:                 End Sub
Public Sub am_PredeclaredIdEnable():         aModule.Active.PredeclaredIDenable:         End Sub
Public Sub am_Duplicate():                   aModule.Active.Duplicate:                   End Sub
Public Sub am_Export():                      aModule.Active.Export PickFolder:           End Sub
Public Sub am_ExportProcedures():            aModule.Active.ExportProcedures PickFolder: End Sub
Public Sub am_HeaderAdd():                   aModule.Active.HeaderAdd:                   End Sub
Public Sub am_Indent():                      aModule.Active.Indent:                      End Sub
Public Sub am_ListProcedures():              aModule.Active.ListProcedures:              End Sub
Public Sub am_ListProceduresPublic():        aModule.Active.ListProceduresPublic:        End Sub
Public Sub am_PrintListOfInclude():          aModule.Active.PrintListOfInclude:          End Sub
Public Sub am_PrintTodoList():               aModule.Active.PrintTodoList:               End Sub
Public Sub am_ProcedureFoldDeclarations():   aModule.Active.ProcedureFoldDeclarations:   End Sub
Public Sub am_ProcedureScopePrivate():       aModule.Active.ProcedureScopePrivate:       End Sub
Public Sub am_ProcedureScopePublic():        aModule.Active.ProcedureScopePublic:        End Sub
Public Sub am_ProceduresNames():             dp aModule.Active.ProceduresNames:          End Sub
Public Sub am_RemoveEmptyLinesButLeaveOne(): aModule.Active.RemoveEmptyLinesButLeaveOne: End Sub
Public Sub am_RemoveEmptyLines():            aModule.Active.RemoveEmptyLines:            End Sub
Public Sub am_SortAZ():                      aModule.Active.ProcedureSortAZ:             End Sub
Public Sub am_SortByKind():                  aModule.Active.ProcedureSortByKind:         End Sub
Public Sub am_SortByScope():                 aModule.Active.ProcedureSortByScope:        End Sub
Public Sub am_UpdateProcedures():            aModule.Active.UpdateProcedures:            End Sub

'____ DESIGNER  Ops _____

Public Sub ad_CenterLabelCaption():            aDesigner.Active.CenterLabelCaption:            End Sub
Public Sub ad_CopyControlProperties():         aDesigner.Active.CopyControlProperties:         End Sub
Public Sub ad_PasteControlProperties():        aDesigner.Active.PasteControlProperties:        End Sub
Public Sub ad_RemoveCaption():                 aDesigner.Active.RemoveCaption:                 End Sub
Public Sub ad_RenameControlAndCode():          aDesigner.Active.RenameControlAndCode:          End Sub
Public Sub ad_ReplaceCommandButtonWithLabel(): aDesigner.Active.ReplaceCommandButtonWithLabel: End Sub
Public Sub ad_SetHandCursor():                 aDesigner.Active.SetHandCursor:                 End Sub
Public Sub ad_SetHandCursorToSubControls():    aDesigner.Active.SetHandCursorToSubControls:    End Sub
Public Sub ad_SortControlsHorizontally():      aDesigner.Active.SortControlsHorizontally:      End Sub
Public Sub ad_SortControlsVertically():        aDesigner.Active.SortControlsVertically:        End Sub
Public Sub ad_SwitchNames():                   aDesigner.Active.SwitchNames:                   End Sub
Public Sub ad_SwitchPositions():               aDesigner.Active.SwitchPositions:               End Sub
Public Sub ad_SideBySide():                    aModules.SideBySide ActiveModule.Name:          End Sub

'____ WORKBOOK  Ops _____
Public Sub aw_AddLinkedLists():                AddLinkedListsToActiveWorkbook:                 End Sub
'Public Sub aw_VersionInitial():                PushVersionInitial:                             End Sub
'Public Sub aw_VersionMajor():                  PushVersionMajor:                               End Sub
'Public Sub aw_VersionMinor():                  PushVersionMinor:                               End Sub
'Public Sub aw_VersionPatch():                  PushVersionPatch:                               End Sub

'____ USERFORMS ____

Public Sub uShow_CodeOnTheFly():     uCodeOnTheFly.Show:    End Sub
Public Sub uShow_ProjectExplorer():  uProjectExplorer.Show: End Sub
Public Sub uShow_ProjectManager():   uProjectManager.Show:  End Sub
Public Sub uShow_References():       uReferences.Show:      End Sub
Public Sub uShow_Skeleton():         uSkeleton.Show:        End Sub
Public Sub uShow_Changelog():        uChangeLog.Show:       End Sub


Public Sub uShow_SnippetsWorkbook()
    ShowInVBE = False
    uSnippets.Show
End Sub

Public Sub uShow_SnippetsVBE()
    ShowInVBE = True
    Application.VBE.MainWindow.Visible = True
    Application.VBE.MainWindow.SetFocus
    uSnippets.Show
End Sub

'____ TXT _____

Public Sub txt_SeparateProcedures(): CallSeparateProcedures: End Sub
Public Sub txt_TxtPrepend(): CallTxtPrependContainedProcedures: End Sub

