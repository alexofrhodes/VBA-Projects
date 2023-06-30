Attribute VB_Name = "A_MAIN"
'NOTE:
'     I'VE BEEN MOVING THINGS TO CLASSES. There be bugs.

'These macros use the new classes and target ACTIVE procedure/module/designer
'There are plenty more procedures in the classes, have a look
'Give feedback

'____ Code Module Formatting _____

Public Sub a_cm_AlignAs():                aCodeModule.Active.AlignAs:                        End Sub
Public Sub a_cm_AlignColumn():            aCodeModule.Active.AlignColumn:                    End Sub
Public Sub a_cm_AlignComments():          aCodeModule.Active.AlignComments:                  End Sub
Public Sub a_cm_AssignEnumValues():       aCodeModule.Active.AssignEnumValues:               End Sub
Public Sub a_cm_BeautifyFunction():       aCodeModule.Active.BeautifyFunction:               End Sub
Public Sub a_cm_BringProcedureHere():     aCodeModule.Active.BringProcedureHere:             End Sub
Public Sub a_cm_CaseLower():              aCodeModule.Active.CaseLower:                      End Sub
Public Sub a_cm_CaseProper():             aCodeModule.Active.CaseProper:                     End Sub
Public Sub a_cm_CaseUpper():              aCodeModule.Active.CaseUpper:                      End Sub
Public Sub a_cm_Comment():                aCodeModule.Active.Comment:                        End Sub
Public Sub a_cm_Copy():                   aCodeModule.Active.Copy:                           End Sub
Public Sub a_cm_Cut():                    aCodeModule.Active.Cut:                            End Sub
Public Sub a_cm_DimMerge():               aCodeModule.Active.DimMerge:                       End Sub
Public Sub a_cm_DimSeparate():            aCodeModule.Active.DimSeparate:                    End Sub
Public Sub a_cm_Duplicate():              aCodeModule.Active.Duplicate:                      End Sub
Public Sub a_cm_EncapsulateParenthesis(): aCodeModule.Active.Encapsulate_Parenthesis:        End Sub
Public Sub a_cm_EncapsulateQuotes():      aCodeModule.Active.Encapsulate_Quotes:             End Sub
Public Sub a_cm_EnumToCase():             aCodeModule.Active.EnumToCase:                     End Sub
Public Sub a_cm_FindCode():               FindCode aCodeModule.Active.Selection:             End Sub
Public Sub a_cm_FoldLine():               aCodeModule.Active.FoldLine:                       End Sub
Public Sub a_cm_FormatVBA7():             aCodeModule.Active.Format_VBA7:                    End Sub
Public Sub a_cm_ImportProcedure():        aCodeModule.Active.ImportProcedure:                End Sub
Public Sub a_cm_Increment():              aCodeModule.Active.Increment:                      End Sub
Public Sub a_cm_InjectArgumentStyle():    aCodeModule.Active.InjectArgumentStyle:            End Sub
Public Sub a_cm_MoveDown():               aCodeModule.Active.Move_Down:                      End Sub
Public Sub a_cm_MoveUp():                 aCodeModule.Active.Move_Up:                        End Sub
Public Sub a_cm_PrintLinesLike():         PrintLinesContaining aCodeModule.Active.Selection: End Sub
Public Sub a_cm_RemAdd():                 aCodeModule.Active.RemAdd:                         End Sub
Public Sub a_cm_RemRemove():              aCodeModule.Active.RemRemove:                      End Sub
Public Sub a_cm_RotateCommas():           aCodeModule.Active.RotateCommas:                   End Sub
Public Sub a_cm_SortComma():              aCodeModule.Active.Sort_Comma:                     End Sub
Public Sub a_cm_SortLines():              aCodeModule.Active.Sort_Lines:                     End Sub
Public Sub a_cm_ToDo():                   aCodeModule.Active.Todo:                           End Sub
Public Sub a_cm_ToggleComments():         aCodeModule.Active.ToggleComments:                 End Sub
Public Sub a_cm_UnFoldLine():             aCodeModule.Active.UnFoldLine:                     End Sub
Public Sub a_cm_Uncomment():              aCodeModule.Active.UnComment:                      End Sub
Public Sub a_cm_injectDivider():          aCodeModule.Active.injectDivider:                  End Sub


'____ PROCEDURE Ops _____

Public Sub a_ap_AddToLinkedTable():            aProcedure.Active.AddToLinkedTable:                                 End Sub
Public Sub a_ap_BringLinkedProceduresHere():   aProcedure.Active.BringLinkedProceduresHere:                        End Sub
Public Sub a_ap_CommentsRemove():              aProcedure.Active.CommentsRemove False, Body_Code:                  End Sub
Public Sub a_ap_CommentsToOwnLine():           aProcedure.Active.CommentsToOwnLine:                                End Sub
Public Sub a_ap_ConvertBlankLinesToDividers(): aProcedure.Active.ConvertBlankLinesToDividers:                      End Sub
Public Sub a_ap_CreateCaller():                On Error Resume Next: aProcedure.Active.CreateCaller InputBoxRange: End Sub
Public Sub a_ap_Export():                      aProcedure.Active.Export:                                           End Sub
Public Sub a_ap_ExportLinkedCode():            aProcedure.Active.ExportLinkedCode:                                 End Sub
Public Sub a_ap_FoldDeclaration():             aProcedure.Active.FoldDeclaration:                                  End Sub
Public Sub a_ap_ImportDependencies():          aProcedure.Active.ImportDependencies:                               End Sub
Public Sub a_ap_Indent():                      aProcedure.Active.Indent:                                           End Sub
Public Sub a_ap_InjectDescription():           aProcedure.Active.InjectDescription:                                End Sub
Public Sub a_ap_InjectLinkedLists():           aProcedure.Active.InjectLinkedLists:                                End Sub
Public Sub a_ap_InjectModification():          aProcedure.Active.InjectModification:                               End Sub
Public Sub a_ap_InjectObjectsRelease():        aProcedure.Active.InjectObjectsRelease:                             End Sub
Public Sub a_ap_InjectTimer():                 aProcedure.Active.InjectTimer:                                      End Sub
Public Sub a_ap_MoveDown():                    aProcedure.Active.MoveDown:                                         End Sub
Public Sub a_ap_MoveToAssignedModule():        aProcedure.Active.MoveToAssignedModule:                             End Sub
Public Sub a_ap_MoveToBottom():                aProcedure.Active.MoveToBottom:                                     End Sub
Public Sub a_ap_MoveToTop():                   aProcedure.Active.MoveToTop:                                        End Sub
Public Sub a_ap_MoveUp():                      aProcedure.Active.MoveUp:                                           End Sub
Public Sub a_ap_NumbersAdd():                  aProcedure.Active.NumbersAdd:                                       End Sub
Public Sub a_ap_NumbersRemove():               aProcedure.Active.NumbersRemove:                                    End Sub
Public Sub a_ap_PrintDims():                   aProcedure.Active.PrintDims:                                        End Sub
Public Sub a_ap_RemoveEmptyLines():            aProcedure.Active.RemoveEmptyLines:                                 End Sub
Public Sub a_ap_RemoveIncludeLines():          aProcedure.Active.RemoveIncludeLines:                               End Sub
Public Sub a_ap_ScopePrivate():                aProcedure.Active.ScopePrivate:                                     End Sub
Public Sub a_ap_ScopePublic():                 aProcedure.Active.ScopePublic:                                      End Sub
Public Sub a_ap_ScopeSuggested():              aProcedure.Active.ScopeSuggested:                                   End Sub
Public Sub a_ap_TestCreate():                  aProcedure.Active.TestCreate:                                       End Sub
Public Sub a_ap_UnfoldDeclaration():           aProcedure.Active.UnfoldDeclaration:                                End Sub
Public Sub a_ap_Update():                      aProcedure.Active.Update:                                           End Sub

'____ MODULE    Ops _____
Sub a_am_CodeRemove():                  aModule.Active.CodeRemove:                  End Sub
Sub a_am_CommentsRemove():              aModule.Active.CommentsRemove:              End Sub
Sub a_am_CommentsToOwnLine():           aModule.Active.CommentsToOwnLine:           End Sub
Sub a_am_DisableDebugPrint():           aModule.Active.DisableDebugPrint:           End Sub
Sub a_am_DisableStop():                 aModule.Active.DisableStop:                 End Sub
Sub a_am_Duplicate():                   aModule.Active.Duplicate:                   End Sub
Sub a_am_EnableDebugPrint():            aModule.Active.EnableDebugPrint:            End Sub
Sub a_am_EnableStop():                  aModule.Active.EnableStop:                  End Sub
Sub a_am_Export():                      aModule.Active.Export PickFolder:           End Sub
Sub a_am_ExportProcedures():            aModule.Active.ExportProcedures PickFolder: End Sub
Sub a_am_HeaderAdd():                   aModule.Active.HeaderAdd:                   End Sub
Sub a_am_Indent():                      aModule.Active.Indent:                      End Sub
Sub a_am_ListProcedures():              aModule.Active.ListProcedures:              End Sub
Sub a_am_ListProceduresPublic():        aModule.Active.ListProceduresPublic:        End Sub
Sub a_am_PrintListOfInclude():          aModule.Active.PrintListOfInclude:          End Sub
Sub a_am_PrintTodoList():               aModule.Active.PrintTodoList:               End Sub
Sub a_am_ProcedureFoldDeclarations():   aModule.Active.ProcedureFoldDeclarations:   End Sub
Sub a_am_ProcedureScopePrivate():       aModule.Active.ProcedureScopePrivate:       End Sub
Sub a_am_ProcedureScopePublic():        aModule.Active.ProcedureScopePublic:        End Sub
Sub a_am_ProceduresNames():             dp aModule.Active.ProceduresNames:          End Sub
Sub a_am_RemoveEmptyLinesButLeaveOne(): aModule.Active.RemoveEmptyLinesButLeaveOne: End Sub
Sub a_am_SortAZ():                      aModule.Active.ProcedureSortAZ:             End Sub
Sub a_am_SortByKind():                  aModule.Active.ProcedureSortByKind:         End Sub
Sub a_am_SortByScope():                 aModule.Active.ProcedureSortByScope:        End Sub
Sub a_am_UpdateProcedures():            aModule.Active.UpdateProcedures:            End Sub

'____ DESIGNER  Ops _____

Public Sub a_ad_CopyControlProperties():    aDesigner.Active.CopyControlProperties:    End Sub
Public Sub a_ad_CreateFrameMenu():          aDesigner.Active.CreateFrameMenu:          End Sub
Public Sub a_ad_PasteControlProperties():   aDesigner.Active.PasteControlProperties:   End Sub
Public Sub a_ad_RemoveCaption():            aDesigner.Active.RemoveCaption:            End Sub
Public Sub a_ad_RenameControlAndCode():     aDesigner.Active.RenameControlAndCode:     End Sub
Public Sub a_ad_SetHandCursor():            aDesigner.Active.SetHandCursor:            End Sub
Public Sub a_ad_SortControlsHorizontally(): aDesigner.Active.SortControlsHorizontally: End Sub
Public Sub a_ad_SortControlsVertivally():   aDesigner.Active.SortControlsVertivally:   End Sub
Public Sub a_ad_SwitchNames():              aDesigner.Active.SwitchNames:              End Sub
Public Sub a_ad_SwitchPositions():          aDesigner.Active.SwitchPositions:          End Sub
Public Sub a_ad_SideBySide():               aModules.SideBySide ActiveModule.Name:     End Sub

'____ WORKBOOK  Ops _____


'____ USERFORMS ____

Public Sub uShow_CodeOnTheFly():     uCodeOnTheFly.Show:    End Sub
Public Sub uShow_ComponentsAdd():    uModulesAdd.Show:      End Sub
Public Sub uShow_ComponentsRemove(): uModulesRemove.Show:   End Sub
Public Sub uShow_ComponentsRename(): uModulesRename.Show:   End Sub
Public Sub uShow_FormBuilder():      uFormBuilder.Show:     End Sub
Public Sub uShow_ProjectExplorer():  uProjectExplorer.Show: End Sub
Public Sub uShow_ProjectManager():   uProjectManager.Show:  End Sub
Public Sub uShow_References():       uReferences.Show:      End Sub
Public Sub uShow_Skeleton():         uSkeleton.Show:        End Sub

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

Public Sub a_txt_SeparateProcedures(): CallSeparateProcedures: End Sub
Public Sub a_txt_TxtPrepend(): CallTxtPrepend: End Sub


'-------------------------

Public Sub RunVbaGui()
    Dim strProgramName As String
    If Not FileExists(strProgramName) Then Exit Sub
    strProgramName = ThisWorkbook.Path & "\vbaGUI.exe"
    Shell """" & strProgramName & """, vbNormalFocus)"
End Sub

