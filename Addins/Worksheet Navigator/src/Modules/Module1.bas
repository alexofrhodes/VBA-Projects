Attribute VB_Name = "Module1"
Sub WorksheetNavigatorButtonClicked(Control As IRibbonControl)
uSheetsNavigator.Show
End Sub
Public Function CLIP(Optional StoreText As String) As String
    Dim X As Variant
    X = StoreText
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(StoreText)
                    .SetData "text", X
                Case Else
                    CLIP = .GetData("text")
            End Select
        End With
    End With
End Function
