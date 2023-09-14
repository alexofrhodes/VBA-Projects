VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uImageMso 
   Caption         =   "UserForm1"
   ClientHeight    =   1500
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7956
   OleObjectBlob   =   "uImageMso.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uImageMso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim Labels() As New clsMsoImage
Dim off As Long
Dim rCount As Long
Dim cCount As Long
Dim Size As Long
Dim matrix As Long
Dim totalImages As Long

Sub SetLabelEvents()
    Dim counter As Integer, obj As MSForms.control
    For Each obj In Me.Controls
        If TypeOf obj Is MSForms.Label Then
            counter = counter + 1
            ReDim Preserve Labels(1 To counter)
            Set Labels(counter).LabelEvents = obj
        End If
    Next
    Set obj = Nothing
End Sub

Private Sub CommandButton1_Click()
ActiveCell.Value = TextBox1.Value
Me.Hide
End Sub

Private Sub UserForm_Initialize()
SetUImageMsoVariables
LoadMSO off
Me.Caption = off + 1 & " to " & off + matrix & " of " & totalImages
SetLabelEvents
End Sub

Sub SetUImageMsoVariables()
    Me.Width = 1020
    Me.Height = 550
    
    PreviousBatch.Top = Me.Height - PreviousBatch.Height - 35
    '    PreviousBatch.left = Me.Width / 2 - PreviousBatch.Width - 50
    NextBatch.Top = Me.Height - NextBatch.Height - 35
    '    NextBatch.left = Me.Width / 2 + NextBatch.Width - 50
    
    CommandButton1.Top = Me.Height - CommandButton1.Height - 35
    '    CommandButton1.left = Me.Width / 2 + CommandButton1.Width + 50
    
    TextBox1.Top = Me.Height - TextBox1.Height - 35
    '    TextBox1.left = Me.Width / 2 - TextBox1.Width - 150
    rCount = 15
    cCount = 30
    matrix = rCount * cCount
    totalImages = WorksheetFunction.CountA(ThisWorkbook.Sheets("msoImages").Range("A1").CurrentRegion)
    Size = 32
    off = 0
End Sub

Private Sub NextBatch_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
off = off + matrix
If off > totalImages - matrix Then off = totalImages - matrix
Restart
LoadMSO off
Me.Caption = off + 1 & " to " & off + matrix & " of " & totalImages
SetLabelEvents
End Sub

Private Sub PreviousBatch_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
off = off - matrix
If off < 0 Then off = 0
Restart
LoadMSO off
Me.Caption = off + 1 & " to " & off + matrix & " of " & totalImages
SetLabelEvents
End Sub

Sub Restart()
    Dim c As MSForms.control
    For Each c In Me.Controls
        If TypeName(c) = "Label" Then Me.Controls.Remove c.Name
    Next
End Sub

Sub LoadMSO(Optional OFFSET As Long = 0)
    Dim cell As Range
    Set cell = ThisWorkbook.Sheets("msoImages").Range("A1")
    Set cell = cell.OFFSET(OFFSET)
    Dim rows As Long, cols As Long
    Dim counter As Long
    Dim newImage As MSForms.Label
    On Error GoTo Skip
    For rows = 1 To rCount
        For cols = 1 To cCount
            Set newImage = Me.Controls.Add("Forms.Label.1", "Image" & counter, True)
            newImage.Top = Size * rows - Size + 5
            newImage.Left = Size * cols - Size + 5
            newImage.Height = 36
            newImage.Width = 36
            newImage.Caption = cell
            newImage.Font.Size = 0
            newImage.Picture = Application.CommandBars.GetImageMso(cell, Size, Size)
Skip:
            Set cell = cell.OFFSET(1)
            If cell = "" Then Exit Sub
        Next
    Next
    Debug.Print Me.Controls.count
End Sub

