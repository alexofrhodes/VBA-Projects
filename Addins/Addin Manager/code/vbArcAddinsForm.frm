VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vbArcAddinsForm 
   Caption         =   "vbArc Addins"
   ClientHeight    =   5124
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4632
   OleObjectBlob   =   "vbArcAddinsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "vbArcAddinsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    DownloadThis
End Sub

Private Sub CommandButton2_Click()

    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            If ListBox1.List(i, 1) = "INSTALLED" Then
                AddinName = Left(ListBox1.List(i, 0), InStrRev(ListBox1.List(i, 0), ".") - 1)
                AddIns(AddinName).Installed = Not AddIns(AddinName).Installed
                ListBox1.List(i, 2) = IIf(AddIns(AddinName).Installed = True, "ENABLED", "DISABLED")
            End If
        End If
    Next

End Sub

Private Sub info_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub UserForm_Initialize()
    Dim v As Variant
    v = vbArcAddins
    Dim i As Long
    On Error Resume Next
    For i = 0 To UBound(v)
        ListBox1.AddItem
        ListBox1.List(i, 0) = Mid(v(i), InStrRev(v(i), "/") + 1)
        Debug.Print ListBox1.List(i, 0)
        ListBox1.List(i, 1) = IIf(FileExists(Application.UserLibraryPath & ListBox1.List(i, 0)) = False, "MISSING", "INSTALLED")
            
        AddinName = Left(ListBox1.List(i, 0), InStrRev(ListBox1.List(i, 0), ".") - 1)
'        Debug.Print AddinName & vbTab & Now()
        Set ad = AddIns(AddinName)
        If Not ad Is Nothing Then
            ListBox1.List(i, 2) = IIf(AddIns(AddinName).Installed = True, "ENABLED", "DISABLED")
        End If
        On Error GoTo 0
        ListBox1.List(i, 3) = v(i)
    Next
End Sub

Sub DownloadThis()
'#IMPORTS DownloadFile
    Dim i As Long
    Dim AddinName As String
    Dim ad As AddIn
    
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            AddinName = Left(ListBox1.List(i, 0), InStrRev(ListBox1.List(i, 0), ".") - 1)
            On Error Resume Next
            Set ad = AddIns(AddinName)
            If Not ad Is Nothing Then
                If AddIns(AddinName).Installed = True Then AddIns(FileName).Installed = False
            End If
            
            Kill Application.UserLibraryPath & ListBox1.List(i, 0)
            On Error GoTo 0
            DownloadFile ListBox1.List(i, 3), Application.UserLibraryPath & ListBox1.List(i, 0)
            AddIns.Add (Application.UserLibraryPath & ListBox1.List(i, 0))
            AddIns(AddinName).Installed = True
            ListBox1.List(i, 1) = "INSTALLED"
            ListBox1.List(i, 2) = "ENABLED"
        End If
    Next

End Sub

