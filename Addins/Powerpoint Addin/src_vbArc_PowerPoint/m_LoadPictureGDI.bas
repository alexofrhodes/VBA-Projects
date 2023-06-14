Attribute VB_Name = "m_LoadPictureGDI"
Option Explicit

'This module provides a LoadPictureGDI function, which can
'be used instead of VBA's LoadPicture, to load a wide variety
'of image types from disk - including png.
'
'The png format is used in Office 2007-2010 to provide images that
'include an alpha channel for each pixel's transparency
'
'Author:    Stephen Bullen
'Date:      31 October, 2006
'Email:     stephen@oaltd.co.uk

'Updated :  30 December, 2010
'By :       Rob Bovey
'Reason :   Also working now in the 64 bit version of Office 2010


'Declare a UDT to store a GUID for the IPicture OLE Interface
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


#If VBA7 Then
    'Declare a UDT to store the bitmap information
    Private Type PICTDESC
        Size As Long
        Type As Long
        hPic As LongPtr
        hPal As LongPtr
    End Type
    
    'Declare a UDT to store the GDI+ Startup information
    Private Type GDIPlusStartupInput
        GdiPlusVersion As Long
        DebugEventCallback As LongPtr
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type
    
    'Windows API calls into the GDI+ library
    Private Declare PtrSafe Function GdiplusStartup Lib "GDIPlus" (token As LongPtr, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
    Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal FileName As LongPtr, BITMAP As LongPtr) As Long
    Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal BITMAP As LongPtr, hbmReturn As LongPtr, ByVal background As LongPtr) As Long
    Private Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" (ByVal image As LongPtr) As Long
    Private Declare PtrSafe Sub GdiplusShutdown Lib "GDIPlus" (ByVal token As LongPtr)
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
#Else
    'Declare a UDT to store the bitmap information
    Private Type PICTDESC
        Size As Long
        Type As Long
        hPic As Long
        hPal As Long
    End Type
    
    'Declare a UDT to store the GDI+ Startup information
    Private Type GDIPlusStartupInput
        GdiPlusVersion As Long
        DebugEventCallback As Long
        SuppressBackgroundThread As Long
        SuppressExternalCodecs As Long
    End Type
    
    'Windows API calls into the GDI+ library
    Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
    Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, bitmap As Long) As Long
    Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As Long
    Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal image As Long) As Long
    Private Declare Sub GdiplusShutdown Lib "GDIPlus" (ByVal token As Long)
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
#End If

' Procedure:    LoadPictureGDI
' Purpose:      Loads an image using GDI+
' Returns:      The image as an IPicture Object
Public Function LoadPictureGDI(ByVal sFilename As String) As IPicture

    Dim uGdiInput As GDIPlusStartupInput
    Dim lResult As Long
#If VBA7 Then
    Dim hGdiPlus As LongPtr
    Dim hGdiImage As LongPtr
    Dim hBitmap As LongPtr
#Else
    Dim hGdiPlus As Long
    Dim hGdiImage As Long
    Dim hBitmap As Long
#End If

    'Initialize GDI+
    uGdiInput.GdiPlusVersion = 1
    lResult = GdiplusStartup(hGdiPlus, uGdiInput)

    If lResult = 0 Then

        'Load the image
        lResult = GdipCreateBitmapFromFile(StrPtr(sFilename), hGdiImage)

        If lResult = 0 Then

            'Create a bitmap handle from the GDI image
            lResult = GdipCreateHBITMAPFromBitmap(hGdiImage, hBitmap, 0)

            'Create the IPicture object from the bitmap handle
            Set LoadPictureGDI = CreateIPicture(hBitmap)

            'Tidy up
            GdipDisposeImage hGdiImage
        End If

        'Shutdown GDI+
        GdiplusShutdown hGdiPlus
    End If

End Function


' Procedure:    CreateIPicture
' Purpose:      Converts a image handle into an IPicture object.
' Returns:      The IPicture object
#If VBA7 Then
Private Function CreateIPicture(ByVal hPic As LongPtr) As IPicture
#Else
Private Function CreateIPicture(ByVal hPic As Long) As IPicture
#End If
    Dim lResult As Long
    Dim uPicinfo As PICTDESC
    Dim IID_IDispatch As GUID
    Dim iPic As IPicture

    'OLE Picture types
    Const PICTYPE_BITMAP = 1

    ' Create the Interface GUID (for the IPicture interface)
    With IID_IDispatch
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With

    ' Fill uPicInfo with necessary parts.
    With uPicinfo
        .Size = Len(uPicinfo)
        .Type = PICTYPE_BITMAP
        .hPic = hPic
        .hPal = 0
    End With

    ' Create the Picture object.
    lResult = OleCreatePictureIndirect(uPicinfo, IID_IDispatch, True, iPic)

    ' Return the new Picture object.
    Set CreateIPicture = iPic

End Function

