Attribute VB_Name = "GDI_Plus"
'***************************************************************************
'GDI+ Interface
'Copyright ©2011-2012 by Tanner Helland
'Created: 1/September/12
'Last updated: 1/September/12
'Last update: initial build
'
'This interface provides a means for interacting with the unnecessarily complex (and overwrought) GDI+ module.  GDI+ is
' primarily used as a fallback for image loading if the FreeImage DLL cannot be found.
'
'These routines are adapted from the work of a number of other talented VB programmers.  Since GDI+ is not well-documented
' for VB users, I have pieced this module together from the following pieces of code:
' Avery P's initial GDI+ deconstruction: http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
' Carles P.V.'s iBMP implementation: http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=42376&lngWId=1
' Robert Rayment's PaintRR implementation: http://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1
' Many thanks to these individuals for their outstanding work on graphics in VB.
'
'***************************************************************************

Option Explicit

'GDI+ Enums
Public Enum GDIPlusStatus
    [OK] = 0
    [GenericError] = 1
    [InvalidParameter] = 2
    [OutOfMemory] = 3
    [ObjectBusy] = 4
    [InsufficientBuffer] = 5
    [NotImplemented] = 6
    [Win32Error] = 7
    [WrongState] = 8
    [Aborted] = 9
    [FileNotFound] = 10
    [ValueOverflow ] = 11
    [AccessDenied] = 12
    [UnknownImageFormat] = 13
    [FontFamilyNotFound] = 14
    [FontStyleNotFound] = 15
    [NotTrueTypeFont] = 16
    [UnsupportedGdiplusVersion] = 17
    [GdiplusNotInitialized ] = 18
    [PropertyNotFound] = 19
    [PropertyNotSupported] = 20
End Enum

'GDI+ required types
Private Type GDIPlusStartupInput
    GDIPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

'OleCreatePictureIndirect types
Private Type PictDesc
    Size       As Long
    Type       As Long
    hBmpOrIcon As Long
    hPal       As Long
End Type

'Start-up and shutdown
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef Token As Long, ByRef InputBuf As GDIPlusStartupInput, Optional ByVal OutputBuffer As Long = 0&) As GDIPlusStatus
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GDIPlusStatus

'Load image from file, process said file, etc.
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal Bitmap As Long, hBmpReturn As Long, ByVal Background As Long) As GDIPlusStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GDIPlusStatus

'OleCreatePictureIndirect is used to convert GDI+ images to VB's preferred StdPicture
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PictDesc, rIid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long

'When GDI+ is initialized, it will assign us a token.  We use this to release GDI+ when the program terminates.
Private GDIPlusToken As Long

'Use GDI+ to load a picture into a StdPicture object - not ideal, as some information will be lost in the transition, but since
' this is only a fallback from FreeImage I'm not going out of my way to improve it.
Public Function GDIPlusLoadPicture(ByVal srcFilename As String) As StdPicture

    'Used to hold the return values of various GDI+ calls
    Dim GDIPlusReturn As Long
      
    'Use GDI+ to load the image
    Dim hImage As Long
    GDIPlusReturn = GdipLoadImageFromFile(StrPtr(srcFilename), hImage)
    
    'Copy the GDI+ image into a standard bitmap
    Dim hBitmap As Long
    GDIPlusReturn = GdipCreateHBITMAPFromBitmap(hImage, hBitmap, vbBlack)
    
    'Now we can release the GDI+ copy of the image
    GDIPlusReturn = GdipDisposeImage(hImage)
    
    'Assuming the load/unload went okay, prepare to copy the bitmap object into an StdPicture
    If (GDIPlusReturn = [OK]) Then

        'Prepare the header required by OleCreatePictureIndirect
        Dim picHeader As PictDesc
        With picHeader
            .Size = Len(picHeader)
            .Type = vbPicTypeBitmap
            .hBmpOrIcon = hBitmap
            .hPal = 0
        End With
        
        'Populate the magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        Dim aGuid(0 To 3) As Long
        aGuid(0) = &H7BF80980
        aGuid(1) = &H101ABF32
        aGuid(2) = &HAA00BB8B
        aGuid(3) = &HAB0C3000
        
        'Using the bitmap indirectly created by GDI+, build an identical StdPicture object
        OleCreatePictureIndirect picHeader, aGuid(0), -1, GDIPlusLoadPicture
        
    End If

End Function

'At start-up, this function is called to determine whether or not we have GDI+ available on this machine.
Public Function isGDIPlusAvailable() As Boolean

    Dim GDICheck As GDIPlusStartupInput
    GDICheck.GDIPlusVersion = 1
    
    If (GdiplusStartup(GDIPlusToken, GDICheck) <> [OK]) Then
        isGDIPlusAvailable = False
    Else
        isGDIPlusAvailable = True
    End If

End Function

'At shutdown, this function must be called to release our GDI+ instance
Public Function releaseGDIPlus()
    GdiplusShutdown GDIPlusToken
End Function
