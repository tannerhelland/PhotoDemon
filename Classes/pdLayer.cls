VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Layer class
'Copyright ©2000-2012 by Tanner Helland
'Created: 29/August/12
'Last updated: 29/August/12
'Last update: initial build
'
'The time has come to break PhotoDemon's reliance on picture boxes.  The new plan is to implement an honest-to-God
' DIB class that will manage all image data in memory.  If built correctly, this thing will provide a whole host of
' benefits, including but not limited to:
'
'- Ability to load large images without problems
'- Alpha channels
'- High bit depths
'- Much faster
'- Eventual implementation of layers and adjustment layers
'
'Those are some serious benefits.
'
'Implementation will be slow and steady at first.  A rough roadmap is as follows:
'- Get image loading working.  This will require a lot of plumbing to already be in place, but I can't do anything else
' until I can get image data into the class.
'- Get viewport rendering working.  If I can't see the image, I can't work on it.
'- Experiment with tying the class's data into GetImageData, which is already well-implemented throughout PD.
'- Replace PicWidthL and PicHeightL with width/height values provided by this class.
'- Remove all direct intervention with ActiveForm.BackBuffer and force it through this class.  That means no unmanaged
' changing of width, height, etc - it has to come through here.
'
'This class will be a subset of pdImage().  Initially there will only be one layer (the image's data), but in the future,
' assuming this class is written well, a pdImage object could theoretically store many layers.
'
'Let the fun begin.
'
'Note: this work is based off a DIB class written by Herman Liu, and his help was invaluable in testing this.  Many thanks.
'
'***************************************************************************

Option Explicit

'DIB API functions
' (Note that these are currently declared in FastDrawing as well)
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BitmapInfo, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BitmapInfo, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'Clean-up API functions
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal HWnd As Long, ByVal hDC As Long) As Long

'DIB Types
Private Type RGBQuad
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Private Type Bitmap
    Type As Long
    Width As Long
    Height As Long
    WidthBytes As Long
    Planes As Integer
    BitsPerPixel As Integer
    Bits As Long
End Type

Private Type BitmapInfoHeader
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    ImageSize As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ColorUsed As Long
    ColorImportant As Long
End Type

Private Type BitmapInfo
    Header As BitmapInfoHeader
    Colors(0 To 255) As RGBQuad
End Type

'Variables related to the DIB
'Public layerDC As Long

'An StdPicture associated with this layer
Private layerPicture As StdPicture

'The image's width and height
Private lWidth As Long, lHeight As Long

'Return this layer's width
Public Function getLayerWidth() As Long
    getLayerWidth = lWidth
End Function

'Return this layer's height
Public Function getLayerheight() As Long
    getLayerWidth = lHeight
End Function


'Give this layer a picture from a standard VB picture object
Public Function setLayerPicture(ByRef srcPicture As StdPicture) As Boolean

    'Erase the current picture object
    If layerPicture <> Nothing Then Set layerPicture = Nothing
    
    'Re-initialize the picture object and copy the image data
    Set layerPicture = New StdPicture
    Set layerPicture = srcPicture

    

    'If we reached this point, the set was successful
    setLayerPicture = True

End Function

'Return a reference to the StdPicture object associated with this layer
Public Function getLayerPicture() As StdPicture
    getLayerPicture = layerPicture
End Function

'Take a given array and fill it with the image's pixel data (DIB-style)
Public Function getLayerArray(ByRef arrayToFill() As Byte) As Boolean

End Function


' Simplified and for 24 or 32-BPP only
Public Function initializeLayer(ByVal iWidth As Long, ByVal iHeight As Long, Optional colorDepth As Long = 24) As Boolean
    
End Function

'This will effectively reset everything related to this layer, including image data.  Use cautiously!
Public Function eraseLayer()

    'If we have image data, clear it out
    If layerDC <> -1 Then
    
    End If

    'Return the associated DIB section DC to zero
    layerDC = -1

End Function

'INITIALIZE class
Private Sub Class_Initialize()

    'Set the layer's DC to -1; this lets us know that it has not been initialized yet
    layerDC = -1

End Sub
