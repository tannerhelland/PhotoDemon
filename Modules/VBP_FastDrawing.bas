Attribute VB_Name = "FastDrawing"
'***************************************************************************
'Fast API Graphics Routines Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 12/June/01
'Last updated: 31/August/12
'Last update: Completed work on prepImageData and finalizeImageData, which are the much-improved successors to
'             Get/SetImageData.  These routines rely on SafeArrays instead of Get/SetDIBits, so they are quite a bit
'             faster.  They are also image independent; passing a "preview" flag will result in the ability to paint the
'             results of a filter to any picture box, and it's all managed internally - meaning the filter/effect routine
'             doesn't have to worry about a thing.  A public "curLayerValues" variable contains everything a filter could
'             ever want to know about the data it's working on.  Most of the values it provides are unused at present, but
'             could be useful once selections/layers are implemented.
'
'This interface provides API support for the main image interaction routines. It assigns memory data
' into a useable array, and later transfers that array back into memory.  Very fast, very compact, can't
' live without it. These functions are arguably the most integral part of PhotoDemon.
'
'If you want to know more about how DIB sections work - and why they're so fast compared to VB's internal
' .PSet and .Point methods - please visit http://www.tannerhelland.com/42/vb-graphics-programming-3/
'
'***************************************************************************

Option Explicit

'BEGIN DIB-RELATED DECLARATIONS
'Private Type Bitmap
'    bmType As Long
'    bmWidth As Long
'    bmHeight As Long
'    bmWidthBytes As Long
'    bmPlanes As Integer
'    bmBitsPixel As Integer
'    bmBits As Long
'End Type
'
'Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
'Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
'Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal DY As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
'
'Private Type RGBQUAD
'    Blue As Byte
'    Green As Byte
'    Red As Byte
'    Alpha As Byte
'End Type
'
'Private Type BITMAPINFOHEADER
'        biSize As Long
'        biWidth As Long
'        biHeight As Long
'        biPlanes As Integer
'        biBitCount As Integer
'        biCompression As Long
'        biSizeImage As Long
'        biXPelsPerMeter As Long
'        biYPelsPerMeter As Long
'        biClrUsed As Long
'        biClrImportant As Long
'End Type
'
'Private Type BITMAPINFO
'        bmiHeader As BITMAPINFOHEADER
'        bmiColors(0 To 255) As RGBQUAD
'End Type
'END DIB DECLARATIONS

Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Long, lpSrc As Long, ByVal byteLength As Long)

Public workingLayer As pdLayer

'prepImageData will fill a variable of this type with everything a filter or effect could possibly want to know
' about the layer it's operating on.  Filters are free to ignore this data, but it is always made available.
Public Type FilterInfo
    Left As Long            'Coordinates of the top-left location the filter is supposed to operate on
    Top As Long
    Right As Long           'Right and Bottom could be inferred, but we do it here to minimize effort on the calling routine's part
    Bottom As Long
    Width As Long           'Dimensions of the area the filter is supposed to operate on
    Height As Long
    MinX As Long            'The lowest coordinate the filter is allowed to check.  This is almost always (0, 0)
    MinY As Long
    MaxX As Long            'The highest coordinate the filter is allowed to check.  This is almost always (width, height)
    MaxY As Long
    colorDepth As Long      'The colorDepth of the current layer; right now, this should always be 24 or 32
    BytesPerPixel As Long   'BPP is colorDepth / 8.  It is provided for convenience.
    LayerX As Long          'Filters shouldn't have to worry about where the layer is physically located, but when it comes
    LayerY As Long          ' time to set the layer back in place, these may be useful (as when previewing, for example)
End Type

Public curLayerValues As FilterInfo


'The new replacement for GetImageData
' prepPixelData's job is to copy the relevant layer into a temporary object, which is what individual filters and effects
' will operate on.  prepPixelData() also populates the relevant SafeArray object and a host of other variables, which
' filters and effects can then copy locally to ensure the fastest possible runtime speed.
'
'If the filter will be rendering a preview only, it can specify the picture box that will receive the preview effect.
' This function will automatically adjust its parameters accordingly, and the filter routine will not have to make any
' modifications to its code.
'
'Finally, the calling routine can optionally specify a different progress bar maximum value.  By default, this is the current
' layer's width, but some routines run vertically and the progress bar needs to be changed accordingly.
Public Sub prepImageData(ByRef tmpSA As SAFEARRAY2D, Optional isPreview As Boolean = False, Optional previewPictureBox As PictureBox, Optional newProgBarMax As Long = -1)

    'Prepare our temporary layer
    Set workingLayer = New pdLayer
    
    'If this is not a preview, simply copy the current layer without modification
    If isPreview = False Then
        workingLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer
    
    'If this IS a preview, more work is involved.
    Else
    
        'Start by calculating the aspect ratio of both the current image and the previewing picture box
        Dim dstWidth As Single, dstHeight As Single
        dstWidth = previewPictureBox.ScaleWidth
        dstHeight = previewPictureBox.ScaleHeight
    
        Dim SrcWidth As Single, SrcHeight As Single
        SrcWidth = pdImages(CurrentImage).mainLayer.getLayerWidth
        SrcHeight = pdImages(CurrentImage).mainLayer.getLayerHeight
    
        Dim srcAspect As Single, dstAspect As Single
        srcAspect = SrcWidth / SrcHeight
        dstAspect = dstWidth / dstHeight
        
        'Now, use that aspect ratio to determine a proper size for our temporary layer
        Dim newWidth As Long, newHeight As Long
    
        If srcAspect > dstAspect Then
            newWidth = dstWidth
            newHeight = CSng(SrcHeight / SrcWidth) * newWidth + 0.5
        Else
            newHeight = dstHeight
            newWidth = CSng(SrcWidth / SrcHeight) * newHeight + 0.5
        End If
        
        'And finally, create our workingLayer using these values
        workingLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer, newWidth, newHeight
        
    End If
    
    'With our temporary layer successfully created, populate the relevant SafeArray variable
    prepSafeArray tmpSA, workingLayer

    'Finally, populate the ubiquitous curLayerValues variable with everything a filter might want to know
    With curLayerValues
        .Left = 0
        .Top = 0
        .Right = workingLayer.getLayerWidth - 1
        .Bottom = workingLayer.getLayerHeight - 1
        .Width = workingLayer.getLayerWidth
        .Height = workingLayer.getLayerHeight
        .MinX = 0
        .MinY = 0
        .MaxX = workingLayer.getLayerWidth - 1
        .MaxY = workingLayer.getLayerHeight - 1
        .colorDepth = workingLayer.getLayerColorDepth
        .BytesPerPixel = (workingLayer.getLayerColorDepth \ 8)
        .LayerX = 0
        .LayerY = 0
    End With

    'Set up the progress bar (only if this is not a preview, mind you)
    If isPreview = False Then
        If newProgBarMax = -1 Then
            SetProgBarMax (curLayerValues.Left + curLayerValues.Width)
        Else
            SetProgBarMax newProgBarMax
        End If
    End If
    
    'MsgBox "prepImageData worked: " & workingLayer.getLayerHeight & ", " & workingLayer.getLayerWidth & " (" & workingLayer.getLayerArrayWidth & ")" & ", " & workingLayer.getLayerDIBits

End Sub

'This function can be used to populate a valid SAFEARRAY2D structure against any layer
Public Sub prepSafeArray(ByRef srcSA As SAFEARRAY2D, ByRef srcLayer As pdLayer)
    
    'With our temporary layer successfully created, populate the relevant SafeArray variable
    With srcSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcLayer.getLayerHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcLayer.getLayerArrayWidth
        .pvData = srcLayer.getLayerDIBits
    End With
    
End Sub

'The counterpart to prepImageData, finalizeImageData copies the working layer back into its source then renders it
' to the screen.  Like prepImageData(), a preview target can also be named.  In this case, finalizeImageData will rely on
' the values calculated by prepImageData(), as it's presumed that preImageData will ALWAYS be called before this routine.
Public Sub finalizeImageData(Optional isPreview As Boolean = False, Optional previewPictureBox As PictureBox)

    'If this is not a preview, our job is simple - get the newly processed DIB rendered to the screen.
    If isPreview = False Then
        
        Message "Rendering image to screen..."
        
        'Paint the working layer over the original layer
        BitBlt pdImages(CurrentImage).mainLayer.getLayerDC, curLayerValues.LayerX, curLayerValues.LayerY, curLayerValues.Width, curLayerValues.Height, workingLayer.getLayerDC, 0, 0, vbSrcCopy
                
        'workingLayer has served its purpose, so erase it from memory
        Set workingLayer = Nothing
                
        'If we're setting data to the screen, we can reasonably assume that the progress bar should be reset
        SetProgBarVal 0
        
        'Pass control to the viewport renderer, which will perform the actual rendering
        ScrollViewport FormMain.ActiveForm
        
        Message "Finished."
        
    Else
    
        'If the current layer is 32bpp, precomposite it against white before rendering
        If workingLayer.getLayerColorDepth = 32 Then workingLayer.compositeBackgroundColor
    
        'Allow workingLayer to paint itself to the target picture box
        workingLayer.renderToPictureBox previewPictureBox
        
        'workingLayer has served its purpose, so erase it from memory
        Set workingLayer = Nothing
        
    End If
    
End Sub

'We only want the progress bar updating when necessary, so this function finds a power of 2 closest to
Public Function findBestProgBarValue() As Long

    'First, figure out what the range of this operation will be using the values in curLayerValues
    Dim progBarRange As Double
    progBarRange = CDbl(getProgBarMax())
    
    'Divide that value by 20.  20 is an arbitrary selection; the value can be set to any value X, where X is the number
    ' of times we want the progress bar to update during a given filter or effect.
    progBarRange = progBarRange / 20
    
    'Find the nearest power of two to that value, rounded down
    Dim nearestP2 As Long
    
    nearestP2 = Log(progBarRange) / Log(2#)
    
    findBestProgBarValue = (2 ^ nearestP2) - 1
    
End Function
