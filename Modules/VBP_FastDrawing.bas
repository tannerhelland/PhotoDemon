Attribute VB_Name = "FastDrawing"
'***************************************************************************
'Fast API Graphics Routines Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 12/June/01
'Last updated: 13/January/07
'Last update: Merged the separate Get/SetImageData routines into a single one.  Now a parameter can
'              be passed to specify using a second array for caching the DIB data.
'
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
Private Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Private Type RGBQUAD
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(0 To 255) As RGBQUAD
End Type
'END DIB DECLARATIONS

'DOHL TEST DECLARATIONS
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

'/DOHC

'To prevent double-image loading errors
Private AllowPreview As Boolean

'I have since implemented two sections of code, one each for two arrays
'This is necessary for implementing certain problematic double-layered effects
'(including all resizing, rotating, flipping, etc.)
Public ImageData() As Byte
Public ImageData2() As Byte

'In the future, these variables could be used to tell calling routines which section of a picture to apply a filter to.
' Right now, they are initialized to default values (0,0-imageWidth,imageHeight)
Public imgStartX As Long, imgStartY As Long
Public imgFinalX As Long, imgFinalY As Long

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
    
        Dim srcWidth As Single, srcHeight As Single
        srcWidth = pdImages(CurrentImage).mainLayer.getLayerWidth
        srcHeight = pdImages(CurrentImage).mainLayer.getLayerHeight
    
        Dim srcAspect As Single, dstAspect As Single
        srcAspect = srcWidth / srcHeight
        dstAspect = dstWidth / dstHeight
        
        'Now, use that aspect ratio to determine a proper size for our temporary layer
        Dim newWidth As Long, newHeight As Long
    
        If srcAspect > dstAspect Then
            newWidth = dstWidth
            newHeight = CSng(srcHeight / srcWidth) * newWidth + 0.5
        Else
            newHeight = dstHeight
            newWidth = CSng(srcWidth / srcHeight) * newHeight + 0.5
        End If
        
        'And finally, create our workingLayer using these values
        workingLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer, newWidth, newHeight
        
    End If
    
    'With our temporary layer successfully created, populate the relevant SafeArray variable
    With tmpSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = workingLayer.getLayerHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = workingLayer.getLayerArrayWidth
        .pvData = workingLayer.getLayerDIBits
    End With

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
    
        'Allow workingLayer to paint itself to the target picture box
        workingLayer.renderToPictureBox previewPictureBox
        
        'workingLayer has served its purpose, so erase it from memory
        Set workingLayer = Nothing
        
    End If
    
End Sub

'Use GetObject to determine an image's width
Public Function GetImageWidth()
    Dim bm As Bitmap
    GetObject FormMain.ActiveForm.BackBuffer.Image, Len(bm), bm
    GetImageWidth = bm.bmWidth
End Function

'Use GetObject to determine an image's height
Public Function GetImageHeight()
    Dim bm As Bitmap
    GetObject FormMain.ActiveForm.BackBuffer.Image, Len(bm), bm
    GetImageHeight = bm.bmHeight
End Function

'GetImageData takes the image data from the active buffer and assigns it to a two-dimensional array.
' It works any color mode, but it will always force image data into a 24-bit color array.
' If you want to work with data of another color depth, PhotoDemon is not the project for you.  ;)
Public Sub GetImageData(Optional ByVal CorrectOrientation As Boolean = False)
   
   
    'Bitmap data types required by the DIB section API calls
    Dim bm As Bitmap
    Dim bmi As BITMAPINFO
    
    'The size of the image array - we need to use some specialized math to ensure the API will work with it
    Dim ArrayWidth As Long, ArrayHeight As Long

    'Use the API to get accurate width and height values
    GetObject FormMain.ActiveForm.BackBuffer.Image, Len(bm), bm
    PicWidthL = bm.bmWidth
    PicHeightL = bm.bmHeight
    
    'Now, build a custom bitmap-type variable with the values we want - specifically, uncompressed 24-bit pixel data
    bmi.bmiHeader.biSize = 40
    bmi.bmiHeader.biWidth = PicWidthL
    bmi.bmiHeader.biHeight = PicHeightL
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 24
    bmi.bmiHeader.biCompression = 0
    
    'Resize the array to a width that's a multiple of 4 (required by the API)
    ArrayWidth = (PicWidthL * 3) - 1
    ArrayWidth = ArrayWidth + (PicWidthL Mod 4)

    'Height doesn't require anything special
    ArrayHeight = PicHeightL
    
    'This array will contain the image data, in the form (X * 3 + Z, Y), where z = 2 for red, 1 for green, 0 for blue
    ReDim ImageData(0 To ArrayWidth, 0 To ArrayHeight) As Byte
    
    'If the calling routine wants us to orient the data in a top-left fashion, assign the image data to a temporary array
    ' (so it can be processed below).  Otherwise, stick it right into ImageData()
    If CorrectOrientation = False Then
        GetDIBits FormMain.ActiveForm.BackBuffer.hDC, FormMain.ActiveForm.BackBuffer.Image, 0, PicHeightL, ImageData(0, 0), bmi, 0
    Else
        ReDim TempArray(0 To ArrayWidth, 0 To ArrayHeight) As Byte
        GetDIBits FormMain.ActiveForm.BackBuffer.hDC, FormMain.ActiveForm.BackBuffer.Image, 0, PicHeightL, TempArray(0, 0), bmi, 0
    End If

    'Because the image processing functions run from 0 to .Width/.Height - 1, adjust the width and height here
    PicWidthL = PicWidthL - 1
    PicHeightL = PicHeightL - 1
    
    'In the future, these values could be used to have calling functions run on a subset of the image (for example, only
    ' the selected area).  Right now, however, they are simply initialized to the full size of the image
    imgStartX = 0
    imgStartY = 0
    imgFinalX = PicWidthL
    imgFinalY = PicHeightL

    'If the user has requested reorientation of the image data (i.e. (0,0) as top-left, (max,max) as bottom right), process that now.
    ' If this option is enabled, we must set the DIB height to negative in the SetImageData routine below
    If CorrectOrientation = True Then
    
        Dim QuickVal As Long
        For x = 0 To PicWidthL
            QuickVal = x * 3
         For y = 0 To PicHeightL
          For z = 0 To 2
            ImageData(QuickVal + z, y) = TempArray(QuickVal + z, PicHeightL - y)
          Next z
         Next y
        Next x
        
        'Clear out the temporary array
        Erase TempArray
        
    End If
    
    'Now that we have valid image data, allow previewing functions to trigger
    AllowPreview = True
    
End Sub

'Take an array created by GetImageData (and probably modified by some sort of filter), and draw it to the active buffer
Public Sub setImageData(Optional ByVal CorrectOrientation As Boolean = False)
    
    Message "Rendering image to screen..."

    'We subtracted one from these values as part of GetImageData - the time has come to return them to their rightful values
    PicWidthL = PicWidthL + 1
    PicHeightL = PicHeightL + 1
    
    'Just like GetImageData, we need to populate a bitmap-type variable with values corresponding to the current image data
    Dim bmi As BITMAPINFO
    bmi.bmiHeader.biSize = 40
    bmi.bmiHeader.biWidth = PicWidthL
    
    'Height must be reversed if the image data has been reoriented in GetImageData
    If CorrectOrientation = False Then bmi.bmiHeader.biHeight = PicHeightL Else bmi.bmiHeader.biHeight = -PicHeightL
    
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 24
    bmi.bmiHeader.biCompression = 0
    
    StretchDIBits FormMain.ActiveForm.BackBuffer.hDC, 0, 0, PicWidthL, PicHeightL, 0, 0, PicWidthL, PicHeightL, ImageData(0, 0), bmi, 0, vbSrcCopy
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    FormMain.ActiveForm.BackBuffer.Refresh
    
    'If we're setting data to the screen, we can reasonably assume that the progress bar should be reset
    SetProgBarVal 0
    
    'Clear out ImageData
    Erase ImageData
    
    'Pass control to the viewport renderer, which will make the new image actually appear on-screen
    ScrollViewport FormMain.ActiveForm
    
    'Restore the image width/heigth variables in case other routines aren't done with them
    PicWidthL = PicWidthL - 1
    PicHeightL = PicHeightL - 1

    Message "Finished. "

End Sub

'Used to draw preview images (for example, on filter forms).  See above GetImageData for comments
Public Sub GetPreviewData(ByRef SrcPic As PictureBox, Optional ByVal CorrectOrientation As Boolean = False)

    Dim bm As Bitmap
    Dim bmi As BITMAPINFO
    Dim ArrayWidth As Long, ArrayHeight As Long

    GetObject SrcPic.Image, Len(bm), bm
    bmi.bmiHeader.biSize = 40
    bmi.bmiHeader.biWidth = bm.bmWidth
    bmi.bmiHeader.biHeight = bm.bmHeight
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 24
    bmi.bmiHeader.biCompression = 0
    
    ArrayWidth = (bm.bmWidth * 3) - 1
    ArrayWidth = ArrayWidth + (bm.bmWidth Mod 4)
    ArrayHeight = bm.bmHeight
    
    ReDim ImageData(0 To ArrayWidth, 0 To ArrayHeight) As Byte
    
    If CorrectOrientation = False Then
        GetDIBits SrcPic.hDC, SrcPic.Image, 0, bm.bmHeight, ImageData(0, 0), bmi, 0
    Else
        ReDim TempArray(0 To ArrayWidth, 0 To ArrayHeight) As Byte
        GetDIBits SrcPic.hDC, SrcPic.Image, 0, bm.bmHeight, TempArray(0, 0), bmi, 0
    End If

    If CorrectOrientation = True Then
        Dim QuickVal As Long
        For x = 0 To bm.bmWidth - 1
            QuickVal = x * 3
         For y = 0 To bm.bmHeight - 1
          For z = 0 To 2
            ImageData(QuickVal + z, y) = TempArray(QuickVal + z, bm.bmHeight - 1 - y)
          Next z
         Next y
        Next x
        
        'Save memory...?
        Erase TempArray
        
    End If
    
End Sub

'Used to draw preview images (for example, on filter forms).  See above SetImageData for comments
Public Sub SetPreviewData(ByRef dstPic As PictureBox, Optional ByVal CorrectOrientation As Boolean = False)
    
    Dim bm As Bitmap
    Dim bmi As BITMAPINFO

    GetObject dstPic.Image, Len(bm), bm

    bmi.bmiHeader.biSize = 40
    bmi.bmiHeader.biWidth = bm.bmWidth
    
    If CorrectOrientation = False Then bmi.bmiHeader.biHeight = bm.bmHeight Else bmi.bmiHeader.biHeight = -bm.bmHeight
    
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 24
    bmi.bmiHeader.biCompression = 0
    
    StretchDIBits dstPic.hDC, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, bm.bmWidth, bm.bmHeight, ImageData(0, 0), bmi, 0, vbSrcCopy
    'dstPic.Picture = dstPic.Image
    dstPic.Refresh
    
End Sub



'REPEAT METHODS FOR SECOND ARRAY (see GetImageData above for relevant comments)
Public Sub GetImageData2(Optional ByVal CorrectOrientation As Boolean = False)
    
    Dim bm As Bitmap
    Dim bmi As BITMAPINFO
    Dim ArrayWidth As Long, ArrayHeight As Long

    GetObject FormMain.ActiveForm.BackBuffer2.Image, Len(bm), bm
    
    PicWidthL = bm.bmWidth
    PicHeightL = bm.bmHeight
    
    bmi.bmiHeader.biSize = 40
    bmi.bmiHeader.biWidth = PicWidthL
    bmi.bmiHeader.biHeight = PicHeightL
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 24
    bmi.bmiHeader.biCompression = 0
    ArrayWidth = (PicWidthL * 3) - 1
    ArrayWidth = ArrayWidth + (PicWidthL Mod 4)
    ArrayHeight = PicHeightL + 1
    
    ReDim ImageData2(0 To ArrayWidth, 0 To ArrayHeight) As Byte
    
    If CorrectOrientation = False Then
        GetDIBits FormMain.ActiveForm.BackBuffer2.hDC, FormMain.ActiveForm.BackBuffer2.Image, 0, PicHeightL, ImageData2(0, 0), bmi, 0
    Else
        ReDim TempArray(0 To (PicWidthL * 3) - 1, 0 To PicHeightL) As Byte
        GetDIBits FormMain.ActiveForm.BackBuffer2.hDC, FormMain.ActiveForm.BackBuffer2.Image, 0, PicHeightL, TempArray(0, 0), bmi, 0
    End If
    
    PicWidthL = PicWidthL - 1
    PicHeightL = PicHeightL - 1

    If CorrectOrientation = True Then
    
        Dim QuickVal As Long
        For x = 0 To PicWidthL
            QuickVal = x * 3
         For y = 0 To PicHeightL
          For z = 0 To 2
            ImageData2(QuickVal + z, y) = TempArray(QuickVal + z, PicHeightL - y)
          Next z
         Next y
        Next x
        
        Erase TempArray
        
    End If
    
End Sub

Public Sub SetImageData2(Optional ByVal CorrectOrientation As Boolean = False)
    
    Message "Rendering image to screen..."
    
    Dim bmi As BITMAPINFO

    PicWidthL = PicWidthL + 1
    PicHeightL = PicHeightL + 1
    bmi.bmiHeader.biSize = 40
    bmi.bmiHeader.biWidth = PicWidthL
    If CorrectOrientation = False Then bmi.bmiHeader.biHeight = PicHeightL Else bmi.bmiHeader.biHeight = -PicHeightL
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 24
    bmi.bmiHeader.biCompression = 0
    StretchDIBits FormMain.ActiveForm.BackBuffer2.hDC, 0, 0, PicWidthL, PicHeightL, 0, 0, PicWidthL, PicHeightL, ImageData2(0, 0), bmi, 0, vbSrcCopy
    FormMain.ActiveForm.BackBuffer2.Picture = FormMain.ActiveForm.BackBuffer2.Image
    FormMain.ActiveForm.BackBuffer2.Refresh
    SetProgBarVal 0
    
    Erase ImageData2
    
    PicWidthL = PicWidthL - 1
    PicHeightL = PicHeightL - 1
    
    ScrollViewport FormMain.ActiveForm
    
    Message "Finished. "

End Sub

'We only want the progress bar updating when necessary, so this function finds a power of 2 closest to
Public Function findBestProgBarValue() As Long

    'First, figure out what the range of this operation will be using the values in curLayerValues
    Dim progBarRange As Single
    progBarRange = curLayerValues.Right - curLayerValues.Left
    
    'Divide that value by 20.  20 is an arbitrary selection; the value can be set to any value X, where X is the number
    ' of times we want the progress bar to update during a given filter or effect.
    progBarRange = progBarRange / 20
    
    'Find the nearest power of two to that value, rounded down
    Dim nearestP2 As Long
    
    nearestP2 = Log(progBarRange) / Log(2#)
    
    findBestProgBarValue = (2 ^ nearestP2) - 1
    
End Function
