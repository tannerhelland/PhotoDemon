Attribute VB_Name = "DIBs"
'***************************************************************************
'DIB Support Functions
'Copyright 2012-2026 by Tanner Helland
'Created: 27/March/15 (though many individual functions are much older!)
'Last updated: 28/October/22
'Last update: improve behavior of alpha-channel thresholding
'
'This module contains support functions for the pdDIB class.  In old versions of PD,
' these functions were provided by pdDIB, but there's no sense cluttering up that class
' with functions that are only used on rare occasions.  As such, I'm moving as many of
' those functions as I can to this module.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Const OBJ_BITMAP As Long = 7
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectW" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Does a given DIB have "binary" transparency, e.g. does it have alpha values of only 0 or 255?
'
'As a convenience, if you want to confirm that an image has a fully opaque alpha channel (all alpha = 255),
' you can set checkForZero to FALSE.  This allows you to quickly validate an alpha channel, which is helpful
' for knowing if you can save time by converting an image to 24-bpp for some operation.
'
'If, on the other hand, you are exporting to a file format like GIF, you probably want to leave checkForZero to TRUE.
' This will actually check for binary alpha values, e.g. TRUE will be returned if the image contains only
' 255 and/or 0 values, both of which are valid for a GIF file.
Public Function IsDIBAlphaBinary(ByRef srcDIB As pdDIB, Optional ByVal checkForZero As Boolean = True) As Boolean
    
    IsDIBAlphaBinary = False
    
    'Make sure the DIB exists
    If (srcDIB Is Nothing) Then Exit Function
    
    'Make sure this DIB is 32bpp. If it isn't, running this function is pointless.
    If (srcDIB.GetDIBColorDepth = 32) Then

        'Make sure this DIB isn't empty
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
    
            'Loop through the image and compare each alpha value against 0 and 255. If a value doesn't
            ' match either of this, this is a non-binary alpha channel and it must be handled specially.
            Dim iData() As Byte, tmpSA As SafeArray1D
            srcDIB.WrapArrayAroundScanline iData, tmpSA, 0
            
            Dim dibPtr As Long, dibStride As Long
            dibPtr = tmpSA.pvData
            dibStride = tmpSA.cElements
            
            Dim x As Long, y As Long
                
            'By default, assume that the image does not have a binary alpha channel. (This is the preferable
            ' default, as we will exit the loop IFF a non-0 or non-255 value is found.)
            Dim notBinary As Boolean
            notBinary = False
            
            Dim chkAlpha As Byte
                
            'Loop through the image, checking alphas as we go
            Dim finalX As Long
            finalX = srcDIB.GetDIBWidth * 4 - 4
            For y = 0 To srcDIB.GetDIBHeight - 1
            
                'Point our array at the current scanline
                tmpSA.pvData = dibPtr + y * dibStride
                
            For x = 0 To finalX Step 4
                
                'Retrieve the alpha value of the current pixel
                chkAlpha = iData(x + 3)
                
                'For optimization reasons, this is stated as two IFs instead of an OR.
                If (chkAlpha <> 255) Then
                
                    If checkForZero Then
                        If (chkAlpha <> 0) Then
                            notBinary = True
                            Exit For
                        End If
                    Else
                        notBinary = True
                        Exit For
                    End If
                
                End If
                
            Next x
                If notBinary Then Exit For
            Next y
    
            srcDIB.UnwrapArrayFromDIB iData
            IsDIBAlphaBinary = Not notBinary
                
        End If
        
    End If
    
End Function

'Is a given DIB grayscale?  Determination is made by scanning each pixel and comparing RGB values to see if they match.
' (32-bpp inputs required.)
Public Function IsDIBGrayscale(ByRef srcDIB As pdDIB) As Boolean
    
    IsDIBGrayscale = False
    
    'Make sure the DIB exists
    If (Not srcDIB Is Nothing) Then
        
        'Make sure this DIB isn't empty
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
        
            'Loop through the image and compare each alpha value against 0 and 255. If a value doesn't
            ' match either of this, this is a non-binary alpha channel and it must be handled specially.
            Dim imgPixels() As Byte, tmpSA As SafeArray1D, dibPtr As Long, dibStride As Long
            srcDIB.WrapArrayAroundScanline imgPixels, tmpSA, 0
            dibPtr = tmpSA.pvData
            dibStride = tmpSA.cElements
            
            Dim x As Long, y As Long
            Dim r As Long, g As Long, b As Long
            
            'Loop through the image, checking alphas as we go
            Dim finalX As Long
            finalX = srcDIB.GetDIBWidth * 4 - 4
            
            For y = 0 To srcDIB.GetDIBHeight - 1
            
                'Point our array at the current scanline
                tmpSA.pvData = dibPtr + y * dibStride
                
            For x = 0 To finalX Step 4
                
                b = imgPixels(x)
                g = imgPixels(x + 1)
                r = imgPixels(x + 2)
                
                If (r <> g) Or (g <> b) Or (r <> b) Then
                    srcDIB.UnwrapArrayFromDIB imgPixels
                    IsDIBGrayscale = False
                    Exit Function
                End If
                    
            Next x
            Next y
        
            'If we scanned all pixels without exiting prematurely, the DIB is grayscale
            srcDIB.UnwrapArrayFromDIB imgPixels
            IsDIBGrayscale = True
            
        End If
            
    End If

End Function

'Is a given DIB one solid color?  Determination is made by scanning each pixel and comparing RGB values to see if they match.
' (32-bpp inputs required.)
Public Function IsDIBSolidColor(ByRef srcDIB As pdDIB) As Boolean
    
    'If we made it to this line, the DIB is blank, so it doesn't matter what value we return
    IsDIBSolidColor = False
    
    'Make sure the DIB exists
    If (Not srcDIB Is Nothing) Then
        
        'Make sure this DIB isn't empty
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
        
            'Loop through the image and compare each alpha value against 0 and 255. If a value doesn't
            ' match either of this, this is a non-binary alpha channel and it must be handled specially.
            Dim imgPixels() As Byte, tmpSA As SafeArray1D, dibPtr As Long, dibStride As Long
            srcDIB.WrapArrayAroundScanline imgPixels, tmpSA, 0
            dibPtr = tmpSA.pvData
            dibStride = tmpSA.cElements
            
            Dim x As Long, y As Long
            Dim r As Long, g As Long, b As Long
            
            'Flag the first color in the image
            Dim srcR As Long, srcG As Long, srcB As Long
            srcB = imgPixels(0)
            srcG = imgPixels(1)
            srcR = imgPixels(2)
            
            'Loop through the image, checking alphas as we go
            Dim finalX As Long
            finalX = srcDIB.GetDIBWidth * 4 - 4
            
            For y = 0 To srcDIB.GetDIBHeight - 1
            
                'Point our array at the current scanline
                tmpSA.pvData = dibPtr + y * dibStride
                
            For x = 0 To finalX Step 4
                
                b = imgPixels(x)
                g = imgPixels(x + 1)
                r = imgPixels(x + 2)
                
                If (r <> srcR) Or (g <> srcG) Or (b <> srcB) Then
                    srcDIB.UnwrapArrayFromDIB imgPixels
                    IsDIBSolidColor = False
                    Exit Function
                End If
                    
            Next x
            Next y
        
            'If we scanned all pixels without exiting prematurely, the DIB is a single solid color
            srcDIB.UnwrapArrayFromDIB imgPixels
            IsDIBSolidColor = True
            
        End If
            
    End If

End Function

'Given a 32-bpp DIB, return TRUE if any of its alpha bytes are non-255.
Public Function IsDIBTransparent(ByRef srcDIB As pdDIB) As Boolean
    
    IsDIBTransparent = False
    
    'Make sure the DIB exists and is 32-bpp
    If (Not srcDIB Is Nothing) Then
    
        If (srcDIB.GetDIBColorDepth = 32) And (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) Then
    
            'Loop through the image and compare each alpha value against 0 and 255. If a value doesn't
            ' match either of this, this is a non-binary alpha channel and it must be handled specially.
            Dim imgPixels() As Byte, tmpSA As SafeArray1D, dibPtr As Long, dibStride As Long
            srcDIB.WrapArrayAroundScanline imgPixels, tmpSA, 0
            dibPtr = tmpSA.pvData
            dibStride = tmpSA.cElements
            
            Dim x As Long, y As Long
            
            'Loop through the image, checking alphas as we go
            Dim finalX As Long
            finalX = srcDIB.GetDIBWidth * 4 - 4
            
            For y = 0 To srcDIB.GetDIBHeight - 1
            
                'Point our array at the current scanline
                tmpSA.pvData = dibPtr + y * dibStride
                
            For x = 0 To finalX Step 4
                
                'Retrieve the alpha value of the current pixel
                If (imgPixels(x + 3) <> 255) Then
                    IsDIBTransparent = True
                    Exit For
                End If
                
            Next x
                If IsDIBTransparent Then Exit For
            Next y
    
            srcDIB.UnwrapArrayFromDIB imgPixels
            
        '/Is DIB 32-bpp and non-zero in size?
        End If
    
    '/Does DIB exist?
    End If
        
End Function

'Given a DIB, return a pdColorCount object describing the contents of said image.
Public Function GetDIBColorCountObject(ByRef srcDIB As pdDIB, ByRef dstColorCount As pdColorCount, Optional ByVal countRGBA As Boolean = True) As Boolean

    Set dstColorCount = New pdColorCount
    dstColorCount.SetAlphaTracking countRGBA
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray1D
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBStride - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    Dim r As Long, g As Long, b As Long, a As Long
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Iterate through all pixels, counting unique values as we go.
    For y = initY To finalY
        srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
    For x = initX To finalX Step 4
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        a = imageData(x + 3)
        dstColorCount.AddColor r, g, b, a
    Next x
    Next y
    
    srcDIB.UnwrapArrayFromDIB imageData
    
    GetDIBColorCountObject = True
    
End Function

'Given a DIB, return a 2D Byte array of the DIB's luminance values.
' The optional toNormalize parameter guarantees that the output stretches from 0 to 255.
'NOTE: to improve performance, this function does not deal with alpha premultiplication *at all*.  It's up to the caller to handle that.
'ALSO NOTE: this function does not support progress reports, by design.
Public Function GetDIBGrayscaleMap(ByRef srcDIB As pdDIB, ByRef dstGrayArray() As Byte, Optional ByVal toNormalize As Boolean = True) As Boolean
    
    'Make sure the DIB exists
    If (srcDIB Is Nothing) Then Exit Function
    
    'Make sure the source DIB isn't empty
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
    
        'Create a local array and point it at the pixel data we want to operate on
        Dim imageData() As Byte, tmpSA As SafeArray1D
        
        'Support both 24-bpp and 32-bpp images
        Dim pxSize As Long
        pxSize = srcDIB.GetDIBColorDepth \ 8
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.GetDIBWidth - 1
        finalY = srcDIB.GetDIBHeight - 1
        
        'Prep the destination array
        ReDim dstGrayArray(initX To finalX, initY To finalY) As Byte
        
        Dim xStride As Long
        
        Dim r As Long, g As Long, b As Long, grayVal As Long
        Dim minVal As Long, maxVal As Long
        minVal = 255
        maxVal = 0
            
        'Now we can loop through each pixel in the image, converting values as we go
        For y = initY To finalY
            srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
        For x = initX To finalX
            
            'Get the source pixel color values
            xStride = x * pxSize
            b = imageData(xStride)
            g = imageData(xStride + 1)
            r = imageData(xStride + 2)
            
            'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
            grayVal = (218 * r + 732 * g + 74 * b) \ 1024
            
            'Cache the value
            dstGrayArray(x, y) = grayVal
            
            'If normalization has been requested, check max/min values now
            If toNormalize Then
                If (grayVal < minVal) Then minVal = grayVal
                If (grayVal > maxVal) Then maxVal = grayVal
            End If
            
        Next x
        Next y
        
        'Safely deallocate imageData()
        srcDIB.UnwrapArrayFromDIB imageData
        
        'If normalization was requested, and the data isn't already normalized, normalize it now
        If toNormalize And ((minVal > 0) Or (maxVal < 255)) Then
            
            Dim curRange As Long
            curRange = maxVal - minVal
            
            'Prevent DBZ errors
            If (curRange = 0) Then curRange = 1
            
            'Build a normalization lookup table
            Dim normalizedLookup(0 To 255) As Byte
            
            For x = 0 To 255
                
                grayVal = (CDbl(x - minVal) / CDbl(curRange)) * 255
                
                If (grayVal < 0) Then
                    grayVal = 0
                ElseIf (grayVal > 255) Then
                    grayVal = 255
                End If
                
                normalizedLookup(x) = grayVal
                
            Next x
            
            For y = initY To finalY
            For x = initX To finalX
                dstGrayArray(x, y) = normalizedLookup(dstGrayArray(x, y))
            Next x
            Next y
        
        End If
                
        GetDIBGrayscaleMap = True
        
    Else
        Debug.Print "WARNING! Non-existent DIB passed to getDIBGrayscaleMap."
        GetDIBGrayscaleMap = False
    End If

End Function

'Given a DIB, return a 2D Byte array of the DIB's luminance values.  An optional number of gray shades can also be specified (max 256).
'NOTE: to improve performance, this function does not deal with alpha premultiplication *at all*.  It's up to the caller to handle that.
'ALSO NOTE: this function does not support progress reports, by design.
Public Function GetDIBGrayscaleMapEx(ByRef srcDIB As pdDIB, ByRef dstGrayArray() As Byte, Optional ByVal numOfShades As Long = 256) As Boolean
    
    'Make sure the DIB exists
    If (srcDIB Is Nothing) Then Exit Function
    
    'Make sure the source DIB isn't empty
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
    
        'Create a local array and point it at the pixel data we want to operate on.
        ' CRITICALLY, this function *must* support both 32-bit and 24-bit DIBs, because the
        ' PNG exporter uses it to produce grayscale maps.
        Dim imageData() As Byte, tmpSA As SafeArray2D, bytesPerPixel As Long
        srcDIB.WrapArrayAroundDIB imageData, tmpSA
        If (srcDIB.GetDIBColorDepth = 32) Then bytesPerPixel = 4 Else bytesPerPixel = 3
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.GetDIBWidth - 1
        finalY = srcDIB.GetDIBHeight - 1
        
        'Prep the destination array
        ReDim dstGrayArray(initX To finalX, initY To finalY) As Byte
        
        Dim xStride As Long
        Dim r As Long, g As Long, b As Long, grayVal As Long
            
        'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
        Dim conversionFactor As Double
        If (numOfShades > 256) Then numOfShades = 256
        If (numOfShades < 2) Then numOfShades = 2
        conversionFactor = (255# / (numOfShades - 1))
        
        'Build a look-up table for our custom grayscale conversion results
        Dim gLookup(0 To 255) As Byte
        For x = 0 To 255
            grayVal = Int((CDbl(x) / conversionFactor) + 0.5) * conversionFactor
            If (grayVal > 255) Then grayVal = 255
            gLookup(x) = grayVal And &HFF&
        Next x
            
        'Now we can loop through each pixel in the image, converting values as we go
        For y = initY To finalY
        For x = initX To finalX
                
            'Get the source pixel color values
            xStride = x * bytesPerPixel
            b = imageData(xStride, y)
            g = imageData(xStride + 1, y)
            r = imageData(xStride + 2, y)
            
            'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
            grayVal = (218 * r + 732 * g + 74 * b) \ 1024
            
            'Cache the value
            dstGrayArray(x, y) = gLookup(grayVal)
            
        Next x
        Next y
        
        'Safely deallocate imageData()
        srcDIB.UnwrapArrayFromDIB imageData
        
        GetDIBGrayscaleMapEx = True
        
    Else
        Debug.Print "WARNING! Non-existent DIB passed to getDIBGrayscaleMap."
        GetDIBGrayscaleMapEx = False
    End If

End Function

'Given a DIB, return a 2D Byte array of the DIB's luminance values, side-by-side with preserved alpha values.
'NOTE: to improve performance, this function does not deal with alpha premultiplication *at all*.  It's up to the caller to handle that.
'ALSO NOTE: this function does not support progress reports, by design.
Public Function GetDIBGrayscaleAndAlphaMap(ByRef srcDIB As pdDIB, ByRef dstGrayAlphaArray() As Byte) As Boolean
    
    'Make sure the DIB exists
    If (srcDIB Is Nothing) Then Exit Function
    
    'Make sure the source DIB isn't empty
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
    
        'Create a local array and point it at the pixel data we want to operate on
        Dim imageData() As Byte, tmpSA As SafeArray2D
        srcDIB.WrapArrayAroundDIB imageData, tmpSA
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.GetDIBWidth - 1
        finalY = srcDIB.GetDIBHeight - 1
        
        'Prep the destination array
        ReDim dstGrayAlphaArray(initX To finalX * 2 + 1, initY To finalY) As Byte
        
        Dim xStride As Long
        
        Dim r As Long, g As Long, b As Long, grayVal As Long
        Dim minVal As Long, maxVal As Long
        minVal = 255
        maxVal = 0
            
        'Now we can loop through each pixel in the image, converting values as we go
        For y = initY To finalY
        For x = initX To finalX
                
            'Get the source pixel color values
            xStride = x * 4
            b = imageData(xStride, y)
            g = imageData(xStride + 1, y)
            r = imageData(xStride + 2, y)
            
            'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
            grayVal = (218 * r + 732 * g + 74 * b) \ 1024
            
            'Cache the value
            dstGrayAlphaArray(x * 2, y) = grayVal
            dstGrayAlphaArray(x * 2 + 1, y) = imageData(xStride + 3, y)
            
        Next x
        Next y
        
        'Safely deallocate imageData()
        srcDIB.UnwrapArrayFromDIB imageData
                
        GetDIBGrayscaleAndAlphaMap = True
        
    Else
        Debug.Print "WARNING! Non-existent DIB passed to getDIBGrayscaleMap."
        GetDIBGrayscaleAndAlphaMap = False
    End If

End Function

'Given a grayscale map (2D byte array), create a matching grayscale DIB from it.  The final DIB will be
' fully opaque, by design.  If you want to treat the grayscale values as alpha values, use the matching
' alpha-based function, below.
'
'(Note: this function does not support progress bar reports, by design.)
Public Function CreateDIBFromGrayscaleMap(ByRef dstDIB As pdDIB, ByRef srcGrayArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long) As Boolean
    
    'Create the DIB
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    If dstDIB.CreateBlank(arrayWidth, arrayHeight, 32, 0, 255) Then
    
        'Point a local array at the DIB
        Dim dstImageData() As RGBQuad, tmpSA As SafeArray1D
        
        Dim x As Long, y As Long, finalX As Long, finalY As Long
        finalX = dstDIB.GetDIBWidth - 1
        finalY = dstDIB.GetDIBHeight - 1
        
        'Prep a LUT
        Dim grayPal() As RGBQuad
        Palettes.GetPalette_Grayscale grayPal
        
        'Now we can loop through each pixel in the image, converting values as we go
        For y = 0 To finalY
            dstDIB.WrapRGBQuadArrayAroundScanline dstImageData, tmpSA, y
        For x = 0 To finalX
            dstImageData(x) = grayPal(srcGrayArray(x, y))
        Next x
        Next y
        
        'Safely deallocate imageData()
        dstDIB.UnwrapRGBQuadArrayFromDIB dstImageData
        dstDIB.SetInitialAlphaPremultiplicationState True
        CreateDIBFromGrayscaleMap = True
        
    Else
        Debug.Print "WARNING! Could not create blank DIB inside createDIBFromGrayscaleMap."
        CreateDIBFromGrayscaleMap = False
    End If

End Function

'Given a grayscale map (2D byte array), create a matching grayscale DIB from it.  The final DIB will have
' variable alpha, as determined by the incoming grayscale values, and the resulting DIB will be premultiplied.
' (Note: this function does not support progress bar reports, by design.)
Public Function CreateDIBFromGrayscaleMap_Alpha(ByRef dstDIB As pdDIB, ByRef srcGrayArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long) As Boolean
    
    'Create the DIB
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    If dstDIB.CreateBlank(arrayWidth, arrayHeight, 32, 0, 0) Then
        
        'Point a local array at the DIB
        Dim dstImageData() As Byte, tmpSA As SafeArray2D
        dstDIB.WrapArrayAroundDIB dstImageData, tmpSA
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = dstDIB.GetDIBWidth - 1
        finalY = dstDIB.GetDIBHeight - 1
        
        Dim xStride As Long, gValue As Byte, aValue As Byte
        
        'Create a lookup table of possible grayscale values, already premultiplied
        Dim gLookup(0 To 255) As Byte
        For x = 0 To 255
            gLookup(x) = Int(CDbl(x) * (CDbl(x) / 255#))
        Next x
        
        'Now we can loop through each pixel in the image, converting values as we go
        For y = initY To finalY
        For x = initX To finalX
            
            aValue = srcGrayArray(x, y)
            gValue = gLookup(aValue)
            
            xStride = x * 4
            dstImageData(xStride, y) = gValue
            dstImageData(xStride + 1, y) = gValue
            dstImageData(xStride + 2, y) = gValue
            dstImageData(xStride + 3, y) = aValue
            
        Next x
        Next y
        
        'Safely deallocate imageData()
        dstDIB.UnwrapArrayFromDIB dstImageData
        dstDIB.SetInitialAlphaPremultiplicationState True
        CreateDIBFromGrayscaleMap_Alpha = True
        
    Else
        Debug.Print "WARNING! Could not create blank DIB inside CreateDIBFromGrayscaleMap_Alpha."
        CreateDIBFromGrayscaleMap_Alpha = False
    End If

End Function

'Create a pdDIB object from a StdPicture object.  Returns TRUE if successful.
Public Function CreateDIBFromStdPicture(ByRef dstDIB As pdDIB, ByRef srcPicture As StdPicture, Optional ByVal forceWhiteBackground As Boolean = False) As Boolean
    
    CreateDIBFromStdPicture = False
    
    'Make sure the picture we're passed isn't empty
    If (Not srcPicture Is Nothing) Then
    
        'Failsafe check to ensure bitmap contents
        If (GetObjectType(srcPicture) = OBJ_BITMAP) Then
        
            'Select the picture's attributes into a bitmap object
            Dim tmpBitmap As GDI_Bitmap
            If (GetObject(srcPicture.Handle, Len(tmpBitmap), tmpBitmap) <> 0) Then
                
                'Use that bitmap object to create a new, blank DIB of the same size
                Dim targetColorDepth As Long
                If (tmpBitmap.BitsPerPixel = 32) Then targetColorDepth = 32 Else targetColorDepth = 24
                
                Set dstDIB = New pdDIB
                If dstDIB.CreateBlank(tmpBitmap.Width, tmpBitmap.Height, targetColorDepth, vbWhite, 255) Then
                
                    'Create a new DC
                    Dim tmpDC As Long
                    tmpDC = GDI.GetMemoryDC()
                    
                    'If successful, select the object into that DC
                    If (tmpDC <> 0) Then
                    
                        'Temporary holder for the object selection
                        Dim oldBitmap As Long
                        oldBitmap = SelectObject(tmpDC, srcPicture.Handle)
                        
                        'Use BitBlt to copy the pixel data to the target DIB
                        GDI.BitBltWrapper dstDIB.GetDIBDC, 0, 0, tmpBitmap.Width, tmpBitmap.Height, tmpDC, 0, 0, vbSrcCopy
                        
                        'Now that we have the pixel data, erase all temporary objects
                        SelectObject tmpDC, oldBitmap
                        GDI.FreeMemoryDC tmpDC
                        
                        'Finally, if the copied image contains an alpha channel (icons, PNGs, etc), it will be set against a
                        ' black background. We typically want the background to be white, so perform a composite if requested.
                        If forceWhiteBackground And (targetColorDepth = 32) Then dstDIB.CompositeBackgroundColor 255, 255, 255
                        
                        'Minimize GDI resources by freeing the auto-created DC
                        dstDIB.FreeFromDC
                        
                        CreateDIBFromStdPicture = True
                        
                    End If
                
                '/Created dstDIB succcessfully
                End If
            
            '/GetObject succeeded
            End If
            
        '/GetObjectType = BITMAP
        End If
    
    '/srcPicture non-null
    End If
    
End Function

'Convert a DIB to its grayscale equivalent.  (Note that this function does not support progress bar reports, by design.)
Public Function MakeDIBGrayscale(ByRef srcDIB As pdDIB, Optional ByVal numOfShades As Long = 256, Optional ByVal ignoreMagicMagenta As Boolean = True) As Boolean
    
    'Make sure the DIB exists
    If (srcDIB Is Nothing) Then Exit Function
    
    'Make sure the source DIB isn't empty
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
    
        'Create a local array and point it at the pixel data we want to operate on
        Dim imageData() As Byte, tmpSA As SafeArray1D
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = (srcDIB.GetDIBWidth - 1) * 4
        finalY = (srcDIB.GetDIBHeight - 1)
        
        Dim r As Long, g As Long, b As Long, a As Long, grayVal As Long
        
        'Premultiplication requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table
        ' for converting single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
        Dim alphaIsPremultiplied As Boolean: alphaIsPremultiplied = srcDIB.GetAlphaPremultiplication
        Dim applyPremult(0 To 255) As Single, removePremult(0 To 255) As Single, tmpAlphaModifier As Single
        
        If alphaIsPremultiplied Then
            For x = 0 To 255
                applyPremult(x) = x / 255!
                If (x <> 0) Then removePremult(x) = 255! / x Else removePremult(x) = 1!
            Next x
        End If
        
        'Validate number of grayscale values before continuing
        If (numOfShades < 2) Then numOfShades = 2
        If (numOfShades > 256) Then numOfShades = 256
        Dim conversionFactor As Double
        conversionFactor = (255 / (numOfShades - 1))
        
        'Build a look-up table for our custom grayscale conversion results
        Dim gLookup(0 To 255) As Byte
        For x = 0 To 255
            grayVal = Int(((CDbl(x) / conversionFactor) + 0.5) * conversionFactor)
            If (grayVal > 255) Then grayVal = 255
            gLookup(x) = grayVal
        Next x
            
        'Now we can loop through each pixel in the image, converting values as we go
        For y = initY To finalY
            srcDIB.WrapArrayAroundScanline imageData, tmpSA, y
        For x = initX To finalX Step 4
                
            'Get the source pixel color values
            b = imageData(x)
            g = imageData(x + 1)
            r = imageData(x + 2)
            
            'Remove premultiplication, as necessary
            If alphaIsPremultiplied Then
                a = imageData(x + 3)
                If (a <> 255) Then
                    tmpAlphaModifier = removePremult(a)
                    r = (r * tmpAlphaModifier)
                    g = (g * tmpAlphaModifier)
                    b = (b * tmpAlphaModifier)
                End If
            End If
            
            'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
            grayVal = (218 * r + 732 * g + 74 * b) \ 1024
            
            'If less than 256 shades are in play, calculate that now as well
            grayVal = gLookup(grayVal)
            
            'If alpha is premultiplied, calculate that now
            If alphaIsPremultiplied Then
                If (a <> 255) Then
                    tmpAlphaModifier = applyPremult(a)
                    grayVal = grayVal * tmpAlphaModifier
                End If
            End If
            
            If ignoreMagicMagenta Then
                imageData(x) = grayVal
                imageData(x + 1) = grayVal
                imageData(x + 2) = grayVal
            Else
                If (r <> 253) Or (g <> 0) Or (b <> 253) Then
                    imageData(x) = grayVal
                    imageData(x + 1) = grayVal
                    imageData(x + 2) = grayVal
                End If
            End If
            
        Next x
            srcDIB.UnwrapArrayFromDIB imageData
        Next y
        
        MakeDIBGrayscale = True
        
    Else
        MakeDIBGrayscale = False
    End If

End Function

'This function will calculate an "alpha-cutoff" for a 32bpp image.  This _Ex version (which is now the
' *only* version supported by PD) requires an input byte array, which will be initialized to the size
' of the image and filled with a copy of the image's new transparency data, if that cut-off were applied.
' (e.g. the return array will only include values of 0 or 255 for each pixel).
'
'Note that - by design - the DIB is not actually modified by this function.
Public Function ApplyAlphaCutoff_Ex(ByRef srcDIB As pdDIB, ByRef dstTransparencyTable() As Byte, Optional ByVal cutOff As Long = 127) As Boolean
    
    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long, finalX As Long, finalY As Long
            finalX = (srcDIB.GetDIBWidth - 1)
            finalY = (srcDIB.GetDIBHeight - 1)
            
            ReDim dstTransparencyTable(0 To finalX, 0 To finalY) As Byte
            
            If (cutOff = 0) Then
                FillMemory VarPtr(dstTransparencyTable(0)), srcDIB.GetDIBWidth * srcDIB.GetDIBHeight, 255
                ApplyAlphaCutoff_Ex = True
                Exit Function
            End If
            
            Dim iData() As Byte, tmpSA As SafeArray1D
            
            Dim chkAlpha As Byte
            
            'Loop through the image, checking alphas as we go
            For y = 0 To finalY
                srcDIB.WrapArrayAroundScanline iData, tmpSA, y
            For x = 0 To finalX
                
                chkAlpha = iData(x * 4 + 3)
                
                'If the alpha value is less than the cutoff, make it transparent.  Otherwise, make it opaque.
                If (chkAlpha < cutOff) Then
                    dstTransparencyTable(x, y) = 0
                Else
                    dstTransparencyTable(x, y) = 255
                End If
                
            Next x
            Next y
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            srcDIB.UnwrapArrayFromDIB iData
            
            ApplyAlphaCutoff_Ex = True
            
        End If
    End If
    
End Function

'This function will calculate and apply an "alpha-cutoff" for a 32bpp image, using rules appropriate
' for export to GIF format.  A transparency table is *not* required - work will be done directly on
' the DIB itself.  (That makes this function destructive, obviously.)
Public Function ApplyAlphaCutoff_Gif(ByRef srcDIB As pdDIB, Optional ByVal cutOff As Long = 127, Optional ByVal newBackgroundColor As Long = vbWhite) As Boolean
    
    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long, finalX As Long, finalY As Long, xLookup As Long
            finalX = (srcDIB.GetDIBWidth - 1)
            finalY = (srcDIB.GetDIBHeight - 1)
            
            Dim iData() As Byte, tmpSA As SafeArray1D
            
            'Premultiplied images can be processed more quickly
            Dim alphaPremultiplied As Boolean: alphaPremultiplied = srcDIB.GetAlphaPremultiplication
            
            'Premultiplication requires a lot of int/float conversions.  To speed things up, we'll use a LUT
            ' for converting single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
            Dim intToFloat(0 To 255) As Single
            Dim i As Long
            For i = 0 To 255
                If alphaPremultiplied Then
                    intToFloat(i) = 1 - (i / 255)
                Else
                    intToFloat(i) = i / 255
                End If
            Next i
            
            Dim chkR As Long, chkG As Long, chkB As Long, chkAlpha As Byte
            Dim tmpAlpha As Double
            
            'Retrieve RGB values from the new background color, which we'll use to composite
            ' any semi-transparent pixels.
            Dim backR As Long, backG As Long, backB As Long
            backR = Colors.ExtractRed(newBackgroundColor)
            backG = Colors.ExtractGreen(newBackgroundColor)
            backB = Colors.ExtractBlue(newBackgroundColor)
            
            'Loop through the image, checking alphas as we go
            For y = 0 To finalY
                srcDIB.WrapArrayAroundScanline iData, tmpSA, y
            For x = 0 To finalX
                
                xLookup = x * 4
                chkAlpha = iData(x * 4 + 3)
                
                'If this pixel is below the cutoff, erase it
                If (chkAlpha < cutOff) Then
                    iData(xLookup) = 0
                    iData(xLookup + 1) = 0
                    iData(xLookup + 2) = 0
                    iData(xLookup + 3) = 0
                
                'Otherwise, make this pixel fully opaque and composite it against the specified backcolor
                Else
                    
                    chkB = iData(xLookup)
                    chkG = iData(xLookup + 1)
                    chkR = iData(xLookup + 2)
                    chkAlpha = iData(xLookup + 3)
                    
                    If (chkAlpha < 255) Then
                        tmpAlpha = intToFloat(chkAlpha)
                        
                        If alphaPremultiplied Then
                            iData(xLookup) = backB * tmpAlpha + chkB
                            iData(xLookup + 1) = backG * tmpAlpha + chkG
                            iData(xLookup + 2) = backR * tmpAlpha + chkR
                        Else
                            iData(xLookup) = Colors.BlendColors(backB, chkB, tmpAlpha)
                            iData(xLookup + 1) = Colors.BlendColors(backG, chkG, tmpAlpha)
                            iData(xLookup + 2) = Colors.BlendColors(backR, chkR, tmpAlpha)
                        End If
                        
                        iData(xLookup + 3) = 255
                        
                    End If
                    
                End If
                
            Next x
            Next y
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            srcDIB.UnwrapArrayFromDIB iData
            
            ApplyAlphaCutoff_Gif = True
        
        '/end DIB contains valid pixel data
        End If
    
    '/end DIB is 32-bpp
    End If
    
End Function

'Given two DIBs and a pre-initialized transparency table, update the transparency table to make
' any duplicate pixels between the two images transparent.  This is used during GIF export to
' minimize file size, by retaining color data from the previous frame and making those pixels
' transparent in the current frame.
'
'IMPORTANTLY: the dstTransparencyTable() array MUST ALREADY BE INITIALIZED to the size of topDIB,
' and topDIB MUST BE smaller (and positioned inclusively) or the same size as bottomDIB.  If it is
' larger or positioned outside of bottomDIB, this function will fail, by design.
'
'Note that - by design - neither DIB is modified by this function.  Only the transparency table
' is modified.
'
'If activated, the optional "autoDenoise" parameter will change the algorithm to *not* blank out
' pixels unless at least two of them are touching (under the assumption that introducing 1-px
' noise will hurt most compression schemes).  Note that the denoiser does not work across
' scanline boundaries, at present, but could be modified to do so.
Public Function ApplyAlpha_DuplicatePixels(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByRef dstTransparencyTable() As Byte, Optional ByVal topOffsetX As Long = 0, Optional ByVal topOffsetY As Long = 0, Optional ByVal autoDenoise As Boolean = False) As Boolean
    
    If (topDIB Is Nothing) Then Exit Function
    If (bottomDIB Is Nothing) Then Exit Function
    
    'Additional failsafe checks
    If (topDIB.GetDIBColorDepth = 32) And (bottomDIB.GetDIBColorDepth = 32) Then
    If (topDIB.GetDIBDC <> 0) And (topDIB.GetDIBWidth <> 0) And (topDIB.GetDIBHeight <> 0) Then
    If (bottomDIB.GetDIBDC <> 0) And (bottomDIB.GetDIBWidth <> 0) And (bottomDIB.GetDIBHeight <> 0) Then
    If (topOffsetX + topDIB.GetDIBWidth <= bottomDIB.GetDIBWidth) And (topOffsetY + topDIB.GetDIBHeight <= bottomDIB.GetDIBHeight) Then
        
        Dim x As Long, y As Long, finalX As Long, finalY As Long
        finalX = (topDIB.GetDIBWidth - 1)
        finalY = (topDIB.GetDIBHeight - 1)
        
        'Failsafe check for single-pixel images
        If (finalX < 1) Then
            ApplyAlpha_DuplicatePixels = True
            Exit Function
        End If
        
        Dim srcDataTop() As Long, tmpSATop As SafeArray1D
        Dim srcDataBottom() As Long, tmpSABottom As SafeArray1D
        Dim origAlpha As Byte
        
        'Loop through the image, checking alphas as we go
        For y = 0 To finalY
            topDIB.WrapLongArrayAroundScanline srcDataTop, tmpSATop, y
            bottomDIB.WrapLongArrayAroundScanline srcDataBottom, tmpSABottom, y + topOffsetY
        For x = 0 To finalX
            
            'We use two strategies here, based on whether autoDenoise is active
            If autoDenoise Then
                
                origAlpha = dstTransparencyTable(x, y)
                
                'First, see if this pixel will be blanked at all
                If srcDataTop(x) = srcDataBottom(x + topOffsetX) Then
                
                    If (x > 0) Then
                        
                        'Check left pixel regardless; if it matches, we can blank the pixel immediately
                        If srcDataTop(x - 1) = srcDataBottom(x + topOffsetX - 1) Then
                            dstTransparencyTable(x, y) = 0
                        
                        'Left pixel doesn't match; try right pixel
                        Else
                            If (x < finalX) Then
                                If srcDataTop(x + 1) = srcDataBottom(x + topOffsetX + 1) Then dstTransparencyTable(x, y) = 0
                            End If
                        End If
                        
                    'x = 0, check right pixel before blanking
                    Else
                        If srcDataTop(x + 1) = srcDataBottom(x + topOffsetX + 1) Then dstTransparencyTable(x, y) = 0
                    End If
                
                '/do nothing if pixels don't match
                End If
                
            'When autoDenoise is disabled, just make matching pixels transparent
            Else
                If srcDataTop(x) = srcDataBottom(x + topOffsetX) Then dstTransparencyTable(x, y) = 0
            End If
            
        Next x
        Next y

        'With our alpha channel complete, point iData() away from the DIB and deallocate it
        topDIB.UnwrapLongArrayFromDIB srcDataTop
        bottomDIB.UnwrapLongArrayFromDIB srcDataBottom
        
        ApplyAlpha_DuplicatePixels = True
    
    'Failsafe checks
    End If
    End If
    End If
    End If
    
End Function

'Given a bottom DIB and a transparency table (for a "top" DIB), return TRUE if the table contains
' transparency where the bottom DIB does not.  This is used during animated GIF/PNG export to
' determine whether we need to clear the previous frame before displaying the current one.
' (By default, we don't clear the previous frame unless necessary, as we can actually share data
' between frames for improved compression.)
'
'IMPORTANTLY: the dstTransparencyTable() array MUST ALREADY BE INITIALIZED to the size of topDIB,
' and topDIB MUST BE smaller (and positioned inclusively) or the same size as bottomDIB.  If it is
' larger or positioned outside of bottomDIB, this function will fail, by design.
'
'RETURNS: TRUE if the top image's alpha is incompatible with the bottom image's; you MUST clear the
'         previous frame if this function returns TRUE.
'
'Note that - by design - this function does not modify any of the input parameters.  They are all CONST.
Public Function CheckAlpha_DuplicatePixels(ByRef bottomDIB As pdDIB, ByRef topDIB As pdDIB, ByRef trnsTable() As Byte, Optional ByVal topOffsetX As Long = 0, Optional ByVal topOffsetY As Long = 0) As Boolean
    
    CheckAlpha_DuplicatePixels = False
    
    If (bottomDIB Is Nothing) Then
        Debug.Print "dib is nothing"
        Exit Function
    End If
    
    'Additional failsafe checks
    If (bottomDIB.GetDIBColorDepth = 32) Then
    If (bottomDIB.GetDIBDC <> 0) And (bottomDIB.GetDIBWidth <> 0) And (bottomDIB.GetDIBHeight <> 0) Then
        
        Dim x As Long, y As Long, finalX As Long, finalY As Long
        finalX = (topDIB.GetDIBWidth - 1)
        finalY = (topDIB.GetDIBHeight - 1)
        
        Dim bottomPixels() As Byte, tmpSA1 As SafeArray1D
        Dim topPixels() As Byte, tmpSA2 As SafeArray1D
        
        Dim xCheck As Long
        
        'Loop through the image, checking alphas as we go
        For y = 0 To finalY
            bottomDIB.WrapArrayAroundScanline bottomPixels, tmpSA1, y + topOffsetY
            topDIB.WrapArrayAroundScanline topPixels, tmpSA2, y
        For x = 0 To finalX
            
            'If the top image has any amount of transparency in this pixel, but the bottom image
            ' *does not*, exit immediately.
            If (trnsTable(x, y) < 255) And (bottomPixels((x + topOffsetX) * 4 + 3) <> 0) Then
                
                'Also confirm that the pixels in both positions are *not* identical (because if they are,
                ' they will be getting blanked anyway!)
                xCheck = (x + topOffsetX) * 4
                If (topPixels(x * 4) <> bottomPixels(xCheck)) Or (topPixels(x * 4 + 1) <> bottomPixels(xCheck + 1)) Or (topPixels(x * 4 + 2) <> bottomPixels(xCheck + 2)) Or (topPixels(x * 4 + 3) <> bottomPixels(xCheck + 3)) Then
                    topDIB.UnwrapArrayFromDIB topPixels
                    bottomDIB.UnwrapArrayFromDIB bottomPixels
                    CheckAlpha_DuplicatePixels = True
                    Exit Function
                End If
            End If
            
        Next x
        Next y
        
        topDIB.UnwrapArrayFromDIB topPixels
        bottomDIB.UnwrapArrayFromDIB bottomPixels
        
    'Failsafe checks
    Else
        Debug.Print "bad DC or dimensions"
    End If
    Else
        Debug.Print "bad color depth"
    End If
    
End Function

'(See comments for ApplyAlphaCutoff_Ex, above; this function is identical, but using a target color instead of an alpha cutoff.)
'Returns: TRUE if at least one pixel was made transparent; FALSE otherwise
Public Function MakeColorTransparent_Ex(ByRef srcDIB As pdDIB, ByRef dstTransparencyTable() As Byte, ByVal srcColor As Long) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long, finalX As Long, finalY As Long, xLookup As Long
            finalX = (srcDIB.GetDIBWidth - 1)
            finalY = (srcDIB.GetDIBHeight - 1)
            
            ReDim dstTransparencyTable(0 To finalX, 0 To finalY) As Byte
            
            Dim chkR As Long, chkG As Long, chkB As Long
            Dim targetR As Long, targetG As Long, targetB As Long
            targetR = Colors.ExtractRed(srcColor)
            targetG = Colors.ExtractGreen(srcColor)
            targetB = Colors.ExtractBlue(srcColor)
            
            'This function requires unpremultiplied alpha
            Dim needToResetPremultiplication As Boolean: needToResetPremultiplication = False
            If srcDIB.GetAlphaPremultiplication Then
                srcDIB.SetAlphaPremultiplication False
                needToResetPremultiplication = True
            End If
            
            Dim iData() As Byte, tmpSA As SafeArray2D
            srcDIB.WrapArrayAroundDIB iData, tmpSA
            
            'We will return TRUE if at least one pixel is made transparent
            Dim stillLookingForTransPixel As Boolean: stillLookingForTransPixel = True
                
            'Loop through the image, checking alphas as we go
            For y = 0 To finalY
            For x = 0 To finalX
                
                xLookup = x * 4
                chkB = iData(xLookup, y)
                chkG = iData(xLookup + 1, y)
                chkR = iData(xLookup + 2, y)
                
                'There are basically two options here:
                ' 1) This pixel matches our target color, and needs to be made fully transparent
                ' 2) This pixel does not match our target color, and needs to be made fully opaque
                If (targetR = chkR) And (targetG = chkG) And (targetB = chkB) Then
                
                    'This pixel is a match!  Make it fully transparent.
                    dstTransparencyTable(x, y) = 0
                    
                    'Remember this location if we haven't already
                    If stillLookingForTransPixel Then stillLookingForTransPixel = False
                    
                Else
                    dstTransparencyTable(x, y) = 255
                End If
                
            Next x
            Next y
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            srcDIB.UnwrapArrayFromDIB iData
            
            If needToResetPremultiplication Then srcDIB.SetAlphaPremultiplication True
            MakeColorTransparent_Ex = Not stillLookingForTransPixel
            
        End If
    Else
        Debug.Print "WARNING!  pdDIB.MakeColorTransparent_Ex() requires a 32-bpp DIB to operate correctly."
    End If
    
End Function

'This function will return a single channel from a 32bpp image.  An input byte array is required; it will be initialized
' to the size of the image and filled with a copy of the image's values in the specified channel.  ChannelOffset must
' be a value between 0 and 3 (obviously); the function will crash otherwise.
'
'Note that - by design - the DIB is not actually modified by this function.
Public Function RetrieveSingleChannel(ByRef srcDIB As pdDIB, ByRef dstChannelTable() As Byte, ByVal channelOffset As Long) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long
            
            Dim finalX As Long, finalY As Long
            finalX = (srcDIB.GetDIBWidth - 1)
            finalY = (srcDIB.GetDIBHeight - 1)
            
            ReDim dstChannelTable(0 To finalX, 0 To finalY) As Byte
            
            Dim iData() As Byte, tmpSA As SafeArray2D
            srcDIB.WrapArrayAroundDIB iData, tmpSA
                
            'Loop through the image, checking alphas as we go
            For y = 0 To finalY
            For x = 0 To finalX
                dstChannelTable(x, y) = iData(x * 4 + channelOffset, y)
            Next x
            Next y
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            srcDIB.UnwrapArrayFromDIB iData
            
            RetrieveSingleChannel = True
            
        End If
    End If
    
End Function

'This function will return the alpha channel of a 32bpp image.  An input byte array is required; it will be initialized
' to the size of the image and filled with a copy of the image's alpha values.  Optionally, you can pass a pointer to a
' RectF struct that defines the target region; if non-null, only that region will be stored - IMPORTANTLY, if you use this
' functionality, you must also supply an identical RectF structure to the ApplyTransparencyTable function.  If you don't,
' the results will be incorrect.
'
'Note that - by design - the DIB is not actually modified by this function.
Public Function RetrieveTransparencyTable(ByRef srcDIB As pdDIB, ByRef dstTransparencyTable() As Byte, Optional ByVal ptrToRectF As Long = 0) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long
            Dim allowedToProceed As Boolean: allowedToProceed = True
            
            Dim initX As Long, initY As Long, finalX As Long, finalY As Long
            If (ptrToRectF = 0) Then
                initX = 0
                initY = 0
                finalX = (srcDIB.GetDIBWidth - 1)
                finalY = (srcDIB.GetDIBHeight - 1)
            Else
            
                Dim tmpRectF As RectF
                CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
                PDMath.GetIntClampedRectF tmpRectF
                initX = tmpRectF.Left
                initY = tmpRectF.Top
                finalX = initX + tmpRectF.Width
                finalY = initY + tmpRectF.Height
                
                'Perform a bunch of failsafe checks on boundaries
                If (initX > (srcDIB.GetDIBWidth - 1)) Then allowedToProceed = False
                If (initY > (srcDIB.GetDIBHeight - 1)) Then allowedToProceed = False
                If (finalX < initX) Then allowedToProceed = False
                If (finalY < initY) Then allowedToProceed = False
                
                If allowedToProceed Then
                    If (initX < 0) Then initX = 0
                    If (initY < 0) Then initY = 0
                    If (finalX > (srcDIB.GetDIBWidth - 1)) Then finalX = srcDIB.GetDIBWidth - 1
                    If (finalY > (srcDIB.GetDIBHeight - 1)) Then finalY = srcDIB.GetDIBHeight - 1
                End If
                
            End If
            
            If allowedToProceed Then
            
                ReDim dstTransparencyTable(initX To finalX, initY To finalY) As Byte
                
                Dim iData() As Byte, tmpSA As SafeArray2D
                srcDIB.WrapArrayAroundDIB iData, tmpSA
                    
                'Loop through the image, checking alphas as we go
                For y = initY To finalY
                For x = initX To finalX
                    dstTransparencyTable(x, y) = iData(x * 4 + 3, y)
                Next x
                Next y
        
                'With our alpha channel complete, point iData() away from the DIB and deallocate it
                srcDIB.UnwrapArrayFromDIB iData
                
            End If
            
            RetrieveTransparencyTable = allowedToProceed
            
        End If
    End If
    
End Function

'Given a transparency table (e.g. a byte array at the same dimensions as the image, containing the desired per-pixel alpha values),
' apply said transparency table to the current DIB.  Optionally, you can pass a pointer to a RectF struct that defines the
' target region; if non-null, only that region will be updated - IMPORTANTLY, if you use this functionality, you must also
' supply an identical RectF structure to the RetrieveTransparencyTable function.  If you don't, the results will be incorrect.
Public Function ApplyTransparencyTable(ByRef srcDIB As pdDIB, ByRef srcTransparencyTable() As Byte, Optional ByVal ptrToRectF As Long = 0) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long
            Dim allowedToProceed As Boolean: allowedToProceed = True
            
            Dim initX As Long, initY As Long, finalX As Long, finalY As Long
            If (ptrToRectF = 0) Then
                initX = 0
                initY = 0
                finalX = (srcDIB.GetDIBWidth - 1)
                finalY = (srcDIB.GetDIBHeight - 1)
            Else
            
                Dim tmpRectF As RectF
                CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
                PDMath.GetIntClampedRectF tmpRectF
                initX = tmpRectF.Left
                initY = tmpRectF.Top
                finalX = initX + tmpRectF.Width
                finalY = initY + tmpRectF.Height
                
                'Perform a bunch of failsafe checks on boundaries
                If (initX > (srcDIB.GetDIBWidth - 1)) Then allowedToProceed = False
                If (initY > (srcDIB.GetDIBHeight - 1)) Then allowedToProceed = False
                If (finalX < initX) Then allowedToProceed = False
                If (finalY < initY) Then allowedToProceed = False
                
                If allowedToProceed Then
                    If (initX < 0) Then initX = 0
                    If (initY < 0) Then initY = 0
                    If (finalX > (srcDIB.GetDIBWidth - 1)) Then finalX = srcDIB.GetDIBWidth - 1
                    If (finalY > (srcDIB.GetDIBHeight - 1)) Then finalY = srcDIB.GetDIBHeight - 1
                End If
                
            End If
            
            If allowedToProceed Then
                
                Dim restorePremultiplication As Boolean: restorePremultiplication = False
                If srcDIB.GetAlphaPremultiplication Then
                    srcDIB.SetAlphaPremultiplication False, , ptrToRectF
                    restorePremultiplication = True
                End If
                
                Dim iData() As Byte, tmpSA As SafeArray1D
                srcDIB.WrapArrayAroundScanline iData, tmpSA, 0
                
                Dim dibPtr As Long, dibStride As Long
                dibPtr = tmpSA.pvData
                dibStride = tmpSA.cElements
                    
                'Loop through the image, checking alphas as we go
                For y = initY To finalY
                    tmpSA.pvData = dibPtr + y * dibStride
                For x = initX To finalX
                    iData(x * 4 + 3) = srcTransparencyTable(x, y)
                Next x
                Next y
                
                'With our alpha channel complete, point iData() away from the DIB and deallocate it
                srcDIB.UnwrapArrayFromDIB iData
                
                If restorePremultiplication Then srcDIB.SetAlphaPremultiplication True, , ptrToRectF
                
            End If
            
            ApplyTransparencyTable = allowedToProceed
            
        End If
    Else
        Debug.Print "WARNING!  pdDIB.ApplyBinaryTransparencyTable() requires a 32-bpp DIB to operate correctly."
    End If
    
End Function

'Given a binary transparency table (e.g. a byte array at the same dimensions as the image, with only values 0 or 255),
' apply said transparency table to the current DIB.  This function contains special optimizations over the generic
' ApplyTransparencyTable function, so use it for a speed boost if you know the transparency table is limited to full
' opacity and/or full transparency.
Public Function ApplyBinaryTransparencyTable(ByRef srcDIB As pdDIB, ByRef srcTransparencyTable() As Byte, Optional ByVal newBackgroundColor As Long = vbWhite) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long, finalX As Long, finalY As Long, xLookup As Long
            finalX = (srcDIB.GetDIBWidth - 1)
            finalY = (srcDIB.GetDIBHeight - 1)
            
            Dim iData() As Byte, tmpSA As SafeArray2D
            srcDIB.WrapArrayAroundDIB iData, tmpSA
            
            Dim alphaPremultiplied As Boolean: alphaPremultiplied = srcDIB.GetAlphaPremultiplication
            
            Dim chkR As Long, chkG As Long, chkB As Long, chkAlpha As Byte
            Dim tmpAlpha As Double
            
            'Premultiplication requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table
            ' for converting single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
            Dim intToFloat() As Single
            ReDim intToFloat(0 To 255) As Single
            Dim i As Long
            For i = 0 To 255
                If alphaPremultiplied Then
                    intToFloat(i) = 1 - (i / 255)
                Else
                    intToFloat(i) = i / 255
                End If
            Next i
            
            'Retrieve RGB values from the new background color, which we'll use to composite semi-transparent pixels
            Dim backR As Long, backG As Long, backB As Long
            backR = Colors.ExtractRed(newBackgroundColor)
            backG = Colors.ExtractGreen(newBackgroundColor)
            backB = Colors.ExtractBlue(newBackgroundColor)
                
            'Loop through the image, checking alphas as we go
            For y = 0 To finalY
            For x = 0 To finalX
                
                xLookup = x * 4
                
                'If the transparency table is 0, erase this pixel
                If srcTransparencyTable(x, y) = 0 Then
                    iData(xLookup, y) = 0
                    iData(xLookup + 1, y) = 0
                    iData(xLookup + 2, y) = 0
                    iData(xLookup + 3, y) = 0
                
                'Otherwise, make this pixel fully opaque and composite it against the specified backcolor
                Else
                    chkB = iData(xLookup, y)
                    chkG = iData(xLookup + 1, y)
                    chkR = iData(xLookup + 2, y)
                    chkAlpha = iData(xLookup + 3, y)
                    
                    If (chkAlpha <> 255) Then
                        tmpAlpha = intToFloat(chkAlpha)
                        
                        If alphaPremultiplied Then
                            iData(xLookup, y) = backB * tmpAlpha + chkB
                            iData(xLookup + 1, y) = backG * tmpAlpha + chkG
                            iData(xLookup + 2, y) = backR * tmpAlpha + chkR
                        Else
                            iData(xLookup, y) = Colors.BlendColors(backB, chkB, tmpAlpha)
                            iData(xLookup + 1, y) = Colors.BlendColors(backG, chkG, tmpAlpha)
                            iData(xLookup + 2, y) = Colors.BlendColors(backR, chkR, tmpAlpha)
                        End If
                        
                        iData(xLookup + 3, y) = 255
                    End If
                    
                End If
                
            Next x
            Next y
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            srcDIB.UnwrapArrayFromDIB iData
            
            ApplyBinaryTransparencyTable = True
            
        End If
    Else
        Debug.Print "WARNING!  pdDIB.ApplyBinaryTransparencyTable() requires a 32-bpp DIB to operate correctly."
    End If
    
End Function

'Given a binary transparency table (e.g. a byte array at the same dimensions as the image, with only values 0 or 255),
' apply said transparency table to the current DIB.  Unlike the normal function, this one replaces transparent pixels
' with a designated transparent color.  (This is useful when exporting PNGs.)
Public Function ApplyBinaryTransparencyTableColor(ByRef srcDIB As pdDIB, ByRef srcTransparencyTable() As Byte, ByVal trnsColor As Long, Optional ByVal newBackgroundColor As Long = vbWhite) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long, finalX As Long, finalY As Long, xLookup As Long
            finalX = (srcDIB.GetDIBWidth - 1)
            finalY = (srcDIB.GetDIBHeight - 1)
            
            Dim iData() As Byte, tmpSA As SafeArray1D
            Dim alphaPremultiplied As Boolean: alphaPremultiplied = srcDIB.GetAlphaPremultiplication
            
            Dim chkR As Long, chkG As Long, chkB As Long, chkAlpha As Byte
            Dim tmpAlpha As Double
            
            'Transparent pixels are replaced with the designated transparent color
            Dim trnsR As Long, trnsG As Long, trnsB As Long
            trnsR = Colors.ExtractRed(trnsColor)
            trnsG = Colors.ExtractGreen(trnsColor)
            trnsB = Colors.ExtractBlue(trnsColor)
            
            'Premultiplication requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table
            ' for converting single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
            Dim intToFloat() As Single
            ReDim intToFloat(0 To 255) As Single
            Dim i As Long
            For i = 0 To 255
                If alphaPremultiplied Then
                    intToFloat(i) = 1! - (CSng(i) / 255!)
                Else
                    intToFloat(i) = CSng(i) / 255!
                End If
            Next i
            
            'Retrieve RGB values from the new background color, which we'll use to composite semi-transparent pixels
            Dim backR As Long, backG As Long, backB As Long
            backR = Colors.ExtractRed(newBackgroundColor)
            backG = Colors.ExtractGreen(newBackgroundColor)
            backB = Colors.ExtractBlue(newBackgroundColor)
            
            'Loop through the image, checking alphas as we go
            For y = 0 To finalY
                srcDIB.WrapArrayAroundScanline iData, tmpSA, y
            For x = 0 To finalX
                
                xLookup = x * 4
                
                'If the transparency table is 0, replace this pixel with the designated transparent color
                If (srcTransparencyTable(x, y) = 0) Then
                    iData(xLookup) = trnsB
                    iData(xLookup + 1) = trnsG
                    iData(xLookup + 2) = trnsR
                    iData(xLookup + 3) = 255
                
                'Otherwise, make this pixel fully opaque and composite it against the specified backcolor
                Else
                    chkB = iData(xLookup)
                    chkG = iData(xLookup + 1)
                    chkR = iData(xLookup + 2)
                    chkAlpha = iData(xLookup + 3)
                    
                    If (chkAlpha <> 255) Then
                    
                        tmpAlpha = intToFloat(chkAlpha)
                        
                        If alphaPremultiplied Then
                            iData(xLookup) = backB * tmpAlpha + chkB
                            iData(xLookup + 1) = backG * tmpAlpha + chkG
                            iData(xLookup + 2) = backR * tmpAlpha + chkR
                        Else
                            iData(xLookup) = Colors.BlendColors(backB, chkB, tmpAlpha)
                            iData(xLookup + 1) = Colors.BlendColors(backG, chkG, tmpAlpha)
                            iData(xLookup + 2) = Colors.BlendColors(backR, chkR, tmpAlpha)
                        End If
                        
                        iData(xLookup + 3) = 255
                        
                    End If
                    
                End If
                
            Next x
            Next y
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            srcDIB.UnwrapArrayFromDIB iData
            
            ApplyBinaryTransparencyTableColor = True
            
        End If
    Else
        Debug.Print "WARNING!  pdDIB.ApplyBinaryTransparencyTable() requires a 32-bpp DIB to operate correctly."
    End If
    
End Function

'Forcibly colorize a DIB.  Alpha is preserved by this function.
'Returns: TRUE if successful; FALSE otherwise
Public Function ColorizeDIB(ByRef srcDIB As pdDIB, ByVal newColor As Long) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long, finalX As Long, finalY As Long
            finalX = (srcDIB.GetDIBWidth - 1)
            finalY = (srcDIB.GetDIBHeight - 1)
            
            Dim chkA As Byte
            
            Dim targetR As Long, targetG As Long, targetB As Long
            targetR = Colors.ExtractRed(newColor)
            targetG = Colors.ExtractGreen(newColor)
            targetB = Colors.ExtractBlue(newColor)
            
            'Construct lookup tables with premultiplied RGB values.  This prevents us from needing
            ' to un-premultiply values in advance, and post-premultiply values afterward.
            Dim lTable(0 To 255) As Long
            
            Dim tmpQuad As RGBQuad, aFloat As Double
            Const ONE_DIV_255 As Double = 1# / 255#
            
            For x = 0 To 255
                tmpQuad.Alpha = x
                aFloat = CDbl(x) * ONE_DIV_255
                tmpQuad.Blue = targetB * aFloat
                tmpQuad.Green = targetG * aFloat
                tmpQuad.Red = targetR * aFloat
                GetMem4 VarPtr(tmpQuad), ByVal VarPtr(lTable(x))
            Next x
            
            'We're now going to do something kinda weird.  Since VB lacks bit-shift operators (ugh),
            ' we're going to wrap both a byte-type and long-type array around the image data.  The byte-type
            ' array is faster for retrieving individual color channels (alpha, in this case), while the
            ' long-type array is faster for setting entire pixel values (all four RGBA bytes at once).
            Dim imgDataL() As Long, tmpSAL As SafeArray1D
            Dim imgDataB() As Byte, tmpSAB As SafeArray1D, dibPtr As Long, dibStride As Long
            srcDIB.WrapArrayAroundScanline imgDataB, tmpSAB, 0
            dibPtr = tmpSAB.pvData
            dibStride = tmpSAB.cElements
            
            srcDIB.WrapLongArrayAroundScanline imgDataL, tmpSAL, 0
            
            'Loop through the image, checking alphas as we go
            For y = 0 To finalY
                tmpSAB.pvData = dibPtr + dibStride * y
                tmpSAL.pvData = dibPtr + dibStride * y
            For x = 0 To finalX
                chkA = imgDataB(x * 4 + 3)
                imgDataL(x) = lTable(chkA)
            Next x
            Next y
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            srcDIB.UnwrapArrayFromDIB imgDataB
            srcDIB.UnwrapLongArrayFromDIB imgDataL
            srcDIB.SetInitialAlphaPremultiplicationState True
            
            ColorizeDIB = True
            
        End If
    Else
        Debug.Print "WARNING!  DIBs.ColorizeDIB() requires a 32-bpp DIB to operate correctly."
    End If
    
End Function

'This function is used specifically for optimizing animation frames.  PD can perform an optimization
' called "pixel blanking", which involves making pixels transparent if they are identical to the
' previous frame's pixels.  It is difficult to predict the cost vs benefit of this optimization
' because sometimes pixel blanking creates a lot of noise, which actually compresses poorly, while
' other times it can provide massive gains.  To try and maximize our benefits from pixel blanking,
' PD's animated GIF exporter will produce two copies of an exported frame: a non-pixel-blanked one,
' and a maximally-pixel-blanked one.  This function will then loop through each scanline and pick
' the one with minimal entropy (using a shorthand estimator by the PNG working group).  Only that
' line gets copied into the destination DIB.  The result is generally a mixed-blanking frame that
' compresses better than either of the source DIBs.
Public Function MakeMinimalEntropyScanlines(ByRef srcData1() As Byte, ByRef srcData2() As Byte, ByVal dataWidth As Long, ByVal dataHeight As Long, ByRef dstData() As Byte) As Boolean

    'Ensure the destination array exists and is the correct size
    ReDim dstData(0 To dataWidth - 1, 0 To dataHeight - 1) As Byte
    
    'We now split into two possible sub-tests, which vary their behavior based on the size of the
    ' incoming dataset(s).  (If the set is too small, DEFLATE is a poor predictor of entropy
    ' because there's not enough data to build a meaningful compression table; in these instances,
    ' we drop back to a simpler RLE scheme.)
    If (dataWidth * dataHeight < 128) Then
        MakeMinimalEntropyScanlines = MakeMinimalEntropy_Small(srcData1, srcData2, dataWidth, dataHeight, dstData)
    Else
        MakeMinimalEntropyScanlines = MakeMinimalEntropy_Big(srcData1, srcData2, dataWidth, dataHeight, dstData)
    End If
    
End Function

Private Function MakeMinimalEntropy_Small(ByRef srcData1() As Byte, ByRef srcData2() As Byte, ByVal dataWidth As Long, ByVal dataHeight As Long, ByRef dstData() As Byte) As Boolean

    'For small data sets, we use a simple RLE-based entropy detector.  Whichever source dataset
    ' currently maintains the longest run of identical bytes gets sent to the destination.
    Dim ent1 As Long, ent2 As Long
    ent1 = 0
    ent2 = 0
    
    Dim cmpPrevious1 As Long, cmpPrevious2 As Long
    Dim pX As Long, pY As Long
    
    'Iterate through the image, tracking consecutive matching pixels as we go
    Dim x As Long, y As Long
    For y = 0 To dataHeight - 1
        
        For x = 0 To dataWidth - 1
            
            'Determine a previous pixel value for both data sets, accounting for scanline wrapping
            ' (compression generally treats the data as a 1D dataset)
            If (x = 0) Then
                If (y > 0) Then
                    pX = dataWidth - 1
                    pY = y - 1
                Else
                    pX = 0
                    pY = 0
                End If
            Else
                pX = x - 1
                pY = y
                
            End If
            
            cmpPrevious1 = srcData1(pX, pY)
            cmpPrevious2 = srcData2(pX, pY)
            
            'If this is *not* the first pixel, store a pixel from whichever data set
            ' has the longest run of identical pixels.
            If (x > 0) Or (y > 0) Then
                If (srcData1(x, y) = cmpPrevious1) Then ent1 = ent1 + 1 Else ent1 = 0
                If (srcData2(x, y) = cmpPrevious2) Then ent2 = ent2 + 1 Else ent2 = 0
                
                'Whichever pixel value is higher determines what we store for the *previous* pixel.
                ' (If both are 0, it doesn't matter what gets stored; the previous pixel doesn't
                ' match either of these ones, so there's no obvious winner.)
                If (ent1 >= ent2) Then
                    dstData(x, y) = srcData1(x, y)
                    If (x = 1) And (y = 0) Then dstData(0, 0) = srcData1(0, 0)
                Else
                    dstData(x, y) = srcData2(x, y)
                    If (x = 1) And (y = 0) Then dstData(0, 0) = srcData2(0, 0)
                End If
                
            End If
            
        Next x
        
    Next y
    
    'Handle the final pixel manually
    pX = dataWidth - 1
    pY = dataHeight - 1
    If (ent1 >= ent2) Then
        dstData(pX, pY) = srcData1(pX, pY)
    Else
        dstData(pX, pY) = srcData2(pX, pY)
    End If
    
    MakeMinimalEntropy_Small = True
    
End Function

Private Function MakeMinimalEntropy_Big(ByRef srcData1() As Byte, ByRef srcData2() As Byte, ByVal dataWidth As Long, ByVal dataHeight As Long, ByRef dstData() As Byte) As Boolean
    
    'On larger data sets, an easy test for entropy is a compression engine (any works).
    ' Just attempt to compress the source data streams and assume whichever compresses better
    ' will produce a similar result in the destination stream.
    Dim chunkSize As Long
    chunkSize = dataWidth
    
    'Wrap 1D arrays around source and destination targets because it makes life much simpler
    Dim totalSize As Long
    totalSize = dataWidth * dataHeight
    
    Dim src1() As Byte, src2() As Byte, dst() As Byte
    Dim srcSA1 As SafeArray1D, srcSA2 As SafeArray1D, dstSA As SafeArray1D
    VBHacks.WrapArrayAroundPtr_Byte src1, srcSA1, VarPtr(srcData1(0, 0)), totalSize
    VBHacks.WrapArrayAroundPtr_Byte src2, srcSA2, VarPtr(srcData2(0, 0)), totalSize
    VBHacks.WrapArrayAroundPtr_Byte dst, dstSA, VarPtr(dstData(0, 0)), totalSize
    
    Dim curOffset As Long
    curOffset = 0
    
    Dim tmpCompress() As Byte, tmpCompressSize As Long
    tmpCompressSize = Compression.GetWorstCaseSize(chunkSize, cf_Lz4)
    ReDim tmpCompress(0 To tmpCompressSize - 1) As Byte
    
    Dim size1 As Long, size2 As Long
    
    'To try and prevent overly-aggressive "flipping" between streams, we apply a slight penalty
    ' to whichever stream was *not* chosen last.  This biases the encoder toward consistently
    ' selecting the same stream (which likely provides better long-term compression benefits)
    ' unless switching to a new stream shows a meaningful compression advantage.
    '
    'The current value of this constant was chosen by trial-and-error.  I am open to modifying
    ' it further pending better data.  Because the modifier is multiplied directly by the
    ' compressed size of the targeted stream, make sure it is > 1 or you'll bias it the
    ' wrong way!
    Const AVOIDANCE_PENALTY_PERCENT As Double = 1.025
    Dim idLastChosenStream As Long
    idLastChosenStream = 0
    
    'Iterate both source arrays and copy over the best-compressing chunks from either
    Do While (curOffset < totalSize)
        
        If (curOffset + chunkSize) > totalSize Then chunkSize = totalSize - curOffset
        
        size1 = tmpCompressSize
        size2 = tmpCompressSize
        Compression.CompressPtrToPtr VarPtr(tmpCompress(0)), size1, VarPtr(src1(curOffset)), chunkSize, cf_Lz4, 1
        Compression.CompressPtrToPtr VarPtr(tmpCompress(0)), size2, VarPtr(src2(curOffset)), chunkSize, cf_Lz4, 1
        
        'Apply a slight penalty to whichever stream was *not* used previously
        If (idLastChosenStream = 1) Then size2 = Int(size2 * AVOIDANCE_PENALTY_PERCENT)
        If (idLastChosenStream = 2) Then size1 = Int(size1 * AVOIDANCE_PENALTY_PERCENT)
        
        'Favor the first input on matches
        If (size1 <= size2) Then
            CopyMemoryStrict VarPtr(dst(curOffset)), VarPtr(src1(curOffset)), chunkSize
            idLastChosenStream = 1
        Else
            CopyMemoryStrict VarPtr(dst(curOffset)), VarPtr(src2(curOffset)), chunkSize
            idLastChosenStream = 2
        End If
        
        If (chunkSize = dataWidth) Then curOffset = curOffset + chunkSize Else curOffset = totalSize
        
    Loop
    
    'Unwrap unsafe arrray wrappers
    VBHacks.UnwrapArrayFromPtr_Byte src1
    VBHacks.UnwrapArrayFromPtr_Byte src2
    VBHacks.UnwrapArrayFromPtr_Byte dst
    
    MakeMinimalEntropy_Big = True
    
End Function

'Outline a 32-bpp DIB.  The outline is drawn along the first-encountered border where transparent and opaque pixels meet.
' The caller must supply the outline pen they want used and optionally, an edge threshold on the range [0, 100].
'Returns: TRUE if successful; FALSE otherwise
Public Function OutlineDIB(ByRef srcDIB As pdDIB, ByRef outlinePen As pd2DPen, Optional ByVal edgeThreshold As Single = 50!) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long, finalX As Long, finalY As Long
            finalX = (srcDIB.GetDIBWidth - 1)
            finalY = (srcDIB.GetDIBHeight - 1)
            
            Dim srcPixels() As Byte, tmpSA As SafeArray2D
            srcDIB.WrapArrayAroundDIB srcPixels, tmpSA
            
            'We first need to construct a byte array that separates the pixels into two groups; the barrier
            ' between these groups will be used to construct our outline.
            Dim edgeThresholdL As Long
            edgeThresholdL = edgeThreshold * 2.55
            If (edgeThresholdL < 0) Then edgeThresholdL = 0
            If (edgeThresholdL >= 255) Then edgeThresholdL = 254
                
            Dim iWidth As Long, iHeight As Long
            iWidth = finalX + 2
            iHeight = finalY + 2
            
            'To spare our edge detector from worrying about edge pixels (which slow down processing due to obnoxious
            ' nested If/Then statements), we declare our input array with a guaranteed list of non-edge pixels on
            ' all sides.
            Dim edgeData() As Byte
            ReDim edgeData(0 To iWidth, 0 To iHeight) As Byte
            
            Dim xOffset As Long, yOffset As Long
            xOffset = 1
            yOffset = 1
            
            For y = 0 To finalY
            For x = 0 To finalX
                If (srcPixels(x * 4 + 3, y) > edgeThresholdL) Then edgeData(x + xOffset, y + yOffset) = 1
            Next x
            Next y
            
            'We no longer need direct access to the source image pixels
            srcDIB.UnwrapArrayFromDIB srcPixels
            
            'With an edge array successfully assembled, prepare an edge detector
            Dim cEdges As pdEdgeDetector
            Set cEdges = New pdEdgeDetector
            
            Dim finalPolygon() As PointFloat, numOfPoints As Long
            
            'Use the edge detector to find an initial (x, y) location to start our path trace
            Dim startX As Long, startY As Long
            If cEdges.FindStartingPoint(edgeData, 1, 1, finalX + 1, finalY + 1, startX, startY) Then
            
                'Run the path analyzer
                If cEdges.FindEdges(edgeData, startX, startY, -xOffset, -yOffset) Then
                
                    'Retrieve the polygon that defines the outer boundary
                    cEdges.RetrieveFinalPolygon finalPolygon, numOfPoints
                
                End If
            
            'No exterior path found, which is basically total failure.  Treat the image boundaries as our
            ' final path outline, instead.
            Else
                numOfPoints = 4
                ReDim finalPolygon(0 To 3) As PointFloat
                finalPolygon(0).x = 0
                finalPolygon(0).y = 0
                finalPolygon(1).x = 0
                finalPolygon(1).y = finalY
                finalPolygon(2).x = finalX
                finalPolygon(2).y = finalY
                finalPolygon(3).x = 0
                finalPolygon(3).y = finalY
            End If
            
            'Use pd2D to render the outline onto the image
            Dim cSurface As pd2DSurface
            Drawing2D.QuickCreateSurfaceFromDC cSurface, srcDIB.GetDIBDC, True
            PD2D.DrawPolygonF cSurface, outlinePen, numOfPoints, VarPtr(finalPolygon(0))
            Set cSurface = Nothing
            
            OutlineDIB = True
            
        End If
    Else
        Debug.Print "WARNING!  DIBs.OutlineDIB() requires a 32-bpp DIB to operate correctly."
    End If
    
End Function

'If a DIB does not already have 256 colors (or less), you can call this function to forcibly return
' an optimized RGB-only palette for the DIB (at the requested color count), a matching one-byte-per-pixel
' palettized image array (with dimensions matching the original image), and a matching one-byte-per-pixel
' mask describing the image's original alpha channel.  This is particularly useful when exporting to
' image formats that support separate RGB and mask layers, like ICOs.
'
'The returned palette will be in RGBA format with A values forced to 255 (so alpha does *not* matter when
' calculating colors); you can downsample to pure RGB triplets without problem.
'
'RETURNS: number of colors in the destination palette (1-based).  If the image already contains less than
' the requested number of colors, the return value is a safe way to identify that.  If the return is 0,
' the function failed.
Public Function GetDIBAs8bpp_RGBMask_Forcibly(ByRef srcDIB As pdDIB, ByRef dstPalette() As RGBQuad, ByRef dstPixels() As Byte, ByRef dstMask() As Byte, Optional ByVal maxSizeOfPalette As Long = 256) As Long

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) And (srcDIB.GetDIBColorDepth = 32) Then
        
        'If image is premultiplied, unpremultiply now
        Dim alphaWasModified As Boolean
        alphaWasModified = srcDIB.GetAlphaPremultiplication()
        If alphaWasModified Then srcDIB.SetAlphaPremultiplication False
        
        'Start by retrieving an optimized RGBA palette for the image in question
        If Palettes.GetOptimizedPalette(srcDIB, dstPalette, maxSizeOfPalette) Then
            
            'Because the result of this will be masked, we need to ensure black is available in the palette.
            Palettes.EnsureBlackAndWhiteInPalette dstPalette, Nothing, True, False
            
            'Find the black index in the palette
            Dim x As Long, y As Long
            Dim idxBlack As Long
            idxBlack = -1
            
            For x = 0 To UBound(dstPalette)
                If (dstPalette(x).Red = 0) And (dstPalette(x).Green = 0) And (dstPalette(x).Blue = 0) Then
                    idxBlack = x
                    Exit For
                End If
            Next x
            
            'A palette was successfully generated.  We now want to match each pixel in the original image
            ' to the palette we've generated, and return the result.
            ReDim dstPixels(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Byte
            ReDim dstMask(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Byte
            
            Dim srcPixels() As Byte, tmpSA As SafeArray1D
            
            Dim pxSize As Long
            pxSize = srcDIB.GetDIBColorDepth \ 8
            
            Dim initX As Long, initY As Long, finalX As Long, finalY As Long, xOffset As Long
            initX = 0
            initY = 0
            finalX = srcDIB.GetDIBWidth - 1
            finalY = srcDIB.GetDIBHeight - 1
            
            'As with normal palette matching, we'll use basic RLE acceleration to try and skip palette
            ' searching for contiguous matching colors.
            Dim lastColor As Long: lastColor = -1
            Dim r As Long, g As Long, b As Long, a As Long
            
            Dim tmpQuad As RGBQuad, newIndex As Long, lastIndex As Long
            tmpQuad.Alpha = 255
            lastIndex = -1
            
            'Build the initial tree
            Dim kdTree As pdKDTree
            Set kdTree = New pdKDTree
            kdTree.BuildTree dstPalette, UBound(dstPalette) + 1
            
            'Start matching pixels
            For y = 0 To finalY
                srcDIB.WrapArrayAroundScanline srcPixels, tmpSA, y
            For x = 0 To finalX
                
                xOffset = x * pxSize
                b = srcPixels(xOffset)
                g = srcPixels(xOffset + 1)
                r = srcPixels(xOffset + 2)
                a = srcPixels(xOffset + 3)
                
                'Store alpha in the mask channel (regardless of value)
                dstMask(x, y) = a
                
                'Manually assign transparent pixels to black
                If (a < 128) And (idxBlack >= 0) Then
                    dstPixels(x, y) = idxBlack
                Else
                    
                    'If this pixel matches the last pixel we tested, reuse our previous match results
                    If (RGB(r, g, b) <> lastColor) Then
                        
                        tmpQuad.Red = r
                        tmpQuad.Green = g
                        tmpQuad.Blue = b
                        
                        'Ask the tree for its best match
                        newIndex = kdTree.GetNearestPaletteIndex(tmpQuad)
                        
                        lastColor = RGB(r, g, b)
                        lastIndex = newIndex
                        
                    Else
                        newIndex = lastIndex
                    End If
                    
                    'Mark the matched index in the destination array
                    dstPixels(x, y) = newIndex
                    
                End If
                    
            Next x
            Next y
            
            srcDIB.UnwrapArrayFromDIB srcPixels
            
            GetDIBAs8bpp_RGBMask_Forcibly = UBound(dstPalette) + 1
    
        End If
        
        'Reset alpha premultiplication, as necessary
        If alphaWasModified Then srcDIB.SetAlphaPremultiplication True
        
    End If
    
End Function

'Assuming a source DIB already contains 256 unique RGBA quads (or less), call this function to return two arrays:
' a palette array (RGB quads), and a one-byte-per-pixel palettized image array (with dimensions matching the
' original image).
'
'At present, this function is *not* optimized.  A naive palette search is used.  Also, the destination palette is in
' RGBA format (so alpha *does* matter when calculating colors.)
'
'RETURNS: number of colors in the destination palette (1-based).  If the return is 257, the function failed because there
' are too many colors in the source image.  Reduce the number of colors, then try again.
Public Function GetDIBAs8bpp_RGBA(ByRef srcDIB As pdDIB, ByRef dstPalette() As RGBQuad, ByRef dstPixels() As Byte) As Long

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) And (srcDIB.GetDIBColorDepth = 32) Then
        
        Dim srcPixels() As Byte, tmpSA As SafeArray2D
        srcDIB.WrapArrayAroundDIB srcPixels, tmpSA
        
        Dim pxSize As Long
        pxSize = srcDIB.GetDIBColorDepth \ 8
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.GetDIBStride - 1
        finalY = srcDIB.GetDIBHeight - 1
        
        ReDim dstPalette(0 To 255) As RGBQuad
        ReDim dstPixels(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Byte
        
        'Always load the first color in advance; this lets us avoid branches in the inner loop
        Dim numColors As Long
        numColors = 1
        With dstPalette(0)
            .Blue = srcPixels(0, 0)
            .Green = srcPixels(1, 0)
            .Red = srcPixels(2, 0)
            .Alpha = srcPixels(3, 0)
        End With
        
        dstPixels(0, 0) = 0
        
        Dim r As Long, g As Long, b As Long, a As Long, i As Long
        Dim matchFound As Boolean
        
        For y = 0 To finalY
        For x = 0 To finalX Step pxSize
            
            b = srcPixels(x, y)
            g = srcPixels(x + 1, y)
            r = srcPixels(x + 2, y)
            a = srcPixels(x + 3, y)
            matchFound = False
            
            'Search the palette for a match
            ' (TODO: optimize this with a hash table)
            For i = 0 To numColors - 1
                If (b = dstPalette(i).Blue) Then
                    If (g = dstPalette(i).Green) Then
                        If (r = dstPalette(i).Red) Then
                            If (a = dstPalette(i).Alpha) Then
                                matchFound = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next i
            
            'Add new colors to the list (until we reach 257; then bail)
            If matchFound Then
                dstPixels(x \ pxSize, y) = i
            Else
                If (numColors = 256) Then
                    numColors = 257
                    Exit For
                End If
                With dstPalette(numColors)
                    .Blue = b
                    .Green = g
                    .Red = r
                    .Alpha = a
                End With
                dstPixels(x \ pxSize, y) = numColors
                numColors = numColors + 1
            End If
            
        Next x
            If (numColors > 256) Then Exit For
        Next y
        
        srcDIB.UnwrapArrayFromDIB srcPixels
        
        GetDIBAs8bpp_RGBA = numColors
        
    End If
    
End Function

'If a DIB does not already have 256 colors (or less), you can call this function to forcibly return an optimized palette for
' the DIB (at the requested color count) and a matching one-byte-per-pixel palettized image array (with dimensions matching the
' original image).
'
'The returned palette will be in RGBA format (so alpha *does* matter when calculating colors.)
'
'RETURNS: number of colors in the destination palette (1-based).  If the image already contains less than the requested
' number of colors, the return value is a safe way to identify that.  If the return is 0, the function failed.
Public Function GetDIBAs8bpp_RGBA_Forcibly(ByRef srcDIB As pdDIB, ByRef dstPalette() As RGBQuad, ByRef dstPixels() As Byte, Optional ByVal maxSizeOfPalette As Long = 256) As Long

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) And (srcDIB.GetDIBColorDepth = 32) Then
        
        'Start by retrieving an optimized RGBA palette for the image in question
        If Palettes.GetNeuquantPalette_RGBA(srcDIB, dstPalette, maxSizeOfPalette) Then
        
            'A palette was successfully generated.  We now want to match each pixel in the original image
            ' to the palette we've generated, and return the result.
            ReDim dstPixels(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Byte
            
            Dim srcPixels() As Byte, tmpSA As SafeArray1D
            
            Dim pxSize As Long
            pxSize = srcDIB.GetDIBColorDepth \ 8
            
            Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long, xOffset As Long
            initX = 0
            initY = 0
            finalX = srcDIB.GetDIBWidth - 1
            finalY = srcDIB.GetDIBHeight - 1
            
            'As with normal palette matching, we'll use basic RLE acceleration to try and skip palette
            ' searching for contiguous matching colors.
            Dim lastColor As Long: lastColor = -1
            Dim lastAlpha As Long: lastAlpha = -1
            Dim r As Long, g As Long, b As Long, a As Long
            
            Dim tmpQuad As RGBQuad, newIndex As Long, lastIndex As Long
            lastIndex = -1
            
            'Build the initial tree
            Dim kdTree As pdKDTree
            Set kdTree = New pdKDTree
            kdTree.BuildTreeIncAlpha dstPalette, UBound(dstPalette) + 1
            
            'Start matching pixels
            For y = 0 To finalY
                srcDIB.WrapArrayAroundScanline srcPixels, tmpSA, y
            For x = 0 To finalX
                
                xOffset = x * pxSize
                b = srcPixels(xOffset)
                g = srcPixels(xOffset + 1)
                r = srcPixels(xOffset + 2)
                If (pxSize = 4) Then a = srcPixels(xOffset + 3) Else a = 255
                
                'If this pixel matches the last pixel we tested, reuse our previous match results
                If ((RGB(r, g, b) <> lastColor) Or (a <> lastAlpha)) Then
                    
                    tmpQuad.Red = r
                    tmpQuad.Green = g
                    tmpQuad.Blue = b
                    tmpQuad.Alpha = a
                    
                    'Ask the tree for its best match
                    newIndex = kdTree.GetNearestPaletteIndexIncAlpha(tmpQuad)
                    
                    lastColor = RGB(r, g, b)
                    lastAlpha = a
                    lastIndex = newIndex
                    
                Else
                    newIndex = lastIndex
                End If
                
                'Mark the matched index in the destination array
                dstPixels(x, y) = newIndex
                
            Next x
            Next y
            
            srcDIB.UnwrapArrayFromDIB srcPixels
            
            GetDIBAs8bpp_RGBA_Forcibly = UBound(dstPalette) + 1
    
        End If
        
    End If
    
End Function

'Assuming a source DIB already contains 256 unique RGBA quads (or less), and you already know its palette
' (common at image export time, especially if e.g. the user has requested "use original file palette" or similar),
' call this function to return a one-byte-per-pixel palettized image array (with dimensions matching the original
' image) with all source bytes matched against their best-case palette equivalent.
'
'RETURNS: number of colors in the supplied palette (1-based).  This is to preserve compatibility with the other
' palette generating functions in this module.
Public Function GetDIBAs8bpp_RGBA_SrcPalette(ByRef srcDIB As pdDIB, ByRef srcPalette() As RGBQuad, ByRef dstPixels() As Byte) As Long

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) And (srcDIB.GetDIBColorDepth = 32) Then
        
        Dim srcPixels() As Byte, tmpSA As SafeArray1D
        
        Dim pxSize As Long
        pxSize = srcDIB.GetDIBColorDepth \ 8
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.GetDIBStride - 1
        finalY = srcDIB.GetDIBHeight - 1
        
        ReDim dstPixels(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Byte
        
        'Copy the supplied palette into a KD-tree; this lets us match colors much more quickly
        Dim cTree As pdKDTree
        Set cTree = New pdKDTree
        cTree.BuildTreeIncAlpha srcPalette, UBound(srcPalette) + 1
        
        Dim r As Long, g As Long, b As Long, a As Long
        Dim tmpQuad As RGBQuad, newIndex As Long, lastIndex As Long
        Dim lastColor As Long: lastColor = -1
        Dim lastAlpha As Long: lastAlpha = -1
        
        'To avoid division on the inner loop, build a lut for x indices
        Dim xLookup() As Long
        ReDim xLookup(0 To finalX) As Long
        For x = 0 To finalX Step pxSize
            xLookup(x) = x \ pxSize
        Next x
    
        For y = 0 To finalY
            srcDIB.WrapArrayAroundScanline srcPixels, tmpSA, y
        For x = 0 To finalX Step pxSize
            
            b = srcPixels(x)
            g = srcPixels(x + 1)
            r = srcPixels(x + 2)
            a = srcPixels(x + 3)
            
            'If this pixel matches the last pixel we tested, reuse our previous match results
            If ((RGB(r, g, b) <> lastColor) Or (a <> lastAlpha)) Then
                
                tmpQuad.Red = r
                tmpQuad.Green = g
                tmpQuad.Blue = b
                tmpQuad.Alpha = a
                
                'Ask the tree for its best match
                newIndex = cTree.GetNearestPaletteIndexIncAlpha(tmpQuad)
                
                lastColor = RGB(r, g, b)
                lastAlpha = a
                lastIndex = newIndex
                
            Else
                newIndex = lastIndex
            End If
            
            dstPixels(xLookup(x), y) = newIndex
            
        Next x
        Next y
        
        srcDIB.UnwrapArrayFromDIB srcPixels
        
        GetDIBAs8bpp_RGBA_SrcPalette = UBound(srcPalette) + 1
        
    End If
    
End Function

'Given a palette, color count, and source palette index array, construct a matching 32-bpp DIB.
'
'IMPORTANT NOTE!  The destination DIB needs to be already constructed to match the source data's width and height.
' (Obviously, it needs to be 32-bpp too!)
'
'RETURNS: TRUE if successful; FALSE otherwise
Public Function GetRGBADIB_FromPalette(ByRef dstDIB As pdDIB, ByVal colorCount As Long, ByRef srcPalette() As RGBQuad, ByRef srcPixels() As Byte) As Boolean

    If (dstDIB Is Nothing) Then Exit Function
    
    If (dstDIB.GetDIBDC <> 0) And (dstDIB.GetDIBWidth <> 0) And (dstDIB.GetDIBHeight <> 0) And (dstDIB.GetDIBColorDepth = 32) Then
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = dstDIB.GetDIBWidth - 1
        finalY = dstDIB.GetDIBHeight - 1
        
        'Construct a lookup table of Long-type RGBA quads; these are faster to apply than individual color bytes
        Dim lTable() As Long
        ReDim lTable(0 To colorCount - 1) As Long
            
        For x = 0 To colorCount - 1
            GetMem4 VarPtr(srcPalette(x)), ByVal VarPtr(lTable(x))
        Next x
        
        'To improve performance, we'll point a Long-type array at each individual line in turn (rather than
        ' addressing the entire thing in 2D)
        Dim imgData() As Long, tmpSA As SafeArray1D, imgPtr As Long, imgStride As Long
        dstDIB.WrapLongArrayAroundScanline imgData, tmpSA, 0
        imgPtr = tmpSA.pvData
        imgStride = tmpSA.cElements * 4
        
        For y = 0 To finalY
            tmpSA.pvData = imgPtr + imgStride * y
        For x = 0 To finalX
            imgData(x) = lTable(srcPixels(x, y))
        Next x
        Next y
        
        dstDIB.UnwrapLongArrayFromDIB imgData
        
        GetRGBADIB_FromPalette = True
        
    End If
    
End Function

'Retrieve a single channel from a DIB as a 2D array.  The 2D array will have guaranteed dimensions
' [0, dibWidth - 1] x [0, dibHeight - 1], unless a source rectangle is passed - then it will have
' guaranteed dimensions [0, rect.width - 1] x [0, rect.height - 1].
'
'No validation is performed on the optional passed rect, so make sure it's valid or this function *will* crash.
Public Function GetSingleChannel_2D(ByRef srcDIB As pdDIB, ByRef dstBytes() As Byte, Optional ByVal channelOffset As Long = 0, Optional ByVal ptrToRectOfInterest As Long = 0) As Boolean
    
    GetSingleChannel_2D = False
    
    If (srcDIB Is Nothing) Then Exit Function
    
    Dim srcOffsetX As Long, srcOffsetY As Long
    Dim newBoundX As Long, newBoundY As Long
    
    If (ptrToRectOfInterest <> 0) Then
        Dim tmpRectF As RectF
        VBHacks.CopyMemoryStrict VarPtr(tmpRectF), ptrToRectOfInterest, 16
        srcOffsetX = Int(tmpRectF.Left)
        srcOffsetY = Int(tmpRectF.Top)
        newBoundX = Int(tmpRectF.Width) - 1
        newBoundY = Int(tmpRectF.Height) - 1
    Else
        srcOffsetX = 0
        srcOffsetY = 0
        newBoundX = srcDIB.GetDIBWidth - 1
        newBoundY = srcDIB.GetDIBHeight - 1
    End If
    
    ReDim dstBytes(0 To newBoundX, 0 To newBoundY) As Byte
    
    Dim srcBytes() As Byte, srcSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB srcBytes, srcSA
    
    Dim x As Long, y As Long
    For y = 0 To newBoundY
    For x = 0 To newBoundX
        dstBytes(x, y) = srcBytes((x + srcOffsetX) * 4 + channelOffset, srcOffsetY + y)
    Next x
    Next y
    
    srcDIB.UnwrapArrayFromDIB srcBytes
    GetSingleChannel_2D = True
    
End Function

'This function returns a DIB, resized to meet a specific pixel count.  This is very helpful for things like image analysis,
' where a full-sized image copy doesn't meaningfully improve heuristics (but requires a hell of a lot longer to analyze).
'
'This function always preserves aspect ratio, and it will return the original image if the image is smaller than the number of
' pixels requested (unless overridden by the optional allowUpsampling parameter).  This simplifies outside functions,
' as you can always call this function prior to running heuristics.
Public Function ResizeDIBByPixelCount(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal numOfPixels As Long, Optional ByVal interpolationType As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal allowUpsampling As Boolean = False) As Boolean

    If (Not srcDIB Is Nothing) Then
        
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        
        'Calculate current megapixel count
        Dim srcWidth As Long, srcHeight As Long
        srcWidth = srcDIB.GetDIBWidth
        srcHeight = srcDIB.GetDIBHeight
        
        'If the source image has less megapixels than the requested amount, just return it as-is
        If (srcWidth * srcHeight < numOfPixels) And (Not allowUpsampling) Then
            dstDIB.CreateFromExistingDIB srcDIB
            ResizeDIBByPixelCount = True
        
        'If the source image is larger than the destination (as it usually will be), calculate an aspect-ratio appropriate
        ' resize that comes close to the target number of pixels
        Else
        
            Dim pxCount As Long
            pxCount = numOfPixels
            
            Dim aspectRatio As Single
            aspectRatio = srcWidth / srcHeight
            
            'Using basic algebra, we can solve for new (x, y) parameters in terms of megapixel count
            Dim newWidth As Long, newHeight As Long
            newHeight = Sqr(numOfPixels / aspectRatio)
            newWidth = aspectRatio * newHeight
            
            dstDIB.CreateBlank newWidth, newHeight, 32, 0, 0
            GDI_Plus.GDIPlus_StretchBlt dstDIB, 0, 0, newWidth, newHeight, srcDIB, 0, 0, srcWidth, srcHeight, , interpolationType, , True, , True
            
            ResizeDIBByPixelCount = True
            
        End If
    
    End If

End Function

'Given a byte array, construct a 32-bpp DIB where each channel is set to the grayscale
' equivalent of the input array.  This is used with selections to generate a transparent
' + grayscale copy of a single byte array.  Note that the DIB *must* already exist as a
' 32-bpp DIB matching the size of the input table.
'
'Returns TRUE if successful, and srcDIB will be filled with a premultiplied 32-bpp DIB matching the input table
Public Function Construct32bppDIBFromByteMap(ByRef srcDIB As pdDIB, ByRef srcMap() As Byte) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long
            Dim initX As Long, initY As Long, finalX As Long, finalY As Long
            initX = 0
            initY = 0
            finalX = srcDIB.GetDIBWidth - 1
            finalY = srcDIB.GetDIBHeight - 1
            
            'Construct a lookup table of premultiplied input values
            Dim lTable(0 To 255) As Long
            Dim tmpA As Single, tmpQuad As RGBQuad
            
            For x = 0 To 255
                tmpQuad.Alpha = x
                tmpA = x * (x / 255)
                tmpQuad.Red = Int(tmpA)
                tmpQuad.Green = Int(tmpA)
                tmpQuad.Blue = Int(tmpA)
                CopyMemoryStrict VarPtr(lTable(x)), VarPtr(tmpQuad), LenB(tmpQuad)
            Next x
            
            Dim imgData() As Long, tmpSA As SafeArray1D
            
            'Loop through the image, checking alphas as we go
            For y = initY To finalY
                srcDIB.WrapLongArrayAroundScanline imgData, tmpSA, y
            For x = initX To finalX
                imgData(x) = lTable(srcMap(x, y))
            Next x
            Next y
            
            srcDIB.UnwrapLongArrayFromDIB imgData
            
            srcDIB.SetInitialAlphaPremultiplicationState True
            
            Construct32bppDIBFromByteMap = True
            
        End If
    Else
        Debug.Print "WARNING!  pdDIB.Construct32bppDIBFromByteMap() requires a 32-bpp DIB to operate correctly."
    End If
    
End Function

'Determine the "rect of interest" in a 32-bpp image.  The "rect of interest" is the smallest rectangle that contains
' all non-transparent pixels in the image.  The source image is not modified.
'
'RETURNS: TRUE if successful, FALSE if the entire image is transparent.  If the rect in question is passed to any
' per-pixel routines, you will want to catch the FALSE case as not doing so may lead to OOB errors.
Public Function GetRectOfInterest(ByRef srcDIB As pdDIB, ByRef dstRectF As RectF) As Boolean
    
    'The image will be analyzed in four steps.  Each edge will be analyzed separately, starting with the top.
    
    'Point an array at the DIB data
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long, xStride As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'The new edges of the image will mark these values for us; at the end of the function, we'll fill the RectF
    ' with these.
    Dim newTop As Long, newBottom As Long, newLeft As Long, newRight As Long
    
    'When a non-transparent pixel is found, this check value will be set to TRUE; we need this to provide a
    ' failsafe for fully transparent images.
    Dim colorFails As Boolean:    colorFails = False
    
    'Scan the image, starting at the top-left and moving right
    For y = 0 To finalY
    For x = 0 To finalX
        
        'If this pixel is transparent, keep scanning.  Otherwise, note that we have found a non-transparent pixel
        ' and exit the loop.
        If (srcImageData(x * 4 + 3, y) <> 0) Then
            colorFails = True
            Exit For
        End If
        
    Next x
        If colorFails Then Exit For
    Next y
    
    'We have now reached one of two conditions:
    '1) The entire image is transparent
    '2) The loop progressed part-way through the image and terminated
    
    'Check for case (1) and warn the user if it occurred
    If (Not colorFails) Then
        srcDIB.UnwrapArrayFromDIB srcImageData
        GetRectOfInterest = False
        Exit Function
    
    'Next, check for case (2)
    Else
        newTop = y
    End If
    
    initY = newTop
    
    'Repeat the above steps, but tracking the left edge instead.  Note also that we will only be scanning from wherever
    ' the top trim failed - this saves processing time.
    colorFails = False
    
    For x = 0 To finalX
        xStride = x * 4
    For y = initY To finalY
    
        If (srcImageData(xStride + 3, y) <> 0) Then
            colorFails = True
            Exit For
        End If
        
    Next y
        If colorFails Then Exit For
    Next x
    
    newLeft = x
    
    'Repeat the above steps, but tracking the right edge instead.  Note also that we will only be scanning from wherever
    ' the top trim failed - this saves processing time.
    colorFails = False
    
    For x = finalX To 0 Step -1
        xStride = x * 4
    For y = initY To finalY
    
        If (srcImageData(xStride + 3, y) <> 0) Then
            colorFails = True
            Exit For
        End If
        
    Next y
        If colorFails Then Exit For
    Next x
    
    newRight = x
    
    'Finally, repeat the steps above for the bottom of the image.  Note also that we will only be scanning from wherever
    ' the left and right trims failed - this saves processing time.
    colorFails = False
    initX = newLeft
    finalX = newRight
    
    For y = finalY To initY Step -1
    For x = initX To finalX
        
        If (srcImageData(x * 4 + 3, y) <> 0) Then
            colorFails = True
            Exit For
        End If
        
    Next x
        If colorFails Then Exit For
    Next y
    
    newBottom = y
    
    'Safely deallocate our temporary array reference
    srcDIB.UnwrapArrayFromDIB srcImageData
    
    'Populate the destination rect (which uses floating-point values, like all PD rects)
    With dstRectF
        .Left = newLeft
        .Top = newTop
        .Width = newRight - newLeft + 1
        .Height = newBottom - newTop + 1
    End With
    
    GetRectOfInterest = True
    
End Function

'Determine the "rect of interest" in a 32-bpp image, using an "overlay" approach where one DIB sits atop some
' other arbitrary DIB.  (Both DIBs need to be the same size!)
'
'The "rect of interest" is the smallest rectangle that contains pixels unique to the *top* image.
' Neither the top nor the bottom DIB is modified by this approach.
'
'RETURNS: TRUE if successful, FALSE if both images are identical (and thus no rect of interest exists).
' If the rect in question is passed to any per-pixel routines, you will want to catch the FALSE case as not
' doing so may lead to OOB errors.
Public Function GetRectOfInterest_Overlay(ByRef topDIB As pdDIB, ByRef bottomDIB As pdDIB, ByRef dstRectF As RectF) As Boolean
    
    GetRectOfInterest_Overlay = False
    
    If (topDIB Is Nothing) Or (bottomDIB Is Nothing) Then
        PDDebug.LogAction "GetRectOfInterest_Overlay failed!  Top or bottom DIB was empty!"
        Exit Function
    End If
    
    If (topDIB.GetDIBWidth <> bottomDIB.GetDIBWidth) Or (topDIB.GetDIBHeight <> bottomDIB.GetDIBHeight) Then
        PDDebug.LogAction "GetRectOfInterest_Overlay failed!  Top and bottom DIB sizes don't match ! (" & topDIB.GetDIBWidth & ", " & bottomDIB.GetDIBWidth & ", " & topDIB.GetDIBHeight & ", " & bottomDIB.GetDIBHeight & ")"
        Exit Function
    End If
    
    'The image will be analyzed in four steps.  Each edge will be analyzed separately, starting with the top.
    
    'Point arrays at both DIBs
    Dim botImageData() As Long, botSA As SafeArray2D
    bottomDIB.WrapLongArrayAroundDIB botImageData, botSA
    
    Dim topImageData() As Long, topSA As SafeArray2D
    topDIB.WrapLongArrayAroundDIB topImageData, topSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = topDIB.GetDIBWidth - 1
    finalY = topDIB.GetDIBHeight - 1
    
    'The new edges of the image will mark these values for us; at the end of the function,
    ' we'll fill the RectF with these.
    Dim newTop As Long, newBottom As Long, newLeft As Long, newRight As Long
    
    'When a non-matching pixel is found, this check value will be set to TRUE; we need this to provide a
    ' failsafe for fully matching images.
    Dim colorFails As Boolean:    colorFails = False
    
    'Scan the image, starting at the top-left and moving right
    For y = 0 To finalY
    For x = 0 To finalX
        
        'If this pixel matches, keep scanning.  Otherwise, note that we found a non-matching pixel
        ' and exit the loop.
        If (topImageData(x, y) <> botImageData(x, y)) Then
            colorFails = True
            Exit For
        End If
        
    Next x
        If colorFails Then Exit For
    Next y
    
    'We have now reached one of two conditions:
    '1) The top image is 100% identical to the bottom image
    '2) The loop progressed part-way through the image and terminated
    
    'Check for case (1) and warn the user if it occurred
    If (Not colorFails) Then
        bottomDIB.UnwrapLongArrayFromDIB botImageData
        topDIB.UnwrapLongArrayFromDIB topImageData
        GetRectOfInterest_Overlay = False
        Exit Function
    
    'Next, check for case (2)
    Else
        newTop = y
    End If
    
    initY = newTop
    
    'Repeat the above steps, but tracking the left edge instead.  Note also that we will only be scanning from wherever
    ' the top trim failed - this saves processing time.
    colorFails = False
    
    For x = 0 To finalX
    For y = initY To finalY
    
        If (topImageData(x, y) <> botImageData(x, y)) Then
            colorFails = True
            Exit For
        End If
        
    Next y
        If colorFails Then Exit For
    Next x
    
    newLeft = x
    
    'Repeat the above steps, but tracking the right edge instead.  Note also that we will only be scanning from wherever
    ' the top trim failed - this saves processing time.
    colorFails = False
    
    For x = finalX To 0 Step -1
    For y = initY To finalY
    
        If (topImageData(x, y) <> botImageData(x, y)) Then
            colorFails = True
            Exit For
        End If
        
    Next y
        If colorFails Then Exit For
    Next x
    
    newRight = x
    
    'Finally, repeat the steps above for the bottom of the image.  Note also that we will only be scanning from wherever
    ' the left and right trims failed - this saves processing time.
    colorFails = False
    initX = newLeft
    finalX = newRight
    
    For y = finalY To initY Step -1
    For x = initX To finalX
        
        If (topImageData(x, y) <> botImageData(x, y)) Then
            colorFails = True
            Exit For
        End If
        
    Next x
        If colorFails Then Exit For
    Next y
    
    newBottom = y
    
    'Safely deallocate our temporary array references
    topDIB.UnwrapLongArrayFromDIB topImageData
    bottomDIB.UnwrapLongArrayFromDIB botImageData
    
    'Populate the destination rect (which uses floating-point values, like all PD rects)
    With dstRectF
        .Left = newLeft
        .Top = newTop
        .Width = newRight - newLeft + 1
        .Height = newBottom - newTop + 1
    End With
    
    GetRectOfInterest_Overlay = True
    
End Function

'Some GDI functions and file formats use reverse-scanline DIBs.  This function provides a fast,
' memory-efficient mechanism for converting a top-down DIB to a bottom-up one (or vice-versa).
Public Sub ReverseScanlines(ByRef dstDIB As pdDIB)
    
    Dim pxScanline() As Byte, scanlineSize As Long
    scanlineSize = dstDIB.GetDIBStride
    ReDim pxScanline(0 To scanlineSize - 1) As Byte
    
    Dim srcPtr As Long, endPtr As Long, numScanlines As Long
    srcPtr = dstDIB.GetDIBPointer
    numScanlines = dstDIB.GetDIBHeight
    endPtr = dstDIB.GetDIBPointer + (scanlineSize * (numScanlines - 1))
    
    Dim numLinesToReverse As Long
    numLinesToReverse = dstDIB.GetDIBHeight \ 2
    
    Dim y As Long
    For y = 0 To numLinesToReverse - 1
        
        'Copy the target scanline into our temporary buffer
        CopyMemoryStrict VarPtr(pxScanline(0)), srcPtr + y * scanlineSize, scanlineSize
        
        'Copy the mirrored scanline over the line we just copied.
        CopyMemoryStrict srcPtr + y * scanlineSize, endPtr - y * scanlineSize, scanlineSize
        
        'Copy the temporary buffer over the mirrored scanline
        CopyMemoryStrict endPtr - y * scanlineSize, VarPtr(pxScanline(0)), scanlineSize
        
    Next y
    
End Sub

'Swap red and blue channels, to convert between RGBA <--> BGRA.
' Premultiplication status is not affected by this operation.
Public Sub SwizzleBR(ByRef srcDIB As pdDIB)

    If (srcDIB Is Nothing) Then Exit Sub
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim numColors As Long
            numColors = srcDIB.GetDIBStride * srcDIB.GetDIBHeight - 1
            
            Dim pxData() As Byte, tmpSA As SafeArray1D
            srcDIB.WrapArrayAroundDIB_1D pxData, tmpSA
                
            Dim x As Long, tmpColor As Byte
            For x = 0 To numColors Step 4
                tmpColor = pxData(x)
                pxData(x) = pxData(x + 2)
                pxData(x + 2) = tmpColor
            Next x
            
            srcDIB.UnwrapArrayFromDIB pxData
            
        '/Ignore empty DIBs
        End If
    
    '/Ignore non-32-bpp data
    End If
    
End Sub

'Reduce an image's alpha channel to on/off only (e.g. values of 0 or 255).  Useful before converting
' to legacy file formats like GIF or ICO.
Public Function ThresholdAlphaChannel(ByRef srcDIB As pdDIB, Optional ByVal alphaCutoff As Long = 127, Optional ByVal ditherMethod As PD_DITHER_METHOD = PDDM_Stucki, Optional ByVal ditherAmount As Single = 50!, Optional ByVal matteColor As Long = vbWhite, Optional ByVal suppressMessages As Boolean = False) As Boolean
    
    ditherAmount = ditherAmount * 0.01
    If (ditherAmount < 0!) Then ditherAmount = 0!
    If (ditherAmount > 1!) Then ditherAmount = 1!
    
    'Ensure the target DIB is *not* using premultiplied alpha
    Dim needToResetAlpha As Boolean
    needToResetAlpha = srcDIB.GetAlphaPremultiplication()
    If needToResetAlpha Then srcDIB.SetAlphaPremultiplication False
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB imageData, tmpSA
    
    Dim x As Long, y As Long, i As Long, j As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        ProgressBars.SetProgBarMax finalY
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Calculating color variables (including luminance)
    Dim a As Long, newA As Long
    Dim ditherTable() As Byte
    Dim xLeft As Long, xRight As Long, yDown As Long
    Dim errorVal As Single, dDivisor As Single
    
    'We also need to manually merge colors with their matte equivalent
    Dim matteR As Byte, matteG As Byte, matteB As Byte
    matteR = Colors.ExtractRed(matteColor)
    matteG = Colors.ExtractGreen(matteColor)
    matteB = Colors.ExtractBlue(matteColor)
    
    'Create a fast lookup table for alpha references
    Dim aLookup(0 To 255) As Single, refAlpha As Single
    Const ONE_DIV_255 As Double = 1# / 255#
    
    For x = 0 To 255
        aLookup(x) = 1! - (x * ONE_DIV_255)
    Next x
    
    'Process the alpha channel based on the dither method requested
    Select Case ditherMethod
        
        'No dither, so just perform a quick and dirty threshold calculation
        Case PDDM_None
    
            For y = initY To finalY
            For x = initX To finalX
                
                xStride = x * 4
                
                'Get the source alpha value
                a = imageData(xStride + 3, y)
                
                'Check the luminance against the threshold, and set new values accordingly
                If (a >= alphaCutoff) Then
                    refAlpha = aLookup(imageData(xStride + 3, y))
                    imageData(xStride, y) = Colors.BlendColors(imageData(xStride, y), matteB, refAlpha)
                    imageData(xStride + 1, y) = Colors.BlendColors(imageData(xStride + 1, y), matteG, refAlpha)
                    imageData(xStride + 2, y) = Colors.BlendColors(imageData(xStride + 2, y), matteR, refAlpha)
                    imageData(xStride + 3, y) = 255
                Else
                    imageData(xStride, y) = 0
                    imageData(xStride + 1, y) = 0
                    imageData(xStride + 2, y) = 0
                    imageData(xStride + 3, y) = 0
                End If
                
            Next x
                If (Not suppressMessages) Then
                    If (y And progBarCheck) = 0 Then
                        If Interface.UserPressedESC() Then Exit For
                        SetProgBarVal y
                    End If
                End If
            Next y
            
            
        'Ordered dither (Bayer 4x4).  Unfortunately, this routine requires a unique set of code owing to its
        ' specialized implementation. Coefficients derived from https://en.wikipedia.org/wiki/Ordered_dithering
        Case 1
        
            'First, prepare a Bayer dither table
            Palettes.GetDitherTable PDDM_Ordered_Bayer4x4, ditherTable, dDivisor, xLeft, xRight, yDown
            
            'Now loop through the image, using the dither values as our threshold
            For y = initY To finalY
            For x = initX To finalX
            
                xStride = x * 4
                
                'Retrieve current alpha
                a = imageData(xStride + 3, y)
                
                'Add the value of the dither table
                a = a + (Int(ditherTable(x And 3, y And 3)) - 127) * ditherAmount
                
                'Check THAT value against the threshold, and set a new value accordingly
                If (a >= alphaCutoff) Then
                    refAlpha = aLookup(imageData(xStride + 3, y))
                    imageData(xStride, y) = Colors.BlendColors(imageData(xStride, y), matteB, refAlpha)
                    imageData(xStride + 1, y) = Colors.BlendColors(imageData(xStride + 1, y), matteG, refAlpha)
                    imageData(xStride + 2, y) = Colors.BlendColors(imageData(xStride + 2, y), matteR, refAlpha)
                    imageData(xStride + 3, y) = 255
                Else
                    imageData(xStride, y) = 0
                    imageData(xStride + 1, y) = 0
                    imageData(xStride + 2, y) = 0
                    imageData(xStride + 3, y) = 0
                End If
                
            Next x
                If (Not suppressMessages) Then
                    If (y And progBarCheck) = 0 Then
                        If Interface.UserPressedESC() Then Exit For
                        SetProgBarVal y
                    End If
                End If
            Next y

        'Ordered dither (Bayer 8x8).  Unfortunately, this routine requires a unique set of code owing to its specialized
        ' implementation. Coefficients derived from https://en.wikipedia.org/wiki/Ordered_dithering
        Case 2
        
            'First, prepare a Bayer dither table
            Palettes.GetDitherTable PDDM_Ordered_Bayer8x8, ditherTable, dDivisor, xLeft, xRight, yDown
            
            'Now loop through the image, using the dither values as our threshold
            For y = initY To finalY
            For x = initX To finalX
                
                xStride = x * 4
                
                'Retrieve current alpha
                a = imageData(xStride + 3, y)
                
                'Add the value of the dither table
                a = a + (Int(ditherTable(x And 7, y And 7)) - 127) * ditherAmount
                
                'Check THAT value against the threshold, and set a new value accordingly
                If (a >= alphaCutoff) Then
                    refAlpha = aLookup(imageData(xStride + 3, y))
                    imageData(xStride, y) = Colors.BlendColors(imageData(xStride, y), matteB, refAlpha)
                    imageData(xStride + 1, y) = Colors.BlendColors(imageData(xStride + 1, y), matteG, refAlpha)
                    imageData(xStride + 2, y) = Colors.BlendColors(imageData(xStride + 2, y), matteR, refAlpha)
                    imageData(xStride + 3, y) = 255
                Else
                    imageData(xStride, y) = 0
                    imageData(xStride + 1, y) = 0
                    imageData(xStride + 2, y) = 0
                    imageData(xStride + 3, y) = 0
                End If
                
            Next x
                If (Not suppressMessages) Then
                    If (y And progBarCheck) = 0 Then
                        If Interface.UserPressedESC() Then Exit For
                        SetProgBarVal y
                    End If
                End If
            Next y
        
        'For all error-diffusion methods, precise dithering table coefficients are retrieved from the
        ' /Modules/Palettes.bas file.  (We do this because other functions also need to retrieve these tables,
        ' e.g. the Effects > Stylize > Palettize menu.)
        
        'Single neighbor.  Simplest form of error-diffusion.
        Case 3
            Palettes.GetDitherTable PDDM_SingleNeighbor, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Genuine Floyd-Steinberg.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 4
            Palettes.GetDitherTable PDDM_FloydSteinberg, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Jarvis, Judice, Ninke.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 5
            Palettes.GetDitherTable PDDM_JarvisJudiceNinke, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Stucki.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 6
            Palettes.GetDitherTable PDDM_Stucki, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Burkes.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 7
            Palettes.GetDitherTable PDDM_Burkes, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Sierra-3.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 8
            Palettes.GetDitherTable PDDM_Sierra3, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Sierra-2.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 9
            Palettes.GetDitherTable PDDM_SierraTwoRow, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Sierra-2-4A.  Coefficients derived from http://www.efg2.com/Lab/Library/ImageProcessing/DHALF.TXT
        Case 10
            Palettes.GetDitherTable PDDM_SierraLite, ditherTable, dDivisor, xLeft, xRight, yDown
            
        'Bill Atkinson's original Hyperdither/HyperScan algorithm.  (Note: Bill invented MacPaint, QuickDraw,
        ' and HyperCard.)  This is the dithering algorithm used on the original Apple Macintosh.
        ' Coefficients derived from http://gazs.github.com/canvas-atkinson-dither/
        Case 11
            Palettes.GetDitherTable PDDM_Atkinson, ditherTable, dDivisor, xLeft, xRight, yDown
            
    End Select
    
    'If we have been asked to use a non-ordered dithering method, apply it now
    If (ditherMethod >= PDDM_SingleNeighbor) Then
    
        'First, we need a dithering table the same size as the image.  Note that we use floats
        ' to prevent rounding errors.
        Dim dErrors() As Single
        ReDim dErrors(0 To finalX, 0 To finalY) As Single
        If (dDivisor <> 0!) Then dDivisor = 1! / dDivisor
        
        Dim xStrideInner As Long, yOffset As Long
        
        'Now loop through the image, calculating errors as we go
        For y = initY To finalY
        For x = initX To finalX
            
            xStride = x * 4
            
            'Retrieve current alpha
            a = imageData(xStride + 3, y)
            
            'Now, for a shortcut: if this pixel is opaque, we want to keep it opaque.
            ' Similarly, if this pixel is transparent, we want to keep it transparent.
            ' Only semi-transparent regions (like drop-shadows) should get dithered.
            If (a < 255) Then
                If (a > 0) Then
                    
                    'Add the value of the error at this location
                    newA = a + dErrors(x, y)
                    
                    'Check our modified luminance value against the threshold, and set new values accordingly
                    If (newA >= alphaCutoff) Then
                        errorVal = newA - 255
                        refAlpha = aLookup(a)
                        imageData(xStride, y) = Colors.BlendColors(imageData(xStride, y), matteB, refAlpha)
                        imageData(xStride + 1, y) = Colors.BlendColors(imageData(xStride + 1, y), matteG, refAlpha)
                        imageData(xStride + 2, y) = Colors.BlendColors(imageData(xStride + 2, y), matteR, refAlpha)
                        imageData(xStride + 3, y) = 255
                    Else
                        errorVal = newA
                        imageData(xStride, y) = 0
                        imageData(xStride + 1, y) = 0
                        imageData(xStride + 2, y) = 0
                        imageData(xStride + 3, y) = 0
                    End If
                    
                    'If there is an error, spread it
                    If (errorVal <> 0) Then
                        
                        errorVal = errorVal * ditherAmount
                        
                        'Now, spread that error across the relevant pixels according to the dither table formula
                        For i = xLeft To xRight
                        For j = 0 To yDown
                        
                            'First, ignore already processed pixels
                            If (j = 0) And (i <= 0) Then GoTo NextDitheredPixel
                            
                            'Second, ignore pixels that have a zero in the dither table
                            If (ditherTable(i, j) = 0) Then GoTo NextDitheredPixel
                            
                            xStrideInner = x + i
                            yOffset = y + j
                            
                            'Next, ignore target pixels that are off the image boundary
                            If (xStrideInner < initX) Then GoTo NextDitheredPixel
                            If (xStrideInner > finalX) Then GoTo NextDitheredPixel
                            If (yOffset > finalY) Then GoTo NextDitheredPixel
                            
                            'If we've made it all the way here, we are able to actually spread the error to this location
                            dErrors(xStrideInner, yOffset) = dErrors(xStrideInner, yOffset) + (errorVal * (CSng(ditherTable(i, j)) * dDivisor))
                        
NextDitheredPixel:             Next j
                        Next i
                    
                    End If
                
                '/a = 0
                ' Ensure transparent pixels are also black in color, as this pixel data is likely being output
                ' as a mask + palette combo, and masking produces unreliable results on non-zero data.
                Else
                    imageData(xStride, y) = 0
                    imageData(xStride + 1, y) = 0
                    imageData(xStride + 2, y) = 0
                End If
                
            '/end a < 255
            End If
                    
        Next x
            If (Not suppressMessages) Then
                If (y And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal y
                End If
            End If
        Next y
    
    End If
    
    'Safely deallocate imageData() before exiting
    srcDIB.UnwrapArrayFromDIB imageData
    
    'Importantly, after thresholding alpha in this function, there is no material difference
    ' between premultiplied and straight alpha (because transparent pixels have been forced to black,
    ' and all other pixels have been forced to full opacity).  As such, we can simply reset the
    ' premultiplication flag *without* actually looping pixels.
    If needToResetAlpha Then srcDIB.SetInitialAlphaPremultiplicationState True
    
    ThresholdAlphaChannel = (Not g_cancelCurrentAction)
    
End Function
