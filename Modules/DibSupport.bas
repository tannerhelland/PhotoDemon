Attribute VB_Name = "DIBs"
'***************************************************************************
'DIB Support Functions
'Copyright 2012-2017 by Tanner Helland
'Created: 27/March/15 (though many of the individual functions are much older!)
'Last updated: 12/June/16
'Last update: continued migration of functions
'
'This module contains support functions for the pdDIB class.  In old versions of PD, these functions were provided by pdDIB,
' but there's no sense cluttering up that class with functions that are only used on rare occasions.  As such, I'm moving
' as many of those functions as I can to this module.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal dstPointer As Long, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)

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
    If srcDIB.GetDIBColorDepth = 32 Then

        'Make sure this DIB isn't empty
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
    
            'Loop through the image and compare each alpha value against 0 and 255. If a value doesn't
            ' match either of this, this is a non-binary alpha channel and it must be handled specially.
            Dim iData() As Byte
            Dim tmpSA As SAFEARRAY2D
            PrepSafeArray tmpSA, srcDIB
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
            Dim x As Long, y As Long, quickX As Long
                
            'By default, assume that the image does not have a binary alpha channel. (This is the preferable
            ' default, as we will exit the loop IFF a non-0 or non-255 value is found.)
            Dim notBinary As Boolean
            notBinary = False
            
            Dim chkAlpha As Byte
                
            'Loop through the image, checking alphas as we go
            For x = 0 To srcDIB.GetDIBWidth - 1
                quickX = x * 4
            For y = 0 To srcDIB.GetDIBHeight - 1
            
                'Retrieve the alpha value of the current pixel
                chkAlpha = iData(quickX + 3, y)
                
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
                
            Next y
                If notBinary Then Exit For
            Next x
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            CopyMemory ByVal VarPtrArray(iData), 0&, 4
            
            IsDIBAlphaBinary = Not notBinary
                
        End If
        
    End If
    
End Function

'Is a given DIB grayscale?  Determination is made by scanning each pixel and comparing RGB values to see if they match.
Public Function IsDIBGrayscale(ByRef srcDIB As pdDIB) As Boolean
    
    'Make sure the DIB exists
    If srcDIB Is Nothing Then Exit Function
    
    'Make sure this DIB isn't empty
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
    
        'Loop through the image and compare RGB values to determine grayscale or not.
        Dim iData() As Byte
        Dim tmpSA As SAFEARRAY2D
        PrepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
        
        Dim x As Long, y As Long, quickX As Long
        Dim r As Long, g As Long, b As Long
        
        Dim qvDepth As Long
        qvDepth = srcDIB.GetDIBColorDepth \ 8
                        
        'Loop through the image, checking alphas as we go
        For x = 0 To srcDIB.GetDIBWidth - 1
            quickX = x * qvDepth
        For y = 0 To srcDIB.GetDIBHeight - 1
            
            r = iData(quickX + 2, y)
            g = iData(quickX + 1, y)
            b = iData(quickX, y)
            
            'For optimization reasons, this is stated as multiple IFs instead of an OR.
            If r <> g Then
                
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
        
                IsDIBGrayscale = False
                Exit Function
                
            ElseIf g <> b Then
            
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
            
                IsDIBGrayscale = False
                Exit Function
                
            ElseIf r <> b Then
            
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
            
                IsDIBGrayscale = False
                Exit Function
                
            End If
                
        Next y
        Next x
    
        'With our alpha channel complete, point iData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        Erase iData
                        
        'If we scanned all pixels without exiting prematurely, the DIB is grayscale
        IsDIBGrayscale = True
        Exit Function
        
    End If
    
    'If we made it to this line, the DIB is blank, so it doesn't matter what value we return
    IsDIBGrayscale = False

End Function

'Given a DIB, return a 2D Byte array of the DIB's luminance values.  An optional preNormalize parameter will guarantee that the output
' stretches from 0 to 255.  (Also note: this function does not support progress bar reports.)
Public Function GetDIBGrayscaleMap(ByRef srcDIB As pdDIB, ByRef dstGrayArray() As Byte, Optional ByVal toNormalize As Boolean = True) As Boolean
    
    'Make sure the DIB exists
    If (srcDIB Is Nothing) Then Exit Function
    
    'Make sure the source DIB isn't empty
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
    
        'Create a local array and point it at the pixel data we want to operate on
        Dim imageData() As Byte
        Dim tmpSA As SAFEARRAY2D
        PrepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
            
        'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.GetDIBWidth - 1
        finalY = srcDIB.GetDIBHeight - 1
        
        'Prep the destination array
        ReDim dstGrayArray(initX To finalX, initY To finalY) As Byte
        
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim quickVal As Long, qvDepth As Long
        qvDepth = srcDIB.GetDIBColorDepth \ 8
        
        Dim r As Long, g As Long, b As Long, grayVal As Long
        Dim minVal As Long, maxVal As Long
        minVal = 255
        maxVal = 0
            
        'Now we can loop through each pixel in the image, converting values as we go
        For x = initX To finalX
            quickVal = x * qvDepth
        For y = initY To finalY
                
            'Get the source pixel color values
            b = imageData(quickVal, y)
            g = imageData(quickVal + 1, y)
            r = imageData(quickVal + 2, y)
            
            'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
            grayVal = (213 * r + 715 * g + 72 * b) \ 1000
            If grayVal > 255 Then grayVal = 255
            
            'Cache the value
            dstGrayArray(x, y) = grayVal
            
            'If normalization has been requested, check max/min values now
            If toNormalize Then
                If (grayVal < minVal) Then
                    minVal = grayVal
                ElseIf (grayVal > maxVal) Then
                    maxVal = grayVal
                End If
            End If
            
        Next y
        Next x
        
        'Safely deallocate imageData()
        CopyMemory ByVal VarPtrArray(imageData), 0&, 4
        
        'If normalization was requested, and the data isn't already normalized, normalize it now
        If toNormalize And ((minVal > 0) Or (maxVal < 255)) Then
            
            Dim curRange As Long
            curRange = maxVal - minVal
            
            'Prevent DBZ errors
            If (curRange = 0) Then curRange = 1
            
            'Build a normalization lookup table
            Dim normalizedLookup() As Byte
            ReDim normalizedLookup(0 To 255) As Byte
            
            For x = 0 To 255
                
                grayVal = (CDbl(x - minVal) / CDbl(curRange)) * 255
                
                If (grayVal < 0) Then
                    grayVal = 0
                ElseIf (grayVal > 255) Then
                    grayVal = 255
                End If
                
                normalizedLookup(x) = grayVal
                
            Next x
            
            For x = initX To finalX
            For y = initY To finalY
                dstGrayArray(x, y) = normalizedLookup(dstGrayArray(x, y))
            Next y
            Next x
        
        End If
                
        GetDIBGrayscaleMap = True
        
    Else
        Debug.Print "WARNING! Non-existent DIB passed to getDIBGrayscaleMap."
        GetDIBGrayscaleMap = False
    End If

End Function

'Given a grayscale map (2D byte array), create a matching grayscale DIB from it.
' (Note: this function does not support progress bar reports, by design.)
Public Function CreateDIBFromGrayscaleMap(ByRef dstDIB As pdDIB, ByRef srcGrayArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long) As Boolean
    
    'Create the DIB
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    If dstDIB.CreateBlank(arrayWidth, arrayHeight, 32, 0, 255) Then
    
        'Point a local array at the DIB
        Dim dstImageData() As Byte
        Dim tmpSA As SAFEARRAY2D
        PrepSafeArray tmpSA, dstDIB
        CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(tmpSA), 4
        
        'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = dstDIB.GetDIBWidth - 1
        finalY = dstDIB.GetDIBHeight - 1
        
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim quickVal As Long, qvDepth As Long
        qvDepth = dstDIB.GetDIBColorDepth \ 8
        
        'Now we can loop through each pixel in the image, converting values as we go
        For x = initX To finalX
            quickVal = x * qvDepth
        For y = initY To finalY
            dstImageData(quickVal, y) = srcGrayArray(x, y)
            dstImageData(quickVal + 1, y) = srcGrayArray(x, y)
            dstImageData(quickVal + 2, y) = srcGrayArray(x, y)
        Next y
        Next x
        
        'Safely deallocate imageData()
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
                
        CreateDIBFromGrayscaleMap = True
        
    Else
        Debug.Print "WARNING! Could not create blank DIB inside createDIBFromGrayscaleMap."
        CreateDIBFromGrayscaleMap = False
    End If

End Function

'Convert a DIB to its grayscale equivalent.  (Note that this function does not support progress bar reports, by design.)
Public Function MakeDIBGrayscale(ByRef srcDIB As pdDIB, Optional ByVal numOfShades As Long = 256, Optional ByVal ignoreMagicMagenta As Boolean = True) As Boolean
    
    'Make sure the DIB exists
    If srcDIB Is Nothing Then Exit Function
    
    'Make sure the source DIB isn't empty
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
    
        'Create a local array and point it at the pixel data we want to operate on
        Dim imageData() As Byte
        Dim tmpSA As SAFEARRAY2D
        PrepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim qvDepth As Long
        qvDepth = srcDIB.GetDIBColorDepth \ 8
        
        'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = (srcDIB.GetDIBWidth - 1) * qvDepth
        finalY = (srcDIB.GetDIBHeight - 1)
        
        Dim r As Long, g As Long, b As Long, a As Long, grayVal As Long
        
        'Premultiplication requires a lot of int/float conversions.  To speed things up, we'll use a persistent look-up table
        ' for converting single bytes on the range [0, 255] to 4-byte floats on the range [0, 1].
        Dim alphaIsPremultiplied As Boolean: alphaIsPremultiplied = srcDIB.GetAlphaPremultiplication
        Dim applyPremult() As Single, removePremult() As Single, tmpAlphaModifier As Single
        ReDim applyPremult(0 To 255) As Single: ReDim removePremult(0 To 255) As Single
        
        If alphaIsPremultiplied Then
            For x = 0 To 255
                applyPremult(x) = x / 255
                If (x <> 0) Then removePremult(x) = 255 / x
            Next x
        End If
        
        'Grayscale values are very look-up friendly
        Dim conversionFactor As Double
        conversionFactor = (255 / (numOfShades - 1))
        
        'Build a look-up table for our custom grayscale conversion results
        Dim gLookUp(0 To 255) As Byte
        For x = 0 To 255
            grayVal = Int((CDbl(x) / conversionFactor) + 0.5) * conversionFactor
            If grayVal > 255 Then grayVal = 255
            gLookUp(x) = CByte(grayVal)
        Next x
            
        'Now we can loop through each pixel in the image, converting values as we go
        For y = initY To finalY
        For x = initX To finalX Step qvDepth
                
            'Get the source pixel color values
            b = imageData(x, y)
            g = imageData(x + 1, y)
            r = imageData(x + 2, y)
            
            'Remove premultiplication, as necessary
            If alphaIsPremultiplied Then
                a = imageData(x + 3, y)
                If (a <> 255) Then
                    tmpAlphaModifier = removePremult(a)
                    r = (r * tmpAlphaModifier)
                    g = (g * tmpAlphaModifier)
                    b = (b * tmpAlphaModifier)
                End If
            End If
            
            'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
            grayVal = (213 * r + 715 * g + 72 * b) \ 1000
            If grayVal > 255 Then grayVal = 255
            
            'If less than 256 shades are in play, calculate that now as well
            grayVal = gLookUp(grayVal)
            
            'If alpha is premultiplied, calculate that now
            If alphaIsPremultiplied Then
                If (a <> 255) Then
                    tmpAlphaModifier = applyPremult(a)
                    grayVal = grayVal * tmpAlphaModifier
                End If
            End If
            
            If ignoreMagicMagenta Then
                imageData(x, y) = grayVal
                imageData(x + 1, y) = grayVal
                imageData(x + 2, y) = grayVal
            Else
                If (r <> 253) Or (g <> 0) Or (b <> 253) Then
                    imageData(x, y) = grayVal
                    imageData(x + 1, y) = grayVal
                    imageData(x + 2, y) = grayVal
                End If
            End If
            
        Next x
        Next y
        
        CopyMemory ByVal VarPtrArray(imageData), 0&, 4
        MakeDIBGrayscale = True
        
    Else
        Debug.Print "WARNING! Non-existent DIB passed to DIBs.MakeDIBGrayscale()."
        MakeDIBGrayscale = False
    End If

End Function

'This function will calculate an "alpha-cutoff" for a 32bpp image.  This _Ex version (which is now the *only* version supported
' by PD) requires an input byte array, which will be initialized to the size of the image and filled with a copy of the image's
' new transparency data, if that cut-off were applied.  (e.g. the return array will only include values of 0 or 255 for each pixel).
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
                Exit Function
            End If
            
            Dim iData() As Byte, tmpSA As SAFEARRAY2D
            srcDIB.WrapArrayAroundDIB iData, tmpSA
            
            Dim chkAlpha As Byte
                
            'Loop through the image, checking alphas as we go
            For y = 0 To finalY
            For x = 0 To finalX
                
                chkAlpha = iData(x * 4 + 3, y)
                
                'If the alpha value is less than the cutoff, mark this pixel for exploration
                If (chkAlpha < cutOff) Then
                    dstTransparencyTable(x, y) = 0
                'If the pixel is not beneath the cut-off, and not fully opaque, composite it against the requested background
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
            
            'This function require unpremultiplied alpha
            Dim needToResetPremultiplication As Boolean: needToResetPremultiplication = False
            If srcDIB.GetAlphaPremultiplication Then
                srcDIB.SetAlphaPremultiplication False
                needToResetPremultiplication = True
            End If
            
            Dim iData() As Byte, tmpSA As SAFEARRAY2D
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
            
                Dim tmpRectF As RECTF
                CopyMemory_Strict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
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
                
                Dim iData() As Byte, tmpSA As SAFEARRAY2D
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
            
                Dim tmpRectF As RECTF
                CopyMemory_Strict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
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
                
                Dim iData() As Byte, tmpSA As SAFEARRAY2D
                srcDIB.WrapArrayAroundDIB iData, tmpSA
                
                Dim restorePremultiplication As Boolean: restorePremultiplication = False
                If srcDIB.GetAlphaPremultiplication Then
                    srcDIB.SetAlphaPremultiplication False, , ptrToRectF
                    restorePremultiplication = True
                End If
                    
                'Loop through the image, checking alphas as we go
                For y = initY To finalY
                For x = initX To finalX
                    iData(x * 4 + 3, y) = srcTransparencyTable(x, y)
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
            
            Dim iData() As Byte, tmpSA As SAFEARRAY2D
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
                            iData(xLookup, y) = Colors.BlendColors(chkB, backB, tmpAlpha)
                            iData(xLookup + 1, y) = Colors.BlendColors(chkG, backG, tmpAlpha)
                            iData(xLookup + 2, y) = Colors.BlendColors(chkR, backR, tmpAlpha)
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

'Forcibly colorize a DIB.  Alpha is preserved by this function.
'Returns: TRUE if successful; FALSE otherwise
Public Function ColorizeDIB(ByRef srcDIB As pdDIB, ByVal newColor As Long) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long, finalX As Long, finalY As Long, xLookup As Long
            finalX = (srcDIB.GetDIBWidth - 1)
            finalY = (srcDIB.GetDIBHeight - 1)
            
            Dim rLookup() As Byte, gLookUp() As Byte, bLookup() As Byte
            ReDim rLookup(0 To 255) As Byte, gLookUp(0 To 255) As Byte, bLookup(0 To 255) As Byte
            Dim chkA As Byte
            
            Dim targetR As Long, targetG As Long, targetB As Long
            targetR = Colors.ExtractRed(newColor)
            targetG = Colors.ExtractGreen(newColor)
            targetB = Colors.ExtractBlue(newColor)
            
            'Construct lookup tables with premultiplied RGB values.  This prevents us from needing
            ' to un-premultiply values in advance, and post-premultiply values afterward.
            Dim aFloat As Double
            For x = 0 To 255
                aFloat = CDbl(x) / 255
                rLookup(x) = targetR * aFloat
                gLookUp(x) = targetG * aFloat
                bLookup(x) = targetB * aFloat
            Next x
            
            Dim iData() As Byte, tmpSA As SAFEARRAY2D
            srcDIB.WrapArrayAroundDIB iData, tmpSA
                
            'Loop through the image, checking alphas as we go
            finalX = finalX * 4
            For y = 0 To finalY
            For x = 0 To finalX Step 4
                chkA = iData(x + 3, y)
                iData(x, y) = bLookup(chkA)
                iData(x + 1, y) = gLookUp(chkA)
                iData(x + 2, y) = rLookup(chkA)
            Next x
            Next y
    
            'With our alpha channel complete, point iData() away from the DIB and deallocate it
            srcDIB.UnwrapArrayFromDIB iData
            srcDIB.SetInitialAlphaPremultiplicationState True
            
            ColorizeDIB = True
            
        End If
    Else
        Debug.Print "WARNING!  DIBs.ColorizeDIB() requires a 32-bpp DIB to operate correctly."
    End If
    
End Function

'Outline a 32-bpp DIB.  The outline is drawn along the first-encountered border where transparent and opaque pixels meet.
' The caller must supply the outline pen they want used and optionally, an edge threshold on the range [0, 100].
'Returns: TRUE if successful; FALSE otherwise
Public Function OutlineDIB(ByRef srcDIB As pdDIB, ByRef outlinePen As pd2DPen, Optional ByVal edgeThreshold As Single = 50#) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long, finalX As Long, finalY As Long, xLookup As Long
            finalX = (srcDIB.GetDIBWidth - 1)
            finalY = (srcDIB.GetDIBHeight - 1)
            
            Dim srcPixels() As Byte, tmpSA As SAFEARRAY2D
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
            
            Dim finalPolygon() As POINTFLOAT, numOfPoints As Long
            
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
                ReDim finalPolygon(0 To 3) As POINTFLOAT
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
            Dim cPainter As pd2DPainter, cSurface As pd2DSurface
            Drawing2D.QuickCreatePainter cPainter
            Drawing2D.QuickCreateSurfaceFromDC cSurface, srcDIB.GetDIBDC, True
            cPainter.DrawPolygonF cSurface, outlinePen, numOfPoints, VarPtr(finalPolygon(0))
            Set cSurface = Nothing: Set cPainter = Nothing
            
            OutlineDIB = True
            
        End If
    Else
        Debug.Print "WARNING!  DIBs.OutlineDIB() requires a 32-bpp DIB to operate correctly."
    End If
    
End Function

'Return the number of unique RGBA entries used in a DIB.
'
'At present, this function is *not* well-optimized.  I only use it during resource file generation, so performance
' isn't a big deal right now.
'
'RETURNS: number of unique RGBA entries in an image, capped at 256 (e.g. 257 is returned if the number of unique
' entries is 257+).
Public Function GetDIBColorCount_RGBA(ByRef srcDIB As pdDIB) As Long

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) And (srcDIB.GetDIBColorDepth = 32) Then
        
        Dim srcPixels() As Byte, tmpSA As SAFEARRAY2D
        srcDIB.WrapArrayAroundDIB srcPixels, tmpSA
        
        Dim pxSize As Long
        pxSize = srcDIB.GetDIBColorDepth \ 8
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.GetDIBStride - 1
        finalY = srcDIB.GetDIBHeight - 1
        
        Dim colorList() As RGBQUAD
        ReDim colorList(0 To 255) As RGBQUAD
        
        'Always load the first color in advance; this lets us avoid branches in the inner loop
        Dim numColors As Long
        numColors = 1
        With colorList(0)
            .Blue = srcPixels(0, 0)
            .Green = srcPixels(1, 0)
            .Red = srcPixels(2, 0)
            .alpha = srcPixels(3, 0)
        End With
        
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
                If (b = colorList(i).Blue) Then
                    If (g = colorList(i).Green) Then
                        If (r = colorList(i).Red) Then
                            If (a = colorList(i).alpha) Then
                                matchFound = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next i
            
            'Add new colors to the list (until we reach 257; then bail)
            If (Not matchFound) Then
                If (numColors = 256) Then
                    numColors = 257
                    Exit For
                End If
                colorList(i).Blue = b
                colorList(i).Green = g
                colorList(i).Red = r
                colorList(i).alpha = a
                numColors = numColors + 1
            End If
            
        Next x
            If (numColors > 256) Then Exit For
        Next y
        
        srcDIB.UnwrapArrayFromDIB srcPixels
        
        GetDIBColorCount_RGBA = numColors
        
    End If
    
End Function

'Assuming a DIB has 256 colors or less (which you can confirm with the function above, if you need to), call this function
' to return two arrays: a palette array, and a one-byte-per-pixel palette array (with dimensions matching the original image).
'
'At present, this function is *not* optimized.  A naive palette search is used.  Also, the destination palette is in
' RGBA format (so alpha *does* matter when calculating colors.)
'
'RETURNS: number of colors in the destination palette (1-based).  If the return is 257, the function failed because there
' are too many colors in the source image.  Reduce the number of colors, then try again.
Public Function GetDIBAs8bpp_RGBA(ByRef srcDIB As pdDIB, ByRef dstPalette() As RGBQUAD, ByRef dstPixels() As Byte) As Long

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) And (srcDIB.GetDIBColorDepth = 32) Then
        
        Dim srcPixels() As Byte, tmpSA As SAFEARRAY2D
        srcDIB.WrapArrayAroundDIB srcPixels, tmpSA
        
        Dim pxSize As Long
        pxSize = srcDIB.GetDIBColorDepth \ 8
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.GetDIBStride - 1
        finalY = srcDIB.GetDIBHeight - 1
        
        ReDim dstPalette(0 To 255) As RGBQUAD
        ReDim dstPixels(0 To srcDIB.GetDIBWidth - 1, 0 To srcDIB.GetDIBHeight - 1) As Byte
        
        'Always load the first color in advance; this lets us avoid branches in the inner loop
        Dim numColors As Long
        numColors = 1
        With dstPalette(0)
            .Blue = srcPixels(0, 0)
            .Green = srcPixels(1, 0)
            .Red = srcPixels(2, 0)
            .alpha = srcPixels(3, 0)
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
                            If (a = dstPalette(i).alpha) Then
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
                    .alpha = a
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

'Given a palette, color count, and source palette index array, construct a matching 32-bpp DIB.
'
'IMPORTANT NOTE!  The destination DIB needs to be already constructed to match the source data's width and height.
' (Obviously, it needs to be 32-bpp too!)
'
'RETURNS: TRUE if successful; FALSE otherwise
Public Function GetRGBADIB_FromPalette(ByRef dstDIB As pdDIB, ByRef colorCount As Long, ByRef srcPalette() As RGBQUAD, ByRef srcPixels() As Byte) As Boolean

    If (dstDIB Is Nothing) Then Exit Function
    
    If (dstDIB.GetDIBDC <> 0) And (dstDIB.GetDIBWidth <> 0) And (dstDIB.GetDIBHeight <> 0) And (dstDIB.GetDIBColorDepth = 32) Then
        
        Dim dstPixels() As Byte, tmpSA As SAFEARRAY2D
        dstDIB.WrapArrayAroundDIB dstPixels, tmpSA
        
        Dim pxSize As Long
        pxSize = dstDIB.GetDIBColorDepth \ 8
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = dstDIB.GetDIBStride - 1
        finalY = dstDIB.GetDIBHeight - 1
        
        Dim colorIndex As Long, numOfColors As Long
        numOfColors = UBound(srcPalette) + 1
        
        For y = 0 To finalY
        For x = 0 To finalX Step pxSize
        
            colorIndex = srcPixels(x \ 4, y)
            
            If (colorIndex < numOfColors) Then
                With srcPalette(colorIndex)
                    dstPixels(x, y) = .Blue
                    dstPixels(x + 1, y) = .Green
                    dstPixels(x + 2, y) = .Red
                    dstPixels(x + 3, y) = .alpha
                End With
            End If
            
        Next x
        Next y
        
        dstDIB.UnwrapArrayFromDIB dstPixels
        
        GetRGBADIB_FromPalette = True
        
    End If
    
End Function

'This function returns a DIB, resized to meet a specific pixel count.  This is very helpful for things like image analysis,
' where a full-sized image copy doesn't meaningfully improve heuristics (but requires a hell of a lot longer to analyze).
'
'This function always preserves aspect ratio, and it will return the original image if the image is smaller than the number of
' pixels requested.  This simplifies outside functions, as you can always call this function prior to running heuristics.
Public Function ResizeDIBByPixelCount(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByVal numOfPixels As Long) As Boolean

    If (Not srcDIB Is Nothing) Then
        
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        
        'Calculate current megapixel count
        Dim srcWidth As Long, srcHeight As Long
        srcWidth = srcDIB.GetDIBWidth
        srcHeight = srcDIB.GetDIBHeight
        
        'If the source image has less megapixels than the requested amount, just return it as-is
        If (srcWidth * srcHeight < numOfPixels) Then
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
            GDI_Plus.GDIPlus_StretchBlt dstDIB, 0, 0, newWidth, newHeight, srcDIB, 0, 0, srcWidth, srcHeight, , GP_IM_HighQualityBicubic, , True, , True
            
            ResizeDIBByPixelCount = True
        
        End If
    
    End If

End Function

'Given a byte array, construct a 32-bpp DIB where each channel is set to the grayscale equivalent of the input array.  This is used
' with selection to generate a transparent + grayscale copy of a single byte array.  Note that the DIB *must* already exist as a
' 32-bpp DIB matching the size of the input table.
'
'Returns TRUE if successful, and srcDIB will be filled with a premultiplied 32-bpp DIB matching the input table
Public Function Construct32bppDIBFromByteMap(ByRef srcDIB As pdDIB, ByRef srcMap() As Byte) As Boolean

    If (srcDIB Is Nothing) Then Exit Function
    
    If (srcDIB.GetDIBColorDepth = 32) Then
        If (srcDIB.GetDIBDC <> 0) And (srcDIB.GetDIBWidth <> 0) And (srcDIB.GetDIBHeight <> 0) Then
            
            Dim x As Long, y As Long
            Dim allowedToProceed As Boolean: allowedToProceed = True
            
            Dim initX As Long, initY As Long, finalX As Long, finalY As Long
            initX = 0
            initY = 0
            finalX = (srcDIB.GetDIBWidth - 1)
            finalY = (srcDIB.GetDIBHeight - 1)
            
            'Construct a lookup table of premultiplied input values
            Dim lTable() As Long
            ReDim lTable(0 To 255) As Long
            
            Dim tmpR As Single, tmpG As Single, tmpB As Single, tmpA As Single
            Dim tmpQuad As RGBQUAD
            
            For x = 0 To 255
                tmpQuad.alpha = x
                tmpA = x * (x / 255)
                tmpQuad.Red = tmpA
                tmpQuad.Green = tmpA
                tmpQuad.Blue = tmpA
                CopyMemory ByVal VarPtr(lTable(x)), ByVal VarPtr(tmpQuad), LenB(tmpQuad)
            Next x
            
            Dim imgData() As Long, tmpSA As SAFEARRAY1D
            
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
