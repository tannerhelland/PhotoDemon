Attribute VB_Name = "DIB_Handler"
'***************************************************************************
'DIB Support Functions
'Copyright 2012-2015 by Tanner Helland
'Created: 27/March/15
'Last updated: 27/March/16
'Last update: start migrating rare functions out of pdDIB and into this module.
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

'Check to see if a 32bpp DIB is really 32bpp. (Basically, scan all pixels in the alpha channel. If all values are set to
' 255 or all values are set to 0, the caller can opt to rebuild the DIB in 24bpp mode.)
Public Function verifyDIBAlphaChannel(ByRef srcDIB As pdDIB) As Boolean

    'Make sure the DIB exists
    If srcDIB Is Nothing Then Exit Function

    'This is only useful for images with alpha channels. Exit if no alpha channel is present.
    If srcDIB.getDIBColorDepth <> 32 Then
        verifyDIBAlphaChannel = True
        Exit Function
    End If
    
    'This routine will fail if the width or height of a DIB is 0
    If srcDIB.getDIBWidth = 0 Or srcDIB.getDIBHeight = 0 Then
        verifyDIBAlphaChannel = True
        Exit Function
    End If

    'Start, as always, with a SafeArray
    Dim iData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
    Dim x As Long, y As Long, QuickX As Long
    Dim checkAlpha As Boolean, initAlpha As Double
    checkAlpha = False
    
    'Determine the alpha value of the top-left pixel. This will be used as our baseline value.
    initAlpha = iData(3, 0)
    
    'If initAlpha is something other than 255 or 0, we don't need to check the image
    If (initAlpha <> 0) And (initAlpha <> 255) Then
        
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        Erase iData
        
        verifyDIBAlphaChannel = True
        Exit Function
        
    End If
        
    'Loop through the image, comparing colors as we go
    For x = 0 To srcDIB.getDIBWidth - 1
        QuickX = x * 4
    For y = 0 To srcDIB.getDIBHeight - 1
        
        'Compare the alpha data for this pixel to the initial pixel. If they DO NOT match, this is a valid alpha channel.
        If initAlpha <> iData(QuickX + 3, y) Then
            checkAlpha = True
            Exit For
        End If
        
    Next y
    
        'If the alpha channel has been verified, exit this loop
        If checkAlpha Then Exit For
        
    Next x
    
    'With our check complete, point iData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(iData), 0&, 4
    Erase iData

    'Return checkAlpha. If varying alpha values were found, this function returns TRUE. If all values were the same,
    ' this function returns FALSE.
    verifyDIBAlphaChannel = checkAlpha

End Function

'Does a given DIB have "binary" transparency, e.g. does it have alpha values of only 0 or 255?
' (This is used to determine how transparency is handled when exporting to file formats like GIF, which do not support variable alpha.)
Public Function isDIBAlphaBinary(ByRef srcDIB As pdDIB, Optional ByVal checkForZero As Boolean = True) As Boolean

    'Make sure the DIB exists
    If srcDIB Is Nothing Then Exit Function

    'Make sure this DIB is 32bpp. If it isn't, running this function is pointless.
    If srcDIB.getDIBColorDepth = 32 Then

        'Make sure this DIB isn't empty
        If (srcDIB.getDIBDC <> 0) And (srcDIB.getDIBWidth <> 0) And (srcDIB.getDIBHeight <> 0) Then
    
            'Loop through the image and compare each alpha value against 0 and 255. If a value doesn't
            ' match either of this, this is a non-binary alpha channel and it must be handled specially.
            Dim iData() As Byte
            Dim tmpSA As SAFEARRAY2D
            prepSafeArray tmpSA, srcDIB
            CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
    
            Dim x As Long, y As Long, QuickX As Long
                
            'By default, assume that the image does not have a binary alpha channel. (This is the preferable
            ' default, as we will exit the loop IFF a non-0 or non-255 value is found.)
            Dim notBinary As Boolean
            notBinary = False
            
            Dim chkAlpha As Byte
                
            'Loop through the image, checking alphas as we go
            For x = 0 To srcDIB.getDIBWidth - 1
                QuickX = x * 4
            For y = 0 To srcDIB.getDIBHeight - 1
            
                'Retrieve the alpha value of the current pixel
                chkAlpha = iData(QuickX + 3, y)
                
                'For optimization reasons, this is stated as two IFs instead of an OR.
                If chkAlpha <> 255 Then
                
                    If checkForZero Then
                    
                        If chkAlpha <> 0 Then
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
            
            Erase iData
                
            'Exit
            isDIBAlphaBinary = Not notBinary
            Exit Function
                
        End If
        
    End If
    
    'If we made it to this line, something went horribly wrong or the user used this function incorrectly
    ' (e.g. calling it on a 24-bpp DIB).
    isDIBAlphaBinary = False

End Function

'Is a given DIB grayscale?  Determination is made by scanning each pixel and comparing RGB values to see if they match.
Public Function isDIBGrayscale(ByRef srcDIB As pdDIB) As Boolean
    
    'Make sure the DIB exists
    If srcDIB Is Nothing Then Exit Function
    
    'Make sure this DIB isn't empty
    If (srcDIB.getDIBDC <> 0) And (srcDIB.getDIBWidth <> 0) And (srcDIB.getDIBHeight <> 0) Then
    
        'Loop through the image and compare RGB values to determine grayscale or not.
        Dim iData() As Byte
        Dim tmpSA As SAFEARRAY2D
        prepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(iData()), VarPtr(tmpSA), 4
        
        Dim x As Long, y As Long, QuickX As Long
                        
        Dim r As Long, g As Long, b As Long
        
        Dim qvDepth As Long
        qvDepth = srcDIB.getDIBColorDepth \ 8
                        
        'Loop through the image, checking alphas as we go
        For x = 0 To srcDIB.getDIBWidth - 1
            QuickX = x * qvDepth
        For y = 0 To srcDIB.getDIBHeight - 1
            
            r = iData(QuickX + 2, y)
            g = iData(QuickX + 1, y)
            b = iData(QuickX, y)
            
            'For optimization reasons, this is stated as multiple IFs instead of an OR.
            If r <> g Then
                
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
        
                isDIBGrayscale = False
                Exit Function
                
            ElseIf g <> b Then
            
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
            
                isDIBGrayscale = False
                Exit Function
                
            ElseIf r <> b Then
            
                CopyMemory ByVal VarPtrArray(iData), 0&, 4
                Erase iData
            
                isDIBGrayscale = False
                Exit Function
                
            End If
                
        Next y
        Next x
    
        'With our alpha channel complete, point iData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(iData), 0&, 4
        Erase iData
                        
        'If we scanned all pixels without exiting prematurely, the DIB is grayscale
        isDIBGrayscale = True
        Exit Function
        
    End If
    
    'If we made it to this line, the DIB is blank, so it doesn't matter what value we return
    isDIBGrayscale = False

End Function

'Convert a source DIB to a 32-bit CMYK DIB.
' TODO: Use an ICC profile for this step, to improve quality.
Public Function createCMYKDIB(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Boolean

    'Make sure the source DIB exists
    If srcDIB Is Nothing Then Exit Function
    
    'Make sure the source DIB isn't empty
    If (srcDIB.getDIBDC <> 0) And (srcDIB.getDIBWidth <> 0) And (srcDIB.getDIBHeight <> 0) Then
    
        'Create the destination DIB as necessary
        If dstDIB Is Nothing Then Set dstDIB = New pdDIB
        
        'Create a 32-bit destination DIB with identical size to the source DIB
        dstDIB.createBlank srcDIB.getDIBWidth, srcDIB.getDIBHeight, 32, 0, 0
        
        'Prepare direct access to the source and destination DIB data
        Dim srcData() As Byte, dstData() As Byte
        Dim srcSA As SAFEARRAY2D, dstSA As SAFEARRAY2D
        prepSafeArray srcSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcData()), VarPtr(srcSA), 4
        
        prepSafeArray dstSA, dstDIB
        CopyMemory ByVal VarPtrArray(dstData()), VarPtr(dstSA), 4
        
        Dim x As Long, y As Long, QuickXSrc As Long, QuickXDst As Long
        Dim cyan As Long, magenta As Long, yellow As Long, k As Long
        
        Dim srcQVDepth As Long
        srcQVDepth = srcDIB.getDIBColorDepth \ 8
                        
        'Loop through the image, checking alphas as we go
        For x = 0 To srcDIB.getDIBWidth - 1
            QuickXSrc = x * srcQVDepth
            QuickXDst = x * 4
        For y = 0 To srcDIB.getDIBHeight - 1
                        
            'Cyan
            cyan = 255 - srcData(QuickXSrc + 2, y)
            
            'Magenta
            magenta = 255 - srcData(QuickXSrc + 1, y)
            
            'Yellow
            yellow = 255 - srcData(QuickXSrc, y)
            
            'Key
            k = CByte(Min3Int(cyan, magenta, yellow))
            dstData(QuickXDst + 3, y) = k
            
            If k = 255 Then
                dstData(QuickXDst, y) = 0
                dstData(QuickXDst + 1, y) = 0
                dstData(QuickXDst + 2, y) = 0
            Else
                dstData(QuickXDst, y) = cyan - k
                dstData(QuickXDst + 1, y) = magenta - k
                dstData(QuickXDst + 2, y) = yellow - k
            End If
                            
        Next y
        Next x
    
        'With our alpha channel complete, point both arrays away from their respective DIBs
        CopyMemory ByVal VarPtrArray(srcData), 0&, 4
        CopyMemory ByVal VarPtrArray(dstData), 0&, 4
        
        createCMYKDIB = True
        
    Else
        createCMYKDIB = False
    End If

End Function

'Given a DIB, return a 2D Byte array of the DIB's luminance values.  An optional preNormalize parameter will guarantee that the output
' stretches from 0 to 255.  (Also note: this function does not support progress bar reports.)
Public Function GetDIBGrayscaleMap(ByRef srcDIB As pdDIB, ByRef dstGrayArray() As Byte, Optional ByVal toNormalize As Boolean = True) As Boolean
    
    'Make sure the DIB exists
    If srcDIB Is Nothing Then Exit Function
    
    'Make sure the source DIB isn't empty
    If (srcDIB.getDIBDC <> 0) And (srcDIB.getDIBWidth <> 0) And (srcDIB.getDIBHeight <> 0) Then
    
        'Create a local array and point it at the pixel data we want to operate on
        Dim ImageData() As Byte
        Dim tmpSA As SAFEARRAY2D
        prepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
            
        'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = srcDIB.getDIBWidth - 1
        finalY = srcDIB.getDIBHeight - 1
        
        'Prep the destination array
        ReDim dstGrayArray(initX To finalX, initY To finalY) As Byte
        
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim QuickVal As Long, qvDepth As Long
        qvDepth = srcDIB.getDIBColorDepth \ 8
        
        Dim r As Long, g As Long, b As Long, grayVal As Long
        Dim minVal As Long, maxVal As Long
        minVal = 255
        maxVal = 0
            
        'Now we can loop through each pixel in the image, converting values as we go
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
                
            'Get the source pixel color values
            r = ImageData(QuickVal + 2, y)
            g = ImageData(QuickVal + 1, y)
            b = ImageData(QuickVal, y)
            
            'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
            grayVal = (213 * r + 715 * g + 72 * b) \ 1000
            If grayVal > 255 Then grayVal = 255
            
            'Cache the value
            dstGrayArray(x, y) = grayVal
            
            'If normalization has been requested, check max/min values now
            If toNormalize Then
                If grayVal < minVal Then
                    minVal = grayVal
                ElseIf grayVal > maxVal Then
                    maxVal = grayVal
                End If
            End If
            
        Next y
        Next x
        
        'With our work complete, point ImageData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
        
        'If normalization was requested, and the data isn't already normalized, normalize it now
        If toNormalize And ((minVal > 0) Or (maxVal < 255)) Then
            
            Dim curRange As Long
            curRange = maxVal - minVal
            
            'Prevent DBZ errors
            If curRange = 0 Then curRange = 1
            
            'Build a normalization lookup table
            Dim normalizedLookup() As Byte
            ReDim normalizedLookup(0 To 255) As Byte
            
            For x = 0 To 255
                
                grayVal = (CDbl(x - minVal) / CDbl(curRange)) * 255
                
                If grayVal < 0 Then
                    grayVal = 0
                ElseIf grayVal > 255 Then
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
Public Function createDIBFromGrayscaleMap(ByRef dstDIB As pdDIB, ByRef srcGrayArray() As Byte, ByVal arrayWidth As Long, ByVal arrayHeight As Long) As Boolean
    
    'Create the DIB
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    If dstDIB.createBlank(arrayWidth, arrayHeight, 32, 0, 255) Then
    
        'Point a local array at the DIB
        Dim dstImageData() As Byte
        Dim tmpSA As SAFEARRAY2D
        prepSafeArray tmpSA, dstDIB
        CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(tmpSA), 4
        
        'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = dstDIB.getDIBWidth - 1
        finalY = dstDIB.getDIBHeight - 1
        
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim QuickVal As Long, qvDepth As Long
        qvDepth = dstDIB.getDIBColorDepth \ 8
        
        'Now we can loop through each pixel in the image, converting values as we go
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
            
            dstImageData(QuickVal, y) = srcGrayArray(x, y)
            dstImageData(QuickVal + 1, y) = srcGrayArray(x, y)
            dstImageData(QuickVal + 2, y) = srcGrayArray(x, y)
            
        Next y
        Next x
        
        'With our work complete, point ImageData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
                
        createDIBFromGrayscaleMap = True
        
    Else
        Debug.Print "WARNING! Could not create blank DIB inside createDIBFromGrayscaleMap."
        createDIBFromGrayscaleMap = False
    End If

End Function

