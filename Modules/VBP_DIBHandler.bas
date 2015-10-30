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

'DC API functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

'Object API functions
Private Const OBJ_BITMAP As Long = 7
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Clipboard interaction
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Const CLIPBOARD_FORMAT_BMP As Long = 2

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

'Copy a given DIB to the clipboard.  If FreeImage is available, PD will also copy the image in PNG format, so things like
' transparency are preserved.
Public Function copyDIBToClipboard(ByRef srcDIB As pdDIB) As Boolean
    
    'Make sure the DIB exists
    If srcDIB Is Nothing Then Exit Function
    
    'Make sure the DIB actually contains an image
    If (srcDIB.getDIBHandle <> 0) And (srcDIB.getDIBWidth <> 0) And (srcDIB.getDIBHeight <> 0) Then
    
        'We are going to copy the image data to the clipboard twice - once in PNG format, then again in standard BMP format.
        ' This maxmimizes operability between major software packages.
        
        'Start by using the vbAccelerator clipboard class, which makes this whole process a bit easier.
        Dim clpObject As cCustomClipboard
        Set clpObject = New cCustomClipboard
        If clpObject.ClipboardOpen(FormMain.hWnd) Then
        
            clpObject.ClearClipboard
        
            'FreeImage is required to perform the PNG transformation.  We could use GDI+, but FreeImage is easier.
            If g_ImageFormats.FreeImageEnabled And srcDIB.getDIBColorDepth = 32 Then
                
                'Most systems willl already have PNG available as a setting; the AddFormat function will detect this
                Dim PNGID As Long
                PNGID = clpObject.AddFormat("PNG")
                
                'Convert our current DIB to a FreeImage-type DIB
                Dim fi_DIB As Long
                fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
                
                'Convert the bitmap to PNG format, save it to an array, and release the original bitmap from memory
                Dim pngArray() As Byte
                Dim fi_Check As Long
                fi_Check = FreeImage_SaveToMemoryEx(FIF_PNG, fi_DIB, pngArray, FISO_PNG_Z_DEFAULT_COMPRESSION, True)
                
                'If the save was successful, hand the new PNG byte array to the clipboard
                If fi_Check Then
                    copyDIBToClipboard = True
                    clpObject.SetBinaryData PNGID, pngArray
                End If
                
            End If
            
            'With a PNG copy successfully saved, proceed to copy a standard 24bpp bitmap to the clipboard
            
            'Get a handle to the current desktop, and create a compatible clipboard device context in it
            Dim desktophWnd As Long
            desktophWnd = GetDesktopWindow
            
            Dim desktopDC As Long, clipboardDC As Long
            desktopDC = GetDC(desktophWnd)
            clipboardDC = CreateCompatibleDC(desktopDC)
            
            'If our temporary DC was created successfully, use it to create a temporary bitmap for the clipboard
            If (clipboardDC <> 0) Then
            
                'Create a bitmap compatible with the current desktop. This will receive the actual pixel data of the current DIB.
                Dim clipboardBMP As Long
                clipboardBMP = CreateCompatibleBitmap(desktopDC, srcDIB.getDIBWidth, srcDIB.getDIBHeight)
                
                If (clipboardBMP <> 0) Then
                    
                    'Place the compatible bitmap within the clipboard device context
                    Dim clipboardOldBMP As Long
                    clipboardOldBMP = SelectObject(clipboardDC, clipboardBMP)
                    
                    '24-bit images can be copied as-is
                    If srcDIB.getDIBColorDepth = 24 Then
                        BitBlt clipboardDC, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight, srcDIB.getDIBDC, 0, 0, vbSrcCopy
                    
                    '32-bit images must be composited against a white background first
                    Else
                    
                        Dim tmpDIB As pdDIB
                        Set tmpDIB = New pdDIB
                        
                        tmpDIB.createFromExistingDIB srcDIB
                        tmpDIB.convertTo24bpp
                        
                        BitBlt clipboardDC, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
                        
                        Set tmpDIB = Nothing
                        
                    End If
                    
                    'Remove the clipboard bitmap from the clipboard DC to leave room for the copy
                    SelectObject clipboardDC, clipboardOldBMP
        
                    'Copy the bitmap to the clipboard, then close and exit
                    clpObject.SetClipboardMemoryHandle CLIPBOARD_FORMAT_BMP, clipboardBMP
                    copyDIBToClipboard = True
                    
                Else
                    copyDIBToClipboard = False
                End If
                
                DeleteDC clipboardDC
                
            Else
                copyDIBToClipboard = False
            End If
            
            'Release (DON'T DELETE!) our control of the current desktop device context
            ReleaseDC desktophWnd, desktopDC
            
            'Release our hold on the clipboard
            clpObject.ClipboardClose
            
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  CopyDIBToClipboard could not open the clipboard."
            #End If
            copyDIBToClipboard = False
        End If
        
    Else
        copyDIBToClipboard = False
    End If

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
Public Function getDIBGrayscaleMap(ByRef srcDIB As pdDIB, ByRef dstGrayArray() As Byte, Optional ByVal toNormalize As Boolean = True) As Boolean
    
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
                
        getDIBGrayscaleMap = True
        
    Else
        Debug.Print "WARNING! Non-existent DIB passed to getDIBGrayscaleMap."
        getDIBGrayscaleMap = False
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

