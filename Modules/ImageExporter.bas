Attribute VB_Name = "ImageExporter"
'***************************************************************************
'Low-level image export interfaces
'Copyright 2001-2016 by Tanner Helland
'Created: 4/15/01
'Last updated: 14/March/16
'Last update: migrate various functions out of the high-level "Saving" module and into this new, format-specific module
'
'This module provides low-level "export" functionality for exporting image files out of PD.  You will not generally
' want to interface with this module directly; instead, rely on the high-level functions in the "Saving" module.
' They will intelligently drop into this module as necessary, sparing you the messy work of handling format-specific
' details (which are many, especially given PD's many "automatic" export features).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Public Enum PD_ALPHA_STATUS
    PDAS_NoAlpha = 0            'All alpha will be forcibly removed, and the image will be composited against a background color
    PDAS_BinaryAlpha = 1        'Alpha will be reduced to just 0 or just 255; semi-transparent pixels will be composited against a background color
    PDAS_ComplicatedAlpha = 2   'Alpha will be left intact (anything on the range [0, 255] is valid)
    PDAS_NewAlphaFromColor = 3  'A new alpha channel will be generated, with the specified color turned fully transparent, and all other pixels composited against a background color
End Enum

#If False Then
    Private Const PDAS_NoAlpha = 0, PDAS_BinaryAlpha = 1, PDAS_ComplicatedAlpha = 2, PDAS_NewAlphaFromColor = 3
#End If

'Given an input DIB, return the most relevant output color depth.  This will be a numeric value like "32" or "24".
' IMPORTANT NOTE: for best results, you must also handle the optional parameter "currentAlphaStatus", which has
'  three possible states.  If you are working with a format (like JPEG) that does not offer alpha support, convert
'  the incoming DIB to 24-bpp *prior* to calling this function; that will improve performance by skipping alpha
'  heuristics entirely.  Similarly, for a format like GIF, this function will return 8-bpp as the recommended
'  color depth, *but you still need to deal with the alpha result*.  You may need to forcibly crop alpha to 0 and 255
'  prior to exporting the GIF; PD provides a dialog for this.
Public Function AutoDetectOutputColorDepth(ByRef srcDIB As pdDIB, ByRef dstFormat As PHOTODEMON_IMAGE_FORMAT, Optional ByRef currentAlphaStatus As PD_ALPHA_STATUS = PDAS_NoAlpha, Optional ByRef uniqueColorCount As Long = 257, Optional ByRef isTrueColor As Boolean = True, Optional ByRef isGrayscale As Boolean = False, Optional ByRef isMonochrome As Boolean = False) As Long
    
    Dim colorCheckSuccessful As Boolean: colorCheckSuccessful = False
    
    'If the incoming image is already 24-bpp, we can skip the alpha-processing steps entirely.  However, it is not
    ' necessary for the caller to do this.  PD will provide correct results either way.
    If srcDIB.getDIBColorDepth = 24 Then
        currentAlphaStatus = PDAS_NoAlpha
        colorCheckSuccessful = AutoDetectColors_24BPPSource(srcDIB, uniqueColorCount, isGrayscale, isMonochrome)
        isTrueColor = CBool(uniqueColorCount > 256)
    
    'If the incoming image is 32-bpp, we will run additional alpha channel heuristics
    Else
        colorCheckSuccessful = AutoDetectColors_32BPPSource(srcDIB, uniqueColorCount, isGrayscale, isMonochrome, currentAlphaStatus)
        isTrueColor = CBool(uniqueColorCount > 256)
    End If
    
    'Any steps beyond this point are identical for 24- and 32-bpp sources.
    If colorCheckSuccessful Then
    
        'Based on the color count, grayscale-ness, and monochromaticity, return an appropriate recommended output depth
        ' for this image format.
        Select Case dstFormat
            
            'BMP files support output depths of 1, 4, 8, 24, and 32.  (16 is also supported, but it will never be auto-recommended.)
            ' Any alpha whatsoever results in a recommendation for 32-bpp, since paletted BMP files are unreliable with alpha data.
            Case PDIF_BMP
                If (currentAlphaStatus <> PDAS_NoAlpha) Then
                    AutoDetectOutputColorDepth = 32
                Else
                    If isTrueColor Then
                        AutoDetectOutputColorDepth = 24
                    Else
                        If isMonochrome Then
                            AutoDetectOutputColorDepth = 1
                        Else
                            If uniqueColorCount <= 16 Then
                                AutoDetectOutputColorDepth = 4
                            Else
                                AutoDetectOutputColorDepth = 8
                            End If
                        End If
                    End If
                End If
            
            'GIF files always recommend an output depth of 8-bpp, regardless of the presence of alpha.  Specific details of
            ' alpha-handling is left to the caller.
            Case PDIF_GIF
                AutoDetectOutputColorDepth = 8
            
            'It's technically pointless to pass HDR files to this function, as they are always output at 96-bpp RGBF
            Case PDIF_HDR
                AutoDetectOutputColorDepth = 96
            
            'JPEG-2000 files support 8-bpp, 24-bpp, and 32-bpp.  Meaningful alpha values result in a recommendation for 32-bpp.
            Case PDIF_JP2
                If (currentAlphaStatus <> PDAS_NoAlpha) Then
                    AutoDetectOutputColorDepth = 32
                Else
                    If isTrueColor Then
                        AutoDetectOutputColorDepth = 24
                    Else
                        AutoDetectOutputColorDepth = 8
                    End If
                End If
            
            'JPEG files are always 24-bpp, unless the source is grayscale.  Then we will recommend 8-bpp.
            Case PDIF_JPEG
                If isGrayscale Then
                    AutoDetectOutputColorDepth = 8
                Else
                    AutoDetectOutputColorDepth = 24
                End If
            
            'JPEG-XR files support 1, 8, 16, 24, and 32-bpp.  Alpha always results in a recommendation of 32-bpp.
            ' 16-bpp is never auto-recommended.  High bit-depths are also suppored, but never (currently) recommended.
            Case PDIF_JXR
                If (currentAlphaStatus <> PDAS_NoAlpha) Then
                    AutoDetectOutputColorDepth = 32
                Else
                    If isTrueColor Then
                        AutoDetectOutputColorDepth = 24
                    Else
                        If isMonochrome Then
                            AutoDetectOutputColorDepth = 1
                        Else
                            AutoDetectOutputColorDepth = 8
                        End If
                    End If
                End If
            
            'PDI files should not technically be passed to this function, as it's a big fat waste of time.  They always
            ' recommend 32-bpp output.
            Case PDIF_PDI
                AutoDetectOutputColorDepth = 32
            
            'PNG files are by far the worst ones to deal with, as they support a bunch of weird, PNG-specific color formats.
            ' Alpha is particularly problematic since it is supported in multiple different ways (transparent key, even in
            ' 24-bpp mode, or a full alpha channel, among others).  As such, the caller must exercise caution, as indexed color
            ' mode may still require 32-bpp saving, if the 256 available colors are mapped to more than 256 variants of color+alpha.
            Case PDIF_PNG
                
                If isTrueColor Then
                    If (currentAlphaStatus = PDAS_NoAlpha) Then
                        AutoDetectOutputColorDepth = 24
                    Else
                        AutoDetectOutputColorDepth = 32
                    End If
                
                'Non-truecolor images are less pleasant to work with, as the presence of alpha complicates everything.
                Else
                    If (currentAlphaStatus = PDAS_ComplicatedAlpha) Then
                        AutoDetectOutputColorDepth = 8
                    
                    'Binary alpha is technically supported by any color-depth, so we don't have to treat it differently from
                    ' non-alpha images
                    Else
                        If isMonochrome Then
                            AutoDetectOutputColorDepth = 1
                        Else
                            If uniqueColorCount <= 16 Then
                                AutoDetectOutputColorDepth = 4
                            Else
                                AutoDetectOutputColorDepth = 8
                            End If
                        End If
                    End If
                End If
            
            'PPM only support non-alpha, 24-bpp images
            Case PDIF_PPM
                AutoDetectOutputColorDepth = 24
            
            'PSD supports multiple bit-depths, but at present, we limit it to 24 or 32 only
            Case PDIF_PSD
                If (currentAlphaStatus = PDAS_NoAlpha) Then
                    AutoDetectOutputColorDepth = 24
                Else
                    AutoDetectOutputColorDepth = 32
                End If
            
            'TGA files support 1, 8, 24, and 32-bpp modes.  Basic GIF-like alpha is supported in 8-bpp mode; anything more
            ' complicated requires 32-bpp.  16-bpp mode is supported, but never recommended.
            Case PDIF_TARGA
                If (currentAlphaStatus <> PDAS_NoAlpha) Then
                    If (currentAlphaStatus = PDAS_ComplicatedAlpha) Then
                        AutoDetectOutputColorDepth = 32
                    Else
                        If isTrueColor Then
                            AutoDetectOutputColorDepth = 32
                        Else
                            AutoDetectOutputColorDepth = 8
                        End If
                    End If
                Else
                    If isTrueColor Then
                        AutoDetectOutputColorDepth = 24
                    Else
                        If isMonochrome Then
                            AutoDetectOutputColorDepth = 1
                        Else
                            AutoDetectOutputColorDepth = 8
                        End If
                    End If
                End If
            
            'TIFF files support 1, 4, 8, 24, and 32-bpp modes.  Basic GIF-like alpha is supported in 8-bpp mode; anything more
            ' complicated requires 32-bpp.  Higher bit-depths are supported, but never recommended.
            Case PDIF_TIFF
                If (currentAlphaStatus <> PDAS_NoAlpha) Then
                    If (currentAlphaStatus = PDAS_ComplicatedAlpha) Then
                        AutoDetectOutputColorDepth = 32
                    Else
                        If isTrueColor Then
                            AutoDetectOutputColorDepth = 32
                        Else
                            AutoDetectOutputColorDepth = 8
                        End If
                    End If
                Else
                    If isTrueColor Then
                        AutoDetectOutputColorDepth = 24
                    Else
                        If isMonochrome Then
                            AutoDetectOutputColorDepth = 1
                        Else
                            If uniqueColorCount <= 16 Then
                                AutoDetectOutputColorDepth = 4
                            Else
                                AutoDetectOutputColorDepth = 8
                            End If
                        End If
                    End If
                End If
            
            'WebP currently supports only 24-bpp and 32-bpp modes, and 32-bpp is forcibly disallowed if alpha is present
            ' (due to the way the FreeImage encoder works, at least - I have no idea if this is to spec or not).
            Case PDIF_WEBP
                If (currentAlphaStatus = PDAS_NoAlpha) Then
                    AutoDetectOutputColorDepth = 24
                Else
                    AutoDetectOutputColorDepth = 32
                End If
            
        End Select
        
    End If

End Function

'Given a 24-bpp source (the source *MUST BE 24-bpp*), fill three inputs:
' 1) netColorCount: an integer on the range [1, 257].  257 = more than 256 unique colors
' 2) isGrayscale: TRUE if the image consists of only gray shades
' 3) isMonochrome: TRUE if the image consists of only black and white
'
'The function as a whole returns TRUE if the source image was scanned correctly; FALSE otherwise.  (FALSE probably means you passed
' it a 32-bpp image!)
Private Function AutoDetectColors_24BPPSource(ByRef srcDIB As pdDIB, ByRef numUniqueColors As Long, ByRef isGrayscale As Boolean, ByRef isMonochrome As Boolean) As Boolean
    
    AutoDetectColors_24BPPSource = False
    
    If srcDIB.getDIBColorDepth = 24 Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Analyzing color count of 24-bpp image..."
        #End If
        
        Dim srcPixels() As Byte, tmpSA As SAFEARRAY2D
        PrepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcPixels()), VarPtr(tmpSA), 4
        
        Dim x As Long, y As Long, finalX As Long, finalY As Long
        finalY = srcDIB.getDIBHeight - 1
        finalX = srcDIB.getDIBWidth - 1
        finalX = finalX * 3
        
        Dim UniqueColors() As Long
        ReDim UniqueColors(0 To 255) As Long
        
        Dim i As Long
        For i = 0 To 255
            UniqueColors(i) = -1
        Next i
        
        numUniqueColors = 0
        
        'Finally, a bunch of variables used in color calculation
        Dim r As Long, g As Long, b As Long
        Dim chkValue As Long
        Dim colorFound As Boolean
            
        'Apply the filter
        For y = 0 To finalY
        For x = 0 To finalX Step 3
            
            b = srcPixels(x, y)
            g = srcPixels(x + 1, y)
            r = srcPixels(x + 2, y)
            
            chkValue = RGB(r, g, b)
            colorFound = False
            
            'Now, loop through the colors we've accumulated thus far and compare this entry against each of them.
            For i = 0 To numUniqueColors - 1
                If UniqueColors(i) = chkValue Then
                    colorFound = True
                    Exit For
                End If
            Next i
            
            'If colorFound is still false, store this value in the array and increment our color counter
            If (Not colorFound) Then
                If (numUniqueColors >= 256) Then
                    numUniqueColors = 257
                    Exit For
                Else
                    UniqueColors(numUniqueColors) = chkValue
                    numUniqueColors = numUniqueColors + 1
                End If
            End If
            
        Next x
            If numUniqueColors > 256 Then Exit For
        Next y
        
        CopyMemory ByVal VarPtrArray(srcPixels), 0&, 4
        
        'By default, we assume that an image is neither monochrome nor grayscale
        isGrayscale = False
        isMonochrome = False
        
        'Further checks are only relevant if the image contains 256 colors or less
        If numUniqueColors <= 256 Then
            
            'Check for grayscale images
            isGrayscale = True
        
            'Loop through all available colors
            For i = 0 To numUniqueColors - 1
            
                r = ExtractR(UniqueColors(i))
                g = ExtractG(UniqueColors(i))
                
                'If any of the components do not match, this is not a grayscale image
                If (r <> g) Then
                    isGrayscale = False
                    Exit For
                Else
                    b = ExtractB(UniqueColors(i))
                    If (b <> r) Or (b <> g) Then
                        isGrayscale = False
                        Exit For
                    End If
                End If
                
            Next i
            
            'If the image is grayscale and it only contains two colors, check for monochrome next
            ' (where monochrome = pure black and pure white, only).
            If isGrayscale And (numUniqueColors = 2) Then
            
                r = ExtractR(UniqueColors(0))
                g = ExtractG(UniqueColors(0))
                b = ExtractB(UniqueColors(0))
                
                If ((r = 0) And (g = 0) And (b = 0)) Or ((r = 255) And (g = 255) And (b = 255)) Then
                    r = ExtractR(UniqueColors(1))
                    g = ExtractG(UniqueColors(1))
                    b = ExtractB(UniqueColors(1))
                    If ((r = 0) And (g = 0) And (b = 0)) Or ((r = 255) And (g = 255) And (b = 255)) Then isGrayscale = True
                End If
            
            'End monochrome check
            End If
        
        'End "If 256 colors or less..."
        End If
        
        AutoDetectColors_24BPPSource = True
        
    End If

End Function


'Given a 32-bpp source (the source *MUST BE 32-bpp*, but its alpha channel can be constant), fill four inputs:
' 1) netColorCount: an integer on the range [1, 257].  257 = more than 256 unique colors
' 2) isGrayscale: TRUE if the image consists of only gray shades
' 3) isMonochrome: TRUE if the image consists of only black and white
' 4) currentAlphaStatus: custom enum describing the alpha channel contents of the image
'
'The function as a whole returns TRUE if the source image was scanned correctly; FALSE otherwise.  (FALSE probably means you passed
' it a 24-bpp image!)
Private Function AutoDetectColors_32BPPSource(ByRef srcDIB As pdDIB, ByRef netColorCount As Long, ByRef isGrayscale As Boolean, ByRef isMonochrome As Boolean, ByRef currentAlphaStatus As PD_ALPHA_STATUS) As Boolean

    AutoDetectColors_32BPPSource = False

    If (srcDIB.getDIBColorDepth = 32) Then

        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Analyzing color count of 32-bpp image..."
        #End If

        Dim srcPixels() As Byte, tmpSA As SAFEARRAY2D
        PrepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcPixels()), VarPtr(tmpSA), 4

        Dim x As Long, y As Long, finalX As Long, finalY As Long
        finalY = srcDIB.getDIBHeight - 1
        finalX = srcDIB.getDIBWidth - 1
        finalX = finalX * 4

        Dim UniqueColors() As RGBQUAD
        ReDim UniqueColors(0 To 255) As RGBQUAD

        Dim i As Long
        For i = 0 To 255
            UniqueColors(i).Red = 1
            UniqueColors(i).Green = 1
            UniqueColors(i).Blue = 0
            UniqueColors(i).alpha = 1
        Next i

        'Total number of unique colors counted so far
        Dim numUniqueColors As Long, non255Alpha As Boolean, nonBinaryAlpha As Boolean
        numUniqueColors = 0
        non255Alpha = False
        nonBinaryAlpha = False
        
        'Finally, a bunch of variables used in color calculation
        Dim r As Long, g As Long, b As Long, a As Long
        Dim chkValue As Long
        Dim colorFound As Boolean

        'Apply the filter
        For y = 0 To finalY
        For x = 0 To finalX Step 4
            
            b = srcPixels(x, y)
            g = srcPixels(x + 1, y)
            r = srcPixels(x + 2, y)
            a = srcPixels(x + 3, y)
            
            If (a < 255) Then
                non255Alpha = True
                If (a > 0) Then nonBinaryAlpha = True
            End If
            
            'Until we find at least 257 unique colors, we need to keep checking individual pixels
            If (numUniqueColors <= 256) Then
                
                colorFound = False
                
                'Now, loop through the colors we've accumulated thus far and compare this entry against each of them.
                For i = 0 To numUniqueColors - 1
                    If (UniqueColors(i).Red = r) Then
                        If (UniqueColors(i).Green = g) Then
                            If (UniqueColors(i).Blue = b) Then
                                If (UniqueColors(i).alpha = a) Then
                                    colorFound = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next i
    
                'If colorFound is still false, store this value in the array and increment our color counter
                If (Not colorFound) Then
                    If (numUniqueColors >= 256) Then
                        numUniqueColors = 257
                        If nonBinaryAlpha Then Exit For
                    Else
                        UniqueColors(numUniqueColors).Red = r
                        UniqueColors(numUniqueColors).Green = g
                        UniqueColors(numUniqueColors).Blue = b
                        UniqueColors(numUniqueColors).alpha = a
                        numUniqueColors = numUniqueColors + 1
                    End If
                End If
                
            End If
            
        Next x
            If (numUniqueColors > 256) And nonBinaryAlpha Then Exit For
        Next y

        CopyMemory ByVal VarPtrArray(srcPixels), 0&, 4

        netColorCount = numUniqueColors

        'By default, we assume that an image is neither monochrome nor grayscale
        isGrayscale = False
        isMonochrome = False

        'Further checks are only relevant if the image contains 256 colors or less
        If numUniqueColors <= 256 Then

            'Check for grayscale images
            If (numUniqueColors <= 256) Then

                isGrayscale = True

                'Loop through all available colors
                For i = 0 To numUniqueColors - 1
                    
                    'If any of the components do not match, this is not a grayscale image
                    If (UniqueColors(i).Red <> UniqueColors(i).Green) Then
                        isGrayscale = False
                        Exit For
                    Else
                        If (UniqueColors(i).Blue <> UniqueColors(i).Red) Or (UniqueColors(i).Blue <> UniqueColors(i).Green) Then
                            isGrayscale = False
                            Exit For
                        End If
                    End If

                Next i

            'End grayscale check
            End If

            'If the image is grayscale and it only contains two colors, check for monochrome next
            ' (where monochrome = pure black and pure white, only).
            If isGrayscale And (numUniqueColors = 2) Then
            
                If ((UniqueColors(i).Red = 0) And (UniqueColors(i).Green = 0) And (UniqueColors(i).Blue = 0)) Or ((UniqueColors(i).Red = 255) And (UniqueColors(i).Green = 255) And (UniqueColors(i).Blue = 255)) Then
                    If ((UniqueColors(i).Red = 0) And (UniqueColors(i).Green = 0) And (UniqueColors(i).Blue = 0)) Or ((UniqueColors(i).Red = 255) And (UniqueColors(i).Green = 255) And (UniqueColors(i).Blue = 255)) Then isGrayscale = True
                End If

            'End monochrome check
            End If

        'End "If 256 colors or less..."
        End If
        
        'Convert our individual alpha trackers into the single "currentAlphaStatus" output, then exit
        If non255Alpha Then
            If nonBinaryAlpha Then
                currentAlphaStatus = PDAS_ComplicatedAlpha
            Else
                currentAlphaStatus = PDAS_BinaryAlpha
            End If
        Else
            currentAlphaStatus = PDAS_NoAlpha
        End If
        
        AutoDetectColors_32BPPSource = True

    End If

End Function

Private Sub ExportDebugMsg(ByVal debugMsg As String)
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction debugMsg
    #End If
End Sub

'Format-specific export functions follow.  A few notes on how these functions work.
' 1) All functions take four input parameters:
'    - [required] srcPDImage: the image to be saved
'    - [required] dstFile: destination path + filename + extension, as a single string
'    - [optional] formatParams: format-specific parameters, in XML format (created via pdParamXML)
'    - [optional] metadataParams: metadata-specific parameters, in XML format (created via pdParamXML)
'
' 2) Format-specific parameters must not be required for saving a proper image.  Default values must be intelligently
'     applied if the format-specific parameter string is missing.
'
' 3) Most formats can ignore the metadataParams string, as metadata handling is typically handled via separate
'     ExifTool-specific functions.  This string primarily exists for formats like JPEG, where metadata handling is
'     messy since some functionality is easier to handle inside FreeImage (like thumbnail generation).  Either way,
'     if a metadata string is generated for a given format, it will be supplied as a parameter, "just in case" the
'     export function needs to parse it.
'
' 4) All functions return success/failure by boolean.  (FreeImage-specific errors are logged and processed externally.)
'
' 5) Because these export functions interface with multiple parts of the program (including the batch processor), it is
'     very important that they maintain identical function signatures.  Any format-specific functionality needs to be
'     handled via the aforementioned XML parameter strings, and not via extra params.
Public Function ExportBMP(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportBMPError
    
    ExportBMP = False
    Dim sFileType As String: sFileType = "BMP"
    
    'Parse all relevant BMP parameters.  (See the BMP export dialog for details on how these are generated.)
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString formatParams
    
    Dim bmpCompression As Boolean, bmpForceGrayscale As Boolean, bmp16bpp_555Mode As Boolean, bmpCustomColors As Long
    bmpCompression = cParams.GetBool("BMPRLECompression", False)
    bmpForceGrayscale = cParams.GetBool("BMPForceGrayscale", False)
    bmp16bpp_555Mode = cParams.GetBool("BMP16bpp555", False)
    bmpCustomColors = cParams.GetLong("BMPIndexedColorCount", 256)
    
    Dim bmpBackgroundColor As Long, bmpFlipRowOrder As Boolean
    bmpBackgroundColor = cParams.GetLong("BMPBackgroundColor", vbWhite)
    bmpFlipRowOrder = cParams.GetBool("BMPFlipRowOrder", False)
    
    'Generate a composited image copy, with alpha automatically un-premultiplied
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, False
    
    'Retrieve the recommended output color depth of the image.
    Dim outputColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, desiredAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
    
    If StrComp(LCase$(cParams.GetString("BMPColorDepth", "Auto")), "auto", vbBinaryCompare) = 0 Then
        outputColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_BMP, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
        ExportDebugMsg "Color depth auto-detection returned " & CStr(outputColorDepth) & "bpp"
        
        'Because BMP files only support alpha in 32-bpp mode, we can ignore binary-alpha mode completely
        If (currentAlphaStatus = PDAS_NoAlpha) Then desiredAlphaStatus = PDAS_NoAlpha Else desiredAlphaStatus = PDAS_ComplicatedAlpha
        
    Else
        outputColorDepth = cParams.GetLong("BMPColorDepth", 32)
        If (outputColorDepth = 32) Then desiredAlphaStatus = PDAS_ComplicatedAlpha
    End If
    
    'BMP files support a number of custom alpha parameters, for legacy compatibility reasons.  These need to be applied manually.
    If (outputColorDepth = 32) Then
        If cParams.GetBool("BMPUseXRGB", False) Then
            tmpImageCopy.ForceNewAlpha 0
        Else
            If cParams.GetBool("BMPPremultiplyAlpha", False) Then tmpImageCopy.SetAlphaPremultiplication True
        End If
    
    'Because bitmaps do not support transparency < 32-bpp, remove transparency immediately if the output depth is < 32-bpp,
    ' and forgo any further alpha handling.
    Else
        tmpImageCopy.convertTo24bpp bmpBackgroundColor
        desiredAlphaStatus = PDAS_NoAlpha
    End If
    
    'If both GDI+ and FreeImage are missing, use our own internal methods to save the BMP file in its current state.
    ' (This is a measure of last resort, as the saved image is unlikely to match the requested output depth.)
    If (Not g_ImageFormats.GDIPlusEnabled) And (Not g_ImageFormats.FreeImageEnabled) Then
        tmpImageCopy.WriteToBitmapFile dstFile
        ExportBMP = True
    Else
    
        If g_ImageFormats.FreeImageEnabled Then
            
            Dim fi_DIB As Long
            fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, desiredAlphaStatus, currentAlphaStatus, , bmpBackgroundColor, isGrayscale Or bmpForceGrayscale, bmpCustomColors, Not bmp16bpp_555Mode)
            If bmpFlipRowOrder Then Outside_FreeImageV3.FreeImage_FlipVertically fi_DIB
            
            'Finally, prepare some BMP save flags.  If the user has requested RLE encoding, and this image is <= 8bpp,
            ' request RLE encoding from FreeImage.
            Dim BMPflags As Long: BMPflags = BMP_DEFAULT
            If (outputColorDepth = 8) And bmpCompression Then BMPflags = BMP_SAVE_RLE
            
            'Use that handle to save the image to BMP format, with required color conversion based on the outgoing color depth
            If (fi_DIB <> 0) Then
                ExportBMP = FreeImage_SaveEx(fi_DIB, dstFile, PDIF_BMP, BMPflags, outputColorDepth, , , , , True)
                If ExportBMP Then
                    ExportDebugMsg "Export to " & sFileType & " appears successful."
                Else
                    Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
                End If
            Else
                Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
                ExportBMP = False
            End If
            
        Else
            ExportBMP = GDIPlusSavePicture(srcPDImage, dstFile, ImageBMP, outputColorDepth)
        End If
    
    End If
    
    Exit Function
    
ExportBMPError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportBMP = False
    
End Function

Public Function ExportGIF(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportGIFError
    
    ExportGIF = False
    Dim sFileType As String: sFileType = "GIF"
    
    'Parse all relevant GIF parameters.  (See the GIF export dialog for details on how these are generated.)
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString formatParams
    
    'Only two parameters are mandatory; the others are used on an as-needed basis
    Dim gifColorMode As String, gifAlphaMode As String
    gifColorMode = cParams.GetString("GIFColorMode", "Auto")
    gifAlphaMode = cParams.GetString("GIFAlphaMode", "Auto")
    
    Dim gifAlphaCutoff As Long, gifColorCount As Long, gifBackgroundColor As Long, gifAlphaColor As Long
    gifAlphaCutoff = cParams.GetLong("GIFAlphaCutoff", 64)
    gifColorCount = cParams.GetLong("GIFColorCount", 256)
    gifBackgroundColor = cParams.GetLong("GIFBackgroundColor", vbWhite)
    gifAlphaColor = cParams.GetLong("GIFAlphaColor", RGB(255, 0, 255))
    
    'Some combinations of parameters invalidate other parameters.  Calculate any overrides now.
    Dim gifForceGrayscale As Boolean
    If StrComp(LCase$(gifColorMode), "gray", vbBinaryCompare) = 0 Then gifForceGrayscale = True Else gifForceGrayscale = False
    If StrComp(LCase$(gifColorMode), "auto", vbBinaryCompare) = 0 Then gifColorCount = 256
    
    Dim desiredAlphaStatus As PD_ALPHA_STATUS
    desiredAlphaStatus = PDAS_BinaryAlpha
    If StrComp(LCase$(gifAlphaMode), "none", vbBinaryCompare) = 0 Then desiredAlphaStatus = PDAS_NoAlpha
    If StrComp(LCase$(gifAlphaMode), "bycolor", vbBinaryCompare) = 0 Then
        desiredAlphaStatus = PDAS_NewAlphaFromColor
        gifAlphaCutoff = gifAlphaColor
    End If
    
    'Generate a composited image copy, with alpha automatically un-premultiplied
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, False
        
    'FreeImage provides the most comprehensive GIF encoder, so we prefer it whenever possible
    If g_ImageFormats.FreeImageEnabled Then
            
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, 8, desiredAlphaStatus, PDAS_ComplicatedAlpha, gifAlphaCutoff, gifBackgroundColor, gifForceGrayscale, gifColorCount)
        
        'Finally, prepare some GIF save flags.  If the user has requested RLE encoding, and this image is <= 8bpp,
        ' request RLE encoding from FreeImage.
        Dim GIFflags As Long: GIFflags = GIF_DEFAULT
        
        'Use that handle to save the image to GIF format, with required color conversion based on the outgoing color depth
        If (fi_DIB <> 0) Then
            ExportGIF = FreeImage_SaveEx(fi_DIB, dstFile, PDIF_GIF, GIFflags, FICD_8BPP, , , , , True)
            If ExportGIF Then
                ExportDebugMsg "Export to " & sFileType & " appears successful."
            Else
                Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            End If
        Else
            Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
            ExportGIF = False
        End If
        
    ElseIf g_ImageFormats.GDIPlusEnabled Then
        ExportGIF = GDIPlusSavePicture(srcPDImage, dstFile, ImageGIF, 8)
    Else
        ExportGIF = False
        Message "No %1 encoder found. Save aborted.", "JPEG"
    End If
    
    
    Exit Function
    
ExportGIFError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportGIF = False
    
End Function

Public Function ExportJPEG(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportJPEGError
    
    ExportJPEG = False
    Dim sFileType As String: sFileType = "JPEG"
    
    'Parse all relevant JPEG parameters.  (See the JPEG export dialog for details on how these are generated.)
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString formatParams
    
    Dim jpegQuality As Long
    jpegQuality = cParams.GetLong("JPEGQuality", 92)
    
    Dim jpegCompression As Long
    Const JPG_CMP_BASELINE = 0, JPG_CMP_OPTIMIZED = 1, JPG_CMP_PROGRESSIVE = 2
    Select Case cParams.GetLong("JPEGCompressionMode", JPG_CMP_OPTIMIZED)
        Case JPG_CMP_BASELINE
            jpegCompression = JPEG_BASELINE
            
        Case JPG_CMP_OPTIMIZED
            jpegCompression = JPEG_OPTIMIZE
            
        Case JPG_CMP_PROGRESSIVE
            jpegCompression = JPEG_OPTIMIZE Or JPEG_PROGRESSIVE
        
    End Select
    
    Dim jpegSubsampling As Long
    Const JPG_SS_444 = 0, JPG_SS_422 = 1, JPG_SS_420 = 2, JPG_SS_411 = 3
    Select Case cParams.GetLong("JPEGSubsampling", JPG_SS_422)
        Case JPG_SS_444
            jpegSubsampling = JPEG_SUBSAMPLING_444
        Case JPG_SS_422
            jpegSubsampling = JPEG_SUBSAMPLING_422
        Case JPG_SS_420
            jpegSubsampling = JPEG_SUBSAMPLING_420
        Case JPG_SS_411
            jpegSubsampling = JPEG_SUBSAMPLING_411
    End Select
    
    'Combine all FreeImage-specific flags into one master flag
    Dim jpegFlags As Long
    jpegFlags = jpegQuality Or jpegCompression Or jpegSubsampling
    
    'Generate a composited image copy, with alpha premultiplied (as we're just going to composite it, anyway)
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, True
    
    'JPEGs do not support alpha, so forcibly flatten the image (regardless of output color depth).
    ' We also apply a custom backcolor here (if one exists; white is used by default).
    Dim jpegBackgroundColor As Long
    jpegBackgroundColor = cParams.GetLong("JPEGBackgroundColor", vbWhite)
    If (tmpImageCopy.getDIBColorDepth = 32) Then tmpImageCopy.convertTo24bpp jpegBackgroundColor
    
    'Retrieve the recommended output color depth of the image.
    Dim outputColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, desiredAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
    Dim forceGrayscale As Boolean
    
    If StrComp(LCase$(cParams.GetString("JPEGColorDepth", "Auto")), "auto", vbBinaryCompare) = 0 Then
        outputColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_JPEG, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
        ExportDebugMsg "Color depth auto-detection returned " & CStr(outputColorDepth) & "bpp"
    Else
        outputColorDepth = cParams.GetLong("JPEGColorDepth", 24)
        If outputColorDepth = 8 Then forceGrayscale = True
    End If
    
    'FreeImage is our preferred export engine, but we can use GDI+ if we have to
    If g_ImageFormats.FreeImageEnabled Then
        
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, PDAS_NoAlpha, PDAS_NoAlpha, , vbWhite, isGrayscale Or forceGrayscale)
        
        'Use that handle to save the image to JPEG format, with required color conversion based on the outgoing color depth
        If (fi_DIB <> 0) Then
            
            'TODO!  Figure out how to best handle thumbnails...
'            'If a thumbnail has been requested, embed it now
'            If cParams.GetBool("JPEGThumbnail", False) Then
'                Dim fThumbnail As Long
'                fThumbnail = FreeImage_MakeThumbnail(fi_DIB, 100)
'                FreeImage_SetThumbnail fi_DIB, fThumbnail
'                FreeImage_Unload fThumbnail
'            End If
            
            'Immediately prior to saving, pass this image's resolution values (if any) to FreeImage.
            ' These values will be embedded in the JFIF header.
            FreeImage_SetResolutionX fi_DIB, srcPDImage.getDPI
            FreeImage_SetResolutionY fi_DIB, srcPDImage.getDPI
            
            ExportJPEG = FreeImage_SaveEx(fi_DIB, dstFile, PDIF_JPEG, jpegFlags, outputColorDepth, , , , , True)
            If ExportJPEG Then
                ExportDebugMsg "Export to " & sFileType & " appears successful."
            Else
                Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            End If
        Else
            Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
            ExportJPEG = False
        End If
        
    ElseIf g_ImageFormats.GDIPlusEnabled Then
        ExportJPEG = GDIPlusSavePicture(srcPDImage, dstFile, ImageJPEG, outputColorDepth, jpegQuality)
    Else
        ExportJPEG = False
        Message "No %1 encoder found. Save aborted.", "JPEG"
    End If
    
    Exit Function
    
ExportJPEGError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportJPEG = False
    
End Function

'Save an HDR (High-Dynamic Range) image
Public Function ExportHDR(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportHDRError
    
    ExportHDR = False
    Dim sFileType As String: sFileType = "HDR"
    
    If g_ImageFormats.FreeImageEnabled Then
        
        'TODO: parse incoming HDR parameters.  (FreeImage doesn't support any HDR export parameters at present, but we could still provide
        ' options for things like gamma correction, background color for 32-bpp images, etc.)
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
        cParams.SetParamString formatParams
        
        'Generate a composited image copy, with alpha automatically un-premultiplied
        Dim tmpImageCopy As pdDIB
        Set tmpImageCopy = New pdDIB
        srcPDImage.GetCompositedImage tmpImageCopy
        
        'HDR does not support alpha-channels, so convert to 24-bpp in advance
        If tmpImageCopy.getDIBColorDepth = 32 Then tmpImageCopy.convertTo24bpp
        
        'HDR only supports one output color depth, so auto-detection is unnecessary
        ExportDebugMsg "HDR format only supports one output depth, so color depth auto-detection was ignored."
            
        'Convert our current DIB to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(tmpImageCopy.getDIBDC)
        Set tmpImageCopy = Nothing
        
        If (fi_DIB <> 0) Then
            
            'Convert the image data to RGBF format
            Dim fi_FloatDIB As Long
            fi_FloatDIB = FreeImage_ConvertToRGBF(fi_DIB)
            FreeImage_Unload fi_DIB
            
            If (fi_FloatDIB <> 0) Then
                
                'Prior to saving, we must account for default 2.2 gamma correction.  We do this by iterating through the source, and modifying gamma
                ' values as we go.  (If we reduce gamma prior to RGBF conversion, quality will obviously be impacted due to clipping.)
                
                'This Single-type array will consistently be updated to point to the current line of pixels in the image (RGBF format, remember!)
                Dim srcImageData() As Single
                Dim srcSA As SAFEARRAY1D
                
                'Iterate through each scanline in the source image, copying it to destination as we go.
                Dim iWidth As Long, iHeight As Long, iScanWidth As Long, iLoopWidth As Long
                iWidth = FreeImage_GetWidth(fi_FloatDIB) - 1
                iHeight = FreeImage_GetHeight(fi_FloatDIB) - 1
                iScanWidth = FreeImage_GetPitch(fi_FloatDIB)
                iLoopWidth = FreeImage_GetWidth(fi_FloatDIB) * 3 - 1
                
                Dim srcF As Single
                
                Dim gammaCorrection As Double
                gammaCorrection = 1# / (1# / 2.2)
                
                Dim x As Long, y As Long
                
                For y = 0 To iHeight
                    
                    'Point a 1D VB array at this scanline
                    With srcSA
                        .cbElements = 4
                        .cDims = 1
                        .lBound = 0
                        .cElements = iScanWidth
                        .pvData = FreeImage_GetScanline(fi_FloatDIB, y)
                    End With
                    CopyMemory ByVal VarPtrArray(srcImageData), VarPtr(srcSA), 4
                    
                    'Iterate through this line, converting values as we go
                    For x = 0 To iLoopWidth
                        
                        'Retrieve the source values
                        srcF = srcImageData(x)
                        
                        'Apply 1/2.2 gamma correction
                        If srcF > 0 Then srcImageData(x) = srcF ^ gammaCorrection
                        
                    Next x
                    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
                Next y
                
                'With gamma properly accounted for, we can finally write the image out to file.
                ExportHDR = FreeImage_Save(PDIF_HDR, fi_FloatDIB, dstFile, 0)
                If ExportHDR Then
                    ExportDebugMsg "Export to " & sFileType & " appears successful."
                Else
                    Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
                End If
                
                FreeImage_Unload fi_FloatDIB
                
            Else
                ExportDebugMsg "HDR save failed; could not convert to RGBF"
                ExportHDR = False
            End If
                
        Else
            Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
            ExportHDR = False
        End If
        
    Else
        If (MacroStatus <> MacroBATCH) Then PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        ExportHDR = False
    End If
    
    Exit Function
        
ExportHDRError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportHDR = False
    
End Function

'Save to PSD (or PSB) format using the FreeImage library
Public Function ExportPSD(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportPSDError
    
    ExportPSD = False
    Dim sFileType As String: sFileType = "PSD"
    
    If g_ImageFormats.FreeImageEnabled Then
    
        'TODO: parse incoming PSD parameters.  (This requires a PSD export dialog, which I haven't constructed yet...)
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
        cParams.SetParamString formatParams
        
        Dim compressRLE As Boolean, usePSBFormat As Boolean
        compressRLE = True
        usePSBFormat = False
        
        If usePSBFormat Then sFileType = "PSB"
        
        'Generate a composited image copy, with alpha automatically un-premultiplied
        Dim tmpImageCopy As pdDIB
        Set tmpImageCopy = New pdDIB
        srcPDImage.GetCompositedImage tmpImageCopy, False
        
        'Retrieve the recommended output color depth of the image.
        ' (TODO: parse incoming params and honor requests for forced color-depths!)
        Dim outputColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, desiredAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
        outputColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_PSD, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
        ExportDebugMsg "Color depth auto-detection returned " & CStr(outputColorDepth) & "bpp"
        
        'Our PSD exporter is only a simplified one, so we can ignore binary-alpha mode completely
        If (currentAlphaStatus = PDAS_NoAlpha) Then desiredAlphaStatus = PDAS_NoAlpha Else desiredAlphaStatus = PDAS_ComplicatedAlpha
        
        'Similarly, because PSD is currently limited to 24-bpp or 32-bpp output, convert any non-transparent images to 24-bpp now
        If (desiredAlphaStatus = PDAS_NoAlpha) Or (outputColorDepth <= 24) Then tmpImageCopy.convertTo24bpp
        
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, desiredAlphaStatus, currentAlphaStatus)
        
        If (fi_DIB <> 0) Then
            
            Dim fi_Flags As Long: fi_Flags = 0&
            If compressRLE Then fi_Flags = fi_Flags Or PSD_RLE Else fi_Flags = fi_Flags Or PSD_NONE
            If usePSBFormat Then fi_Flags = fi_Flags Or PSD_PSB
            
            ExportPSD = FreeImage_SaveEx(fi_DIB, dstFile, PDIF_PSD, fi_Flags, outputColorDepth, , , , , True)
            If ExportPSD Then
                ExportDebugMsg "Export to " & sFileType & " appears successful."
            Else
                Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
            ExportPSD = False
        End If
    Else
        If (MacroStatus <> MacroBATCH) Then PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbExclamation + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Message "Save cannot be completed without FreeImage library."
        ExportPSD = False
    End If
    
    Exit Function
    
ExportPSDError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportPSD = False
    
End Function
