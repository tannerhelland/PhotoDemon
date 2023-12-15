Attribute VB_Name = "ImageExporter"
'***************************************************************************
'Low-level image export interfaces
'Copyright 2001-2023 by Tanner Helland
'Created: 4/15/01
'Last updated: 22/June/22
'Last update: convert esoteric format exporters to safe-overwrite strategy
'
'This module provides low-level "export" functionality for exporting image files out of PD.
'
'You should not (generally) interface with this module directly; instead, rely on the high-level
' functions in the "Saving" module. They will intelligently drop into this module as necessary,
' sparing you the messy work of handling format-specific details (which are many, especially given
' PD's many "automatic detection" export features).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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

'Given an input DIB, return the most relevant output color depth.
'This will be a numeric value like "32" or "24".
'
'IMPORTANT NOTE: for best results, you must also handle the optional parameter "currentAlphaStatus",
' which has three possible states.  If you are working with a format like JPEG that doesn't support
' alpha channels, convert the incoming DIB to 24-bpp *prior* to calling this function; this improves
' performance by skipping alpha heuristics entirely.  Similarly, for legacy formats that only support
' sub-24-bpp color, this function will return 8-bpp as the recommended color depth *but you may still
' need to deal with the alpha result*, by e.g. thresholding alpha to binary on/off values.
'
'ANOTHER IMPORTANT NOTE: for some formats, this function is superceded by per-format logic.
' For example, PNG's color-depth representations are convoluted and PNG-specific, so PD's PNG exporter
' doesn't rely on this function at all.  This is why some PD export formats are not covered here.
Public Function AutoDetectOutputColorDepth(ByRef srcDIB As pdDIB, ByRef dstFormat As PD_IMAGE_FORMAT, Optional ByRef currentAlphaStatus As PD_ALPHA_STATUS = PDAS_NoAlpha, Optional ByRef uniqueColorCount As Long = 257, Optional ByRef isTrueColor As Boolean = True, Optional ByRef isGrayscale As Boolean = False, Optional ByRef isMonochrome As Boolean = False, Optional ByRef goodTransparentColor As Long = vbBlack) As Long
    
    Dim colorCheckSuccessful As Boolean: colorCheckSuccessful = False
    
    'If the outgoing image has 256 colors or less, we want to populate a color table with the auto-detected palette.
    ' This can be used to assess things like "4-bit grayscale" modes, which require us to validate individual
    ' grayscale values (to ensure they match 4-bit encoding patterns).
    Dim outPalette() As RGBQuad
    
    'If the incoming image is already 24-bpp, we can skip the alpha-processing steps entirely.  However, it is not
    ' necessary for the caller to do this.  PD will provide correct results either way.
    If (srcDIB.GetDIBColorDepth = 24) Then
        currentAlphaStatus = PDAS_NoAlpha
        colorCheckSuccessful = AutoDetectColors_24BPPSource(srcDIB, uniqueColorCount, isGrayscale, isMonochrome)
        isTrueColor = (uniqueColorCount > 256)
    
    'If the incoming image is 32-bpp, we will run additional alpha channel heuristics
    Else
        colorCheckSuccessful = AutoDetectColors_32BPPSource(srcDIB, uniqueColorCount, isGrayscale, isMonochrome, currentAlphaStatus, outPalette)
        isTrueColor = (uniqueColorCount > 256)
    End If
    
    'Any steps beyond this point are identical for 24- and 32-bpp sources.
    If colorCheckSuccessful Then
    
        'Based on the color count, grayscale-ness, and monochromaticity, return an appropriate recommended output depth
        ' for this image format.
        Select Case dstFormat
            
            'To be completely honest, I'm not sure what export depths should be used with our current strategy of
            ' PD > PNG > AVIF.  I currently limit output to 24- or 32-bit, but this likely needs to be revisited
            ' pending testing.
            Case PDIF_AVIF
                If (currentAlphaStatus = PDAS_NoAlpha) Then
                    AutoDetectOutputColorDepth = 24
                Else
                    AutoDetectOutputColorDepth = 32
                End If
            
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
                            If (uniqueColorCount <= 16) Then
                                AutoDetectOutputColorDepth = 4
                            Else
                                AutoDetectOutputColorDepth = 8
                            End If
                        End If
                    End If
                End If
            
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
            
            'PNM supports only non-alpha modes, but the file extension should really be changed to match the output depth
            Case PDIF_PNM
                If isTrueColor Then
                    AutoDetectOutputColorDepth = 24
                Else
                    If isMonochrome Then
                        AutoDetectOutputColorDepth = 1
                    Else
                        AutoDetectOutputColorDepth = 8
                    End If
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
            
            'WebP currently supports only 24-bpp and 32-bpp modes, and 32-bpp is forcibly disallowed if alpha is not present
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
    
    If srcDIB.GetDIBColorDepth = 24 Then
        
        PDDebug.LogAction "Analyzing color count of 24-bpp image..."
        
        Dim srcPixels() As Byte, tmpSA As SafeArray2D
        srcDIB.WrapArrayAroundDIB srcPixels, tmpSA
        
        Dim x As Long, y As Long, finalX As Long, finalY As Long
        finalY = srcDIB.GetDIBHeight - 1
        finalX = srcDIB.GetDIBWidth - 1
        finalX = finalX * 3
        
        Dim i As Long, uniqueColors(0 To 255) As Long
        For i = 0 To 255
            uniqueColors(i) = -1
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
                If uniqueColors(i) = chkValue Then
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
                    uniqueColors(numUniqueColors) = chkValue
                    numUniqueColors = numUniqueColors + 1
                End If
            End If
            
        Next x
            If (numUniqueColors > 256) Then Exit For
        Next y
        
        srcDIB.UnwrapArrayFromDIB srcPixels
        
        'By default, we assume that an image is neither monochrome nor grayscale
        isGrayscale = False
        isMonochrome = False
        
        'Further checks are only relevant if the image contains 256 colors or less
        If (numUniqueColors <= 256) Then
            
            'Check for grayscale images
            isGrayscale = True
        
            'Loop through all available colors
            For i = 0 To numUniqueColors - 1
            
                r = Colors.ExtractRed(uniqueColors(i))
                g = Colors.ExtractGreen(uniqueColors(i))
                
                'If any of the components do not match, this is not a grayscale image
                If (r <> g) Then
                    isGrayscale = False
                    Exit For
                Else
                    b = Colors.ExtractBlue(uniqueColors(i))
                    If (b <> r) Or (b <> g) Then
                        isGrayscale = False
                        Exit For
                    End If
                End If
                
            Next i
            
            'If the image is grayscale and it only contains two colors, check for monochrome next
            ' (where monochrome = pure black and pure white, only).
            If isGrayscale And (numUniqueColors <= 2) Then
            
                r = Colors.ExtractRed(uniqueColors(0))
                g = Colors.ExtractGreen(uniqueColors(0))
                b = Colors.ExtractBlue(uniqueColors(0))
                
                If ((r = 0) And (g = 0) And (b = 0)) Or ((r = 255) And (g = 255) And (b = 255)) Then
                    r = Colors.ExtractRed(uniqueColors(1))
                    g = Colors.ExtractGreen(uniqueColors(1))
                    b = Colors.ExtractBlue(uniqueColors(1))
                    If ((r = 0) And (g = 0) And (b = 0)) Or ((r = 255) And (g = 255) And (b = 255)) Then isMonochrome = True
                End If
            
            'End monochrome check
            End If
        
        'End "If 256 colors or less..."
        End If
        
        AutoDetectColors_24BPPSource = True
        
    End If

End Function

'Given a 32-bpp source (the source *MUST BE 32-bpp*, but its alpha channel can be constant), fill various critical
' pieces of information about the image's color+opacity makeup:
' 1) netColorCount: an integer on the range [1, 257].  257 = more than 256 unique colors
' 2) isGrayscale: TRUE if the image consists of only gray shades
' 3) isMonochrome: TRUE if the image consists of only black and white
' 4) currentAlphaStatus: custom enum describing the alpha channel contents of the image
' 5) uniqueColors(): if the image contains 256 unique color + opacity combinations (or less), this will return an exact palette
'
'The function as a whole returns TRUE if the source image was scanned correctly; FALSE otherwise.  (FALSE probably means you passed
' it a 24-bpp image!)
Private Function AutoDetectColors_32BPPSource(ByRef srcDIB As pdDIB, ByRef netColorCount As Long, ByRef isGrayscale As Boolean, ByRef isMonochrome As Boolean, ByRef currentAlphaStatus As PD_ALPHA_STATUS, ByRef uniqueColors() As RGBQuad) As Boolean

    AutoDetectColors_32BPPSource = False

    If (srcDIB.GetDIBColorDepth = 32) Then

        PDDebug.LogAction "Analyzing color count of 32-bpp image..."
        
        Dim srcPixels() As Byte, tmpSA As SafeArray2D
        srcDIB.WrapArrayAroundDIB srcPixels, tmpSA
        
        Dim x As Long, y As Long, finalX As Long, finalY As Long
        finalY = srcDIB.GetDIBHeight - 1
        finalX = srcDIB.GetDIBWidth - 1
        finalX = finalX * 4
        
        'Use a dedicated color counting class to collect a palette for this image
        Dim cColorTree As pdColorCount
        Set cColorTree = New pdColorCount
        cColorTree.SetAlphaTracking True
        
        Dim i As Long
        
        'Total number of unique colors counted so far
        Dim numUniqueColors As Long, non255Alpha As Boolean, nonBinaryAlpha As Boolean
        numUniqueColors = 0
        non255Alpha = False
        nonBinaryAlpha = False
        
        'Finally, a bunch of variables used in color calculation
        Dim r As Long, g As Long, b As Long, a As Long
        
        'Look for unique colors
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
                If cColorTree.AddColor(r, g, b, a) Then numUniqueColors = numUniqueColors + 1
                
                'Once more than 256 colors have been found, we no longer need to count colors, because we
                ' already know the image must be exported as 24-bit (or higher)
                If (numUniqueColors > 256) Then
                    numUniqueColors = 257
                    If nonBinaryAlpha Then Exit For
                End If
                
            End If
            
        Next x
            If (numUniqueColors > 256) And nonBinaryAlpha Then Exit For
        Next y
        
        srcDIB.UnwrapArrayFromDIB srcPixels
        
        netColorCount = numUniqueColors

        'By default, we assume that an image is neither monochrome nor grayscale
        isGrayscale = False
        isMonochrome = False

        'Further checks are only relevant if the image contains 256 colors or less
        If (numUniqueColors <= 256) Then
            
            'Retrieve the current color palette for this image
            cColorTree.GetPalette uniqueColors
            
            'Next, we want to see if the image is grayscale
            isGrayscale = True

            'Loop through all palette entries
            For i = 0 To numUniqueColors - 1
                
                'If any of the components do not match, this is not a grayscale image
                If (uniqueColors(i).Red <> uniqueColors(i).Green) Then
                    isGrayscale = False
                    Exit For
                Else
                    If (uniqueColors(i).Blue <> uniqueColors(i).Red) Or (uniqueColors(i).Blue <> uniqueColors(i).Green) Then
                        isGrayscale = False
                        Exit For
                    End If
                End If

            Next i
            
            'Grayscale images have some restrictions that paletted images do not (e.g. they cannot have
            ' variable per-index alpha values).  Check for these now.
            If isGrayscale Then
                
                'In the case of PNGs, grayscale images are not allowed to have variable transparency values.
                ' (This is likely true for other image formats as well - or at least, it's universal enough
                ' that we don't need to deviate according to image format.)  Note that values of 0 may be okay;
                ' these can be encoded using an alternate tRNS chunk.
                '
                'If there are any alpha values on the range [1, 254], consider this a non-grayscale image.
                If nonBinaryAlpha Then
                    isGrayscale = False
                
                'If the image doesn't contain weird alpha values, look for monochrome data specifically.
                Else
                    
                    'Check monochrome; monochrome images must only contain pure black and pure white.
                    If (numUniqueColors <= 2) Then
                        If ((uniqueColors(0).Red = 0) And (uniqueColors(0).Green = 0) And (uniqueColors(0).Blue = 0)) Or ((uniqueColors(0).Red = 255) And (uniqueColors(0).Green = 255) And (uniqueColors(0).Blue = 255)) Then
                            If (numUniqueColors = 2) Then
                                If ((uniqueColors(1).Red = 0) And (uniqueColors(1).Green = 0) And (uniqueColors(1).Blue = 0)) Or ((uniqueColors(1).Red = 255) And (uniqueColors(1).Green = 255) And (uniqueColors(1).Blue = 255)) Then isMonochrome = True
                            Else
                                isMonochrome = True
                            End If
                        End If
                    End If
                    
                End If
                
            'End "special" grayscale mode checks
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

Private Sub ExportDebugMsg(ByRef srcDebugMsg As String)
    PDDebug.LogAction srcDebugMsg
End Sub

'Format-specific export functions follow.  A few notes on how these functions work.
' 1) All functions take four input parameters:
'    - [required] srcPDImage: the image to be saved
'    - [required] dstFile: destination path + filename + extension, as a single string
'    - [optional] formatParams: format-specific parameters, in XML format (created via pdSerialize)
'    - [optional] metadataParams: metadata-specific parameters, in XML format (created via pdSerialize)
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

Public Function ExportAVIF(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportAVIFError
    
    ExportAVIF = False
    Dim sFileType As String: sFileType = "AVIF"
    
    'If this system is 64-bit capable but libavif doesn't exist, ask if we can download a copy
    If OS.OSSupports64bitExe And (Not Plugin_AVIF.IsAVIFExportAvailable()) Then
        
        If (Not Plugin_AVIF.PromptForLibraryDownload_AVIF()) Then GoTo ExportAVIFError
        
        'Downloading the AVIF plugins will raise new messages in the status bar; restore the original
        ' "saving %1 image" text
        Message "Saving %1 file...", sFileType
        
    End If
    
    'Failsafe check before proceeding
    If (Not Plugin_AVIF.IsAVIFExportAvailable()) Then GoTo ExportAVIFError
    
    'Generate a composited image copy, with alpha automatically un-premultiplied
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, False
    
    'Parse all relevant AVIF parameters.  (See the AVIF export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    'Retrieve target AVIF quality
    Dim avifQuality As Long
    avifQuality = cParams.GetLong("avif-quality", 0)    '0=lossless
    
    'PD's AVIF interface requires us to first save a PNG file; the external AVIF engine
    ' will then convert this to an AVIF file.
    Dim cPNG As pdPNG
    Set cPNG = New pdPNG
    
    'For performance reasons, write an uncompressed PNG
    Const PNG_COMPRESS As Long = 0
    
    Dim imgSavedOK As Boolean
    imgSavedOK = False
    
    'Generate a temporary filename for the intermediary PNG file.
    Dim tmpFilename As String
    tmpFilename = OS.UniqueTempFilename(customExtension:="png")
    
    'PD now uses its own custom-built PNG encoder.  This encoder is capable of much better compression
    ' and format coverage than either FreeImage or GDI+.  Use it to dump a lossless copy of the current image
    ' to file.
    If (Not imgSavedOK) Then
        PDDebug.LogAction "Using internal PNG encoder for this operation..."
        imgSavedOK = (cPNG.SavePNG_ToFile(tmpFilename, tmpImageCopy, srcPDImage, png_AutoColorType, 0, PNG_COMPRESS, vbNullString, True) < png_Failure)
    
        'If other mechanisms failed, attempt a failsafe export using GDI+.  (This should never trigger, but is
        ' a holdover from when PD's PNG encoder was in its infancy and reliability was not yet real-world-confirmed.)
        If (Not imgSavedOK) Then imgSavedOK = GDIPlusSavePicture(srcPDImage, tmpFilename, P2_FFE_PNG, 32)
        
    End If
    
    'We now have a temporary PNG file saved.  Shell avifenc with the proper parameters to generate a
    ' valid AVIF (at the requested filename).
    If imgSavedOK Then ExportAVIF = Plugin_AVIF.ConvertStandardImageToAVIF(tmpFilename, dstFile, avifQuality)
    
    'With the AVIF generated, we can now erase our temporary PNG file
    Files.FileDeleteIfExists tmpFilename
    
    Exit Function
    
ExportAVIFError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportAVIF = False
    
End Function

Public Function ExportBMP(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportBMPError
    
    ExportBMP = False
    Dim sFileType As String: sFileType = "BMP"
    
    'Parse all relevant BMP parameters.  (See the BMP export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    Dim bmpCompression As Boolean, bmpForceGrayscale As Boolean, bmp16bpp_555Mode As Boolean, bmpCustomColors As Long
    bmpCompression = cParams.GetBool("bmp-rle", False)
    bmpForceGrayscale = cParams.GetBool("bmp-force-gray", False)
    bmp16bpp_555Mode = cParams.GetBool("bmp-16bpp-555", False)
    bmpCustomColors = cParams.GetLong("bmp-indexed-color-count", 256)
    
    Dim bmpBackgroundColor As Long, bmpFlipRowOrder As Boolean
    bmpBackgroundColor = cParams.GetLong("bmp-backcolor", vbWhite)
    bmpFlipRowOrder = cParams.GetBool("bmp-flip-vertical", False)
    
    'Generate a composited image copy, with alpha automatically un-premultiplied
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, False
    
    'Retrieve the recommended output color depth of the image.
    Dim outputColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, desiredAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
    
    If Strings.StringsEqual(cParams.GetString("bmp-color-depth", "auto"), "auto", True) Then
        outputColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_BMP, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
        ExportDebugMsg "Color depth auto-detection returned " & CStr(outputColorDepth) & "bpp"
        
        'Because BMP files only support alpha in 32-bpp mode, we can ignore binary-alpha mode completely
        If (currentAlphaStatus = PDAS_NoAlpha) Then desiredAlphaStatus = PDAS_NoAlpha Else desiredAlphaStatus = PDAS_ComplicatedAlpha
        
    Else
        outputColorDepth = cParams.GetLong("bmp-color-depth", 32)
        If (outputColorDepth = 32) Then desiredAlphaStatus = PDAS_ComplicatedAlpha
    End If
    
    'BMP files support a number of custom alpha parameters, for legacy compatibility reasons.  These need to be applied manually.
    If (outputColorDepth = 32) Then
        If cParams.GetBool("bmp-use-xrgb", False) Then
            tmpImageCopy.ForceNewAlpha 0
        Else
            If cParams.GetBool("bmp-use-pargb", False) Then tmpImageCopy.SetAlphaPremultiplication True
        End If
    
    'Because bitmaps do not support transparency < 32-bpp, remove transparency immediately if the output depth is < 32-bpp,
    ' and forgo any further alpha handling.
    Else
        tmpImageCopy.ConvertTo24bpp bmpBackgroundColor
        desiredAlphaStatus = PDAS_NoAlpha
    End If
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
    ' the original file remains untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
        
    'If both GDI+ and FreeImage are missing, use our own internal methods to save the BMP file in its current state.
    ' (This is a measure of last resort, as the saved image is unlikely to match the requested output depth.)
    If (Not Drawing2D.IsRenderingEngineActive(P2_GDIPlusBackend)) And (Not ImageFormats.IsFreeImageEnabled) Then
        tmpImageCopy.WriteToBitmapFile tmpFilename
        ExportBMP = True
    Else
    
        If ImageFormats.IsFreeImageEnabled Then
            
            Dim fi_DIB As Long
            fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, desiredAlphaStatus, currentAlphaStatus, , bmpBackgroundColor, isGrayscale Or bmpForceGrayscale, bmpCustomColors, Not bmp16bpp_555Mode)
            If bmpFlipRowOrder Then Outside_FreeImageV3.FreeImage_FlipVertically fi_DIB
            
            'Finally, prepare some BMP save flags.  If the user has requested RLE encoding, and this image is <= 8bpp,
            ' request RLE encoding from FreeImage.
            Dim BMPflags As Long: BMPflags = BMP_DEFAULT
            If (outputColorDepth = 8) And bmpCompression Then BMPflags = BMP_SAVE_RLE
            
            'BMP supports DPI information, so append that immediately prior to saving
            FreeImage_SetResolutionX fi_DIB, srcPDImage.GetDPI
            FreeImage_SetResolutionY fi_DIB, srcPDImage.GetDPI
            
            'Use that handle to save the image to BMP format, with required color conversion based on the outgoing color depth
            If (fi_DIB <> 0) Then
                ExportBMP = FreeImage_SaveEx(fi_DIB, tmpFilename, PDIF_BMP, BMPflags, outputColorDepth, True)
                If ExportBMP Then
                    ExportDebugMsg "Export to " & sFileType & " appears successful."
                Else
                    PDDebug.LogAction "WARNING: FreeImage_SaveEx silent fail"
                    Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
                End If
            Else
                PDDebug.LogAction "WARNING: FreeImage returned blank handle"
                Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
                ExportBMP = False
            End If
            
        Else
            ExportBMP = GDIPlusSavePicture(srcPDImage, tmpFilename, P2_FFE_BMP, outputColorDepth)
        End If
    
    End If
    
    'If the original file already existed, attempt to replace it now
    If ExportBMP And Strings.StringsNotEqual(dstFile, tmpFilename) Then
        ExportBMP = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
        If (Not ExportBMP) Then
            Files.FileDelete tmpFilename
            PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
        End If
    End If
    
    Exit Function
    
ExportBMPError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportBMP = False
    
End Function

'GIF is a unique outlier because the format is so complicated (particularly animated GIF),
' so I've moved it into its own module to avoid cluttering up this one.  Check there for details.
Public Function ExportGIF(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    ExportGIF = ImageFormats_GIF.ExportGIF_LL(srcPDImage, dstFile, formatParams, metadataParams)
End Function

Public Function ExportGIF_Animated(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    ExportGIF_Animated = ImageFormats_GIF.ExportGIF_Animated_LL(srcPDImage, dstFile, formatParams, metadataParams)
End Function

'Save to JP2 format using the FreeImage library
Public Function ExportJP2(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportJP2Error
    
    ExportJP2 = False
    Dim sFileType As String: sFileType = "JP2"
    
    If ImageFormats.IsFreeImageEnabled Then
    
        'Parse incoming JP2 parameters
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.SetParamString formatParams
        
        'The only output parameter JP2 supports is compression level
        Dim jp2Quality As Long
        jp2Quality = cParams.GetLong("jp2-quality", 1)
        
        'Generate a composited image copy, with alpha automatically un-premultiplied
        Dim tmpImageCopy As pdDIB
        Set tmpImageCopy = New pdDIB
        srcPDImage.GetCompositedImage tmpImageCopy, False
        
        'Retrieve the recommended output color depth of the image.
        ' (TODO: parse incoming params and honor requests for forced color-depths!)
        Dim outputColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, desiredAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
        outputColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_JP2, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
        ExportDebugMsg "Color depth auto-detection returned " & CStr(outputColorDepth) & "bpp"
        
        'Our JP2 exporter is a simplified one, so ignore special alpha modes
        If (currentAlphaStatus = PDAS_NoAlpha) Then
            desiredAlphaStatus = PDAS_NoAlpha
        Else
            desiredAlphaStatus = PDAS_ComplicatedAlpha
            outputColorDepth = 32
        End If
        
        'To save us some time, auto-convert any non-transparent images to 24-bpp now
        If (desiredAlphaStatus = PDAS_NoAlpha) Then tmpImageCopy.ConvertTo24bpp
        
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, desiredAlphaStatus, currentAlphaStatus)
        
        If (fi_DIB <> 0) Then
            
            Dim fi_Flags As Long: fi_Flags = 0&
            fi_Flags = fi_Flags Or jp2Quality
            
            'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
            ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
            ' the original file remains untouched).
            Dim tmpFilename As String
            If Files.FileExists(dstFile) Then
                Do
                    tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
                Loop While Files.FileExists(tmpFilename)
            Else
                tmpFilename = dstFile
            End If
            
            ExportJP2 = FreeImage_Save(FIF_JP2, fi_DIB, tmpFilename, fi_Flags)
            If ExportJP2 Then
            
                ExportDebugMsg "Export to " & sFileType & " appears successful."
                
                'If the original file already existed, attempt to replace it now
                If Strings.StringsNotEqual(dstFile, tmpFilename) Then
                    If (Files.FileReplace(dstFile, tmpFilename) <> FPR_SUCCESS) Then
                        Files.FileDelete tmpFilename
                        PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
                    End If
                End If
                
            Else
                PDDebug.LogAction "WARNING: FreeImage_Save silent fail"
                Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            PDDebug.LogAction "WARNING: FreeImage returned blank handle"
            Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            ExportJP2 = False
        End If
    Else
        RaiseFreeImageWarning
        ExportJP2 = False
    End If
    
    Exit Function
    
ExportJP2Error:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportJP2 = False
    
End Function

Public Function ExportJPEG(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportJPEGError
    
    ExportJPEG = False
    Dim sFileType As String: sFileType = "JPEG"
    
    'Parse all relevant JPEG parameters.  (See the JPEG export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    'Some JPEG information (like embedding a thumbnail) is handled by the metadata parameter string.
    Dim cParamsMetadata As pdSerialize
    Set cParamsMetadata = New pdSerialize
    cParamsMetadata.SetParamString metadataParams
    
    Dim jpegQuality As Long
    jpegQuality = cParams.GetLong("jpg-quality", 92)
    
    Dim jpegCompression As Long
    Const JPG_CMP_BASELINE = 0, JPG_CMP_OPTIMIZED = 1, JPG_CMP_PROGRESSIVE = 2
    Select Case cParams.GetLong("jpg-compression-mode", JPG_CMP_OPTIMIZED)
        Case JPG_CMP_BASELINE
            jpegCompression = JPEG_BASELINE
            
        Case JPG_CMP_OPTIMIZED
            jpegCompression = JPEG_OPTIMIZE
            
        Case JPG_CMP_PROGRESSIVE
            jpegCompression = JPEG_OPTIMIZE Or JPEG_PROGRESSIVE
        
    End Select
    
    Dim jpegSubsampling As Long
    Const JPG_SS_444 = 0, JPG_SS_422 = 1, JPG_SS_420 = 2, JPG_SS_411 = 3
    Select Case cParams.GetLong("jpg-subsampling", JPG_SS_422)
        Case JPG_SS_444
            jpegSubsampling = JPEG_SUBSAMPLING_444
        Case JPG_SS_422
            jpegSubsampling = JPEG_SUBSAMPLING_422
        Case JPG_SS_420
            jpegSubsampling = JPEG_SUBSAMPLING_420
        Case JPG_SS_411
            jpegSubsampling = JPEG_SUBSAMPLING_411
    End Select
    
    'Combine all FreeImage-specific flags into one merged flag
    Dim jpegFlags As Long
    jpegFlags = jpegQuality Or jpegCompression Or jpegSubsampling
    
    'Generate a composited image copy, with alpha premultiplied (as we're just going to composite it, anyway)
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, True
    
    'JPEGs do not support alpha, so forcibly flatten the image (regardless of output color depth).
    ' We also apply a custom backcolor here (if one exists; white is used by default).
    Dim jpegBackgroundColor As Long
    jpegBackgroundColor = cParams.GetLong("jpg-backcolor", vbWhite)
    If (tmpImageCopy.GetDIBColorDepth = 32) Then tmpImageCopy.ConvertTo24bpp jpegBackgroundColor
    
    'Retrieve the recommended output color depth of the image.
    Dim outputColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
    Dim forceGrayscale As Boolean
    
    If StrComp(LCase$(cParams.GetString("jpg-color-depth", "auto")), "auto", vbBinaryCompare) = 0 Then
        outputColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_JPEG, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
        ExportDebugMsg "Color depth auto-detection returned " & CStr(outputColorDepth) & "bpp"
    Else
        outputColorDepth = cParams.GetLong("jpg-color-depth", 24)
        If outputColorDepth = 8 Then forceGrayscale = True
    End If
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
    ' the original file remains untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    'FreeImage is our preferred export engine
    If ImageFormats.IsFreeImageEnabled Then
        
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, PDAS_NoAlpha, PDAS_NoAlpha, , vbWhite, isGrayscale Or forceGrayscale)
        
        'Use that handle to save the image to JPEG format, with required color conversion based on the outgoing color depth
        If (fi_DIB <> 0) Then
            
            'Next, we need to see if thumbnail embedding is enabled.  If it is, we need to write out a tiny copy
            ' of the main image, which ExifTool will use to generate a thumbnail metadata entry
            If cParams.GetBool("MetadataExportAllowed", True) And cParamsMetadata.GetBool("MetadataEmbedThumbnail", False) Then
                
                Dim fThumbnail As Long, tmpThumbFile As String
                fThumbnail = FreeImage_MakeThumbnail(fi_DIB, 100)
                tmpThumbFile = cParamsMetadata.GetString("MetadataTempFilename")
                
                If (LenB(tmpThumbFile) <> 0) Then
                    Files.FileDeleteIfExists tmpThumbFile
                    FreeImage_SaveEx fThumbnail, tmpThumbFile, FIF_JPEG, FISO_JPEG_BASELINE Or FISO_JPEG_QUALITYNORMAL, FICD_24BPP
                End If
                
                FreeImage_Unload fThumbnail
                
            End If

            'Immediately prior to saving, pass this image's resolution values (if any) to FreeImage.
            ' These values will be embedded in the JFIF header.
            FreeImage_SetResolutionX fi_DIB, srcPDImage.GetDPI
            FreeImage_SetResolutionY fi_DIB, srcPDImage.GetDPI
            
            ExportJPEG = FreeImage_SaveEx(fi_DIB, tmpFilename, PDIF_JPEG, jpegFlags, outputColorDepth, True)
            If ExportJPEG Then
                ExportDebugMsg "Export to " & sFileType & " appears successful."
            Else
                PDDebug.LogAction "WARNING: FreeImage_SaveEx silent fail"
                Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            PDDebug.LogAction "WARNING: FreeImage returned blank handle"
            Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            ExportJPEG = False
        End If
    
    'If FreeImage is unavailable, fall back to GDI+
    Else
        ExportJPEG = GDIPlusSavePicture(srcPDImage, tmpFilename, P2_FFE_JPEG, outputColorDepth, jpegQuality)
    End If
    
    'If the original file already existed, attempt to replace it now
    If ExportJPEG And Strings.StringsNotEqual(dstFile, tmpFilename) Then
        ExportJPEG = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
        If (Not ExportJPEG) Then
            Files.FileDelete tmpFilename
            PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
        End If
    End If
        
    Exit Function
    
ExportJPEGError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportJPEG = False
    
End Function

'Save to JPEG XL format using the official libjxl library
Public Function ExportJXL(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportJXLError
    
    ExportJXL = False
    Dim sFileType As String: sFileType = "JXL"
    
    'If this system is post-XP but libjxl doesn't exist, ask if we can download a copy
    If OS.IsVistaOrLater And (Not Plugin_jxl.IsJXLExportAvailable()) Then
        
        'If the user doesn't allow download, we can't export
        If (Not Plugin_jxl.PromptForLibraryDownload_JXL()) Then GoTo ExportJXLError
        
        'Downloading libjxl will raise new messages in the status bar; restore the original
        ' "saving %1 image" text
        Message "Saving %1 file...", sFileType
        
    End If
    
    'Failsafe check before proceeding
    If (Not Plugin_jxl.IsJXLExportAvailable()) Then GoTo ExportJXLError
    
    'JXL exporting leans on libjxl
    If Plugin_jxl.IsJXLExportAvailable() Then
        
        'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
        ' and if it's saved successfully, overwrite the original file *then* - this way, if an error occurs
        ' mid-save, the original file is left untouched).
        Dim tmpFilename As String
        If Files.FileExists(dstFile) Then
            Do
                tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
            Loop While Files.FileExists(tmpFilename)
        Else
            tmpFilename = dstFile
        End If
        
        'Use libjxl to save (note: PNG is currently used as an intermediary format)
        If Plugin_jxl.SaveJXL_ToFile(srcPDImage, formatParams, tmpFilename) Then
        
            If Strings.StringsEqual(dstFile, tmpFilename) Then
                ExportJXL = True
            
            'If we wrote our data to a temp file, attempt to replace the original file
            Else
            
                ExportJXL = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
                
                If (Not ExportJXL) Then
                    Files.FileDelete tmpFilename
                    PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
                End If
                
            End If
        
        Else
            ExportJXL = False
            ExportDebugMsg "WARNING!  ExportJXL() failed for reasons unknown; check the debug log for additional details"
        End If
        
        Exit Function
    
    Else
        PDDebug.LogAction "libjxl missing or broken; JXL export abandoned"
        ExportJXL = False
    End If
    
    Exit Function
    
ExportJXLError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportJXL = False
    
End Function

'libjxl also provides animated JXL support (using APNG as an intermediary wrapper)
Public Function ExportJXL_Animated(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportJXLError
    
    ExportJXL_Animated = False
    Dim sFileType As String: sFileType = "JXL"
    
    'If this system is post-XP but libjxl doesn't exist, ask if we can download a copy
    If OS.IsVistaOrLater And (Not Plugin_jxl.IsJXLExportAvailable()) Then
        
        'If the user doesn't allow download, we can't export
        If (Not Plugin_jxl.PromptForLibraryDownload_JXL()) Then GoTo ExportJXLError
        
        'Downloading libjxl will raise new messages in the status bar; restore the original
        ' "saving %1 image" text
        Message "Saving %1 file...", sFileType
        
    End If
    
    'Failsafe check before proceeding
    If (Not Plugin_jxl.IsJXLExportAvailable()) Then GoTo ExportJXLError
    
    'JXL exporting leans on libjxl
    If Plugin_jxl.IsJXLExportAvailable() Then
        
        'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
        ' and if it's saved successfully, overwrite the original file *then* - this way, if an error occurs
        ' mid-save, the original file is left untouched).
        Dim tmpFilename As String
        If Files.FileExists(dstFile) Then
            Do
                tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
            Loop While Files.FileExists(tmpFilename)
        Else
            tmpFilename = dstFile
        End If
        
        'Use libjxl to save (note: PNG is currently used as an intermediary format)
        If Plugin_jxl.SaveJXL_ToFile_Animated(srcPDImage, formatParams, tmpFilename) Then
        
            If Strings.StringsEqual(dstFile, tmpFilename) Then
                ExportJXL_Animated = True
            
            'If we wrote our data to a temp file, attempt to replace the original file
            Else
            
                ExportJXL_Animated = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
                
                If (Not ExportJXL_Animated) Then
                    Files.FileDelete tmpFilename
                    PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
                End If
                
            End If
        
        Else
            ExportJXL_Animated = False
            ExportDebugMsg "WARNING!  ExportJXL() failed for reasons unknown; check the debug log for additional details"
        End If
        
        Exit Function
    
    Else
        PDDebug.LogAction "libjxl missing or broken; JXL export abandoned"
        ExportJXL_Animated = False
    End If
    
    Exit Function
    
ExportJXLError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportJXL_Animated = False
    
End Function

'Save to JXR format using the FreeImage library
Public Function ExportJXR(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportJXRError
    
    ExportJXR = False
    Dim sFileType As String: sFileType = "JXR"
    
    If ImageFormats.IsFreeImageEnabled Then
    
        'Parse incoming JXR parameters
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.SetParamString formatParams
        
        'The only output parameter JXR supports is compression level
        Dim jxrQuality As Long, jxrProgressive As Boolean
        jxrQuality = cParams.GetLong("jxr-quality", 1)
        jxrProgressive = cParams.GetBool("jxr-progressive", False)
        
        'Generate a composited image copy, with alpha automatically un-premultiplied
        Dim tmpImageCopy As pdDIB
        Set tmpImageCopy = New pdDIB
        srcPDImage.GetCompositedImage tmpImageCopy, False
        
        'Retrieve the recommended output color depth of the image.
        ' (TODO: parse incoming params and honor requests for forced color-depths!)
        Dim outputColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, desiredAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
        outputColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_JXR, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
        ExportDebugMsg "Color depth auto-detection returned " & CStr(outputColorDepth) & "bpp"
        
        'Our JXR exporter is a simplified one, so ignore special alpha modes
        If (currentAlphaStatus = PDAS_NoAlpha) Then
            desiredAlphaStatus = PDAS_NoAlpha
        Else
            desiredAlphaStatus = PDAS_ComplicatedAlpha
            outputColorDepth = 32
        End If
        
        'To save us some time, auto-convert any non-transparent images to 24-bpp now
        If (desiredAlphaStatus = PDAS_NoAlpha) Then tmpImageCopy.ConvertTo24bpp
        
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, desiredAlphaStatus, currentAlphaStatus)
        
        If (fi_DIB <> 0) Then
            
            Dim fi_Flags As Long: fi_Flags = 0&
            fi_Flags = fi_Flags Or jxrQuality
            If jxrProgressive Then fi_Flags = fi_Flags Or JXR_PROGRESSIVE
            
            'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
            ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
            ' the original file remains untouched).
            Dim tmpFilename As String
            If Files.FileExists(dstFile) Then
                Do
                    tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
                Loop While Files.FileExists(tmpFilename)
            Else
                tmpFilename = dstFile
            End If
            
            ExportJXR = FreeImage_Save(FIF_JXR, fi_DIB, tmpFilename, fi_Flags)
            
            If ExportJXR Then
            
                ExportDebugMsg "Export to " & sFileType & " appears successful."
                
                'If the original file already existed, attempt to replace it now
                If Strings.StringsNotEqual(dstFile, tmpFilename) Then
                    If (Files.FileReplace(dstFile, tmpFilename) <> FPR_SUCCESS) Then
                        Files.FileDelete tmpFilename
                        PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
                    End If
                End If
                
            Else
                PDDebug.LogAction "WARNING: FreeImage_Save silent fail"
                Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            PDDebug.LogAction "WARNING: FreeImage returned blank handle"
            Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            ExportJXR = False
        End If
    Else
        RaiseFreeImageWarning
        ExportJXR = False
    End If
    
    Exit Function
    
ExportJXRError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportJXR = False
    
End Function

'Save an HDR (High-Dynamic Range) image
Public Function ExportHDR(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportHDRError
    
    ExportHDR = False
    Dim sFileType As String: sFileType = "HDR"
    
    If ImageFormats.IsFreeImageEnabled Then
        
        'TODO: parse incoming HDR parameters.  (FreeImage doesn't support any HDR export parameters
        ' at present, but we could still provide options for things like gamma correction,
        ' background color for 32-bpp images, etc.)
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.SetParamString formatParams
        
        'Generate a composited image copy, with alpha automatically un-premultiplied
        Dim tmpImageCopy As pdDIB
        Set tmpImageCopy = New pdDIB
        srcPDImage.GetCompositedImage tmpImageCopy
        
        'HDR does not support alpha-channels, so convert to 24-bpp in advance
        If (tmpImageCopy.GetDIBColorDepth = 32) Then tmpImageCopy.ConvertTo24bpp
        
        'HDR only supports one output color depth, so auto-detection is unnecessary
        ExportDebugMsg "HDR format only supports one output depth, so color depth auto-detection was ignored."
            
        'Convert our current DIB to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(tmpImageCopy.GetDIBDC)
        Set tmpImageCopy = Nothing
        
        If (fi_DIB <> 0) Then
            
            'Convert the image data to RGBF format
            Dim fi_FloatDIB As Long
            fi_FloatDIB = FreeImage_ConvertToRGBF(fi_DIB)
            FreeImage_Unload fi_DIB
            
            If (fi_FloatDIB <> 0) Then
                
                'Prior to saving, we must account for default 2.2 gamma correction.
                ' We do this by iterating through the source, and modifying gamma values as we go.
                ' (If we reduce gamma prior to RGBF conversion, quality will obviously be impacted due to clipping.)
                Dim srcImageData() As Single, srcSA As SafeArray1D
                
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
                    VBHacks.WrapArrayAroundPtr_Float srcImageData, srcSA, FreeImage_GetScanline(fi_FloatDIB, y), iScanWidth * 4
                    
                    'Iterate through this line, converting values as we go
                    For x = 0 To iLoopWidth
                        
                        'Retrieve the source values
                        srcF = srcImageData(x)
                        
                        'Apply 1/2.2 gamma correction
                        If (srcF > 0!) Then srcImageData(x) = srcF ^ gammaCorrection
                        
                    Next x
                    
                    'Safely unalias the VB array object
                    VBHacks.UnwrapArrayFromPtr_Float srcImageData
                    
                Next y
                
                'With gamma properly accounted for, we can finally write the image out to file.
                
                'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
                ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
                ' the original file remains untouched).
                Dim tmpFilename As String
                If Files.FileExists(dstFile) Then
                    Do
                        tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
                    Loop While Files.FileExists(tmpFilename)
                Else
                    tmpFilename = dstFile
                End If
                
                ExportHDR = FreeImage_Save(PDIF_HDR, fi_FloatDIB, tmpFilename, 0)
                If ExportHDR Then
                    
                    ExportDebugMsg "Export to " & sFileType & " appears successful."
                    
                    'If the original file already existed, attempt to replace it now
                    If Strings.StringsNotEqual(dstFile, tmpFilename) Then
                        If (Files.FileReplace(dstFile, tmpFilename) <> FPR_SUCCESS) Then
                            Files.FileDelete tmpFilename
                            PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
                        End If
                    End If
                    
                Else
                    PDDebug.LogAction "WARNING: FreeImage_Save silent fail"
                    Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
                End If
                
                FreeImage_Unload fi_FloatDIB
                
            Else
                ExportDebugMsg "HDR save failed; could not convert to RGBF"
                ExportHDR = False
            End If
                
        Else
            PDDebug.LogAction "WARNING: FreeImage returned blank handle"
            Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            ExportHDR = False
        End If
        
    Else
        RaiseFreeImageWarning
        ExportHDR = False
    End If
    
    Exit Function
        
ExportHDRError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportHDR = False
    
End Function

'Export a Windows Icon (ICO) file
Public Function ExportICO(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportICOError
    
    ExportICO = False
    Dim sFileType As String: sFileType = "ICO"
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    'TEST: ensure we can retrieve all icon parameters
    Dim numIcons As Long
    numIcons = cParams.GetLong("icon-count", 0, True)
    If (numIcons = 0) Then Exit Function
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
    ' the original file remains untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    'PD uses its own custom-built ICO encoder to create icon files.
    PDDebug.LogAction "Using internal ICO encoder for this operation..."
    
    Dim cICO As pdICO
    Set cICO = New pdICO
    ExportICO = cICO.SaveICO_ToFile(tmpFilename, srcPDImage, formatParams)
    
    'If we wrote the ICO to a temp file, attempt to replace the original file with it now
    If ExportICO And Strings.StringsNotEqual(dstFile, tmpFilename) Then
        
        ExportICO = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
        
        If (Not ExportICO) Then
            Files.FileDelete tmpFilename
            PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
        End If
        
    End If
    
    Exit Function
    
ExportICOError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportICO = False
    
End Function

Public Function ExportORA(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean

    On Error GoTo ExportORAError

    ExportORA = False
    Dim sFileType As String: sFileType = "ORA"
    
    'OpenRaster has a straightforward spec based on a zip file container:
    ' https://www.openraster.org/
    
    'Most of the heavy lifting for the save will be performed by our pdOpenRaster class
    Dim cORA As pdOpenRaster
    Set cORA = New pdOpenRaster
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file then - this way, if an error occurs
    ' mid-save, the original file is left untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    If cORA.SaveORA(srcPDImage, tmpFilename) Then
    
        If Strings.StringsEqual(dstFile, tmpFilename) Then
            ExportORA = True
        
        'If we wrote our data to a temp file, attempt to replace the original file
        Else
        
            ExportORA = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
            
            If (Not ExportORA) Then
                Files.FileDelete tmpFilename
                PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
            End If
            
        End If
    
    Else
        ExportORA = False
        ExportDebugMsg "WARNING!  pdOpenRaster.SaveORA() failed for reasons unknown; check the debug log for additional details"
    End If
    
    Exit Function
    
ExportORAError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportORA = False
    
End Function

Public Function ExportPNG(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportPNGError
    
    ExportPNG = False
    Dim sFileType As String: sFileType = "PNG"
    
    'Generate a composited image copy, with alpha automatically un-premultiplied
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, False
    
    'Parse all relevant PNG parameters.  (See the PNG export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    Dim cParamsDepth As pdSerialize
    Set cParamsDepth = New pdSerialize
    cParamsDepth.SetParamString cParams.GetString("png-color-depth")
    
    Dim cPNG As pdPNG
    Set cPNG = New pdPNG
    
    'The only settings we need to extract here is compression level; everything else is handled automatically
    ' by the PNG export class.
    Dim pngCompressionLevel As Long
    pngCompressionLevel = cParams.GetLong("png-compression-level", 9)
    
    Dim imgSavedOK As Boolean
    imgSavedOK = False
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
    ' the original file remains untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    'PD now uses its own custom-built PNG encoder.  This encoder is capable of much better compression
    ' and format coverage than either FreeImage or GDI+.
    If (Not imgSavedOK) Then
        PDDebug.LogAction "Using internal PNG encoder for this operation..."
        imgSavedOK = (cPNG.SavePNG_ToFile(tmpFilename, tmpImageCopy, srcPDImage, png_AutoColorType, 0, pngCompressionLevel, formatParams, True) < png_Failure)
    End If
    
    'If other mechanisms failed, attempt a failsafe export using GDI+.  (Note that this pathway is *not* preferred,
    ' as GDI+ forcibly writes problematic color data chunks and it performs no adaptive filtering so file sizes
    ' are enormous, but hey - it's better than not writing a PNG at all, right?)
    If (Not imgSavedOK) Then
        PDDebug.LogAction "WARNING: pdPNG failed!"
        imgSavedOK = GDIPlusSavePicture(srcPDImage, tmpFilename, P2_FFE_PNG, 32)
    End If
    
    'If the original file already existed, attempt to replace it now
    If imgSavedOK And Strings.StringsNotEqual(dstFile, tmpFilename) Then
        imgSavedOK = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
        If (Not imgSavedOK) Then
            Files.FileDelete tmpFilename
            PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
        End If
    End If
    
    ExportPNG = imgSavedOK
    
    Exit Function
    
ExportPNGError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportPNG = False
    
End Function

Public Function ExportPNG_Animated(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportPNGError
    
    ExportPNG_Animated = False
    Dim sFileType As String: sFileType = "APNG"
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    'The only settings we need to extract here is compression level; everything else is handled automatically
    ' by the PNG export class.
    Dim pngCompressionLevel As Long
    pngCompressionLevel = cParams.GetLong("compression-level", 9)
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
    ' the original file remains untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    'PD uses its own custom-built PNG encoder to create APNG files.  (Neither FreeImage nor GDI+ support APNGs,
    ' and we use a comprehensive optimization tree that produces much better files than those would anyway!)
    PDDebug.LogAction "Using internal PNG encoder for this operation..."
        
    Dim cPNG As pdPNG
    Set cPNG = New pdPNG
    ExportPNG_Animated = (cPNG.SaveAPNG_ToFile(tmpFilename, srcPDImage, png_AutoColorType, 0, pngCompressionLevel, formatParams) < png_Failure)
    
    'If we wrote the APNG to a temp file, attempt to replace the original file with it now
    If ExportPNG_Animated And Strings.StringsNotEqual(dstFile, tmpFilename) Then
        
        ExportPNG_Animated = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
        
        If (Not ExportPNG_Animated) Then
            Files.FileDelete tmpFilename
            PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
        End If
        
    End If
    
    Exit Function
    
ExportPNGError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportPNG_Animated = False
    
End Function

Public Function ExportPNM(ByRef srcPDImage As pdImage, ByRef dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportPNMError
    
    ExportPNM = False
    Dim sFileType As String: sFileType = "PNM"
    
    'Parse all relevant PNM parameters.  (See the PNM export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    Dim pnmColorModel As String, pnmColorDepth As String
    pnmColorModel = cParams.GetString("pnm-color-model", "auto")
    pnmColorDepth = cParams.GetString("pnm-color-depth", "standard")
    
    Dim pnmForceExtension As Boolean, pnmUseASCII As Boolean
    pnmForceExtension = cParams.GetBool("pnm-change-extension", True)
    pnmUseASCII = cParams.GetBool("pnm-use-ascii", True)
    
    Dim pnmBackColor As Long
    pnmBackColor = cParams.GetLong("pnm-background-color", vbWhite)
    
    'Generate a composited image copy, with alpha premultiplied (as we're just going to composite it, anyway)
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, True
    
    'PNMs do not support alpha, so forcibly flatten the image (regardless of output color depth).
    If (tmpImageCopy.GetDIBColorDepth = 32) Then tmpImageCopy.ConvertTo24bpp pnmBackColor
    
    'If any "auto" parameters are present, calculate their ideal values now
    Dim outputColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
    Dim forceGrayscale As Boolean
    
    If ParamsEqual(pnmColorModel, "auto") Then
        outputColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_PNM, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
        ExportDebugMsg "Color depth auto-detection returned " & CStr(outputColorDepth) & "bpp"
    Else
        If ParamsEqual(pnmColorModel, "color") Then
            outputColorDepth = 24
        ElseIf ParamsEqual(pnmColorModel, "gray") Then
            outputColorDepth = 8
        ElseIf ParamsEqual(pnmColorModel, "monochrome") Then
            outputColorDepth = 1
        Else
            outputColorDepth = 24
        End If
        forceGrayscale = (outputColorDepth = 8)
    End If
    
    'If the user wants us to modify the output file extension to match the selected encoding, apply it now
    If pnmForceExtension Then
    
        Dim newExtension As String
        If ParamsEqual(pnmColorDepth, "float") Then
            newExtension = "pfm"
        Else
            Select Case outputColorDepth
                Case 1
                    newExtension = "pbm"
                
                Case 8
                    newExtension = "pgm"
                
                Case Else
                    newExtension = "ppm"
            End Select
        End If
        
        Dim tmpFilename As String
        tmpFilename = Files.FileGetName(dstFile, True)
        dstFile = Files.FileGetPath(dstFile) & tmpFilename & "." & newExtension
        
    End If
    
    'The caller can request HDR or float color-depths; calculate those now
    Dim finalColorDepth As Long
    If ParamsEqual(pnmColorDepth, "hdr") Then
        finalColorDepth = outputColorDepth * 2
    ElseIf ParamsEqual(pnmColorDepth, "float") Then
        finalColorDepth = outputColorDepth * 4
    Else
        finalColorDepth = outputColorDepth
    End If
    
    'Failsafe check for monochrome images
    If (outputColorDepth = 1) Then finalColorDepth = 1
    
    'FreeImage is required for pixmap writing
    If ImageFormats.IsFreeImageEnabled Then
        
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, finalColorDepth, PDAS_NoAlpha, PDAS_NoAlpha, , pnmBackColor, isGrayscale Or forceGrayscale)
        
        'Use that handle to save the image to PNM format, with required color conversion based on the outgoing color depth
        If (fi_DIB <> 0) Then
            
            'From the input parameters, determine a matching FreeImage output constant
            Dim fif_Final As FREE_IMAGE_FORMAT
            If ParamsEqual(pnmColorDepth, "float") Then
                fif_Final = FIF_PFM
            Else
                If (outputColorDepth = 1) Then
                    'On 25/May/16 I discovered that FreeImage's ASCII encoding is broken for PBM files.
                    ' We now default to binary encoding until the bug is fixed.
                    'If pnmUseASCII Then fif_Final = FIF_PBM Else fif_Final = FIF_PBMRAW
                    fif_Final = FIF_PBMRAW
                    FreeImage_Invert fi_DIB
                ElseIf (outputColorDepth = 8) Then
                    If pnmUseASCII Then fif_Final = FIF_PGM Else fif_Final = FIF_PGMRAW
                Else
                    If pnmUseASCII Then fif_Final = FIF_PPM Else fif_Final = FIF_PPMRAW
                End If
            End If
            
            Dim fi_Flags As FREE_IMAGE_SAVE_OPTIONS
            If (fif_Final = FIF_PBM) Or (fif_Final = FIF_PGM) Or (fif_Final = FIF_PPM) Then
                fi_Flags = FISO_PNM_SAVE_ASCII
            Else
                fi_Flags = FISO_PNM_SAVE_RAW
            End If
            
            'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
            ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
            ' the original file remains untouched).
            If Files.FileExists(dstFile) Then
                Do
                    tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
                Loop While Files.FileExists(tmpFilename)
            Else
                tmpFilename = dstFile
            End If
            
            ExportPNM = FreeImage_Save(fif_Final, fi_DIB, tmpFilename, fi_Flags)
            If ExportPNM Then
                
                ExportDebugMsg "Export to " & sFileType & " appears successful."
                
                'If the original file already existed, attempt to replace it now
                If Strings.StringsNotEqual(dstFile, tmpFilename) Then
                    ExportPNM = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
                    If (Not ExportPNM) Then
                        Files.FileDelete tmpFilename
                        PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
                    End If
                End If
        
            Else
                PDDebug.LogAction "WARNING: FreeImage_Save silent fail"
                Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            PDDebug.LogAction "WARNING: FreeImage returned blank handle"
            Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            ExportPNM = False
        End If
        
    Else
        ExportPNM = False
        PDDebug.LogAction "No PNM encoder found. Save aborted."
    End If
    
    Exit Function
    
ExportPNMError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportPNM = False
    
End Function

'Save to PSD (or PSB) format using our own internal PSD encoder
Public Function ExportPSD(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportPSDError

    ExportPSD = False
    Dim sFileType As String: sFileType = "PSD"
    
    'Parse all relevant PSD parameters.  (See the PSD export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    Dim useMaxCompatibility As Boolean
    useMaxCompatibility = cParams.GetBool("max-compatibility", True)
    
    'Compression defaults to 1 - PackBits (RLE), same as Photoshop
    Dim compressionType As Long
    compressionType = cParams.GetLong("compression", 1)
    
    'Most of the heavy lifting for the save will be performed by our pdPSD class
    Dim cPSD As pdPSD
    Set cPSD = New pdPSD
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file then - this way, if an error occurs
    ' mid-save, the original file is left untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    If cPSD.SavePSD(srcPDImage, tmpFilename, useMaxCompatibility, compressionType, False) Then
    
        If Strings.StringsEqual(dstFile, tmpFilename) Then
            ExportPSD = True
        
        'If we wrote our data to a temp file, attempt to replace the original file
        Else
        
            ExportPSD = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
            
            If (Not ExportPSD) Then
                Files.FileDelete tmpFilename
                PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
            End If
            
        End If
    
    Else
        ExportPSD = False
        ExportDebugMsg "WARNING!  pdPSD.SavePSD() failed for reasons unknown; check the debug log for additional details"
    End If
    
    Exit Function
    
ExportPSDError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportPSD = False
    
End Function

'Save to PSP (Paintshop Pro) format using PD's internal PSP encoder
Public Function ExportPSP(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportPSPError

    ExportPSP = False
    Dim sFileType As String: sFileType = "PSP"
    
    'Parse all relevant PSP parameters.  (See the PSP export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    'Figure out which PSP version to target.  Format specs are only publicly available through
    ' PSP 8, so target versions larger than this are unsupported.  (Similarly, the PSP format
    ' was invented for PSP 5 but v5 files are vastly different from later versions, so PD only
    ' attempts to support v6 at the earliest.)
    Dim strPSPVersion As String, targetPSPVersion As Long
    strPSPVersion = cParams.GetString("compatibility-target", "auto", True)
    If Strings.StringsEqual(strPSPVersion, "auto", True) Then
        targetPSPVersion = 8
    Else
        If TextSupport.IsNumberLocaleUnaware(strPSPVersion) Then
            targetPSPVersion = strPSPVersion
        Else
            targetPSPVersion = 8
        End If
    End If
    
    If (targetPSPVersion > 8) Then
        targetPSPVersion = 8
    ElseIf (targetPSPVersion < 6) Then
        targetPSPVersion = 6
    End If
    
    'PSP files use zLib compression.  Figure out which compression level to use.
    Dim cmpLevel As Long
    cmpLevel = cParams.GetLong("compression-level", 9, True)
    If (cmpLevel > Compression.GetMaxCompressionLevel(cf_Zlib)) Then
        cmpLevel = Compression.GetMaxCompressionLevel(cf_Zlib)
    ElseIf (cmpLevel < 0) Then
        cmpLevel = Compression.GetDefaultCompressionLevel(cf_Zlib)
    End If
    
    'Most of the heavy lifting for the save will be performed by the pdPSP class
    Dim cPSP As pdPSP
    Set cPSP = New pdPSP
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file then - this way, if an error occurs
    ' mid-save, the original file is left untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    If cPSP.SavePSP(srcPDImage, tmpFilename, targetPSPVersion, True, cmpLevel, True) Then

        If Strings.StringsEqual(dstFile, tmpFilename) Then
            ExportPSP = True

        'If we wrote our data to a temp file, attempt to replace the original file
        Else

            ExportPSP = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)

            If (Not ExportPSP) Then
                Files.FileDelete tmpFilename
                PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
            End If

        End If

    Else
        ExportPSP = False
        ExportDebugMsg "WARNING!  pdPSP.SavePSP() failed for reasons unknown; check the debug log for additional details"
    End If
    
    Exit Function
    
ExportPSPError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportPSP = False
    
End Function

'QOI uses native codecs (implemented in the pdQOI class)
Public Function ExportQOI(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportQOIError
    
    ExportQOI = False
    Dim sFileType As String: sFileType = "QOI"
    
    'Generate a composited image copy, with alpha automatically un-premultiplied
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, False
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file then - this way, if an error occurs
    ' mid-save, the original file is left untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    'All export details are handled by a pdQOI instance
    Dim cQOI As pdQOI
    Set cQOI = New pdQOI
    If cQOI.SaveQOI_ToFile(tmpFilename, srcPDImage, tmpImageCopy) Then
        
        'If we are not overwriting an existing file, exit immediately; otherwise, attempt to replace the original file
        If Strings.StringsEqual(dstFile, tmpFilename) Then
            ExportQOI = True
        Else
            ExportQOI = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
            If (Not ExportQOI) Then
                Files.FileDeleteIfExists tmpFilename
                PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
            End If
        End If
        
    End If
    
    Exit Function
    
ExportQOIError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportQOI = False
    
End Function

'Save to TGA format using the FreeImage library
Public Function ExportTGA(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportTGAError
    
    ExportTGA = False
    Dim sFileType As String: sFileType = "TGA"
    
    If ImageFormats.IsFreeImageEnabled Then
    
        'TODO: parse incoming TGA parameters.  (This requires a TGA export dialog, which I haven't constructed yet...)
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.SetParamString formatParams
        
        'The only output parameter TGA supports is whether to enable basic RLE compression
        Dim compressRLE As Boolean
        compressRLE = True
        
        'Generate a composited image copy, with alpha automatically un-premultiplied
        Dim tmpImageCopy As pdDIB
        Set tmpImageCopy = New pdDIB
        srcPDImage.GetCompositedImage tmpImageCopy, False
        
        'Retrieve the recommended output color depth of the image.
        ' (TODO: parse incoming params and honor requests for forced color-depths!)
        Dim outputColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, desiredAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
        outputColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_TARGA, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
        ExportDebugMsg "Color depth auto-detection returned " & CStr(outputColorDepth) & "bpp"
        
        'Our TGA exporter is a simplified one, so ignore special alpha modes
        If (currentAlphaStatus = PDAS_NoAlpha) Then
            desiredAlphaStatus = PDAS_NoAlpha
        Else
            If (currentAlphaStatus = PDAS_BinaryAlpha) And (outputColorDepth = 8) Then
                desiredAlphaStatus = PDAS_BinaryAlpha
            Else
                desiredAlphaStatus = PDAS_ComplicatedAlpha
                outputColorDepth = 32
            End If
        End If
        
        'To save us some time, auto-convert any non-transparent images to 24-bpp now
        If (desiredAlphaStatus = PDAS_NoAlpha) Then tmpImageCopy.ConvertTo24bpp
        
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, desiredAlphaStatus, currentAlphaStatus)
        
        If (fi_DIB <> 0) Then
            
            Dim fi_Flags As Long: fi_Flags = 0&
            If compressRLE Then fi_Flags = fi_Flags Or TARGA_SAVE_RLE
            
            'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
            ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
            ' the original file remains untouched).
            Dim tmpFilename As String
            If Files.FileExists(dstFile) Then
                Do
                    tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
                Loop While Files.FileExists(tmpFilename)
            Else
                tmpFilename = dstFile
            End If
            
            ExportTGA = FreeImage_Save(FIF_TARGA, fi_DIB, tmpFilename, fi_Flags)
            If ExportTGA Then
            
                ExportDebugMsg "Export to " & sFileType & " appears successful."
                        
                'If the original file already existed, attempt to replace it now
                If Strings.StringsNotEqual(dstFile, tmpFilename) Then
                    ExportTGA = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
                    If (Not ExportTGA) Then
                        Files.FileDelete tmpFilename
                        PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
                    End If
                End If
                
            Else
                PDDebug.LogAction "WARNING: FreeImage_Save silent fail"
                Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            PDDebug.LogAction "WARNING: FreeImage returned blank handle"
            Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            ExportTGA = False
        End If
    Else
        RaiseFreeImageWarning
        ExportTGA = False
    End If
    
    Exit Function
    
ExportTGAError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportTGA = False
    
End Function

Public Function ExportTIFF(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportTIFFError
    
    ExportTIFF = False
    Dim sFileType As String: sFileType = "TIFF"
    
    'Parse all relevant TIFF parameters.  (See the TIFF export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString formatParams
    
    Dim cParamsDepth As pdSerialize
    Set cParamsDepth = New pdSerialize
    cParamsDepth.SetParamString cParams.GetString("tiff-color-depth")
    
    'First come generic TIFF settings (compression methods, basically)
    Dim tiffCompressionColor As String, tiffCompressionMono As String
    tiffCompressionColor = cParams.GetString("tiff-compression-color", "LZW")
    tiffCompressionMono = cParams.GetString("tiff-compression-mono", "Fax4")
    
    'This value is not currently supplied by the source dialog; I've left it here in case we
    ' decide to make it user-adjustable in the future
    Dim tiffBackgroundColor As Long
    tiffBackgroundColor = cParams.GetLong("tiff-backcolor", vbWhite)
        
    'Next come the various color-depth and alpha modes
    Dim outputColorModel As String
    outputColorModel = cParamsDepth.GetString("cd-color-model", "auto")
    
    'If the output color model is "gray", note that we will apply a forcible grayscale conversion prior to export
    Dim forceGrayscale As Boolean
    forceGrayscale = ParamsEqual(outputColorModel, "gray")
    
    'From the color depth requests, calculate an actual, numeric color depth.
    ' (This includes checks like -- if we are forcibly outputting a grayscale image, set the bit-depth to 8-bpp to match.)
    Dim outputColorDepth As Long, outputPaletteSize As Long, outputColorDepthName As String
    If forceGrayscale Then
    
        outputColorDepthName = cParamsDepth.GetString("cd-gray-depth", "gray-standard")
        
        If ParamsEqual(outputColorDepthName, "gray-hdr") Then
            outputColorDepth = 16
        ElseIf ParamsEqual(outputColorDepthName, "gray-monochrome") Then
            outputColorDepth = 1
        Else
            outputColorDepth = 8
        End If
        
    Else
    
        outputColorDepthName = cParamsDepth.GetString("cd-color-depth", "color-standard")
        
        If ParamsEqual(outputColorDepthName, "color-hdr") Then
            outputColorDepth = 48
        ElseIf ParamsEqual(outputColorDepthName, "color-indexed") Then
            outputColorDepth = 8
        Else
            outputColorDepth = 24
        End If
        
    End If
    
    outputPaletteSize = cParamsDepth.GetLong("cd-palette-size", 256)
    
    'PD supports multiple alpha output modes; some of these modes (like "binary" alpha, which consists of only 0 or 255 values),
    ' require additional settings.  We always retrieve all values, even if we don't plan on using them.
    Dim outputAlphaModel As String
    outputAlphaModel = cParamsDepth.GetString("cd-alpha-model", "auto")
    
    Dim outputTiffCutoff As Long, outputTiffColor As Long
    outputTiffCutoff = cParams.GetLong("cd-alpha-cutoff", PD_DEFAULT_ALPHA_CUTOFF)
    outputTiffColor = cParams.GetLong("cd-alpha-color", vbMagenta)
    
    'If "automatic" mode is selected for either color space or transparency, we need to determine appropriate
    ' color-depth and alpha-detection values now.
    Dim autoColorModeActive As Boolean, autoTransparencyModeActive As Boolean
    autoColorModeActive = ParamsEqual(outputColorModel, "auto")
    autoTransparencyModeActive = ParamsEqual(outputAlphaModel, "auto")
    
    Dim autoColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, desiredAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
    
    'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
    ' and if it's saved successfully, overwrite the original file - this way, if an error occurs mid-save,
    ' the original file remains untouched).
    Dim tmpFilename As String
    If Files.FileExists(dstFile) Then
        Do
            tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
        Loop While Files.FileExists(tmpFilename)
    Else
        tmpFilename = dstFile
    End If
    
    Dim TIFFflags As Long: TIFFflags = TIFF_DEFAULT
    
    'Next comes the multipage settings, which is crucial as we have to use a totally different codepath for multipage images
    Dim writeMultipage As Boolean
    writeMultipage = cParams.GetBool("tiff-multipage", False)
    
    'Multipage TIFFs use their own custom path (this is due to the way the FreeImage API works; it's convoluted!)
    If writeMultipage And ImageFormats.IsFreeImageEnabled And (srcPDImage.GetNumOfVisibleLayers > 1) Then
        
        'Multipage files use a fairly simple format:
        ' 1) Iterate through each visible layer
        ' 2) Convert each layer to a null-padded layer at the size of the current image
        ' 3) Create a FreeImage copy of the null-padded layer
        ' 4) Insert that layer into a running FreeImage Multipage object
        ' 5) When all layers are finished, write the TIFF out to file
        
        'Start by creating a blank multipage object
        Dim fi_MultipageHandle As Long
        fi_MultipageHandle = FreeImage_OpenMultiBitmap(PDIF_TIFF, tmpFilename, True)
        
        'If all pages are monochrome, we can encode the final TIFF object using monochrome compression settings, but if even
        ' one page is color, it complicates that.
        Dim allPagesMonochrome As Boolean: allPagesMonochrome = True
        
        Dim fi_PageHandle As Long
        Dim tmpLayerDIB As pdDIB, tmpLayer As pdLayer
        Dim pageColorDepth As Long, pageForceGrayscale As Boolean
        
        Dim i As Long
        For i = 0 To srcPDImage.GetNumOfLayers - 1
            
            If srcPDImage.GetLayerByIndex(i).GetLayerVisibility Then
                
                'Clone the current layer
                If (tmpLayer Is Nothing) Then Set tmpLayer = New pdLayer
                tmpLayer.CopyExistingLayer srcPDImage.GetLayerByIndex(i)
                
                'Rasterize as necessary
                If (Not tmpLayer.IsLayerRaster) Then tmpLayer.RasterizeVectorData
                
                'Convert the layer to a flat, null-padded layer at the same size as the original image
                tmpLayer.ConvertToNullPaddedLayer srcPDImage.Width, srcPDImage.Height, True
                
                'Un-premultiply alpha, if any
                tmpLayer.GetLayerDIB.SetAlphaPremultiplication False
                
                'Point a DIB wrapper at the fully processed layer
                Set tmpLayerDIB = tmpLayer.GetLayerDIB
                
                If autoColorModeActive Or autoTransparencyModeActive Then
                    autoColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpLayerDIB, PDIF_TIFF, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
                    ExportDebugMsg "Color depth auto-detection returned " & CStr(autoColorDepth) & "bpp"
                    If (currentAlphaStatus = PDAS_BinaryAlpha) Then currentAlphaStatus = PDAS_ComplicatedAlpha
                Else
                    currentAlphaStatus = PDAS_ComplicatedAlpha
                End If
                
                'From the automatic values, construct matching output values
                If autoColorModeActive Then
                    pageColorDepth = autoColorDepth
                    pageForceGrayscale = isGrayscale
                    If (Not isTrueColor) Then outputPaletteSize = netColorCount
                Else
                    pageColorDepth = outputColorDepth
                    pageForceGrayscale = forceGrayscale
                End If
        
                'Convert the auto-detected transparency mode to a usable string parameter.  (We need this later in the function,
                ' so we can combine color depth and alpha depth into a single usable bit-depth.)
                If autoTransparencyModeActive Then
                    desiredAlphaStatus = currentAlphaStatus
                    If desiredAlphaStatus = PDAS_NoAlpha Then
                        outputAlphaModel = "none"
                    ElseIf desiredAlphaStatus = PDAS_BinaryAlpha Then
                        outputAlphaModel = "by-cutoff"
                    ElseIf desiredAlphaStatus = PDAS_NewAlphaFromColor Then
                        outputAlphaModel = "by-color"
                    ElseIf desiredAlphaStatus = PDAS_ComplicatedAlpha Then
                        outputAlphaModel = "full"
                    Else
                        outputAlphaModel = "full"
                    End If
                End If
        
                'Use the current transparency mode (whether auto-created or manually requested) to construct a new output
                ' depth that correctly represents the combination of color depth + alpha depth.  Note that this also requires
                ' us to workaround some FreeImage deficiencies, so these depths may not match what TIFF formally supports.
                If ParamsEqual(outputAlphaModel, "full") Then
                
                    desiredAlphaStatus = PDAS_ComplicatedAlpha
                    
                    'PNG supports 8-bpp grayscale + 8-bpp alpha as a valid channel combination.  Unfortunately, FreeImage has
                    ' no way of generating such an image.  We must fall back to 32-bpp mode.
                    If (Not forceGrayscale) Then
                        If (pageColorDepth = 24) Then pageColorDepth = 32
                        If (pageColorDepth = 48) Then pageColorDepth = 64
                    End If
                    
                ElseIf ParamsEqual(outputAlphaModel, "none") Then
                    desiredAlphaStatus = PDAS_NoAlpha
                    If (Not pageForceGrayscale) Then
                        If (pageColorDepth = 64) Then pageColorDepth = 48
                        If (pageColorDepth = 32) Then pageColorDepth = 24
                    End If
                    outputTiffCutoff = 0
            
                ElseIf ParamsEqual(outputAlphaModel, "by-cutoff") Then
                    desiredAlphaStatus = PDAS_BinaryAlpha
                    If (Not pageForceGrayscale) Then
                        If (pageColorDepth = 24) Then pageColorDepth = 32
                        If (pageColorDepth = 48) Then pageColorDepth = 64
                    End If
                    
                ElseIf ParamsEqual(outputAlphaModel, "by-color") Then
                    desiredAlphaStatus = PDAS_NewAlphaFromColor
                    outputTiffCutoff = outputTiffColor
                    If (Not pageForceGrayscale) Then
                        If (pageColorDepth = 24) Then pageColorDepth = 32
                        If (pageColorDepth = 48) Then pageColorDepth = 64
                    End If
                End If
                    
                'Monochrome depths require special treatment if alpha is active
                If (pageColorDepth = 1) And (desiredAlphaStatus <> PDAS_NoAlpha) Then
                    pageColorDepth = 8
                    outputPaletteSize = 2
                End If
                
                If (pageColorDepth <> 1) Then allPagesMonochrome = False
                
                'We now have enough information to create a FreeImage copy of this DIB
                fi_PageHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpLayerDIB, pageColorDepth, desiredAlphaStatus, currentAlphaStatus, outputTiffCutoff, tiffBackgroundColor, pageForceGrayscale, outputPaletteSize, , (desiredAlphaStatus <> PDAS_NoAlpha))
                
                If (fi_PageHandle <> 0) Then
                
                    'Insert this page at the *end* of the current multipage file, then free our copy of it
                    FreeImage_AppendPage fi_MultipageHandle, fi_PageHandle
                    Plugin_FreeImage.ReleaseFreeImageObject fi_PageHandle
                    
                Else
                    PDDebug.LogAction "WARNING!  PD was unable to create a FreeImage handle for layer # " & i
                End If
                
            'End "is layer visible?"
            End If
            
        Next i
        
        'With all pages inserted, we can now write the multipage TIFF out to file
        If allPagesMonochrome Then
            TIFFflags = TIFFflags Or GetFreeImageTIFFConstant(tiffCompressionMono)
        Else
            TIFFflags = TIFFflags Or GetFreeImageTIFFConstant(tiffCompressionColor)
        End If
        
        ExportTIFF = FreeImage_CloseMultiBitmap(fi_MultipageHandle, TIFFflags)
        If ExportTIFF Then
            ExportDebugMsg "Export to " & sFileType & " appears successful."
        Else
            PDDebug.LogAction "WARNING: FreeImage mystery fail"
            Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
        End If
        
        'FreeImage unloads the multipage bitmap automatically when it is closed; this is different from single-page bitmaps,
        ' which must be manually unloaded.
        
    'Single-page TIFFs are simpler to write
    Else
        
        'Generate a composited image copy, with alpha automatically un-premultiplied
        Dim tmpImageCopy As pdDIB
        Set tmpImageCopy = New pdDIB
        srcPDImage.GetCompositedImage tmpImageCopy, False
        
        If autoColorModeActive Or autoTransparencyModeActive Then
            autoColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_TIFF, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
            ExportDebugMsg "Color depth auto-detection returned " & CStr(autoColorDepth) & "bpp"
        Else
            currentAlphaStatus = PDAS_ComplicatedAlpha
        End If
        
        'From the automatic values, construct matching output values
        If autoColorModeActive Then
            outputColorDepth = autoColorDepth
            forceGrayscale = isGrayscale
            If (Not isTrueColor) Then outputPaletteSize = netColorCount
        End If
        
        'Convert the auto-detected transparency mode to a usable string parameter.  (We need this later in the function,
        ' so we can combine color depth and alpha depth into a single usable bit-depth.)
        If autoTransparencyModeActive Then
            desiredAlphaStatus = currentAlphaStatus
            If desiredAlphaStatus = PDAS_NoAlpha Then
                outputAlphaModel = "none"
            ElseIf desiredAlphaStatus = PDAS_BinaryAlpha Then
                outputAlphaModel = "by-cutoff"
            ElseIf desiredAlphaStatus = PDAS_NewAlphaFromColor Then
                outputAlphaModel = "by-color"
            ElseIf desiredAlphaStatus = PDAS_ComplicatedAlpha Then
                outputAlphaModel = "full"
            Else
                outputAlphaModel = "full"
            End If
        End If
        
        'Use the current transparency mode (whether auto-created or manually requested) to construct a new output
        ' depth that correctly represents the combination of color depth + alpha depth.  Note that this also requires
        ' us to workaround some FreeImage deficiencies, so these depths may not match what TIFF formally supports.
        If ParamsEqual(outputAlphaModel, "full") Then
        
            desiredAlphaStatus = PDAS_ComplicatedAlpha
            
            If (Not forceGrayscale) Then
                If (outputColorDepth = 24) Then outputColorDepth = 32
                If (outputColorDepth = 48) Then outputColorDepth = 64
            End If
            
        ElseIf ParamsEqual(outputAlphaModel, "none") Then
            desiredAlphaStatus = PDAS_NoAlpha
            If (Not forceGrayscale) Then
                If (outputColorDepth = 64) Then outputColorDepth = 48
                If (outputColorDepth = 32) Then outputColorDepth = 24
            End If
            outputTiffCutoff = 0
            
        ElseIf ParamsEqual(outputAlphaModel, "by-cutoff") Then
            desiredAlphaStatus = PDAS_BinaryAlpha
            If (Not forceGrayscale) Then
                If (outputColorDepth = 24) Then outputColorDepth = 32
                If (outputColorDepth = 48) Then outputColorDepth = 64
            End If
            
        ElseIf ParamsEqual(outputAlphaModel, "by-color") Then
            desiredAlphaStatus = PDAS_NewAlphaFromColor
            outputTiffCutoff = outputTiffColor
            If (Not forceGrayscale) Then
                If (outputColorDepth = 24) Then outputColorDepth = 32
                If (outputColorDepth = 48) Then outputColorDepth = 64
            End If
        End If
            
        'Monochrome depths require special treatment if alpha is active
        If (outputColorDepth = 1) And (desiredAlphaStatus <> PDAS_NoAlpha) Then
            outputColorDepth = 8
            outputPaletteSize = 2
        End If
        
        'The TIFF export engine supports both FreeImage and GDI+.  Note that many, *many* features are disabled under GDI+,
        ' so the FreeImage path is absolutely preferred.
        If ImageFormats.IsFreeImageEnabled Then
            
            Dim fi_DIB As Long
            fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, desiredAlphaStatus, currentAlphaStatus, outputTiffCutoff, tiffBackgroundColor, forceGrayscale, outputPaletteSize, , (desiredAlphaStatus <> PDAS_NoAlpha))
            
            'Finally, prepare some TIFF save flags.  If the user has requested RLE encoding, and this image is <= 8bpp,
            ' request RLE encoding from FreeImage.
            If (outputColorDepth = 1) Then
                TIFFflags = TIFFflags Or GetFreeImageTIFFConstant(tiffCompressionMono)
            Else
                TIFFflags = TIFFflags Or GetFreeImageTIFFConstant(tiffCompressionColor)
            End If
                    
            'Use that handle to save the image to TIFF format, with required color conversion based on the outgoing color depth
            If (fi_DIB <> 0) Then
                ExportTIFF = FreeImage_Save(PDIF_TIFF, fi_DIB, tmpFilename, TIFFflags)
                FreeImage_Unload fi_DIB
                If ExportTIFF Then
                    ExportDebugMsg "Export to " & sFileType & " appears successful."
                Else
                    PDDebug.LogAction "WARNING: FreeImage_Save silent fail"
                    Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
                End If
            Else
                PDDebug.LogAction "WARNING: FreeImage returned blank handle"
                Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
                ExportTIFF = False
            End If
            
        Else
            ExportTIFF = GDIPlusSavePicture(srcPDImage, tmpFilename, P2_FFE_TIFF, outputColorDepth)
        End If
        
    End If
    
    'If the original file already existed, attempt to replace it now
    If ExportTIFF And Strings.StringsNotEqual(dstFile, tmpFilename) Then
        ExportTIFF = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
        If (Not ExportTIFF) Then
            Files.FileDelete tmpFilename
            PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
        End If
    End If
    
    Exit Function
    
ExportTIFFError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportTIFF = False
    
End Function

Private Function GetFreeImageTIFFConstant(ByRef compressionName As String) As Long
    If ParamsEqual(compressionName, "LZW") Then
        GetFreeImageTIFFConstant = TIFF_LZW
    ElseIf ParamsEqual(compressionName, "ZIP") Then
        GetFreeImageTIFFConstant = TIFF_ADOBE_DEFLATE
    ElseIf ParamsEqual(compressionName, "Fax4") Then
        GetFreeImageTIFFConstant = TIFF_CCITTFAX4
    ElseIf ParamsEqual(compressionName, "Fax3") Then
        GetFreeImageTIFFConstant = TIFF_CCITTFAX3
    ElseIf ParamsEqual(compressionName, "none") Then
        GetFreeImageTIFFConstant = TIFF_NONE
    End If
End Function

'Save to WebP format using Google's official libwebp library
Public Function ExportWebP(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportWebPError
    
    ExportWebP = False
    Dim sFileType As String: sFileType = "WebP"
    
    'WebP exporting leans on libwebp via pdWebP
    If Plugin_WebP.IsWebPEnabled() Then
        
        'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
        ' and if it's saved successfully, overwrite the original file *then* - this way, if an error occurs
        ' mid-save, the original file is left untouched).
        Dim tmpFilename As String
        If Files.FileExists(dstFile) Then
            Do
                tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
            Loop While Files.FileExists(tmpFilename)
        Else
            tmpFilename = dstFile
        End If
        
        'Use pdWebP to save the WebP file
        Dim cWebP As pdWebP
        Set cWebP = New pdWebP
        If cWebP.SaveWebP_ToFile(srcPDImage, formatParams, tmpFilename) Then
        
            If Strings.StringsEqual(dstFile, tmpFilename) Then
                ExportWebP = True
            
            'If we wrote our data to a temp file, attempt to replace the original file
            Else
            
                ExportWebP = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
                
                If (Not ExportWebP) Then
                    Files.FileDelete tmpFilename
                    PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
                End If
                
            End If
        
        Else
            ExportWebP = False
            ExportDebugMsg "WARNING!  pdWebP.SaveWebP_ToFile() failed for reasons unknown; check the debug log for additional details"
        End If
        
        Exit Function
    
    End If
    
    'If we're still here, libwebp is missing or broken.  We can still attempt to save via FreeImage (if available)
    ' but many WebP features will no longer work (and the following code is not actively maintained, so "you get what you get")
    PDDebug.LogAction "libwebp missing or broken"
    
    If ImageFormats.IsFreeImageEnabled Then
    
        'Parse incoming WebP parameters
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.SetParamString formatParams
        
        'The only output parameter WebP supports is compression level
        Dim webPQuality As Long
        webPQuality = cParams.GetLong("webp-quality", 100)
        
        'Generate a composited image copy, with alpha automatically un-premultiplied
        Dim tmpImageCopy As pdDIB
        Set tmpImageCopy = New pdDIB
        srcPDImage.GetCompositedImage tmpImageCopy, False
        
        'Retrieve the recommended output color depth of the image.
        ' (TODO: parse incoming params and honor requests for forced color-depths!)
        Dim outputColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, desiredAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
        outputColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_WEBP, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
        ExportDebugMsg "Color depth auto-detection returned " & CStr(outputColorDepth) & "bpp"
        
        'WebP only supports 24-bpp and 32-bpp outputs, so check for transparency now
        If (currentAlphaStatus = PDAS_NoAlpha) Then
            desiredAlphaStatus = PDAS_NoAlpha
            outputColorDepth = 24
        Else
            desiredAlphaStatus = PDAS_ComplicatedAlpha
            outputColorDepth = 32
        End If
        
        'To save us some time, auto-convert any non-transparent images to 24-bpp now
        If (desiredAlphaStatus = PDAS_NoAlpha) Then tmpImageCopy.ConvertTo24bpp
        
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, desiredAlphaStatus, currentAlphaStatus)
        
        If (fi_DIB <> 0) Then
            
            Dim fi_Flags As Long: fi_Flags = 0&
            fi_Flags = fi_Flags Or webPQuality
            
            ExportWebP = FreeImage_Save(FIF_WEBP, fi_DIB, dstFile, fi_Flags)
            If ExportWebP Then
                ExportDebugMsg "Export to " & sFileType & " appears successful."
            Else
                PDDebug.LogAction "WARNING: FreeImage_Save silent fail"
                Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            PDDebug.LogAction "WARNING: FreeImage returned blank handle"
            Message "%1 save failed. Please report this error using Help -> Submit Bug Report.", sFileType
            ExportWebP = False
        End If
    Else
        RaiseFreeImageWarning
        ExportWebP = False
    End If
    
    Exit Function
    
ExportWebPError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportWebP = False
    
End Function

'Same as "Save to WebP format using Google's official libwebp library", above, but for animated WebP targets
Public Function ExportWebP_Animated(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportWebPError
    
    ExportWebP_Animated = False
    Dim sFileType As String: sFileType = "WebP"
    
    'Exporting animated WebP files requires libwebp availability
    If Plugin_WebP.IsWebPEnabled() Then
        
        'If the target file already exists, use "safe" file saving (e.g. write the save data to a new file,
        ' and if it's saved successfully, overwrite the original file *then* - this way, if an error occurs
        ' mid-save, the original file is left untouched).
        Dim tmpFilename As String
        If Files.FileExists(dstFile) Then
            Do
                tmpFilename = dstFile & Hex$(PDMath.GetCompletelyRandomInt()) & ".pdtmp"
            Loop While Files.FileExists(tmpFilename)
        Else
            tmpFilename = dstFile
        End If
        
        'Use pdWebP to save the WebP file
        Dim cWebP As pdWebP
        Set cWebP = New pdWebP
        If cWebP.SaveAnimatedWebP_ToFile(srcPDImage, formatParams, tmpFilename) Then
        
            If Strings.StringsEqual(dstFile, tmpFilename) Then
                ExportWebP_Animated = True
            
            'If we wrote our data to a temp file, attempt to replace the original file
            Else
                ExportWebP_Animated = (Files.FileReplace(dstFile, tmpFilename) = FPR_SUCCESS)
                If (Not ExportWebP_Animated) Then
                    Files.FileDelete tmpFilename
                    PDDebug.LogAction "WARNING!  Safe save did not overwrite original file (is it open elsewhere?)"
                End If
            End If
        
        Else
            ExportWebP_Animated = False
            ExportDebugMsg "WARNING!  pdWebP.SaveWebPAnimation_ToFile() failed for reasons unknown; check the debug log for additional details"
        End If
        
    End If
    
    Exit Function
    
ExportWebPError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportWebP_Animated = False
    
End Function

'Many export functions require FreeImage.  If it doesn't exist, a generic warning will be raised when the user tries to
' export to a FreeImage-based format.  (Note that the warning is suppressed during batch processing, by design.)
Private Sub RaiseFreeImageWarning()
    If (Macros.GetMacroStatus <> MacroBATCH) Then PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbCritical Or vbOKOnly, "Error"
    Message "Save cannot be completed without FreeImage library."
End Sub

'Basic case-insensitive string comparison function
Private Function ParamsEqual(ByRef param1 As String, ByRef param2 As String) As Boolean
    ParamsEqual = Strings.StringsEqual(param1, param2, True)
End Function
