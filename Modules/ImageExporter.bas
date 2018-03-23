Attribute VB_Name = "ImageExporter"
'***************************************************************************
'Low-level image export interfaces
'Copyright 2001-2018 by Tanner Helland
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
Public Function AutoDetectOutputColorDepth(ByRef srcDIB As pdDIB, ByRef dstFormat As PD_IMAGE_FORMAT, Optional ByRef currentAlphaStatus As PD_ALPHA_STATUS = PDAS_NoAlpha, Optional ByRef uniqueColorCount As Long = 257, Optional ByRef isTrueColor As Boolean = True, Optional ByRef isGrayscale As Boolean = False, Optional ByRef isMonochrome As Boolean = False) As Long
    
    Dim colorCheckSuccessful As Boolean: colorCheckSuccessful = False
    
    'If the incoming image is already 24-bpp, we can skip the alpha-processing steps entirely.  However, it is not
    ' necessary for the caller to do this.  PD will provide correct results either way.
    If (srcDIB.GetDIBColorDepth = 24) Then
        currentAlphaStatus = PDAS_NoAlpha
        colorCheckSuccessful = AutoDetectColors_24BPPSource(srcDIB, uniqueColorCount, isGrayscale, isMonochrome)
        isTrueColor = (uniqueColorCount > 256)
    
    'If the incoming image is 32-bpp, we will run additional alpha channel heuristics
    Else
        colorCheckSuccessful = AutoDetectColors_32BPPSource(srcDIB, uniqueColorCount, isGrayscale, isMonochrome, currentAlphaStatus)
        isTrueColor = (uniqueColorCount > 256)
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
                
                    'If the image contains meaningful alpha channel data, we have two output options
                    If (currentAlphaStatus <> PDAS_NoAlpha) Then
                        
                        'If the alpha is "complicated" (meaning it contains more values than just 0 or 255), we must fall back to
                        ' 32-bpp output modes, regardless of color status.  (PNG supports more fine-grained results than this,
                        ' but FreeImage does not, so our hands are tied.)
                        If (currentAlphaStatus = PDAS_ComplicatedAlpha) Then
                            AutoDetectOutputColorDepth = 32
                        
                        'If the alpha is *not* complicated - meaning it consists of only 0 or 255 values - we can use
                        ' an 8-bpp output mode, with a designated transparent color.
                        Else
                            AutoDetectOutputColorDepth = 8
                        End If
                    
                    Else
                        If isMonochrome Then
                            AutoDetectOutputColorDepth = 1
                        Else
                            'I'm debating whether to provide 4-bpp as an output depth.  It has limited usage, and there
                            ' are complications with binary alpha... this is marked as TODO for now
                            If (uniqueColorCount <= 16) Then
                                AutoDetectOutputColorDepth = 4
                            Else
                                AutoDetectOutputColorDepth = 8
                            End If
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
        
        pdDebug.LogAction "Analyzing color count of 24-bpp image..."
        
        Dim srcPixels() As Byte, tmpSA As SafeArray2D
        PrepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcPixels()), VarPtr(tmpSA), 4
        
        Dim x As Long, y As Long, finalX As Long, finalY As Long
        finalY = srcDIB.GetDIBHeight - 1
        finalX = srcDIB.GetDIBWidth - 1
        finalX = finalX * 3
        
        Dim uniqueColors() As Long
        ReDim uniqueColors(0 To 255) As Long
        
        Dim i As Long
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

    If (srcDIB.GetDIBColorDepth = 32) Then

        pdDebug.LogAction "Analyzing color count of 32-bpp image..."
        
        Dim srcPixels() As Byte, tmpSA As SafeArray2D
        PrepSafeArray tmpSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcPixels()), VarPtr(tmpSA), 4

        Dim x As Long, y As Long, finalX As Long, finalY As Long
        finalY = srcDIB.GetDIBHeight - 1
        finalX = srcDIB.GetDIBWidth - 1
        finalX = finalX * 4

        Dim uniqueColors() As RGBQuad
        ReDim uniqueColors(0 To 255) As RGBQuad

        Dim i As Long
        For i = 0 To 255
            uniqueColors(i).Red = 1
            uniqueColors(i).Green = 1
            uniqueColors(i).Blue = 0
            uniqueColors(i).Alpha = 1
        Next i

        'Total number of unique colors counted so far
        Dim numUniqueColors As Long, non255Alpha As Boolean, nonBinaryAlpha As Boolean
        numUniqueColors = 0
        non255Alpha = False
        nonBinaryAlpha = False
        
        'Finally, a bunch of variables used in color calculation
        Dim r As Long, g As Long, b As Long, a As Long
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
                    If (uniqueColors(i).Red = r) Then
                        If (uniqueColors(i).Green = g) Then
                            If (uniqueColors(i).Blue = b) Then
                                If (uniqueColors(i).Alpha = a) Then
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
                        uniqueColors(numUniqueColors).Red = r
                        uniqueColors(numUniqueColors).Green = g
                        uniqueColors(numUniqueColors).Blue = b
                        uniqueColors(numUniqueColors).Alpha = a
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

            'End grayscale check
            End If

            'If the image is grayscale and it only contains two colors, check for monochrome next
            ' (where monochrome = pure black and pure white, only).
            If isGrayscale And (numUniqueColors <= 2) Then
                
                If ((uniqueColors(0).Red = 0) And (uniqueColors(0).Green = 0) And (uniqueColors(0).Blue = 0)) Or ((uniqueColors(0).Red = 255) And (uniqueColors(0).Green = 255) And (uniqueColors(0).Blue = 255)) Then
                    If ((uniqueColors(1).Red = 0) And (uniqueColors(1).Green = 0) And (uniqueColors(1).Blue = 0)) Or ((uniqueColors(1).Red = 255) And (uniqueColors(1).Green = 255) And (uniqueColors(1).Blue = 255)) Then isMonochrome = True
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
    pdDebug.LogAction debugMsg
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
        tmpImageCopy.ConvertTo24bpp bmpBackgroundColor
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
            ExportBMP = GDIPlusSavePicture(srcPDImage, dstFile, P2_FFE_BMP, outputColorDepth)
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
    gifForceGrayscale = Strings.StringsEqual(gifColorMode, "gray", True)
    If Strings.StringsEqual(gifColorMode, "auto", True) Then gifColorCount = 256
    
    Dim desiredAlphaStatus As PD_ALPHA_STATUS
    desiredAlphaStatus = PDAS_BinaryAlpha
    If Strings.StringsEqual(gifAlphaMode, "none", True) Then desiredAlphaStatus = PDAS_NoAlpha
    If Strings.StringsEqual(gifAlphaMode, "bycolor", True) Then
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
        ExportGIF = GDIPlusSavePicture(srcPDImage, dstFile, P2_FFE_GIF, 8)
    Else
        ExportGIF = False
        Message "No %1 encoder found. Save aborted.", "JPEG"
    End If
    
    
    Exit Function
    
ExportGIFError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportGIF = False
    
End Function

'Save to JP2 format using the FreeImage library
Public Function ExportJP2(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportJP2Error
    
    ExportJP2 = False
    Dim sFileType As String: sFileType = "JP2"
    
    If g_ImageFormats.FreeImageEnabled Then
    
        'Parse incoming JP2 parameters
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
        cParams.SetParamString formatParams
        
        'The only output parameter JP2 supports is compression level
        Dim jp2Quality As Long
        jp2Quality = cParams.GetLong("JP2Quality", 1)
        
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
            
            ExportJP2 = FreeImage_Save(FIF_JP2, fi_DIB, dstFile, fi_Flags)
            If ExportJP2 Then
                ExportDebugMsg "Export to " & sFileType & " appears successful."
            Else
                Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
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
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString formatParams
    
    'Some JPEG information (like embedding a thumbnail) is handled by the metadata parameter string.
    Dim cParamsMetadata As pdParamXML
    Set cParamsMetadata = New pdParamXML
    cParamsMetadata.SetParamString metadataParams
    
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
    If (tmpImageCopy.GetDIBColorDepth = 32) Then tmpImageCopy.ConvertTo24bpp jpegBackgroundColor
    
    'Retrieve the recommended output color depth of the image.
    Dim outputColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
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
            
            'Next, we need to see if thumbnail embedding is enabled.  If it is, we need to write out a tiny copy
            ' of the main image, which ExifTool will use to generate a thumbnail metadata entry
            If cParams.GetBool("MetadataExportAllowed", True) And cParamsMetadata.GetBool("MetadataEmbedThumbnail", False) Then
                
                Dim fThumbnail As Long, tmpFile As String
                fThumbnail = FreeImage_MakeThumbnail(fi_DIB, 100)
                tmpFile = cParamsMetadata.GetString("MetadataTempFilename")
                
                If (Len(tmpFile) <> 0) Then
                    Files.FileDeleteIfExists tmpFile
                    FreeImage_SaveEx fThumbnail, tmpFile, FIF_JPEG, FISO_JPEG_BASELINE Or FISO_JPEG_QUALITYNORMAL, FICD_24BPP
                End If
                
                FreeImage_Unload fThumbnail
                
            End If

            'Immediately prior to saving, pass this image's resolution values (if any) to FreeImage.
            ' These values will be embedded in the JFIF header.
            FreeImage_SetResolutionX fi_DIB, srcPDImage.GetDPI
            FreeImage_SetResolutionY fi_DIB, srcPDImage.GetDPI
            
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
        ExportJPEG = GDIPlusSavePicture(srcPDImage, dstFile, P2_FFE_JPEG, outputColorDepth, jpegQuality)
    Else
        ExportJPEG = False
        Message "No %1 encoder found. Save aborted.", "JPEG"
    End If
    
    Exit Function
    
ExportJPEGError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportJPEG = False
    
End Function

'Save to JXR format using the FreeImage library
Public Function ExportJXR(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportJXRError
    
    ExportJXR = False
    Dim sFileType As String: sFileType = "JXR"
    
    If g_ImageFormats.FreeImageEnabled Then
    
        'Parse incoming JXR parameters
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
        cParams.SetParamString formatParams
        
        'The only output parameter JXR supports is compression level
        Dim jxrQuality As Long, jxrProgressive As Boolean
        jxrQuality = cParams.GetLong("JXRQuality", 1)
        jxrProgressive = cParams.GetBool("JXRProgressive", False)
        
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
            
            ExportJXR = FreeImage_Save(FIF_JXR, fi_DIB, dstFile, fi_Flags)
            If ExportJXR Then
                ExportDebugMsg "Export to " & sFileType & " appears successful."
            Else
                Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
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
                
                'Prior to saving, we must account for default 2.2 gamma correction.  We do this by iterating through the source, and modifying gamma
                ' values as we go.  (If we reduce gamma prior to RGBF conversion, quality will obviously be impacted due to clipping.)
                
                'This Single-type array will consistently be updated to point to the current line of pixels in the image (RGBF format, remember!)
                Dim srcImageData() As Single
                Dim srcSA As SafeArray1D
                
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
        RaiseFreeImageWarning
        ExportHDR = False
    End If
    
    Exit Function
        
ExportHDRError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportHDR = False
    
End Function

Public Function ExportPNG(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportPNGError
    
    ExportPNG = False
    Dim sFileType As String: sFileType = "PNG"
    
    'Generate a composited image copy, with alpha automatically un-premultiplied
    Dim tmpImageCopy As pdDIB
    Set tmpImageCopy = New pdDIB
    srcPDImage.GetCompositedImage tmpImageCopy, False
    
    Dim fi_DIB As Long
    
    'Parse all relevant PNG parameters.  (See the PNG export dialog for details on how these are generated.)
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString formatParams
    
    Dim cParamsDepth As pdParamXML
    Set cParamsDepth = New pdParamXML
    cParamsDepth.SetParamString cParams.GetString("PNGColorDepth")
    
    Dim useWebOptimizedPath As Boolean
    useWebOptimizedPath = cParams.GetBool("PNGCreateWebOptimized", False)
    
    'Web-optimized PNGs use their own path, and they supply their own special variables
    If useWebOptimizedPath And (g_ImageFormats.pngQuantEnabled Or PluginManager.IsPluginCurrentlyEnabled(CCP_OptiPNG)) Then
    
        Dim pngLossyEnabled As Boolean, pngLossyQuality As Long
        pngLossyEnabled = cParams.GetBool("PNGOptimizeLossy", True)
        pngLossyQuality = cParams.GetLong("PNGOptimizeLossyQuality", 80)
        
        Dim pngLossyPerformance As Long, pngLossyDithering As Boolean
        pngLossyPerformance = cParams.GetLong("PNGOptimizeLossyPerformance", 3)
        pngLossyDithering = cParams.GetBool("PNGOptimizeLossyDithering", True)
        
        Dim pngLosslessPerformance As Long
        pngLosslessPerformance = cParams.GetLong("PNGOptimizeLosslessPerformance")
        
        'Quickly dump out a PNG file; we don't need to spend time here finding optimal outputs, as subsequent
        ' optimization passes will find the most appropriate color depth for us.
        fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, 32, PDAS_ComplicatedAlpha, PDAS_ComplicatedAlpha)
        If FreeImage_Save(FIF_PNG, fi_DIB, dstFile, FISO_PNG_Z_BEST_SPEED) Then
            FreeImage_Unload fi_DIB
            
            'Start with pngquant's lossy optimization, if it's enabled
            If pngLossyEnabled Then
                If Plugin_PNGQuant.ApplyPNGQuantToFile_Synchronous(dstFile, pngLossyQuality, pngLossyPerformance, pngLossyDithering, False) Then
                    ExportDebugMsg "pngquant pass successful!"
                End If
            End If
            
            'We always finish with at least one OptiPNG pass
            If PluginManager.IsPluginCurrentlyEnabled(CCP_OptiPNG) And (pngLosslessPerformance > 0) Then
                Plugin_OptiPNG.ApplyOptiPNGToFile_Synchronous dstFile, pngLosslessPerformance
                ExportDebugMsg "OptiPNG pass successful!"
            End If
            
            ExportPNG = True
            
        Else
            ExportDebugMsg "WARNING!  GDI+ failed to save an initial PNG copy.  Subsequent optimizations were not performed."
            ExportPNG = False
        End If
        
    'As of 7.0, standard-mode PNGs support a ton of user-editable parameters.
    Else
        
        'First come the PNG-specific settings (compression level, chunks, etc)
        Dim pngCompressionLevel As Long, pngInterlacing As Boolean
        pngCompressionLevel = cParams.GetLong("PNGCompressionLevel", 3)
        pngInterlacing = cParams.GetBool("PNGInterlacing", False)
        
        Dim pngBackgroundColor As Long, pngCreateBkgdChunk As Boolean
        pngBackgroundColor = cParams.GetLong("PNGBackgroundColor", vbWhite)
        pngCreateBkgdChunk = cParams.GetBool("PNGCreateBkgdChunk", False)
        
        Dim pngStandardOptimizeLevel As Long
        pngStandardOptimizeLevel = cParams.GetLong("PNGStandardOptimization", 1)
        
        'If we're applying some measure of optimization, reset the PNG compression level (as we're just going to
        ' overwrite it during the optimization step)
        If (pngStandardOptimizeLevel >= 2) Then pngCompressionLevel = 1
        
        'Next come the various color-depth and alpha modes
        Dim outputColorModel As String
        outputColorModel = cParamsDepth.GetString("ColorDepth_ColorModel", "Auto")
        
        'If the output color model is "gray", note that we will apply a forcible grayscale conversion prior to export
        Dim forceGrayscale As Boolean
        forceGrayscale = ParamsEqual(outputColorModel, "gray")
        
        'From the color depth requests, calculate an actual, numeric color depth.
        ' (This includes checks like -- if we are forcibly outputting a grayscale image, set the bit-depth to 8-bpp to match.)
        Dim outputColorDepth As Long, outputPaletteSize As Long, outputColorDepthName As String
        If forceGrayscale Then
        
            outputColorDepthName = cParamsDepth.GetString("ColorDepth_GrayDepth", "Gray_Standard")
            
            If ParamsEqual(outputColorDepthName, "Gray_HDR") Then
                outputColorDepth = 16
            ElseIf ParamsEqual(outputColorDepthName, "Gray_Monochrome") Then
                outputColorDepth = 1
            Else
                outputColorDepth = 8
            End If
            
        Else
        
            outputColorDepthName = cParamsDepth.GetString("ColorDepth_ColorDepth", "Color_Standard")
            
            If ParamsEqual(outputColorDepthName, "Color_HDR") Then
                outputColorDepth = 48
            ElseIf ParamsEqual(outputColorDepthName, "Color_Indexed") Then
                outputColorDepth = 8
            Else
                outputColorDepth = 24
            End If
            
        End If
        
        outputPaletteSize = cParamsDepth.GetLong("ColorDepth_PaletteSize", 256)
        
        'PD supports multiple alpha output modes; some of these modes (like "binary" alpha, which consists of only 0 or 255 values),
        ' require additional settings.  We always retrieve all values, even if we don't plan on using them.
        Dim outputAlphaModel As String
        outputAlphaModel = cParamsDepth.GetString("ColorDepth_AlphaModel", "Auto")
        
        Dim outputPNGCutoff As Long, outputPNGColor As Long
        outputPNGCutoff = cParams.GetLong("ColorDepth_AlphaCutoff", PD_DEFAULT_ALPHA_CUTOFF)
        outputPNGColor = cParams.GetLong("ColorDepth_AlphaColor", vbMagenta)
        
        'If "automatic" mode is selected for either color space or transparency, we need to determine appropriate
        ' color-depth and alpha-detection values now.
        Dim autoColorModeActive As Boolean, autoTransparencyModeActive As Boolean
        autoColorModeActive = ParamsEqual(outputColorModel, "auto")
        autoTransparencyModeActive = ParamsEqual(outputAlphaModel, "auto")
        
        Dim autoColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, desiredAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
        If autoColorModeActive Or autoTransparencyModeActive Then
            autoColorDepth = ImageExporter.AutoDetectOutputColorDepth(tmpImageCopy, PDIF_PNG, currentAlphaStatus, netColorCount, isTrueColor, isGrayscale, isMonochrome)
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
                outputAlphaModel = "bycutoff"
            ElseIf desiredAlphaStatus = PDAS_NewAlphaFromColor Then
                outputAlphaModel = "bycolor"
            ElseIf desiredAlphaStatus = PDAS_ComplicatedAlpha Then
                outputAlphaModel = "full"
            Else
                outputAlphaModel = "full"
            End If
        End If
        
        'Use the current transparency mode (whether auto-created or manually requested) to construct a new output
        ' depth that correctly represents the combination of color depth + alpha depth.  Note that this also requires
        ' us to workaround some FreeImage deficiencies, so these depths may not match what PNG formally supports.
        If ParamsEqual(outputAlphaModel, "full") Then
        
            desiredAlphaStatus = PDAS_ComplicatedAlpha
            
            'PNG supports 8-bpp grayscale + 8-bpp alpha as a valid channel combination.  Unfortunately, FreeImage has
            ' no way of generating such an image.  We must fall back to 32-bpp mode.
            If forceGrayscale Then
                outputColorDepth = 32
                forceGrayscale = False
            Else
                If (outputColorDepth = 24) Then outputColorDepth = 32
                If (outputColorDepth = 48) Then outputColorDepth = 64
            End If
            
        ElseIf ParamsEqual(outputAlphaModel, "none") Then
            desiredAlphaStatus = PDAS_NoAlpha
            If (Not forceGrayscale) Then
                If outputColorDepth = 64 Then outputColorDepth = 48
                If outputColorDepth = 32 Then outputColorDepth = 24
            End If
            outputPNGCutoff = 0
            
        ElseIf ParamsEqual(outputAlphaModel, "bycutoff") Then
            desiredAlphaStatus = PDAS_BinaryAlpha
            If (Not forceGrayscale) Then
                If outputColorDepth = 24 Then outputColorDepth = 32
                If outputColorDepth = 48 Then outputColorDepth = 64
            End If
            
        ElseIf ParamsEqual(outputAlphaModel, "bycolor") Then
            desiredAlphaStatus = PDAS_NewAlphaFromColor
            outputPNGCutoff = outputPNGColor
            If (Not forceGrayscale) Then
                If outputColorDepth = 24 Then outputColorDepth = 32
                If outputColorDepth = 48 Then outputColorDepth = 64
            End If
        End If
            
        'Monochrome depths require special treatment if alpha is active
        If (outputColorDepth = 1) And (desiredAlphaStatus <> PDAS_NoAlpha) Then
            outputColorDepth = 8
            outputPaletteSize = 2
        End If
        
        'The PNG export engine supports both FreeImage and GDI+.  Note that many, *many* features are disabled under GDI+,
        ' so the FreeImage path is definitely preferred!
        If g_ImageFormats.FreeImageEnabled Then
            
            fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, desiredAlphaStatus, currentAlphaStatus, outputPNGCutoff, pngBackgroundColor, forceGrayscale, outputPaletteSize, , (desiredAlphaStatus <> PDAS_NoAlpha))
            
            'FreeImage supports the embedding of a bkgd chunk; this doesn't make a lot of sense in modern image development,
            ' but it is part of the PNG spec, so we provide it as an option
            If pngCreateBkgdChunk Then
                Dim rQuad As RGBQuad
                rQuad.Red = Colors.ExtractRed(pngBackgroundColor)
                rQuad.Green = Colors.ExtractGreen(pngBackgroundColor)
                rQuad.Blue = Colors.ExtractBlue(pngBackgroundColor)
                FreeImage_SetBackgroundColor fi_DIB, rQuad
            End If
            
            'Finally, prepare some PNG save flags.  If the user has requested RLE encoding, and this image is <= 8bpp,
            ' request RLE encoding from FreeImage.
            Dim PNGflags As Long: PNGflags = PNG_DEFAULT
            If pngCompressionLevel = 0 Then PNGflags = PNGflags Or PNG_Z_NO_COMPRESSION Else PNGflags = PNGflags Or pngCompressionLevel
            If pngInterlacing Then PNGflags = PNGflags Or PNG_INTERLACED
                    
            'Use that handle to save the image to PNG format, with required color conversion based on the outgoing color depth
            If (fi_DIB <> 0) Then
                ExportPNG = FreeImage_Save(PDIF_PNG, fi_DIB, dstFile, PNGflags)
                FreeImage_Unload fi_DIB
                If ExportPNG Then
                    ExportDebugMsg "Export to " & sFileType & " appears successful."
                    
                    'There are some color+alpha variants that PNG supports, but FreeImage cannot write.  OptiPNG is capable
                    ' of converting existing PNG images to these more compact formats.  Engage it now.
                    If PluginManager.IsPluginCurrentlyEnabled(CCP_OptiPNG) And (pngStandardOptimizeLevel > 0) Then
                        Plugin_OptiPNG.ApplyOptiPNGToFile_Synchronous dstFile, pngStandardOptimizeLevel
                    End If
                    
                Else
                    Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
                End If
            Else
                Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
                ExportPNG = False
            End If
            
        Else
            ExportPNG = GDIPlusSavePicture(srcPDImage, dstFile, P2_FFE_PNG, outputColorDepth)
        End If
        
    End If
    
    Exit Function
    
ExportPNGError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportPNG = False
    
End Function

Public Function ExportPNM(ByRef srcPDImage As pdImage, ByRef dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportPNMError
    
    ExportPNM = False
    Dim sFileType As String: sFileType = "PNM"
    
    'Parse all relevant PNM parameters.  (See the PNM export dialog for details on how these are generated.)
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString formatParams
    
    Dim pnmColorModel As String, pnmColorDepth As String
    pnmColorModel = cParams.GetString("PNMColorModel", "Auto")
    pnmColorDepth = cParams.GetString("PNMColorDepth", "Standard")
    
    Dim pnmForceExtension As Boolean, pnmUseASCII As Boolean
    pnmForceExtension = cParams.GetBool("PNMChangeExtensionToMatch", True)
    pnmUseASCII = cParams.GetBool("PNMUseASCII", True)
    
    Dim pnmBackColor As Long
    pnmBackColor = cParams.GetLong("PNMBackgroundColor", vbWhite)
    
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
    If g_ImageFormats.FreeImageEnabled Then
        
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
                    'On 25/May/16 I discovered that FreeImage's ASCII encoding is broken for PBM files.  We now default to binary encoding
                    ' until the bug is fixed.
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
            
            ExportPNM = FreeImage_Save(fif_Final, fi_DIB, dstFile, fi_Flags)
            If ExportPNM Then
                ExportDebugMsg "Export to " & sFileType & " appears successful."
            Else
                Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
            ExportPNM = False
        End If
        
    Else
        ExportPNM = False
        Message "No %1 encoder found. Save aborted.", "PNM"
    End If
    
    Exit Function
    
ExportPNMError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportPNM = False
    
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
        If (desiredAlphaStatus = PDAS_NoAlpha) Or (outputColorDepth <= 24) Then tmpImageCopy.ConvertTo24bpp
        
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
        RaiseFreeImageWarning
        ExportPSD = False
    End If
    
    Exit Function
    
ExportPSDError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportPSD = False
    
End Function

'Save to TGA format using the FreeImage library
Public Function ExportTGA(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportTGAError
    
    ExportTGA = False
    Dim sFileType As String: sFileType = "TGA"
    
    If g_ImageFormats.FreeImageEnabled Then
    
        'TODO: parse incoming TGA parameters.  (This requires a TGA export dialog, which I haven't constructed yet...)
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
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
            
            ExportTGA = FreeImage_Save(FIF_TARGA, fi_DIB, dstFile, fi_Flags)
            If ExportTGA Then
                ExportDebugMsg "Export to " & sFileType & " appears successful."
            Else
                Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
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
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString formatParams
    
    Dim cParamsDepth As pdParamXML
    Set cParamsDepth = New pdParamXML
    cParamsDepth.SetParamString cParams.GetString("TIFFColorDepth")
    
    'First come generic TIFF settings (compression methods, basically)
    Dim TIFFCompressionColor As String, TIFFCompressionMono As String
    TIFFCompressionColor = cParams.GetString("TIFFCompressionColor", "LZW")
    TIFFCompressionMono = cParams.GetString("TIFFCompressionMono", "Fax4")
    
    Dim TIFFBackgroundColor As Long
    TIFFBackgroundColor = cParams.GetLong("TIFFBackgroundColor", vbWhite)
        
    'Next come the various color-depth and alpha modes
    Dim outputColorModel As String
    outputColorModel = cParamsDepth.GetString("ColorDepth_ColorModel", "Auto")
    
    'If the output color model is "gray", note that we will apply a forcible grayscale conversion prior to export
    Dim forceGrayscale As Boolean
    forceGrayscale = ParamsEqual(outputColorModel, "gray")
    
    'From the color depth requests, calculate an actual, numeric color depth.
    ' (This includes checks like -- if we are forcibly outputting a grayscale image, set the bit-depth to 8-bpp to match.)
    Dim outputColorDepth As Long, outputPaletteSize As Long, outputColorDepthName As String
    If forceGrayscale Then
    
        outputColorDepthName = cParamsDepth.GetString("ColorDepth_GrayDepth", "Gray_Standard")
        
        If ParamsEqual(outputColorDepthName, "Gray_HDR") Then
            outputColorDepth = 16
        ElseIf ParamsEqual(outputColorDepthName, "Gray_Monochrome") Then
            outputColorDepth = 1
        Else
            outputColorDepth = 8
        End If
        
    Else
    
        outputColorDepthName = cParamsDepth.GetString("ColorDepth_ColorDepth", "Color_Standard")
        
        If ParamsEqual(outputColorDepthName, "Color_HDR") Then
            outputColorDepth = 48
        ElseIf ParamsEqual(outputColorDepthName, "Color_Indexed") Then
            outputColorDepth = 8
        Else
            outputColorDepth = 24
        End If
        
    End If
    
    outputPaletteSize = cParamsDepth.GetLong("ColorDepth_PaletteSize", 256)
    
    'PD supports multiple alpha output modes; some of these modes (like "binary" alpha, which consists of only 0 or 255 values),
    ' require additional settings.  We always retrieve all values, even if we don't plan on using them.
    Dim outputAlphaModel As String
    outputAlphaModel = cParamsDepth.GetString("ColorDepth_AlphaModel", "Auto")
    
    Dim outputTiffCutoff As Long, outputTiffColor As Long
    outputTiffCutoff = cParams.GetLong("ColorDepth_AlphaCutoff", PD_DEFAULT_ALPHA_CUTOFF)
    outputTiffColor = cParams.GetLong("ColorDepth_AlphaColor", vbMagenta)
    
    'If "automatic" mode is selected for either color space or transparency, we need to determine appropriate
    ' color-depth and alpha-detection values now.
    Dim autoColorModeActive As Boolean, autoTransparencyModeActive As Boolean
    autoColorModeActive = ParamsEqual(outputColorModel, "auto")
    autoTransparencyModeActive = ParamsEqual(outputAlphaModel, "auto")
    
    Dim autoColorDepth As Long, currentAlphaStatus As PD_ALPHA_STATUS, desiredAlphaStatus As PD_ALPHA_STATUS, netColorCount As Long, isTrueColor As Boolean, isGrayscale As Boolean, isMonochrome As Boolean
    
    Dim TIFFflags As Long: TIFFflags = TIFF_DEFAULT
    
    'Next comes the multipage settings, which is crucial as we have to use a totally different codepath for multipage images
    Dim writeMultipage As Boolean
    writeMultipage = cParams.GetBool("TIFFMultipage", False)
    
    'Multipage TIFFs use their own custom path (this is due to the way the FreeImage API works; it's convoluted!)
    If writeMultipage And g_ImageFormats.FreeImageEnabled And (srcPDImage.GetNumOfVisibleLayers > 1) Then
        
        'Multipage files use a fairly simple format:
        ' 1) Iterate through each visible layer
        ' 2) Convert each layer to a null-padded layer at the size of the current image
        ' 3) Create a FreeImage copy of the null-padded layer
        ' 4) Insert that layer into a running FreeImage Multipage object
        ' 5) When all layers are finished, write the TIFF out to file
        
        'Start by creating a blank multipage object
        Files.FileDeleteIfExists dstFile
        
        Dim fi_MasterHandle As Long
        fi_MasterHandle = FreeImage_OpenMultiBitmap(PDIF_TIFF, dstFile, True, False, False)
        
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
                
                'Convert the layer to a flat, null-padded layer at the same size as the master image
                tmpLayer.ConvertToNullPaddedLayer srcPDImage.Width, srcPDImage.Height, True
                
                'Un-premultiply alpha, if any
                tmpLayer.layerDIB.SetAlphaPremultiplication False
                
                'Point a DIB wrapper at the fully processed layer
                Set tmpLayerDIB = tmpLayer.layerDIB
                
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
                        outputAlphaModel = "bycutoff"
                    ElseIf desiredAlphaStatus = PDAS_NewAlphaFromColor Then
                        outputAlphaModel = "bycolor"
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
            
                ElseIf ParamsEqual(outputAlphaModel, "bycutoff") Then
                    desiredAlphaStatus = PDAS_BinaryAlpha
                    If (Not pageForceGrayscale) Then
                        If (pageColorDepth = 24) Then pageColorDepth = 32
                        If (pageColorDepth = 48) Then pageColorDepth = 64
                    End If
                    
                ElseIf ParamsEqual(outputAlphaModel, "bycolor") Then
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
                fi_PageHandle = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpLayerDIB, pageColorDepth, desiredAlphaStatus, currentAlphaStatus, outputTiffCutoff, TIFFBackgroundColor, pageForceGrayscale, outputPaletteSize, , (desiredAlphaStatus <> PDAS_NoAlpha))
                
                If (fi_PageHandle <> 0) Then
                
                    'Insert this page at the *end* of the current multipage file, then free our copy of it
                    FreeImage_AppendPage fi_MasterHandle, fi_PageHandle
                    Plugin_FreeImage.ReleaseFreeImageObject fi_PageHandle
                    
                Else
                    pdDebug.LogAction "WARNING!  PD was unable to create a FreeImage handle for layer # " & i
                End If
                
            'End "is layer visible?"
            End If
            
        Next i
        
        'With all pages inserted, we can now write the multipage TIFF out to file
        If allPagesMonochrome Then
            TIFFflags = TIFFflags Or GetFreeImageTIFFConstant(TIFFCompressionMono)
        Else
            TIFFflags = TIFFflags Or GetFreeImageTIFFConstant(TIFFCompressionColor)
        End If
        
        ExportTIFF = FreeImage_CloseMultiBitmap(fi_MasterHandle, TIFFflags)
        If ExportTIFF Then
            ExportDebugMsg "Export to " & sFileType & " appears successful."
        Else
            Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
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
                outputAlphaModel = "bycutoff"
            ElseIf desiredAlphaStatus = PDAS_NewAlphaFromColor Then
                outputAlphaModel = "bycolor"
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
            
        ElseIf ParamsEqual(outputAlphaModel, "bycutoff") Then
            desiredAlphaStatus = PDAS_BinaryAlpha
            If (Not forceGrayscale) Then
                If (outputColorDepth = 24) Then outputColorDepth = 32
                If (outputColorDepth = 48) Then outputColorDepth = 64
            End If
            
        ElseIf ParamsEqual(outputAlphaModel, "bycolor") Then
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
        If g_ImageFormats.FreeImageEnabled Then
            
            Dim fi_DIB As Long
            fi_DIB = Plugin_FreeImage.GetFIDib_SpecificColorMode(tmpImageCopy, outputColorDepth, desiredAlphaStatus, currentAlphaStatus, outputTiffCutoff, TIFFBackgroundColor, forceGrayscale, outputPaletteSize, , (desiredAlphaStatus <> PDAS_NoAlpha))
            
            'Finally, prepare some TIFF save flags.  If the user has requested RLE encoding, and this image is <= 8bpp,
            ' request RLE encoding from FreeImage.
            If (outputColorDepth = 1) Then
                TIFFflags = TIFFflags Or GetFreeImageTIFFConstant(TIFFCompressionMono)
            Else
                TIFFflags = TIFFflags Or GetFreeImageTIFFConstant(TIFFCompressionColor)
            End If
                    
            'Use that handle to save the image to TIFF format, with required color conversion based on the outgoing color depth
            If (fi_DIB <> 0) Then
                ExportTIFF = FreeImage_Save(PDIF_TIFF, fi_DIB, dstFile, TIFFflags)
                FreeImage_Unload fi_DIB
                If ExportTIFF Then
                    ExportDebugMsg "Export to " & sFileType & " appears successful."
                Else
                    Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
                End If
            Else
                Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
                ExportTIFF = False
            End If
            
        Else
            ExportTIFF = GDIPlusSavePicture(srcPDImage, dstFile, P2_FFE_TIFF, outputColorDepth)
        End If
        
    End If
    
    Exit Function
    
ExportTIFFError:
    ExportDebugMsg "Internal VB error encountered in " & sFileType & " routine.  Err #" & Err.Number & ", " & Err.Description
    ExportTIFF = False
    
End Function

Private Function GetFreeImageTIFFConstant(ByVal compressionName As String) As Long
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

'Save to WebP format using the FreeImage library
Public Function ExportWebP(ByRef srcPDImage As pdImage, ByVal dstFile As String, Optional ByVal formatParams As String = vbNullString, Optional ByVal metadataParams As String = vbNullString) As Boolean
    
    On Error GoTo ExportWebPError
    
    ExportWebP = False
    Dim sFileType As String: sFileType = "WebP"
    
    If g_ImageFormats.FreeImageEnabled Then
    
        'Parse incoming WebP parameters
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
        cParams.SetParamString formatParams
        
        'The only output parameter WebP supports is compression level
        Dim webPQuality As Long
        webPQuality = cParams.GetLong("WebPQuality", 100)
        
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
                Message "%1 save failed (FreeImage_SaveEx silent fail). Please report this error using Help -> Submit Bug Report.", sFileType
            End If
            
        Else
            Message "%1 save failed (FreeImage returned blank handle). Please report this error using Help -> Submit Bug Report.", sFileType
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

'Many export functions require FreeImage.  If it doesn't exist, a generic warning will be raised when the user tries to
' export to a FreeImage-based format.  (Note that the warning is suppressed during batch processing, by design.)
Private Sub RaiseFreeImageWarning()
    If (Macros.GetMacroStatus <> MacroBATCH) Then PDMsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload the program.", vbCritical Or vbOKOnly, "FreeImage Interface Error"
    Message "Save cannot be completed without FreeImage library."
End Sub

'Basic case-insensitive string comparison function
Private Function ParamsEqual(ByVal param1 As String, ByVal param2 As String) As Boolean
    ParamsEqual = Strings.StringsEqual(param1, param2, True)
End Function
