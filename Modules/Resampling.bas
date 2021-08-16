Attribute VB_Name = "Resampling"
'***************************************************************************
'Image Resampling engine
'Copyright 2021-2021 by Tanner Helland
'Created: 16/August/21
'Last updated: 16/August/21
'Last update: instead of relying on 3rd-party code for resampling, let's write our own!
'
'This module is currently a WIP.
'
'Resampling algorithms in this article include heavily modified versions of code originally written by Libor Tinka.
' Libor shared his original C# implementation under a Code Project Open License (CPOL):
'  https://www.codeproject.com/info/cpol10.aspx
' His original, unmodified source code is available here (link good as of Aug 2021):
'  https://www.codeproject.com/Articles/11143/Image-Resizing-outperform-GDI
' Many thanks to Libor for his original example of universal image resampling in C#.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

'Currently available resamplers
Public Enum PD_ResamplingFilter
    rf_Box = 0
    rf_Triangle
    rf_Hermite
    rf_Bell
    rf_CubicBSpline
    rf_Lanczos3
    rf_Mitchell
    rf_Cosine
    rf_CatmullRom
    rf_Quadratic
    rf_QuadraticBSpline
    rf_CubicConvolution
    rf_Lanczos8
End Enum

#If False Then
    Private Const rf_Box = 0, rf_Triangle = 1, rf_Hermite = 2, rf_Bell = 3, rf_CubicBSpline = 4, rf_Lanczos3 = 5, rf_Mitchell = 6
    Private Const rf_Cosine = 7, rf_CatmullRom = 8, rf_Quadratic = 9, rf_QuadraticBSpline = 10, rf_CubicConvolution = 11, rf_Lanczos8 = 12
#End If

'Weight calculation
Private Type Contributor
    pixel As Long
    weight As Double
End Type
    
Private Type ContributorEntry
    n As Long
    p() As Contributor
    wsum As Double
End Type

'Our current resampling approach uses an intermediate copy of the image; this allows us to handle x and y
' resampling independently (which improves performance and greatly simplifies the code, at some trade-off to
' memory consumption).  This intermediate array will be reused on subsequent calls, and can also be manually
' freed when bulk resizing completes.
Private m_tmpPixels() As Byte
Private m_tmpPixelWidth As Long, m_tmpPixelHeight As Long

'Resample an image using the supplied algorith.  A few notes...
' 1) Resampling requires an intermediary image copy to store a copy of resampled data.  This allows you to resize in
'    two dimensions simultaneously (actually this will take two passes, but that's invisible to the caller).
' 2) When two-dimensional resampling is required, the x-dimension will be resampled first.
' 3) 32-bpp inputs are required.  All channels will be resampled using identical code and weights.
Public Function ResampleImage(ByRef dstDIB As pdDIB, ByRef srcDIB As pdDIB, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal rsFilter As PD_ResamplingFilter) As Boolean
    
    ResampleImage = False
    
    Const FUNC_NAME As String = "ResampleImage"
    
    'Validate all inputs
    If (srcDIB Is Nothing) Then
        InternalError FUNC_NAME, "null source"
        Exit Function
    End If
    
    If (dstWidth <= 0) Or (dstHeight <= 0) Then
        InternalError FUNC_NAME, "bad width/height: " & dstWidth & ", " & dstHeight
        Exit Function
    End If
    
    'Inputs look good.  Prepare intermediary data structs.  Custom types are used to improve memory locality.
    Dim srcWidth As Long: srcWidth = srcDIB.GetDIBWidth
    Dim srcHeight As Long: srcHeight = srcDIB.GetDIBHeight
    Dim totalSize As Long: totalSize = srcDIB.GetDIBWidth * srcDIB.GetDIBHeight
    
    'Allocate the intermediary "working" copy of the image width dimensions [dstWidth, srcHeight]
    If (dstWidth > m_tmpPixelWidth) Or (srcHeight > m_tmpPixelHeight) Then
        ReDim m_tmpPixels(0 To (dstWidth * 4) - 1, 0 To srcHeight - 1)
    Else
        VBHacks.ZeroMemory VarPtr(m_tmpPixels(0, 0)), m_tmpPixelWidth * m_tmpPixelHeight * 4
    End If
    
    'Calculate x/y scales; this provides a simple mechanism for checking up- vs downsampling
    Dim xScale As Double, yScale As Double
    xScale = dstWidth / srcWidth
    yScale = dstHeight / srcHeight
    
    'Prep array of contributors (which contain per-pixel weights)
    Dim contrib() As ContributorEntry
    
    'Reset contributor table (one entry per column for horizontal resampling)
    ReDim contrib(0 To dstWidth - 1) As ContributorEntry
    
    Dim wdth As Double, center As Double, weight As Double
    Dim intensityR As Double, intensityG As Double, intensityB As Double, intensityA As Double
    Dim left As Long, right As Long, i As Long, j As Long, k As Long
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Now, calculate all input weights
    
    'Horizontal downsampling
    If (xScale < 1#) Then
        
        'The source width is larger than the destination width
        wdth = (GetDefaultRadius(rsFilter) / xScale)

        For i = 0 To dstWidth - 1
            
            'Initialize contributor weight table
            contrib(i).n = 0
            ReDim contrib(i).p(0 To Int(2 * wdth + 1)) As Contributor
            contrib(i).wsum = 0#
            
            'Calculate center/left/right for this column
            center = ((i + 0.5) / xScale)
            left = Fix(center - wdth)
            right = Fix(center + wdth)
            
            For j = left To right
                
                'Ignore OOB pixels; they will not contribute to weighting
                If (j < 0) Then GoTo NextJXlt1
                If (j >= srcWidth) Then GoTo NextJXlt1
                
                weight = GetValue(rsFilter, (center - j - 0.5) * xScale)
                If (weight <> 0#) Then
                    contrib(i).p(contrib(i).n).pixel = j
                    contrib(i).p(contrib(i).n).weight = weight
                    contrib(i).wsum = contrib(i).wsum + weight
                    contrib(i).n = contrib(i).n + 1
                End If
NextJXlt1:
            Next j
            
            'Could exit here if user requests cancellation
            
        Next i
    
    'Horizontal upsampling
    'TODO: special case for unchanged width?
    Else
        
        'The source width is smaller than the destination width
        For i = 0 To dstWidth - 1
            
            'Initialize contributor weight table
            contrib(i).n = 0
            ReDim contrib(i).p(0 To Int(2 * GetDefaultRadius(rsFilter) + 1)) As Contributor
            contrib(i).wsum = 0#
            
            'Calculate center/left/right for this column
            center = ((i + 0.5) / xScale)
            left = Fix(center - GetDefaultRadius(rsFilter))
            right = Fix(center + GetDefaultRadius(rsFilter) + 0.99999999999)

            For j = left To right
                
                'Ignore OOB pixels; they will not contribute to weighting
                If (j < 0) Then GoTo NextJXgt1
                If (j >= srcWidth) Then GoTo NextJXgt1
                
                weight = GetValue(rsFilter, center - j - 0.5)
                If (weight <> 0#) Then
                    contrib(i).p(contrib(i).n).pixel = j
                    contrib(i).p(contrib(i).n).weight = weight
                    contrib(i).wsum = contrib(i).wsum + weight
                    contrib(i).n = contrib(i).n + 1
                End If
NextJXgt1:
            Next j
            
            'Could exit here if user requests cancellation
            
        Next i
    
    '/End up- vs downsampling
    End If
    
    'With weights successfully calculated, we can now filter horizontally from the input image
    ' to the temporary "working" copy.
    Dim srcImageData() As Byte, srcSA As SafeArray1D
    Dim idxPixel As Long
    
    'Each row (source image height)...
    For k = 0 To srcHeight - 1
        
        'Wrap a VB array around the image at this line
        srcDIB.WrapArrayAroundScanline srcImageData, srcSA, k
        
        'Each column (destination image width)...
        For i = 0 To dstWidth - 1

            intensityB = 0#
            intensityG = 0#
            intensityR = 0#
            intensityA = 0#
            
            'Generate weighted result for each color component
            For j = 0 To contrib(i).n - 1
                weight = contrib(i).p(j).weight
                If (weight <> 0#) Then
                    idxPixel = contrib(i).p(j).pixel * 4
                    intensityB = intensityB + (srcImageData(idxPixel) * weight)
                    intensityG = intensityG + (srcImageData(idxPixel + 1) * weight)
                    intensityR = intensityR + (srcImageData(idxPixel + 2) * weight)
                    intensityA = intensityA + (srcImageData(idxPixel + 3) * weight)
                End If
            Next j
            
            'Weight and clamp final RGBA values
            b = Int(intensityB / contrib(i).wsum + 0.5)
            If (b > 255) Then b = 255
            If (b < 0) Then b = 0
            
            g = Int(intensityG / contrib(i).wsum + 0.5)
            If (g > 255) Then g = 255
            If (g < 0) Then g = 0
            
            r = Int(intensityR / contrib(i).wsum + 0.5)
            If (r > 255) Then r = 255
            If (r < 0) Then r = 0
            
            a = Int(intensityA / contrib(i).wsum + 0.5)
            If (a > 255) Then a = 255
            If (a < 0) Then a = 0
            
            'Assign new RGBA values to the working data array
            idxPixel = i * 4
            m_tmpPixels(idxPixel) = b
            m_tmpPixels(idxPixel + 1) = g
            m_tmpPixels(idxPixel + 2) = r
            m_tmpPixels(idxPixel + 3) = a
            
        'Next pixel in row...
        Next i
        
        'Could check abort status here?
        
    'Next row in image...
    Next k
    
    'Free any unsafe references
    srcDIB.UnwrapArrayFromDIB srcImageData
    
    'Horizontal sampling is now complete.
    
    'Next, we need to perform nearly identical sampling from the "working" copy to the destination image
    
    'Reset contributor weight table (one entry per row for vertical resampling)
    ReDim contrib(0 To dstHeight - 1) As ContributorEntry
    
    'Vertical downsampling
    If (yScale < 1#) Then
        
        'The source height is larger than the destination height
        wdth = GetDefaultRadius(rsFilter) / yScale
        
        'Iterate through each row in the image
        For i = 0 To dstHeight - 1
          
            contrib(i).n = 0
            ReDim contrib(i).p(0 To Fix(2 * wdth + 1)) As Contributor
            contrib(i).wsum = 0#
          
            center = (i + 0.5) / yScale
            left = Fix(center - wdth)
            right = Fix(center + wdth)
            
            'Precalculate all weights for this column (technically these are not left/right values
            ' but up/down ones, remember)
            For j = left To right

                'Skip OOB pixels
                If (j < 0) Then GoTo NextJYlt1
                If (j >= srcHeight) Then GoTo NextJYlt1
                
                weight = GetValue(rsFilter, (center - j - 0.5) * yScale)
                If (weight <> 0#) Then
                    contrib(i).p(contrib(i).n).pixel = j
                    contrib(i).p(contrib(i).n).weight = weight
                    contrib(i).wsum = contrib(i).wsum + weight
                    contrib(i).n = contrib(i).n + 1
                End If
NextJYlt1:
            Next j
        
        'Next row...
        Next i
    
    'Vertical upsampling
    'TODO: special case for unchanged height?
    Else
        
        'The source height is smaller than the destination height
        For i = 0 To dstHeight - 1
            
            contrib(i).n = 0
            ReDim contrib(i).p(0 To Int(2 * GetDefaultRadius(rsFilter) + 1)) As Contributor
            contrib(i).wsum = 0#
            
            center = ((i + 0.5) / yScale)
            left = Fix(center - GetDefaultRadius(rsFilter))
            right = Fix(center + GetDefaultRadius(rsFilter) + 0.9999999999)

            For j = left To right
                
                If (j < 0) Then GoTo NextJYgt1
                If (j >= srcHeight) Then GoTo NextJYgt1
                
                weight = GetValue(rsFilter, center - j - 0.5)
                If (weight <> 0#) Then
                    contrib(i).p(contrib(i).n).pixel = j
                    contrib(i).p(contrib(i).n).weight = weight
                    contrib(i).wsum = contrib(i).wsum + weight
                    contrib(i).n = contrib(i).n + 1
                End If
NextJYgt1:
            Next j
            
            'Could check abort status here...
        
        'Next row...
        Next i
        
    '/End up- vs downsampling
    End If
    
    'With weights successfully calculated, we can now filter vertically from the "working copy"
    ' to the destination image.
    
    'Because we need to access pixels across rows, it's easiest to just wrap a single array around
    'the entire image.
    dstDIB.WrapArrayAroundDIB_1D srcImageData, srcSA
    
    'Each column (new image width)...
    For k = 0 To dstWidth - 1

        'Each row (destination image height)...
        For i = 0 To dstHeight - 1
            
            intensityB = 0#
            intensityG = 0#
            intensityR = 0#
            intensityA = 0#
            
            'Generate weighted result for each color component
            For j = 0 To contrib(i).n - 1
                weight = contrib(i).p(j).weight
                If (weight <> 0#) Then
                    idxPixel = (contrib(i).p(j).pixel * dstWidth * 4) + (k * 4)
                    intensityB = intensityB + (m_tmpPixels(idxPixel) * weight)
                    intensityG = intensityG + (m_tmpPixels(idxPixel + 1) * weight)
                    intensityR = intensityR + (m_tmpPixels(idxPixel + 2) * weight)
                    intensityA = intensityA + (m_tmpPixels(idxPixel + 3) * weight)
                End If
            Next j
            
            'Weight and clamp final RGBA values
            b = Int(intensityB / contrib(i).wsum + 0.5)
            If (b > 255) Then b = 255
            If (b < 0) Then b = 0
            
            g = Int(intensityG / contrib(i).wsum + 0.5)
            If (g > 255) Then g = 255
            If (g < 0) Then g = 0
            
            r = Int(intensityR / contrib(i).wsum + 0.5)
            If (r > 255) Then r = 255
            If (r < 0) Then r = 0
            
            a = Int(intensityA / contrib(i).wsum + 0.5)
            If (a > 255) Then a = 255
            If (a < 0) Then a = 0
            
            'Assign new RGBA values to the working data array
            idxPixel = (k * 4) + (i * dstWidth * 4)
            srcImageData(idxPixel) = b
            srcImageData(idxPixel + 1) = g
            srcImageData(idxPixel + 2) = r
            srcImageData(idxPixel + 3) = a
            
        'Next row...
        Next i
        
        'Could check abort status here?
        
    'Next column...
    Next k
    
    'Release all unsafe references
    dstDIB.UnwrapArrayFromDIB srcImageData
    
    'Resampling complete!
    ResampleImage = True
    
End Function

Private Function GetDefaultRadius(ByVal rsType As PD_ResamplingFilter) As Double

    Select Case rsType
        Case rf_Box
            GetDefaultRadius = 0.5
        Case rf_Triangle
            GetDefaultRadius = 1#
        Case rf_Hermite
            GetDefaultRadius = 1#
        Case rf_Bell
            GetDefaultRadius = 1.5
        Case rf_CubicBSpline
            GetDefaultRadius = 2#
        Case rf_Lanczos3
            GetDefaultRadius = 3#
        Case rf_Mitchell
            GetDefaultRadius = 2#
        Case rf_Cosine
            GetDefaultRadius = 1#
        Case rf_CatmullRom
            GetDefaultRadius = 2#
        Case rf_Quadratic
            GetDefaultRadius = 1.5
        Case rf_QuadraticBSpline
            GetDefaultRadius = 1.5
        Case rf_CubicConvolution
            GetDefaultRadius = 3#
        Case rf_Lanczos8
            GetDefaultRadius = 8#
    End Select

End Function

Private Function GetValue(ByVal rsType As PD_ResamplingFilter, ByVal x As Double) As Double
    
    Dim temp As Double
    
    Select Case rsType
        
        Case rf_Box
            If (x < 0#) Then x = -x
            If (x <= 0.5) Then GetValue = 1# Else GetValue = 0#
      
        Case rf_Triangle
            If (x < 0#) Then x = -x
            If (x < 1#) Then
                GetValue = (1# - x)
            Else
                GetValue = 0#
            End If
            
        Case rf_Hermite
            If (x < 0#) Then x = -x
            If (x < 1#) Then
                GetValue = ((2# * x - 3#) * x * x + 1#)
            Else
                GetValue = 0#
            End If
        
        Case rf_Bell
            If (x < 0#) Then x = -x
            If (x < 0.5) Then
                GetValue = (0.75 - x * x)
            ElseIf (x < 1.5) Then
                temp = x - 1.5
                GetValue = (0.5 * temp * temp)
            Else
                GetValue = 0#
            End If
      
        Case rf_CubicBSpline
            If (x < 0#) Then x = -x
            If (x < 1#) Then
                temp = x * x
                GetValue = (0.5 * temp * x - temp + 0.666666666666667)
            ElseIf (x < 2#) Then
                x = 2# - x
                GetValue = (x * x * x) / 6#
            Else
                GetValue = 0#
            End If
            
        Case rf_Lanczos3
            If (x < 0#) Then x = -x
            If (x < 3#) Then
                GetValue = (SinC(x) * SinC(x * 0.333333333333333))
            Else
                GetValue = 0#
            End If
      
        Case rf_Mitchell
            Const MC As Double = 0.333333333333333
            
            If (x < 0#) Then x = -x
            temp = x * x
            
            If (x < 1#) Then
                x = (((12 - 9 * MC - 6 * MC) * (x * temp)) + ((-18 + 12 * MC + 6 * MC) * temp) + (6 - 2 * MC))
                GetValue = (x / 6)
            ElseIf (x < 2#) Then
                x = (((-MC - 6 * MC) * (x * temp)) + ((6 * MC + 30 * MC) * temp) + ((-12 * MC - 48 * MC) * x) + (8 * MC + 24 * MC))
                GetValue = (x / 6)
            Else
                GetValue = 0#
            End If
            
        Case rf_Cosine
            If ((x >= -1) And (x <= 1)) Then
                GetValue = (Cos(x * PI) + 1#) * 0.5
            Else
                GetValue = 0#
            End If
            
        Case rf_CatmullRom
            If (x < 0#) Then x = -x
            temp = x * x
            If (x <= 1#) Then
                GetValue = (1.5 * temp * x - 2.5 * temp + 1)
            ElseIf (x <= 2#) Then
                GetValue = (-0.5 * temp * x + 2.5 * temp - 4 * x + 2)
            Else
                GetValue = 0#
            End If
            
        Case rf_Quadratic
            If (x < 0#) Then x = -x
            If (x <= 0.5) Then
                GetValue = (-2 * x * x + 1)
            ElseIf (x <= 1.5) Then
                GetValue = (x * x - 2.5 * x + 1.5)
            Else
                GetValue = 0#
            End If
            
        Case rf_QuadraticBSpline
            If (x < 0#) Then x = -x
            If (x <= 0.5) Then
                GetValue = (-x * x + 0.75)
            ElseIf (x <= 1.5) Then
                GetValue = 0.5 * x * x - 1.5 * x + 1.125
            Else
                GetValue = 0#
            End If
            
        Case rf_CubicConvolution
            If (x < 0#) Then x = -x
            temp = x * x
            If (x <= 1#) Then
                GetValue = ((4# / 3#) * temp * x - (7# / 3#) * temp + 1#)
            ElseIf (x <= 2#) Then
                GetValue = (-(7# / 12#) * temp * x + 3# * temp - (59# / 12#) * x + 2.5)
            ElseIf (x <= 3#) Then
                GetValue = ((1# / 12#) * temp * x - (2# / 3#) * temp + 1.75 * x - 1.5)
            Else
                GetValue = 0#
            End If
            
        Case rf_Lanczos8
            If (x < 0#) Then x = -x
            If (x < 8#) Then
                GetValue = (SinC(x) * SinC(x * 0.25))
            Else
                GetValue = 0#
            End If
            
    End Select

End Function

Private Function SinC(ByVal x As Double) As Double
    If (x <> 0!) Then
        x = x * PI
        SinC = Sin(x) / x
    Else
        SinC = 1#
    End If
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    
    Dim errText As String
    errText = "Resampling." & funcName & "() reported an error: " & errDescription
    
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction errText
    Else
        Debug.Print errText
    End If
    
End Sub
