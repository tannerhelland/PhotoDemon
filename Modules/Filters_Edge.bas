Attribute VB_Name = "Filters_Edge"
'***************************************************************************
'Edge Filter Collection
'Copyright 2003-2026 by Tanner Helland
'Created: sometime 2003?
'Last updated: 14/September/22
'Last update: start migrating some edge filters here
'
'Container module for PD's edge filter collection.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a relief filter, which gives the image a pseudo-3D appearance.
'
'To use as a standard PD effect, do *not* pass srcDIB.  (If srcDIB is empty, PD will auto-retrieve the
' current working DIB.)
Public Sub Filter_Edge_Relief(ByVal effectParams As String, Optional ByRef srcDIB As pdDIB = Nothing, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Carving image relief..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim eDistance As Double, eAngle As Double, eDepth As Double
    
    With cParams
        eDistance = .GetDouble("distance", 1!)
        eAngle = .GetDouble("angle", 0!)
        eDepth = .GetDouble("depth", 10!)
    End With
    
    'As a divisor, we can't allow distance to be 0
    If eDistance = 0# Then eDistance = 0.01
       
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D, dstSA1D As SafeArray1D
    If (srcDIB Is Nothing) Then
        EffectPrep.PrepImageData dstSA, toPreview, dstPic
    Else
        Set workingDIB = srcDIB
    End If
    
    'Create a copy of the current image; we will use it as our source reference.
    Dim copyOfSrcDIB As pdDIB
    Set copyOfSrcDIB = New pdDIB
    copyOfSrcDIB.CreateFromExistingDIB workingDIB
    
    'At present, stride is always width * 4 (32-bit RGBA)
    Dim xStride As Long
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.SetDistortParameters pdeo_Clamp, True, curDIBValues.maxX, curDIBValues.maxY
    
    'During previews, adjust the distance parameter to compensate for preview size
    If toPreview And (srcDIB Is Nothing) Then eDistance = eDistance * curDIBValues.previewModifier
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tG As Long, tB As Long
    Dim reliefOffset As Double
    
    'Convert the rotation angle to radians
    eAngle = eAngle * (PI / 180#)
    
    'X and Y offsets are hard-coded per the current angle
    Dim xOffset As Double, yOffset As Double
    xOffset = Cos(eAngle) * eDistance
    yOffset = Sin(eAngle) * eDistance
    
    Dim tmpQuad As RGBQuad
    fSupport.AliasTargetDIB copyOfSrcDIB
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline dstImageData, dstSA1D, y
    For x = initX To finalX
    
        'Retrieve original RGB values
        xStride = x * 4
        b = dstImageData(xStride)
        g = dstImageData(xStride + 1)
        r = dstImageData(xStride + 2)
        
        'Use the filter support class to interpolate and edge-clamp pixels as necessary
        ' on the source pixel at the pre-calculated offset
        tmpQuad = fSupport.GetColorsFromSource(x + xOffset, y + yOffset, x, y)
        tB = tmpQuad.Blue
        tG = tmpQuad.Green
        tR = tmpQuad.Red
        
        'Calculate a single grayscale relief value (and emphasize green similar to a
        ' luminance calculation - this lets us use bit-shifts for division)
        reliefOffset = ((r - tR) + (g - tG) * 2 + (b - tB)) \ 4
        reliefOffset = reliefOffset * eDepth
        
        'Apply the relief to each channel
        r = r + reliefOffset
        g = g + reliefOffset
        b = b + reliefOffset
                
        'Clamp RGB values
        If (r < 0) Then r = 0
        If (r > 255) Then r = 255
        If (g < 0) Then g = 0
        If (g > 255) Then g = 255
        If (b < 0) Then b = 0
        If (b > 255) Then b = 255
        
        dstImageData(xStride) = b
        dstImageData(xStride + 1) = g
        dstImageData(xStride + 2) = r
        
        'Leave alpha as-is for this effect
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate all image arrays
    fSupport.UnaliasTargetDIB
    workingDIB.UnwrapArrayFromDIB dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    If (srcDIB Is Nothing) Then
        EffectPrep.FinalizeImageData toPreview, dstPic
    Else
        Set workingDIB = Nothing
    End If
 
End Sub
