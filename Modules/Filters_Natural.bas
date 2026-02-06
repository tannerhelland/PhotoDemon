Attribute VB_Name = "Filters_Natural"
'***************************************************************************
'"Natural" Filters
'Copyright 2002-2026 by Tanner Helland
'Created: 8/April/02
'Last updated: 17/October/17
'Last update: migrate "metal" effect here, as it's useful in a number of other settings
'
'Runs all nature-type filters.  Includes water, steel, burn, rainbow, etc.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Given two DIBs, fill one with a "chrome-filtered" version of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
' This operation is performed in-place, so no separate destination DIB is required.
Public Function GetChromeDIB(ByRef srcDIB As pdDIB, ByVal steelDetail As Long, ByVal steelSmoothness As Double, Optional ByVal shadowColor As Long = vbBlack, Optional ByVal highlightColor As Long = vbWhite, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Decompose the shadow and highlight colors into their individual color components
    Dim rShadow As Long, gShadow As Long, bShadow As Long
    Dim rHighlight As Long, gHighlight As Long, bHighlight As Long
    
    rShadow = Colors.ExtractRed(shadowColor)
    gShadow = Colors.ExtractGreen(shadowColor)
    bShadow = Colors.ExtractBlue(shadowColor)
    
    rHighlight = Colors.ExtractRed(highlightColor)
    gHighlight = Colors.ExtractGreen(highlightColor)
    bHighlight = Colors.ExtractBlue(highlightColor)
    
    'Retrieve a normalized luminance map of the current image
    Dim grayMap() As Byte
    DIBs.GetDIBGrayscaleMap srcDIB, grayMap, True
    
    'If the user specified a non-zero smoothness, apply it now
    If (steelSmoothness > 0) Then Filters_ByteArray.GaussianBlur_AM_ByteArray grayMap, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, steelSmoothness, 2
        
    'Re-normalize the data (this ends up not being necessary, but it could be exposed to the user in a future update)
    'Filters_ByteArray.normalizeByteArray grayMap, workingDIB.getDIBWidth, workingDIB.getDIBHeight
    
    'Next, we need to generate a sinusoidal octave lookup table for the graymap.  This causes the luminance of the map to
    ' vary evenly between the number of detail points requested by the user.
    
    'Detail cannot be lower than 2, but it is presented to the user as [0, (arbitrary upper bound)], so add two to the total now
    steelDetail = steelDetail + 2
    
    'We will be using pdFilterLUT to generate corresponding RGB lookup tables, which means we need to use POINTFLOAT arrays
    Dim rCurve() As PointFloat, gCurve() As PointFloat, bCurve() As PointFloat
    ReDim rCurve(0 To steelDetail) As PointFloat
    ReDim gCurve(0 To steelDetail) As PointFloat
    ReDim bCurve(0 To steelDetail) As PointFloat
    
    Dim detailModifier As Double
    detailModifier = 1# / CDbl(steelDetail)
    
    'For all channels, X values are evenly distributed from 0 to 255
    Dim i As Long
    For i = 0 To steelDetail
        rCurve(i).x = CDbl(i) * detailModifier * 255#
        gCurve(i).x = CDbl(i) * detailModifier * 255#
        bCurve(i).x = CDbl(i) * detailModifier * 255#
    Next i
    
    'Y values alternate between the shadow and highlight colors; these are calculated on a per-channel basis
    For i = 0 To steelDetail
        
        If (i Mod 2) = 0 Then
            rCurve(i).y = rShadow
            gCurve(i).y = gShadow
            bCurve(i).y = bShadow
        Else
            rCurve(i).y = rHighlight
            gCurve(i).y = gHighlight
            bCurve(i).y = bHighlight
        End If
        
    Next i
    
    'Convert our point array into color curves
    Dim rLookup() As Byte, gLookup() As Byte, bLookup() As Byte
    
    Dim cLUT As pdFilterLUT
    Set cLUT = New pdFilterLUT
    cLUT.FillLUT_Curve rLookup, rCurve
    cLUT.FillLUT_Curve gLookup, gCurve
    cLUT.FillLUT_Curve bLookup, bCurve
        
    'We are now ready to apply the final curve to the image!
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB dstImageData, dstSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalX Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim grayVal As Long
    
    'Apply the filter
    For x = initX To finalX
        xStride = x * 4
    For y = initY To finalY
        
        grayVal = grayMap(x, y)
        
        dstImageData(xStride, y) = bLookup(grayVal)
        dstImageData(xStride + 1, y) = gLookup(grayVal)
        dstImageData(xStride + 2, y) = rLookup(grayVal)
        
    Next y
        If (Not suppressMessages) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
        
    'Safely deallocate imageData()
    srcDIB.UnwrapArrayFromDIB dstImageData
    
    GetChromeDIB = 1
    
End Function
