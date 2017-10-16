Attribute VB_Name = "Filters_Natural"
'***************************************************************************
'"Natural" Filters
'Copyright 2002-2017 by Tanner Helland
'Created: 8/April/02
'Last updated: Summer '14
'Last update: see comments below
'
'NOTE!!  As of summer 2014, I am working on rewriting all "Nature" filters as full-blown filters, and not these
' crappy little one-click variants.  Any that can't be reworked will be dropped for good.
'
'Runs all nature-type filters.  Includes water, steel, burn, rainbow, etc.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply a "water" effect to an image.  (See the "ocean" effect below for a similar approach.)
Public Sub MenuWater()
    
    Message "Submerging image in artificial water..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent diffused pixels from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SafeArray2D
    
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'Because interpolation may be used, it's necessary to keep pixel values within special ranges
    Dim xLimit As Long, yLimit As Long
    xLimit = finalX - 1
    yLimit = finalY - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
          
    'This wave transformation requires specialized variables
    Dim xWavelength As Double
    xWavelength = 31
    
    Dim xAmplitude As Double
    xAmplitude = 10
        
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
        
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long, a As Long
    Dim grayVal As Long
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For x = 0 To 765
        gLookup(x) = CByte(x \ 3)
    Next x
                 
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Calculate new source pixel locations
        srcX = x + Sin(y / xWavelength) * xAmplitude
        srcY = y
        
        'Make sure the source coordinates are in-bounds
        If srcX < 0 Then srcX = 0
        If srcX > xLimit Then srcX = xLimit
        If srcY > yLimit Then srcY = yLimit
        
        'Interpolate the source pixel for better results
        r = GetInterpolatedVal(srcX, srcY, srcImageData, 2, qvDepth)
        g = GetInterpolatedVal(srcX, srcY, srcImageData, 1, qvDepth)
        b = GetInterpolatedVal(srcX, srcY, srcImageData, 0, qvDepth)
        If qvDepth = 4 Then a = GetInterpolatedVal(srcX, srcY, srcImageData, 3, qvDepth)
            
        'Now, modify the colors to give a bluish-green tint to the image
        grayVal = gLookup(r + g + b)
        
        r = gray - g - b
        g = gray - r - b
        b = gray - r - g
        
        'Keep all values in range
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
            
        'Write the colors (and alpha, if necessary) out to the destination image's data
        dstImageData(quickVal + 2, y) = r
        dstImageData(quickVal + 1, y) = g
        dstImageData(quickVal, y) = b
        If qvDepth = 4 Then dstImageData(quickVal + 3, y) = a
            
    Next y
        If (x And progBarCheck) = 0 Then
            If Interface.UserPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
    
    'Safely deallocate all image arrays
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub

'Apply a strange, lava-ish transformation to an image
Public Sub MenuLava()
    
    'NEW APPROACH!  This is just testing phase now...
    Dim effectDIB As pdDIB
    Set effectDIB = New pdDIB
    
    Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, False, Nothing
    
    Filters_Render.GetCloudDIB workingDIB, 10#, 4, Rnd
    
    EffectPrep.FinalizeImageData False, Nothing, False
    
    Exit Sub
    
    Message "Exploding imaginary volcano..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    'Dim tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For x = 0 To 765
        gLookup(x) = CByte(x \ 3)
    Next x
        
    'Apply the filter
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
        
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
        
        grayVal = gLookup(r + g + b)
        
        r = grayVal
        g = Abs(b - 128)
        b = Abs(b - 128)
        
        imageData(quickVal + 2, y) = r
        imageData(quickVal + 1, y) = g
        imageData(quickVal, y) = b
        
    Next y
        If (x And progBarCheck) = 0 Then
            If Interface.UserPressedESC() Then Exit For
            SetProgBarVal x
        End If
    Next x
        
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
    
End Sub

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
    If (steelSmoothness > 0) Then Filters_ByteArray.GaussianBlur_IIR_ByteArray grayMap, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, steelSmoothness, 2
        
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
    
    Dim cLut As pdFilterLUT
    Set cLut = New pdFilterLUT
    cLut.FillLUT_Curve rLookup, rCurve
    cLut.FillLUT_Curve gLookup, gCurve
    cLut.FillLUT_Curve bLookup, bCurve
        
    'We are now ready to apply the final curve to the image!
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    srcDIB.WrapArrayAroundDIB dstImageData, dstSA
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
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
        quickVal = x * qvDepth
    For y = initY To finalY
        
        grayVal = grayMap(x, y)
        
        dstImageData(quickVal, y) = bLookup(grayVal)
        dstImageData(quickVal + 1, y) = gLookup(grayVal)
        dstImageData(quickVal + 2, y) = rLookup(grayVal)
        
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
