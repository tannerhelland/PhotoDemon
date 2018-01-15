Attribute VB_Name = "Filters_Render"
'***************************************************************************
'Render Filter Collection
'Copyright 2017-2018 by Tanner Helland
'Created: 14/October/17
'Last updated: 14/October/17
'Last update: start migrating render-specific functions here
'
'Container module for PD's render filter collection.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Render a "cloud" effect to an arbitrary DIB.  The DIB must already exist and be sized to whatever dimensions
' the caller requires.
Public Function GetCloudDIB(ByRef dstDIB As pdDIB, ByVal fxScale As Double, Optional ByVal fxQuality As Long = 4, Optional ByVal fxRndSeed As Double = 0#, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    'Quality is passed on a [1, 8] scale; rework it to [0, 7] now
    fxQuality = fxQuality - 1
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Long, dstSA As SafeArray1D
    dstDIB.WrapLongArrayAroundScanline dstImageData, dstSA, 0
    
    Dim dibPtr As Long, dibStride As Long
    dibPtr = dstDIB.GetDIBPointer
    dibStride = dstDIB.GetDIBStride
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBWidth - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalX Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Scale is used as a fraction of the image's smallest dimension.
    If (finalX > finalY) Then
        fxScale = (fxScale * 0.01) * dstDIB.GetDIBHeight
    Else
        fxScale = (fxScale * 0.01) * dstDIB.GetDIBWidth
    End If
    
    If (fxScale > 0#) Then fxScale = 1# / fxScale
    
    'A pdNoise instance handles the actual noise generation
    Dim cNoise As pdNoise
    Set cNoise = New pdNoise
    
    'To generate "random" values despite using a fixed 2D noise generator, we calculate random offsets
    ' into the "infinite grid" of possible noise values.  This yields (perceptually) random results.
    Dim rndOffsetX As Double, rndOffsetY As Double
    
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_Float fxRndSeed
    rndOffsetX = cRandom.GetRandomFloat_WH * 10000000# - 5000000#
    rndOffsetY = cRandom.GetRandomFloat_WH * 10000000# - 5000000#
    
    'Some values can be cached in the interior loop to speed up processing time
    Dim pNoiseCache As Double, xScaleCache As Double, yScaleCache As Double
    
    'Finally, an integer displacement will be used to actually calculate the RGB values at any point in the fog
    Dim pDisplace As Long, i As Long
    
    'The bulk of the processing time for this function occurs when we set up the initial cloud table; rather than
    ' doing this as part of the RGB assignment array, I've separated it into its own step (in hopes the compiled
    ' will be better able to optimize it!)
    Dim p2Lookup() As Single, p2InvLookup() As Single
    ReDim p2Lookup(0 To fxQuality) As Single, p2InvLookup(0 To fxQuality) As Single
    
    'The fractal noise approach we use requires successive sums of 2 ^ n and 2 ^ -n; we calculate these in advance
    ' as the POW operator is so hideously slow.
    For i = 0 To fxQuality
        p2Lookup(i) = 2 ^ i
        p2InvLookup(i) = 1# / (2 ^ i)
    Next i
    
    'Generate a displacement lookup table.  Because we ignore alpha values at present, and clouds are always
    ' produced as grayscale, we have a fixed set of possible output values.
    Dim dispLookup() As Long
    ReDim dispLookup(0 To 255) As Long
    
    Dim tmpRGBA As RGBQuad, dispFinal As Byte
    
    For i = 0 To 255
        With tmpRGBA
            .Red = i
            .Green = i
            .Blue = i
            .Alpha = 255
        End With
        CopyMemory ByVal VarPtr(dispLookup(i)), ByVal VarPtr(tmpRGBA), 4&
    Next i
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        dstSA.pvData = dibPtr + dibStride * y
        yScaleCache = CDbl(y) * fxScale
    For x = initX To finalX
    
        'Calculate an x-displacement for this point.  (Note that y-displacements are calculated in the outer loop.)
        xScaleCache = CDbl(x) * fxScale
        pNoiseCache = 0#
        
        'Fractal noise works by summing successively smaller noise values taken from successively larger
        ' amplitudes of the original function.
        For i = 0 To fxQuality
            pNoiseCache = pNoiseCache + p2InvLookup(i) * cNoise.SimplexNoise2d(rndOffsetX + xScaleCache * p2Lookup(i), rndOffsetY + yScaleCache * p2Lookup(i))
        Next i
        
        'Convert the calculated noise value to RGB range and cache it
        pDisplace = 127 + (pNoiseCache * 127#)
        If (pDisplace > 255) Then pDisplace = 255 Else If (pDisplace < 0) Then pDisplace = 0
        
        'Write all RGBA bytes at once
        dstImageData(x) = dispLookup(pDisplace)
          
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'tmpFogDIB now contains a grayscale representation of our fog data
    dstDIB.UnwrapLongArrayFromDIB dstImageData
    dstDIB.SetInitialAlphaPremultiplicationState True
    
    GetCloudDIB = True
        
End Function

