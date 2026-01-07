Attribute VB_Name = "Filters_Render"
'***************************************************************************
'Render Filter Collection
'Copyright 2017-2026 by Tanner Helland
'Created: 14/October/17
'Last updated: 15/August/21
'Last update: wrap up work on Truchet tile rendering
'
'Container module for PD's render filter collection.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Truchet tiles support various shapes
Public Enum PD_TruchetShape
    ts_Arc = 0
    ts_Line = 1
    ts_Maze = 2
    ts_Triangle = 3
    ts_Octagon = 4
    ts_Circle = 5
    ts_Max = 6
End Enum

#If False Then
    Private Const ts_Arc = 0, ts_Line = 1, ts_Maze = 2, ts_Triangle = 3, ts_Octagon = 4, ts_Circle = 5, ts_Max = 6
#End If

'You can also generate tiles in various patterns.  These bare some similarity to tiles in PD's Voronoi tool,
' primarily so that we can share localization text between the two.
Public Enum PD_TruchetPattern
    tp_Random = 0
    tp_BaseImage = 1
    tp_Repeat = 2
    tp_Wave = 3
    tp_Quilt = 4
    tp_Chain = 5
    tp_Weave = 6
    tp_Max = 7
End Enum

#If False Then
    Private Const tp_Random = 0, tp_BaseImage = 1, tp_Repeat = 2, tp_Wave = 3, tp_Quilt = 4, tp_Chain = 5, tp_Weave = 6
    Private Const tp_Max = 7
#End If

'Render a "cloud" effect to an arbitrary DIB.  The DIB must already exist and be sized to whatever dimensions
' the caller requires.
Public Function GetCloudDIB(ByRef dstDIB As pdDIB, ByVal fxScale As Double, ByVal ptrToPalette As Long, ByVal numPalColors As Long, Optional ByVal noiseGenerator As PD_NoiseGenerator = ng_Simplex, Optional ByVal fxQuality As Long = 4, Optional ByVal fxRndSeed As Double = 0#, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    'Quality is passed on a [1, 8] scale; rework it to [0, 7] now
    fxQuality = fxQuality - 1
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Long, dstSA As SafeArray1D
    dstDIB.WrapLongArrayAroundScanline dstImageData, dstSA, 0
    
    Dim dibPtr As Long, dibStride As Long
    dibPtr = dstDIB.GetDIBPointer
    dibStride = dstDIB.GetDIBStride
    
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
    
    'Generate a displacement lookup table.  Because we don't need to assign individual RGBA values,
    ' it's faster to alias our incoming palette (type RGBQuad) into a Long-type array, because we
    ' can then assign all four RGBA lookup values at once.
    Dim dispLookup() As Long
    ReDim dispLookup(0 To numPalColors - 1) As Long
    CopyMemoryStrict VarPtr(dispLookup(0)), ptrToPalette, 4& * numPalColors
    
    Dim lookupMaxI As Long, halfLookupF As Long
    lookupMaxI = numPalColors - 1
    halfLookupF = lookupMaxI / 2#
    
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
        If (noiseGenerator = ng_Perlin) Then
            For i = 0 To fxQuality
                pNoiseCache = pNoiseCache + p2InvLookup(i) * cNoise.PerlinNoise2d(rndOffsetX + xScaleCache * p2Lookup(i), rndOffsetY + yScaleCache * p2Lookup(i))
            Next i
        ElseIf (noiseGenerator = ng_Simplex) Then
            For i = 0 To fxQuality
                pNoiseCache = pNoiseCache + p2InvLookup(i) * cNoise.SimplexNoise2d(rndOffsetX + xScaleCache * p2Lookup(i), rndOffsetY + yScaleCache * p2Lookup(i))
            Next i
        Else
            For i = 0 To fxQuality
                pNoiseCache = pNoiseCache + p2InvLookup(i) * cNoise.OpenSimplexNoise2d(rndOffsetX + xScaleCache * p2Lookup(i), rndOffsetY + yScaleCache * p2Lookup(i))
            Next i
        End If
        
        'Convert the calculated noise value to RGB range and cache it
        pDisplace = Int(halfLookupF + (pNoiseCache * halfLookupF) + 0.5)
        If (pDisplace > lookupMaxI) Then pDisplace = lookupMaxI
        If (pDisplace < 0&) Then pDisplace = 0&
        
        'Write all RGBA bytes at once
        dstImageData(x) = dispLookup(pDisplace)
          
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal modifyProgBarOffset + y
            End If
        End If
    Next y
    
    'tmpFogDIB now contains a grayscale representation of our fog data
    dstDIB.UnwrapLongArrayFromDIB dstImageData
    dstDIB.SetInitialAlphaPremultiplicationState True
    
    GetCloudDIB = True
        
End Function

Public Function GetNameOfTruchetPattern(ByVal truchetID As PD_TruchetPattern, Optional ByVal getLocalizedName As Boolean = True) As String
    
    'A (sloppy) branch is used so that the language translation tool finds the various strings in this function
    If getLocalizedName Then
        Select Case truchetID
            Case tp_Random
                GetNameOfTruchetPattern = g_Language.TranslateMessage("random")
            Case tp_BaseImage
                GetNameOfTruchetPattern = g_Language.TranslateMessage("image")
            Case tp_Repeat
                GetNameOfTruchetPattern = g_Language.TranslateMessage("repeat")
            Case tp_Wave
                GetNameOfTruchetPattern = g_Language.TranslateMessage("waves")
            Case tp_Quilt
                GetNameOfTruchetPattern = g_Language.TranslateMessage("quilt")
            Case tp_Chain
                GetNameOfTruchetPattern = g_Language.TranslateMessage("chain")
            Case tp_Weave
                GetNameOfTruchetPattern = g_Language.TranslateMessage("weave")
        End Select
    Else
        Select Case truchetID
            Case tp_Random
                GetNameOfTruchetPattern = "random"
            Case tp_BaseImage
                GetNameOfTruchetPattern = "image"
            Case tp_Repeat
                GetNameOfTruchetPattern = "repeat"
            Case tp_Wave
                GetNameOfTruchetPattern = "waves"
            Case tp_Quilt
                GetNameOfTruchetPattern = "quilt"
            Case tp_Chain
                GetNameOfTruchetPattern = "chain"
            Case tp_Weave
                GetNameOfTruchetPattern = "weave"
        End Select
    End If
    
End Function

Public Function GetNameOfTruchetShape(ByVal truchetID As PD_TruchetShape, Optional ByVal getLocalizedName As Boolean = True) As String
    
    'A (sloppy) branch is used so that the language translation tool finds the various strings in this function
    If getLocalizedName Then
        Select Case truchetID
            Case ts_Arc
                GetNameOfTruchetShape = g_Language.TranslateMessage("arc")
            Case ts_Line
                GetNameOfTruchetShape = g_Language.TranslateMessage("line")
            Case ts_Maze
                GetNameOfTruchetShape = g_Language.TranslateMessage("maze")
            Case ts_Triangle
                GetNameOfTruchetShape = g_Language.TranslateMessage("triangle")
            Case ts_Octagon
                GetNameOfTruchetShape = g_Language.TranslateMessage("octagon")
            Case ts_Circle
                GetNameOfTruchetShape = g_Language.TranslateMessage("circle")
        End Select
    Else
        Select Case truchetID
            Case ts_Arc
                GetNameOfTruchetShape = "arc"
            Case ts_Line
                GetNameOfTruchetShape = "line"
            Case ts_Maze
                GetNameOfTruchetShape = "maze"
            Case ts_Triangle
                GetNameOfTruchetShape = "triangle"
            Case ts_Octagon
                GetNameOfTruchetShape = "octagon"
            Case ts_Circle
                GetNameOfTruchetShape = "circle"
        End Select
    End If
End Function

'Name *MUST* be the unmodified, en-US name!
Public Function GetTruchetPatternFromName(ByRef truchetName As String) As PD_TruchetPattern
    Select Case LCase$(truchetName)
        Case "random"
            GetTruchetPatternFromName = tp_Random
        Case "image"
            GetTruchetPatternFromName = tp_BaseImage
        Case "repeat"
            GetTruchetPatternFromName = tp_Repeat
        Case "waves"
            GetTruchetPatternFromName = tp_Wave
        Case "quilt"
            GetTruchetPatternFromName = tp_Quilt
        Case "chain"
            GetTruchetPatternFromName = tp_Chain
        Case "weave"
            GetTruchetPatternFromName = tp_Weave
    End Select
End Function

'Name *MUST* be the unmodified, en-US name!
Public Function GetTruchetShapeFromName(ByRef truchetName As String) As PD_TruchetShape
    Select Case LCase$(truchetName)
        Case "arc"
            GetTruchetShapeFromName = ts_Arc
        Case "line"
            GetTruchetShapeFromName = ts_Line
        Case "maze"
            GetTruchetShapeFromName = ts_Maze
        Case "triangle"
            GetTruchetShapeFromName = ts_Triangle
        Case "octagon"
            GetTruchetShapeFromName = ts_Octagon
        Case "circle"
            GetTruchetShapeFromName = ts_Circle
    End Select
End Function

'Render a "fiber" effect to an arbitrary DIB.  A two-color system (a la Photoshop) is used.
' The DIB must already exist and be sized to whatever dimensions the caller requires.
Public Function RenderFibers_TwoColor(ByRef dstDIB As pdDIB, ByVal firstColorRGBA As Long, ByVal secondColorRGBA As Long, ByVal fxStrength As Double, Optional ByVal fxRndSeed As Double = 0#, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Long, dstSA As SafeArray2D
    dstDIB.WrapLongArrayAroundDIB dstImageData, dstSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long, yStep As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBWidth - 1
    finalY = dstDIB.GetDIBHeight - 1
    yStep = 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Prep a randomizer
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_Float fxRndSeed
    
    'Set the initial color randomly
    Dim lastColor As Long, newColor As Long, tmpColor As Long
    If (cRandom.GetRandomFloat_WH() > 0.5) Then
        lastColor = firstColorRGBA
        newColor = secondColorRGBA
    Else
        lastColor = secondColorRGBA
        newColor = firstColorRGBA
    End If
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
    For y = initY To finalY Step yStep
        
        If (cRandom.GetRandomFloat_WH() < fxStrength) Then
            tmpColor = lastColor
            lastColor = newColor
            newColor = tmpColor
        End If
        
        'Write all RGBA bytes at once
        dstImageData(x, y) = lastColor
          
    Next y
        If (Not suppressMessages) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal modifyProgBarOffset + x
            End If
        End If
        
        'Switch direction on each iteration (serpentine)
        If (yStep > 0) Then
            initY = finalY
            finalY = 0
            yStep = -1
        Else
            finalY = initY
            initY = 0
            yStep = 1
        End If
        
    Next x
    
    'tmpFogDIB now contains a grayscale representation of our fog data
    dstDIB.UnwrapLongArrayFromDIB dstImageData
    dstDIB.SetInitialAlphaPremultiplicationState True
    
    RenderFibers_TwoColor = True
        
End Function

'Render a "fiber" effect to an arbitrary DIB.  An arbitrary lookup-table system is used.
' The DIB must also already exist and be sized to whatever dimensions the caller requires.
Public Function RenderFibers_LUT(ByRef dstDIB As pdDIB, ByRef cLUT() As Long, ByVal fxStrength As Double, Optional ByVal fxRndSeed As Double = 0#, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Long, dstSA As SafeArray2D
    dstDIB.WrapLongArrayAroundDIB dstImageData, dstSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = dstDIB.GetDIBWidth - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Prep a randomizer
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_Float fxRndSeed
    
    'Set the initial color randomly
    Dim lastColor As Long, lutLimit As Long
    lutLimit = UBound(cLUT)
    lastColor = cLUT(Int(cRandom.GetRandomFloat_WH() * lutLimit + 0.9999))
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
    For y = initY To finalY
        
        If (cRandom.GetRandomFloat_WH() < fxStrength) Then lastColor = cLUT(Int(cRandom.GetRandomFloat_WH() * lutLimit + 0.9999))
        
        'Write all RGBA bytes at once
        dstImageData(x, y) = lastColor
          
    Next y
        If (Not suppressMessages) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal modifyProgBarOffset + x
            End If
        End If
    Next x
    
    'tmpFogDIB now contains a grayscale representation of our fog data
    dstDIB.UnwrapLongArrayFromDIB dstImageData
    dstDIB.SetInitialAlphaPremultiplicationState True
    
    RenderFibers_LUT = True
        
End Function

'Render "Truchet tiles" to an arbitrary DIB.  The DIB must already exist and be sized to whatever dimensions
' the caller requires.
Public Function GetTruchetDIB(ByRef dstDIB As pdDIB, ByVal fxScale As Long, ByVal fxLineWidthRatio As Single, ByVal fxForegroundColor As Long, ByVal fxForegroundOpacity As Single, ByVal fxBackgroundColor As Long, ByVal fxBackgroundOpacity As Single, Optional ByVal fxShape As PD_TruchetShape = ts_Triangle, Optional ByVal fxPattern As PD_TruchetPattern = tp_Random, Optional ByVal fxRndSeed As Double = 0#, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Boolean
    
    'Basic validation of input params
    If (fxScale > dstDIB.GetDIBWidth) Then fxScale = dstDIB.GetDIBWidth
    If (fxScale > dstDIB.GetDIBHeight) Then fxScale = dstDIB.GetDIBHeight
    If (fxScale < 3) Then fxScale = 3
    
    'Create an array of source images.  (The count varies by shape.)
    ' Some images may be generated using an array of floating-point coordinates.  Initialize that count
    ' here as well, as needed.
    Dim srcImages() As pdDIB, numImages As Long
    Dim fxPoints() As PointFloat, numPoints As Long
    
    Select Case fxShape
        Case ts_Arc
            numImages = 2
        Case ts_Line
            numImages = 2
        Case ts_Maze
            numImages = 2
            numPoints = 2
        Case ts_Triangle
            numImages = 4
            numPoints = 3
        Case ts_Octagon
            numImages = 4
            numPoints = 4
        Case ts_Circle
            numImages = 4
            numPoints = 1
    End Select
    
    ReDim srcImages(0 To numImages - 1) As pdDIB
    If (numPoints > 0) Then ReDim fxPoints(0 To numPoints - 1) As PointFloat
    
    'Initialize the initial image collection to the specified background color and opacity
    Dim x As Long, y As Long
    For x = 0 To numImages - 1
        Set srcImages(x) = New pdDIB
        srcImages(x).CreateBlank fxScale, fxScale, 32, fxBackgroundColor, Int(fxBackgroundOpacity * 2.55 + 0.5)
        If (Not srcImages(x).GetAlphaPremultiplication) Then srcImages(x).SetAlphaPremultiplication True
    Next x
    
    'Create a brush and/or pen for the foreground color
    Dim foreBrush As pd2DBrush, forePen As pd2DPen
    Set foreBrush = New pd2DBrush
    foreBrush.SetBrushMode P2_BM_Solid
    
    Set forePen = New pd2DPen
    forePen.SetPenStyle P2_DS_Solid
    
    Select Case fxShape
        Case ts_Arc
            forePen.SetPenColor fxForegroundColor
            forePen.SetPenOpacity fxForegroundOpacity
            forePen.SetPenLineCap P2_LC_Round
            
        Case ts_Line, ts_Maze
            forePen.SetPenColor fxForegroundColor
            forePen.SetPenOpacity fxForegroundOpacity
            forePen.SetPenLineCap P2_LC_Square
            
        Case Else
            foreBrush.SetBrushColor fxForegroundColor
            foreBrush.SetBrushOpacity fxForegroundOpacity
        
    End Select
    
    'Pen-based shapes also need to specify line width.  This is passed as a ratio on the scale [0, 100] and we
    ' transform it as a ratio of the current tile size.
    Dim penWidth As Single
    
    If (fxLineWidthRatio > 100!) Then fxLineWidthRatio = 100!
    If (fxLineWidthRatio < 1!) Then fxLineWidthRatio = 1!
    fxLineWidthRatio = fxLineWidthRatio / 100!
    penWidth = fxLineWidthRatio * (fxScale / 3!)
    
    'Failsafe check on pen width
    If (penWidth < 1!) Then penWidth = 1!
    forePen.SetPenWidth penWidth
    
    'More complex shapes may use a path object for rendering
    Dim tmpPath As pd2DPath
            
    'Generate all tile images (approach varies by shape)
    Dim dstSurface As pd2DSurface
    Set dstSurface = New pd2DSurface
    
    For x = 0 To numImages - 1
        
        'Start by wrapping a pd2D surface around the target tile image
        dstSurface.WrapSurfaceAroundPDDIB srcImages(x)
        dstSurface.SetSurfaceAntialiasing P2_AA_HighQuality
        dstSurface.SetSurfacePixelOffset P2_PO_Half
        
        Select Case fxShape
            
            Case ts_Arc
                If (x = 0) Then
                    PD2D.DrawArcF dstSurface, forePen, 0!, 0!, fxScale / 2!, 0!, 91!
                    PD2D.DrawArcF dstSurface, forePen, fxScale, fxScale, fxScale / 2!, 180!, 91!
                Else
                    PD2D.DrawArcF dstSurface, forePen, fxScale, 0!, fxScale / 2!, 90!, 91!
                    PD2D.DrawArcF dstSurface, forePen, 0!, fxScale, fxScale / 2!, 270!, 91!
                End If
            
            Case ts_Line
                If (x = 0) Then
                    PD2D.DrawLineF dstSurface, forePen, 0!, fxScale / 2!, fxScale / 2!, 0!
                    PD2D.DrawLineF dstSurface, forePen, fxScale / 2!, fxScale, fxScale, fxScale / 2!
                Else
                    PD2D.DrawLineF dstSurface, forePen, fxScale / 2!, 0!, fxScale, fxScale / 2!
                    PD2D.DrawLineF dstSurface, forePen, 0!, fxScale / 2!, fxScale / 2!, fxScale
                End If
            
            Case ts_Maze
                If (x = 0) Then
                    PD2D.DrawLineF dstSurface, forePen, 0!, 0!, fxScale, fxScale
                Else
                    PD2D.DrawLineF dstSurface, forePen, 0!, fxScale, fxScale, 0!
                End If
                
            Case ts_Triangle
                fxPoints(0).x = 0!: fxPoints(0).y = 0!
                fxPoints(1).x = fxScale: fxPoints(1).y = 0!
                fxPoints(2).x = 0!: fxPoints(2).y = fxScale
                    
                Set tmpPath = New pd2DPath
                tmpPath.AddLines 3, VarPtr(fxPoints(0))
                tmpPath.CloseCurrentFigure
                
                'Rotate by [90 * n] degrees
                If (x > 0) Then tmpPath.RotatePathAroundItsCenter 90! * x
                
                PD2D.FillPath dstSurface, foreBrush, tmpPath
                
            Case ts_Octagon
                fxPoints(0).x = 0!: fxPoints(0).y = 0!
                fxPoints(1).x = fxScale: fxPoints(1).y = 0!
                PDMath.RotatePointAroundPoint fxScale, 0!, 0!, 0!, PDMath.DegreesToRadians(45#), fxPoints(2).x, fxPoints(2).y
                fxPoints(3).x = 0!: fxPoints(3).y = fxScale
                    
                Set tmpPath = New pd2DPath
                tmpPath.AddLines 4, VarPtr(fxPoints(0))
                tmpPath.CloseCurrentFigure
                
                'Rotate by [90 * n] degrees
                If (x > 0) Then tmpPath.RotatePathAroundItsCenter 90! * x
                
                PD2D.FillPath dstSurface, foreBrush, tmpPath
                
            Case ts_Circle
                If (x = 0) Then
                    fxPoints(0).x = 0!: fxPoints(0).y = 0!
                ElseIf (x = 1) Then
                    fxPoints(0).x = fxScale: fxPoints(0).y = 0!
                ElseIf (x = 2) Then
                    fxPoints(0).x = fxScale: fxPoints(0).y = fxScale
                Else
                    fxPoints(0).x = 0!: fxPoints(0).y = fxScale
                End If
                
                PD2D.FillCircleI dstSurface, foreBrush, fxPoints(0).x, fxPoints(0).y, fxScale
                
                
        End Select
            
    Next x
    
    Set dstSurface = Nothing
    
    'Source tile images are now ready!
    
    'Figure out how many tiles we'll need to paint in each direction
    Dim numTilesX As Long, numTilesY As Long
    numTilesX = Int(dstDIB.GetDIBWidth / fxScale + 0.5)
    numTilesY = Int(dstDIB.GetDIBHeight / fxScale + 0.5)
    
    'If random number mode is in use, initialize the random number generator
    Dim cRandom As pdRandomize, cLUT() As Byte, tmpDIB As pdDIB
    If (fxPattern = tp_Random) Then
        Set cRandom = New pdRandomize
        cRandom.SetSeed_Float fxRndSeed
        cRandom.SetRndIntegerBounds 0, numImages - 1
    
    'If the pattern should instead be based on the base image, generate a mini version of the
    ' image and produce corresponding lookup tables.
    ElseIf (fxPattern = tp_BaseImage) Then
    
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromExistingDIB dstDIB, numTilesX + 1, numTilesY + 1, GP_IM_HighQualityBilinear
        
        'Convert the mini reference image to grayscale
        DIBs.GetDIBGrayscaleMap tmpDIB, cLUT, True
        Set tmpDIB = Nothing
        
        'Reduce the map to [numImages] discrete values
        Dim scaleFactor As Double
        scaleFactor = (255# / (numImages - 1))
        
        For y = 0 To numTilesY
        For x = 0 To numTilesX
            cLUT(x, y) = Int(cLUT(x, y) / scaleFactor + 0.5)
        Next x
        Next y
        
    End If
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If (modifyProgBarMax = -1) Then SetProgBarMax numTilesY Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim idxSrcImage As Long
    
    'Iterate through each destination tile, painting a matching source tile based on the random number generator
    For y = 0 To numTilesY
    For x = 0 To numTilesX
        
        'Painting is the same regardless of shape, but the system we use for generating the initial lookup table
        ' varies by pattern type.
        Select Case fxPattern
            Case tp_Random
                idxSrcImage = cRandom.GetRandomInt_WH()
            Case tp_BaseImage
                idxSrcImage = cLUT(x, y)
            Case tp_Repeat
                idxSrcImage = x And 1
                If (numImages > 2) Then
                    If (y And 1) Then idxSrcImage = Abs(idxSrcImage - 3)
                End If
            Case tp_Wave
                idxSrcImage = x And 1
                If (numImages > 2) Then
                    If (y And 1) Then idxSrcImage = idxSrcImage + 2
                End If
            Case tp_Quilt
                If (y And 1) Then
                    idxSrcImage = x And 3
                Else
                    idxSrcImage = (x + 2) And 3
                End If
                If (idxSrcImage > 1) Then idxSrcImage = Abs(idxSrcImage - 5)
            Case tp_Chain
                idxSrcImage = x And 1
                If (numImages > 2) Then
                    If (y And 1) Then idxSrcImage = idxSrcImage + 2
                End If
                If (x And 2) Then idxSrcImage = 3 - idxSrcImage
            Case tp_Weave
                If (y And 1) Then
                    idxSrcImage = x And 3
                Else
                    idxSrcImage = (x + 2) And 3
                End If
        End Select
        
        'As a failsafe, mask against the number of images to ensure no overflow
        idxSrcImage = idxSrcImage And (numImages - 1)
        
        GDI.BitBltWrapper dstDIB.GetDIBDC, x * fxScale, y * fxScale, fxScale, fxScale, srcImages(idxSrcImage).GetDIBDC, 0, 0, vbSrcCopy
          
    Next x
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal modifyProgBarOffset + y
            End If
        End If
    Next y
    
    GetTruchetDIB = True
        
End Function
