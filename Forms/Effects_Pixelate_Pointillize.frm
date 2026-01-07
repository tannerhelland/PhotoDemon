VERSION 5.00
Begin VB.Form FormPointillize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Pointillize"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12210
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   814
   Begin PhotoDemon.pdRandomizeUI rndUI 
      Height          =   855
      Left            =   5880
      TabIndex        =   7
      Top             =   4560
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1508
      Caption         =   "random seed"
   End
   Begin PhotoDemon.pdSlider sldOpacity 
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   873
      Max             =   100
      Value           =   100
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdColorSelector csBackground 
      Height          =   855
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1508
      Caption         =   "background color and opacity"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12210
      _ExtentX        =   21537
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdSlider sldSize 
      Height          =   855
      Left            =   5880
      TabIndex        =   4
      Top             =   1680
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1508
      Caption         =   "cell size"
      Min             =   3
      Max             =   1000
      ScaleStyle      =   1
      Value           =   15
      GradientColorRight=   1703935
      NotchValueCustom=   15
   End
   Begin PhotoDemon.pdSlider sldDensity 
      Height          =   855
      Left            =   5880
      TabIndex        =   5
      Top             =   2640
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1508
      Caption         =   "density"
      Min             =   1
      Max             =   100
      SigDigits       =   1
      Value           =   100
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.pdSlider sldRandom 
      Height          =   855
      Left            =   5880
      TabIndex        =   6
      Top             =   3600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   1508
      Caption         =   "color variance"
      Max             =   100
      Value           =   30
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   30
   End
End
Attribute VB_Name = "FormPointillize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Pointillize Effect UI
'Copyright 2019-2026 by Tanner Helland
'Created: 25/November/19
'Last updated: 24/April/25
'Last update: fix preview fit/1:1 zoom toggle not working
'Dependencies: pdPoissonDisc (for generating the initial collection of filter nodes)
'
'PD's Pointillize filter is designed to mimic Photoshop's classic pointillize filter pretty closely
' (with the usual caveat that we give the user much more control over filter settings).  This means
' that a somewhat weird approach is used; a set of random points are generated using the
' pdPoissonDisc class, and those points are then matched against underlying pixels to produce
' small "circles".  Photoshop, however, does not use circles directly - instead, "blob" shapes are
' allowed to form where neighboring points overlap.  This is a more Voronoi-diagram like approach,
' and it means special antialiasing code is required to handling neighboring nodes.
'
'The end result is a high-quality approximation of Photoshop's implementation (and several times
' faster, too!).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'We cache some variables at module-level to improve preview performance
Private m_PreviewDIB As pdDIB

'OK button
Private Sub cmdBar_OKClick()
    Process "Pointillize", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub csBackground_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews until everything is loaded
    cmdBar.SetPreviewStatus False
     
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Custom diffuse effect
' Inputs: diameter in x direction, diameter in y direction, whether or not to wrap edge pixels, and optional preview settings
Public Sub Pointillize(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Splattering canvas with paint..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim BackgroundColor As Long, backgroundOpacity As Single
    Dim cellSize As Single, cellDensity As Single, colorVariation As Single
    Dim rndSeed As String
    
    With cParams
        BackgroundColor = .GetLong("background-color", vbWhite)
        backgroundOpacity = .GetSingle("background-opacity", 100!)
        cellSize = .GetSingle("cell-size", 10!)
        cellDensity = .GetSingle("cell-density", 100!)
        colorVariation = .GetSingle("color-variation", 0!)
        rndSeed = .GetString("random-seed", vbNullString)
    End With
    
    'Density controls the limiting radius that we pass to the poisson disc generator.
    ' As such, HIGHER density equals a LOWER limiting radius (with a minimum value of
    ' 1, where we simply use the input radius as-is).
    cellDensity = cellDensity * 0.01!
    cellDensity = 1! + (1! - cellDensity) * 2!
    
    'Color varies hue alone; as such, rescale it to the range [0, 1]
    colorVariation = colorVariation * 0.01!
    
    'Retrieve a working copy of the image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, , , True
    
    'Prep a new destination DIB with the requested background and color value (premultiplied!)
    backgroundOpacity = backgroundOpacity * 0.01
    Dim rNew As Single, gNew As Single, bNew As Single
    rNew = Colors.ExtractRed(BackgroundColor)
    gNew = Colors.ExtractGreen(BackgroundColor)
    bNew = Colors.ExtractBlue(BackgroundColor)
    
    If (m_PreviewDIB Is Nothing) Then Set m_PreviewDIB = New pdDIB
    m_PreviewDIB.CreateFromExistingDIB workingDIB
    m_PreviewDIB.FillWithColor RGB(rNew, gNew, bNew), backgroundOpacity * 100#
    m_PreviewDIB.SetInitialAlphaPremultiplicationState True
    
    'Rescale cell size to the current preview, as necessary.  Note that the minimum
    ' 2.5 value is used because below this point, antialiased circles start to look
    ' pretty gnarly; so even in previews, we want to limit below this radius.
    If toPreview Then
        cellSize = cellSize * curDIBValues.previewModifier
        If (cellSize < 2.5) Then cellSize = 2.5
    End If
    
    'pd2D is used for rendering
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundPDDIB m_PreviewDIB
    cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
    cSurface.SetSurfacePixelOffset P2_PO_Half
    
    Dim cBrush As pd2DBrush
    Set cBrush = New pd2DBrush
    
    'We'll need to sample pixels from the source image
    Dim srcPixels() As RGBQuad
    workingDIB.WrapRGBQuadArrayAroundDIB srcPixels, dstSA
    
    'pdRandomize handles random number duties
    Dim cRandom As pdRandomize
    Set cRandom = New pdRandomize
    cRandom.SetSeed_String rndSeed
    
    'Use our poisson disc sampler to generate the set of points to paint
    Dim cPoints As pdPoissonDisc
    Set cPoints = New pdPoissonDisc
    
    Dim listOfPoints() As PointFloat, numOfPoints As Long
    Dim ptGrid() As Long, gridWidth As Long, gridHeight As Long
    
    Dim limitRadius As Single
    limitRadius = cellSize * Sqr(2#) * cellDensity
    cPoints.GetDisc listOfPoints, numOfPoints, ptGrid, gridWidth, gridHeight, limitRadius, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight
    
    'Fast lut for opacity translations
    Dim i As Long, j As Long, lutOpacity() As Byte
    ReDim lutOpacity(0 To 255) As Byte
    For i = 0 To 255
        lutOpacity(i) = i * (100# / 255#)
    Next i
    
    'Build a list of colors for all points in the list; this way, we don't have to dive into the source
    ' image for sampling on each point.
    Dim listOfColors() As RGBQuad
    ReDim listOfColors(0 To numOfPoints - 1) As RGBQuad
    
    Dim tmpH As Double, tmpS As Double, tmpL As Double, tmpR As Double, tmpG As Double, tmpB As Double
    Const ONE_DIV_255 As Double = 1# / 255#
    
    For i = 0 To numOfPoints - 1
        
        listOfColors(i) = srcPixels(Int(listOfPoints(i).x), Int(listOfPoints(i).y))
        
        'Randomize, as required by user settings
        If (colorVariation > 0!) Then
            
            tmpR = listOfColors(i).Red * ONE_DIV_255
            tmpG = listOfColors(i).Green * ONE_DIV_255
            tmpB = listOfColors(i).Blue * ONE_DIV_255
            
            Colors.fRGBtoHSV tmpR, tmpG, tmpB, tmpH, tmpS, tmpL
            tmpH = tmpH + (cRandom.GetRandomFloat_WH() - 0.5) * colorVariation
            Colors.fHSVtoRGB tmpH, tmpS, tmpL, tmpR, tmpG, tmpB
            
            listOfColors(i).Red = Int(tmpR * 255# + 0.5)
            listOfColors(i).Green = Int(tmpG * 255# + 0.5)
            listOfColors(i).Blue = Int(tmpB * 255# + 0.5)
            
        End If
        
    Next i
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim uiUpdate As Long, idxPoint As Long, srcColor As RGBQuad
    
    'Two branches; we currently use a Photoshop-imitation algorithm that requires specialized
    ' handling of antialiasing (due to the way it handles neighboring cells).  You can deactivate
    ' this flag to simply "splatter" all points against the canvas.
    Dim usePSStyle As Boolean
    usePSStyle = True
    
    If usePSStyle Then
        
        'For the PS-style pointillize, we basically need to construct a Voronoi map;
        ' each pixel gets mapped to its nearest active point.
        Dim x As Long, y As Long, xFinal As Long, yFinal As Long
        xFinal = workingDIB.GetDIBWidth - 1
        yFinal = workingDIB.GetDIBHeight - 1
        
        If (Not toPreview) Then ProgressBars.SetProgBarMax yFinal
        uiUpdate = ProgressBars.FindBestProgBarValue()
        
        'Use the same grid box limiter as the poisson disc class
        Dim cellBoxSize As Long
        cellBoxSize = Int(limitRadius / Sqr(2#))
        
        'Build a quick LUT for x/y grid indices; this avoids divisions on the inner loop
        Dim xLutIndex() As Long, yLutIndex() As Long
        ReDim xLutIndex(0 To xFinal) As Long
        ReDim yLutIndex(0 To yFinal) As Long
        For x = 0 To xFinal
            xLutIndex(x) = x \ cellBoxSize
        Next x
        For y = 0 To yFinal
            yLutIndex(y) = y \ cellBoxSize
        Next y
        
        'Boundary checking *will* be required on the inner loop, unfortunately
        Dim maxGridWidth As Long, maxGridHeight As Long
        maxGridWidth = gridWidth - 1
        maxGridHeight = gridHeight - 1
        
        'Colors may need to be blended as part of antialiasing; we also need to track
        ' current "best-match" point and the second-best point (for antialiasing purposes),
        ' when a point lies evenly between two control points.  Note that we also generate
        ' some variables purely for perf reasons - like precalculating inverses, so we can
        ' avoid divisions on the inner loop.
        Dim gL As Long, gR As Long, gU As Long, gD As Long
        Dim nearestIndex As Long, nearestDist As Double, testDist As Double
        Dim nearestIndex2 As Long, nearestDist2 As Double
        Dim cellSizeSquared As Double, cellSizeAA As Double, cellSizeAADiff As Double, invCellSizeAADiff As Double
        cellSizeSquared = cellSize * cellSize
        cellSizeAA = (cellSize - 1) * (cellSize - 1)
        cellSizeAADiff = (cellSizeSquared - cellSizeAA)
        invCellSizeAADiff = 1# / cellSizeAADiff
        
        'Calculate similar values for points that lie very nearly in-between two control points;
        ' these are antialiased slightly stronger (as it makes the end-result closer to PS's)
        Dim cellSizeAADiff2 As Double, invCellSizeAADiff2 As Double, blendRatio As Double
        cellSizeAADiff2 = cellSizeAADiff * 1.5
        invCellSizeAADiff2 = 1# / cellSizeAADiff2
        
        'Pixels will only be acessed one scanline at a time, so a 1D array suffices
        Dim dstPixels() As RGBQuad, dstSA1D As SafeArray1D
        
        For y = 0 To yFinal
            m_PreviewDIB.WrapRGBQuadArrayAroundScanline dstPixels, dstSA1D, y
        For x = 0 To xFinal
            
            'Search a 3x3 grid (if possible; edge pixels may search less) for the
            ' nearest and 2nd-nearest grid points.
            
            'Calculate inner loop bounds
            gL = xLutIndex(x) - 1
            gR = gL + 2
            gU = yLutIndex(y) - 1
            gD = gU + 2
            If (gL < 0) Then gL = 0
            If (gR > maxGridWidth) Then gR = maxGridWidth
            If (gU < 0) Then gU = 0
            If (gD > maxGridHeight) Then gD = maxGridHeight
            
            'Reset nearest distance trackers
            nearestDist = DOUBLE_MAX
            nearestDist2 = DOUBLE_MAX
            
            'Search grid
            For i = gL To gR
            For j = gU To gD
                
                idxPoint = ptGrid(i, j)
                
                'Make sure cell is not empty (-1 is used to flag empty cells)
                If (idxPoint >= 0) Then
                    testDist = PDMath.DistanceTwoPointsShortcut(x, y, listOfPoints(idxPoint).x, listOfPoints(idxPoint).y)
                    If (testDist < nearestDist) Then
                        nearestDist2 = nearestDist
                        nearestIndex2 = nearestIndex
                        nearestDist = testDist
                        nearestIndex = idxPoint
                    End If
                End If
                
            Next j
            Next i
            
            'See if a point from the collection is within the radius of this point
            If (nearestDist <= cellSizeSquared) Then
                
                'A point is nearby!
                
                'See if edge AA is required for this pixel
                If (nearestDist <= cellSizeAA) Then
                    
                    'See if it overlaps a neighboring cell; if it does, we need to antialias it against
                    ' *that* cell's color.
                    blendRatio = (nearestDist2 - nearestDist)
                    If (blendRatio > 0#) And (blendRatio < cellSizeAADiff2) Then
                        blendRatio = blendRatio * invCellSizeAADiff2
                        dstPixels(x).Blue = Colors.BlendColors(listOfColors(nearestIndex2).Blue, listOfColors(nearestIndex).Blue, blendRatio)
                        dstPixels(x).Green = Colors.BlendColors(listOfColors(nearestIndex2).Green, listOfColors(nearestIndex).Green, blendRatio)
                        dstPixels(x).Red = Colors.BlendColors(listOfColors(nearestIndex2).Red, listOfColors(nearestIndex).Red, blendRatio)
                        dstPixels(x).Alpha = Colors.BlendColors(listOfColors(nearestIndex2).Alpha, listOfColors(nearestIndex).Alpha, blendRatio)
                    
                    'No neighboring control points matter; render this color as-is
                    Else
                        dstPixels(x) = listOfColors(nearestIndex)
                    End If
                
                'This pixel lies in the outermost 1px of this shape.  Antialias it against
                ' the background color.
                Else

                    blendRatio = (nearestDist - cellSizeAA) * invCellSizeAADiff
                    dstPixels(x).Blue = Colors.BlendColors(listOfColors(nearestIndex).Blue, dstPixels(x).Blue, blendRatio)
                    dstPixels(x).Green = Colors.BlendColors(listOfColors(nearestIndex).Green, dstPixels(x).Green, blendRatio)
                    dstPixels(x).Red = Colors.BlendColors(listOfColors(nearestIndex).Red, dstPixels(x).Red, blendRatio)
                    dstPixels(x).Alpha = Colors.BlendColors(listOfColors(nearestIndex).Alpha, dstPixels(x).Alpha, blendRatio)
                    
                End If
            
            '/end pointillize point is near this pixel
            End If
            
        Next x
            If toPreview Then
                If ((y And uiUpdate) = 0) Then ProgressBars.SetProgBarVal y
            End If
        Next y
        
        m_PreviewDIB.UnwrapRGBQuadArrayFromDIB dstPixels
    
    'A non-Photoshop approach would just "splatter" points against the canvas.
    ' This block of code does just that (and is much faster than the PS algorithm).
    Else
        
        If (Not toPreview) Then ProgressBars.SetProgBarMax numOfPoints
        uiUpdate = ProgressBars.FindBestProgBarValue()
        
        'Paint all points in random order
        Do While (numOfPoints > 0)
            
            'Generate a random point index
            idxPoint = cRandom.GetRandomIntRange_WH(0, numOfPoints - 1)
            
            'Paint said point
            srcColor = listOfColors(idxPoint)
            cBrush.SetBrushColor RGB(srcColor.Red, srcColor.Green, srcColor.Blue)
            cBrush.SetBrushOpacity lutOpacity(srcColor.Alpha)
            
            PD2D.FillCircleF cSurface, cBrush, listOfPoints(idxPoint).x, listOfPoints(idxPoint).y, cellSize
            
            'Remove this point from the collection
            listOfPoints(idxPoint) = listOfPoints(numOfPoints - 1)
            numOfPoints = numOfPoints - 1
            
            'Update UI periodically
            If toPreview Then
                If ((i And uiUpdate) = 0) Then ProgressBars.SetProgBarVal i
            End If
            
        Loop
        
    End If
    
    'Safely deallocate all image arrays
    workingDIB.UnwrapRGBQuadArrayFromDIB srcPixels
    workingDIB.CreateFromExistingDIB m_PreviewDIB
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
     
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.Pointillize GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "background-color", csBackground.Color
        .AddParam "background-opacity", sldOpacity.Value
        .AddParam "cell-size", sldSize.Value
        .AddParam "cell-density", sldDensity.Value
        .AddParam "color-variation", sldRandom.Value
        .AddParam "random-seed", rndUI.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Sub rndUI_Change()
    UpdatePreview
End Sub

Private Sub sldDensity_Change()
    UpdatePreview
End Sub

Private Sub sldOpacity_Change()
    UpdatePreview
End Sub

Private Sub sldRandom_Change()
    UpdatePreview
End Sub

Private Sub sldSize_Change()
    UpdatePreview
End Sub
