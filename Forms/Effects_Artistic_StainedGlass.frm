VERSION 5.00
Begin VB.Form FormStainedGlass 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Stained glass"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11775
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
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   785
   Visible         =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6030
      Width           =   11775
      _ExtentX        =   21325
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5745
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdButtonStrip btsOptions 
      Height          =   1080
      Left            =   6000
      TabIndex        =   2
      Top             =   4800
      Width           =   5595
      _ExtentX        =   10504
      _ExtentY        =   1905
      Caption         =   "options"
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   4695
      Index           =   0
      Left            =   5880
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7858
      Begin PhotoDemon.pdDropDown cboPattern 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   1296
         Caption         =   "pattern"
      End
      Begin PhotoDemon.pdSlider sltSize 
         Height          =   705
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   5535
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "cell size"
         Min             =   3
         Max             =   200
         Value           =   20
         DefaultValue    =   20
      End
      Begin PhotoDemon.pdSlider sltShadeQuality 
         Height          =   705
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   5535
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "shading quality"
         Min             =   1
         Max             =   6
         Value           =   6
         NotchPosition   =   2
         NotchValueCustom=   6
      End
      Begin PhotoDemon.pdSlider sltTurbulence 
         Height          =   705
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   5535
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "turbulence"
         Max             =   1
         SigDigits       =   2
         Value           =   0.5
         DefaultValue    =   0.5
      End
   End
   Begin PhotoDemon.pdContainer pnlOptions 
      Height          =   4695
      Index           =   1
      Left            =   5880
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7858
      Begin PhotoDemon.pdDropDown cboDistance 
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   5535
         _ExtentX        =   10398
         _ExtentY        =   1296
         Caption         =   "distance method"
      End
      Begin PhotoDemon.pdSlider sltEdge 
         Height          =   705
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   5535
         _ExtentX        =   10398
         _ExtentY        =   1270
         Caption         =   "edge thickness"
         Max             =   1
         SigDigits       =   2
      End
      Begin PhotoDemon.pdDropDown cboColorSampling 
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   5535
         _ExtentX        =   10398
         _ExtentY        =   1296
         Caption         =   "color sampling"
      End
      Begin PhotoDemon.pdSlider sldOpacity 
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   3960
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   873
         Max             =   100
         Value           =   100
         GradientColorRight=   1703935
         NotchPosition   =   2
         NotchValueCustom=   100
      End
      Begin PhotoDemon.pdColorSelector csBackground 
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   1508
         Caption         =   "edge color and opacity"
         curColor        =   0
      End
   End
End
Attribute VB_Name = "FormStainedGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Stained Glass Effect Interface
'Copyright 2014-2026 by Tanner Helland
'Created: 14/July/14
'Last updated: 07/December/20
'Last update: 2x performance improvements! yay!
'Dependencies: pdVoronoi, pdRandomize
'
'PhotoDemon's crystallize effect is implemented using Worley Noise...
' (https://en.wikipedia.org/wiki/Worley_noise)
' ...which is basically a special algorithmic approach to Voronoi diagrams...
' (https://en.wikipedia.org/wiki/Voronoi_diagram)
'
'The associated pdVoronoi class does most the heavy lifting for this effect.  The main
' fxStainedGlass function basically forwards all relevant parameters to a pdVoronoi instance,
' applies a first pass over the image, caching matching Voronoi indices as it goes, then uses
' those indices in a second pass to recolor the image.
'
'Parameters are currently available for a number of tweaks; these will be refined further as
' the tool nears completion.  (As a warning, some methods may be dropped in the interest of
' simplifying the dialog.)
'
'Finally, note that multiple lookup tables are used to improve the performance of this function.
' While these may seem excessive, the fact that we can produce the entire effect without copying
' the current image is pretty awesome, so despite the many lookup tables, this actually uses
' less RAM than many other effects in PD.
'
'Finally, many thanks to Robert Rayment, who did extensive profiling and research on various
' Voronoi implementations before I started work on this class.  His comments were invaluable in
' determining the shape and style of this class's interface.  (FYI, Robert's PaintRR app has a
' much simpler approach to this filter - check it out if PD's method seems like overkill!
' Link here: http://rrprogs.com/)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To make sure the function looks similar in the preview and final image, we cache the random seed used
Private m_Random As pdRandomize

'Voronoi patterns have a non-zero initialization cost.  It can be minimized by reusing
' a persistent instance
Private m_Voronoi As pdVoronoi

'Background color/opacity is handled by a basic merge op
Private m_BackgroundDIB As pdDIB

'Apply a Stained Glass effect to an image
' Inputs:
'  cellSize = size, in pixels, of each initial grid box in the Voronoi array.  Do not make this less than 3.
'  fxTurbulence = how much to distort cell shape, range [0, 1], 0 = perfect grid
'  colorSamplingMethod = how to determine cell color (0 = just use pixel at Voronoi point, 1 = average all pixels in cell)
'  shadeQuality = how detailed to shade each cell (1 = flat, 5 = detailed non-linear depth rendering)
'  distanceMethod = 0 - Cartesian, 1 - Manhattan, 2 - Chebyshev
Public Sub fxStainedGlass(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Carving image from stained glass..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim cellSize As Long, shadeQuality As Long, colorSamplingMethod As Long, distanceMethod As Long
    Dim fxTurbulence As Double, edgeThreshold As Double, fxPatternName As String
    Dim newBackColor As Long, newBackOpacity As Single
    
    With cParams
        cellSize = .GetLong("size", sltSize.Value)
        fxTurbulence = .GetDouble("turbulence", sltTurbulence.Value)
        colorSamplingMethod = .GetLong("color", cboColorSampling.ListIndex)
        shadeQuality = .GetLong("shading", sltShadeQuality.Value)
        edgeThreshold = .GetDouble("edges", sltEdge.Value)
        distanceMethod = .GetLong("distance", cboDistance.ListIndex)
        fxPatternName = .GetString("pattern", "square")
        newBackColor = .GetLong("background-color", vbBlack)
        newBackOpacity = .GetSingle("background-opacity", 100!)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D, dstSA1D As SafeArray1D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'Because this is a two-pass filter, we have to manually change the progress bar maximum to 2x
    If (Not toPreview) Then
        ProgressBars.SetProgBarMax finalY * 2
    
    'If this is a preview, reduce cell size to better portray how the final image will look
    Else
        cellSize = cellSize * curDIBValues.previewModifier
        If (cellSize < 3) Then cellSize = 3
    End If
    
    'Failsafe check for cellsize
    Dim minDimension As Long
    If (curDIBValues.Width < curDIBValues.Height) Then minDimension = curDIBValues.Width Else minDimension = curDIBValues.Height
    minDimension = (minDimension - 1) \ 2
    If (cellSize >= minDimension) Then cellSize = minDimension
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Create a Voronoi class to help us with processing; it does all the messy Voronoi work for us.
    If (m_Voronoi Is Nothing) Then Set m_Voronoi = New pdVoronoi
    
    'Pass all meaningful input parameters on to the Voronoi class
    m_Voronoi.InitPoints cellSize, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight
    m_Voronoi.RandomizePoints fxTurbulence, m_Random.GetSeed
    m_Voronoi.SetDistanceMode distanceMethod
    m_Voronoi.SetShadingMode shadeQuality
    m_Voronoi.SetInitialPattern m_Voronoi.GetPatternIDFromName(fxPatternName)
    m_Voronoi.SetDensity 1#
    m_Voronoi.FinalizeParameters
    
    'Finally, we will also make two image-sized look-up tables that store the nearest
    ' and second-nearest Voronoi point index for each pixel in the image.  While this
    ' consumes a lot of memory, it makes our second pass through the image
    ' (the recoloring pass) much faster than it would otherwise be.  (Note also that
    ' it greatly improves data locality to store these points next to each other instead
    ' of in separate arrays - in this case, I just use the x/y members of a standard
    ' int-point-struct)
    Dim vLookup() As PointLong
    ReDim vLookup(initX To finalX, initY To finalY) As PointLong
    
    'We need a list of Voronoi points in the image; colors corresponding to each point
    ' may need to be average together to arrive at a "final" color for each cell
    Dim numVoronoiPoints As Long
    numVoronoiPoints = m_Voronoi.GetTotalNumOfVoronoiPoints() - 1
    
    'Create several look-up tables, specifically:
    ' One table for each color channel (RGBA)
    ' One table for number of pixels in each Voronoi cell
    Dim rLookup() As Long, gLookup() As Long, bLookup() As Long, aLookup() As Long
    ReDim rLookup(0 To numVoronoiPoints) As Long
    ReDim gLookup(0 To numVoronoiPoints) As Long
    ReDim bLookup(0 To numVoronoiPoints) As Long
    ReDim aLookup(0 To numVoronoiPoints) As Long
    
    Dim numPixels() As Long
    ReDim numPixels(0 To numVoronoiPoints) As Long
    
    'Color values must be individually processed to account for shading, so we need to declare them
    Dim r As Long, g As Long, b As Long, a As Long
    
    'To support variable background color/opacity, we apply all effects to a second image copy,
    ' then merge the result onto the specified background RGBA
    If (m_BackgroundDIB Is Nothing) Then Set m_BackgroundDIB = New pdDIB
    m_BackgroundDIB.CreateFromExistingDIB workingDIB
    
    'The Voronoi approach we use requires knowledge of the distance to the nearest Voronoi point, and depending on
    ' shading quality, distance to the second-nearest point as well.
    Dim nearestPoint As Long, secondNearestPoint As Long
    
    'Loop through each pixel in the image, calculating nearest Voronoi points as we go
    For y = initY To finalY
        m_BackgroundDIB.WrapArrayAroundScanline dstImageData, dstSA1D, y
    For x = initX To finalX
        xStride = x * 4
        
        'Use the Voronoi class to find the nearest points to this pixel
        nearestPoint = m_Voronoi.GetNearestPointIndex(x, y, secondNearestPoint)
        
        'Store the nearest and second-nearest point indices in our central lookup table
        vLookup(x, y).x = nearestPoint
        vLookup(x, y).y = secondNearestPoint
        
        'If the user has elected to recolor each cell using the average color for the cell,
        ' we need to track color values.  This is no different from a histogram approach,
        ' except in this case, each histogram bucket corresponds to one Voronoi cell.
        If (colorSamplingMethod = 0) Then
        
            'Retrieve RGBA values for this pixel
            b = dstImageData(xStride)
            g = dstImageData(xStride + 1)
            r = dstImageData(xStride + 2)
            a = dstImageData(xStride + 3)
            
            'Store those RGBA values into their respective lookup "bin"
            rLookup(nearestPoint) = rLookup(nearestPoint) + r
            gLookup(nearestPoint) = gLookup(nearestPoint) + g
            bLookup(nearestPoint) = bLookup(nearestPoint) + b
            aLookup(nearestPoint) = aLookup(nearestPoint) + a
            
            'Increment the count of all pixels who share this Voronoi point as their nearest point
            numPixels(nearestPoint) = numPixels(nearestPoint) + 1
            
        End If
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    m_BackgroundDIB.UnwrapArrayFromDIB dstImageData
    
    'All lookup tables are now properly initialized.  Depending on the user's color sampling choice, calculate
    ' cell colors now.
    Dim numPixelsCache As Long, thisPoint As PointFloat
    m_BackgroundDIB.WrapArrayAroundDIB dstImageData, dstSA
    
    For x = 0 To numVoronoiPoints
    
        Select Case colorSamplingMethod
        
            'Accurate
            Case 0
                
                'The user wants us to find the average color for each cell.
                ' This is effectively just a blur operation; for each bin in the lookup table,
                ' divide the total RGBA values by the number of pixels in that bin.
                numPixelsCache = numPixels(x)
                
                If (numPixelsCache > 0) Then
                    rLookup(x) = rLookup(x) \ numPixelsCache
                    gLookup(x) = gLookup(x) \ numPixelsCache
                    bLookup(x) = bLookup(x) \ numPixelsCache
                    aLookup(x) = aLookup(x) \ numPixelsCache
                End If
            
            'Fast
            Case 1
                
                'The user wants a "fast and dirty" approach to coloring.
                ' For each cell, use only the color of the Voronoi point pixel.
                
                'Retrieve the location of this Voronoi point
                thisPoint = m_Voronoi.GetVoronoiCoordinates(x)
                
                'Validate its bounds
                If (thisPoint.x < initX) Then thisPoint.x = initX
                If (thisPoint.x > finalX) Then thisPoint.x = finalX
                
                If (thisPoint.y < initX) Then thisPoint.y = initY
                If (thisPoint.y > finalY) Then thisPoint.y = finalY
                
                'Retrieve the color at this Voronoi point's location, and assign it to the lookup arrays
                xStride = Int(thisPoint.x + 0.5!) * 4
                y = Int(thisPoint.y + 0.5!)
                bLookup(x) = dstImageData(xStride, y)
                gLookup(x) = dstImageData(xStride + 1, y)
                rLookup(x) = dstImageData(xStride + 2, y)
                aLookup(x) = dstImageData(xStride + 3, y)
            
            'Random
            Case 2
                rLookup(x) = m_Random.GetRandomIntRange_WH(0, 255)
                gLookup(x) = m_Random.GetRandomIntRange_WH(0, 255)
                bLookup(x) = m_Random.GetRandomIntRange_WH(0, 255)
                aLookup(x) = 255
        
        End Select
    
    Next x
    
    m_BackgroundDIB.UnwrapArrayFromDIB dstImageData
    
    'Our pixel count cache is now unneeded; free it
    Erase numPixels
    
    'Shading requires a number of specialized variables
    Dim shadeAdjustment As Single, shadeThreshold As Single, edgeAdjustment As Single, maxDistance As Single
            
    'Loop through the image, changing colors to match our calculated Voronoi values
    For y = initY To finalY
        m_BackgroundDIB.WrapArrayAroundScanline dstImageData, dstSA1D, y
    For x = initX To finalX
        
        xStride = x * 4
        
        'Use the lookup table from step 1 to find the nearest and second-nearest Voronoi point indices for this pixel.
        ' (NOTE: this step could be rewritten to simply re-request a distance calculation from the Voronoi class,
        '        but that would slow the function considerably.)
        nearestPoint = vLookup(x, y).x
        secondNearestPoint = vLookup(x, y).y
        
        'Retrieve the RGB values from the relevant Voronoi cell bin
        b = bLookup(nearestPoint)
        r = rLookup(nearestPoint)
        g = gLookup(nearestPoint)
        a = aLookup(nearestPoint)
        
        'If the user is using a custom edge value, we need to perform a number of extra calculations.  If they are
        ' just doing a generic filter, however, we can greatly shortcut the process.
        If (edgeThreshold = 0) Then
        
            If (shadeQuality <> vs_NoShade) Then
                
                'Retrieve a shade value on the scale [0, 1] from the Voronoi class; it will calculate this
                ' value using the relationship between this point's distance to the nearest Voronoi point,
                ' and the maximum shading value for this cell.
                shadeAdjustment = m_Voronoi.GetShadingValue(x, y, nearestPoint, secondNearestPoint)
                
                'Modify the alpha value for this pixel by the retrieved shading adjustment
                a = a * shadeAdjustment
                
            End If
        
        'The user has modified the edge threshold.  Break out your mathbook!
        Else
            
            'We will now proceed to calculate and apply an edge modification on top of the pixel's existing shading value.
            ' Basically, the edge parameter controls an artificial fade for pixels whose distances fall below the edge
            ' threshold value.  Pixels above the threshold value are untouched (meaning they receive only their default
            ' shading adjustment).
            
            'We can shortcut the edge calculation process for the basic, non-shade method.
            If (shadeQuality = vs_NoShade) Then
                shadeAdjustment = 1!
            Else
                shadeAdjustment = m_Voronoi.GetShadingValue(x, y, nearestPoint, secondNearestPoint)
            End If
            
            'Retrieve the maximum distance for this Voronoi cell, and use that to calculate a cell threshold value.
            maxDistance = m_Voronoi.GetMaxDistanceForCell(nearestPoint)
            shadeThreshold = (edgeThreshold * maxDistance) + 0.000001!  'DBZ workaround
            
            'Different shading methods require different calculations to make the edge algorithm work similarly.
            ' Sort by shade method, and calculate only a relevant edge adjustment value.
            If (shadeQuality < vs_ShadeF2MinusF1) Then
                edgeAdjustment = maxDistance - m_Voronoi.GetDistance(x, y, nearestPoint)
            Else
                
                If (shadeQuality = vs_ShadeF2MinusF1) Then
                    edgeAdjustment = (m_Voronoi.GetDistance(x, y, secondNearestPoint) - m_Voronoi.GetDistance(x, y, nearestPoint))
                Else
                    edgeAdjustment = shadeAdjustment
                End If
                
            End If
            
            'If our calculated adjustment falls below the shading threshold we calculated,
            ' this pixel is a candidate for edge enhancement.
            If (edgeAdjustment < shadeThreshold) Then
                edgeAdjustment = edgeAdjustment / shadeThreshold
                
                'To provide a slightly better look, we actually use an n^2 fall-off instead of a linear one
                shadeAdjustment = shadeAdjustment * edgeAdjustment * edgeAdjustment
                
                'To avoid potential overflow errors, make sure our edge parameter only shrinks RGB values.
                ' (This case should never occur, but given the number of parameters at play in this tool,
                '  it doesn't hurt to exert a little extra caution!)
                If (shadeAdjustment > 1!) Then shadeAdjustment = 1!
                
            End If
            
            'With our shade adjustment finalized, we can finally calculate a final
            ' alpha value for this pixel.
            If (shadeAdjustment < 0!) Then shadeAdjustment = 0!
            a = a * shadeAdjustment
            
        End If
        
        'Set the new RGBA values to the image
        dstImageData(xStride) = b
        dstImageData(xStride + 1) = g
        dstImageData(xStride + 2) = r
        dstImageData(xStride + 3) = a
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal finalY + y
            End If
        End If
    Next y
    
    m_BackgroundDIB.UnwrapArrayFromDIB dstImageData
    
    'Fill workingDIB with the requested background color, then premultiply everything
    workingDIB.FillWithColor newBackColor, newBackOpacity
    If (Not workingDIB.GetAlphaPremultiplication) Then workingDIB.SetAlphaPremultiplication True
    m_BackgroundDIB.SetAlphaPremultiplication True
    
    'Perform the final blend
    m_BackgroundDIB.AlphaBlendToDC workingDIB.GetDIBDC
    
    'For fun, you can uncomment the code block below to render the calculated Voronoi points onto the image.
'    For x = 0 To numVoronoiPoints
'        thisPoint = m_Voronoi.GetVoronoiCoordinates(x)
'        GDIPlusDrawCircleToDC workingDIB.GetDIBDC, thisPoint.x, thisPoint.y, 2, RGB(255, 0, 255)
'    Next x
        
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub btsOptions_Click(ByVal buttonIndex As Long)
    UpdatePanelVisibility
End Sub

Private Sub cboColorSampling_Click()
    UpdatePreview
End Sub

Private Sub cboDistance_Click()
    UpdatePreview
End Sub

Private Sub cboPattern_Click()
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Stained glass", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    csBackground.Color = vbBlack
End Sub

Private Sub csBackground_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    Set m_Voronoi = New pdVoronoi
    
    'Disable previews until the dialog is fully initialized
    cmdBar.SetPreviewStatus False
    
    'Initial pattern options
    cboPattern.Clear
    Dim i As Long
    For i = 0 To m_Voronoi.GetPatternCount - 1
        cboPattern.AddItem m_Voronoi.GetPatternUINameFromID(i), i
    Next i
    cboPattern.ListIndex = 0
    
    'Provide with user with several color sampling options
    cboColorSampling.Clear
    cboColorSampling.AddItem "accurate"
    cboColorSampling.AddItem "fast"
    cboColorSampling.AddItem "random"
    cboColorSampling.ListIndex = 0
        
    'Provide three experimental distance functions
    cboDistance.Clear
    cboDistance.AddItem "Cartesian (traditional)"
    cboDistance.AddItem "Manhattan (walking)"
    cboDistance.AddItem "Chebyshev (chessboard)"
    cboDistance.ListIndex = 0
    
    'Calculate a random turbulence seed
    Set m_Random = New pdRandomize
    m_Random.SetSeed_AutomaticAndRandom
        
    'Activate only one options panel (basic/advanced)
    btsOptions.AddItem "basic", 0
    btsOptions.AddItem "advanced", 1
    btsOptions.ListIndex = 0
    UpdatePanelVisibility
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    
    'Request a preview
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = pnlOptions.lBound To pnlOptions.UBound
        pnlOptions(i).Visible = (btsOptions.ListIndex = i)
    Next i
End Sub

'Redraw the effect preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then fxStainedGlass GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sldOpacity_Change()
    UpdatePreview
End Sub

Private Sub sltEdge_Change()
    UpdatePreview
End Sub

Private Sub sltShadeQuality_Change()
    UpdatePreview
End Sub

Private Sub sltSize_Change()
    UpdatePreview
End Sub

Private Sub sltTurbulence_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "size", sltSize.Value
        
        'Shape options use string values for future-proofed expansion possibilities
        If (m_Voronoi Is Nothing) Then Set m_Voronoi = New pdVoronoi
        .AddParam "pattern", m_Voronoi.GetPatternNameFromID(cboPattern.ListIndex)
        
        .AddParam "turbulence", sltTurbulence.Value
        .AddParam "color", cboColorSampling.ListIndex
        .AddParam "shading", sltShadeQuality.Value
        .AddParam "edges", sltEdge.Value
        .AddParam "distance", cboDistance.ListIndex
        .AddParam "background-color", csBackground.Color
        .AddParam "background-opacity", sldOpacity.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
