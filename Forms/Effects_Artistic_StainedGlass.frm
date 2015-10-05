VERSION 5.00
Begin VB.Form FormStainedGlass 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Stained glass"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ComboBox cboDistance 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5205
      Width           =   5775
   End
   Begin VB.ComboBox cboColorSampling 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4260
      Width           =   5775
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltSize 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "cell size"
      Min             =   3
      Max             =   200
      Value           =   50
   End
   Begin PhotoDemon.sliderTextCombo sltTurbulence 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "turbulence"
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.sliderTextCombo sltEdge 
      Height          =   720
      Left            =   6000
      TabIndex        =   8
      Top             =   3000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "edge thickness"
      Max             =   1
      SigDigits       =   2
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin PhotoDemon.sliderTextCombo sltShadeQuality 
      Height          =   720
      Left            =   6000
      TabIndex        =   9
      Top             =   2040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "shading quality"
      Min             =   1
      Max             =   5
      Value           =   5
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "distance method"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   5
      Left            =   6000
      TabIndex        =   7
      Top             =   4845
      Width           =   1740
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "color sampling"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   4
      Top             =   3900
      Width           =   1560
   End
End
Attribute VB_Name = "FormStainedGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Stained Glass Effect Interface
'Copyright 2014-2015 by Tanner Helland
'Created: 14/July/14
'Last updated: 23/June/15
'Last update: move randomize responsibilities over to pdRandomize
'Dependencies: pdRandomize
'
'PhotoDemon's stained glass effect is implemented using Worley Noise (http://en.wikipedia.org/wiki/Worley_noise),
' which is basically a special algorithmic approach to Voronoi diagrams (http://en.wikipedia.org/wiki/Voronoi_diagram).
'
'The associated pdVoronoi class does most the heavy lifting for this effect.  The main fxStainedGlass function basically
' forwards all relevant parameters to a pdVoronoi instance, applies a first pass over the image, caching matching
' Voronoi indices as it goes, then using those indices in a second pass to recolor the image.
'
'Parameters are currently available for a number of tweaks; these will be refined further as the tool nears completion.
' (As a warning, some methods may be dropped in the interest of simplifying the dialog.)
'
'Finally, note that multiple lookup tables are used to improve the performance of this function.  While these may
' seem excessive, the fact that we can produce the entire effect without copying the current image is pretty awesome,
' so despite the many lookup tables, this actually uses less RAM than many other effects in PD.
'
'I should give a special thanks to Robert Rayment, who did extensive profiling and research on this filter before I ever
' began work on it.  His comments were invaluable in determining the shape and style of this class.  FYI, Robert's PaintRR
' project has a faster and simpler version of this routine worth checking out if PD's methods seem like overkill!
' Link here: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=66991&lngWId=1
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'To make sure the function looks similar in the preview and final image, we cache the random seed used
Private cRandom As pdRandomize

'Apply a Stained Glass effect to an image
' Inputs:
'  cellSize = size, in pixels, of each initial grid box in the Voronoi array.  Do not make this less than 3.
'  fxTurbulence = how much to distort cell shape, range [0, 1], 0 = perfect grid
'  colorSamplingMethod = how to determine cell color (0 = just use pixel at Voronoi point, 1 = average all pixels in cell)
'  shadeQuality = how detailed to shade each cell (1 = flat, 5 = detailed non-linear depth rendering)
'  distanceMethod = 0 - Cartesian, 1 - Manhattan, 2 - Chebyshev
Public Sub fxStainedGlass(ByVal cellSize As Long, ByVal fxTurbulence As Double, ByVal colorSamplingMethod As Long, ByVal shadeQuality As Long, ByVal edgeThreshold As Double, ByVal distanceMethod As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Carving image from stained glass..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'Because this is a two-pass filter, we have to manually change the progress bar maximum to 2 * width
    If Not toPreview Then
        SetProgBarMax finalX * 2
    
    'If this is a preview, reduce cell size to better portray how the final image will look
    Else
        cellSize = cellSize * curDIBValues.previewModifier
        If cellSize < 3 Then cellSize = 3
    End If
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Create a Voronoi class to help us with processing; it does all the messy Voronoi work for us.
    Dim cVoronoi As pdVoronoi
    Set cVoronoi = New pdVoronoi
    
    'Pass all meaningful input parameters on to the Voronoi class
    cVoronoi.initPoints cellSize, workingDIB.getDIBWidth, workingDIB.getDIBHeight
    cVoronoi.randomizePoints fxTurbulence, cRandom.getSeed
    cVoronoi.setDistanceMode distanceMethod
    cVoronoi.setShadingMode shadeQuality
    
    'Create several look-up tables, specifically:
    ' One table for each color channel (RGBA)
    ' One table for number of pixels in each Voronoi cell
    Dim rLookup() As Long, gLookUp() As Long, bLookup() As Long, aLookup() As Long
    Dim numPixels() As Long
    
    'Finally, we will also make two image-sized look-up tables that store the nearest Voronoi point index for
    ' each pixel in the image, and if certain shading types are active, the second-nearest Voronoi point as well.
    ' While this consumes a lot of memory, it makes our second pass through the image (the recoloring pass) much
    ' faster than it would otherwise be.
    Dim vLookup() As Long, vLookup2() As Long
    
    'Size all pixels to match the number of possible Voronoi points; the nearest Voronoi point for each pixel
    ' will be used to determine the relevant point in the lookup tables.
    Dim numVoronoiPoints As Long
    numVoronoiPoints = cVoronoi.getTotalNumOfVoronoiPoints() - 1
    
    'If the number of unique Voronoi points is less than 32767 (the limit for an unsigned Int), we could get away with
    ' using an Integer lookup table instead of Longs, but as VB6 doesn't provide an easy way to change array types
    ' post-declaration, we are stuck using the worst-case scenario of Longs.
    ReDim rLookup(0 To numVoronoiPoints) As Long
    ReDim gLookUp(0 To numVoronoiPoints) As Long
    ReDim bLookup(0 To numVoronoiPoints) As Long
    ReDim aLookup(0 To numVoronoiPoints) As Long
    ReDim numPixels(0 To numVoronoiPoints) As Long
    ReDim vLookup(initX To finalX, initY To finalY) As Long
    If shadeQuality > SHADE_F1 Then
        ReDim vLookup2(initX To finalX, initY To finalY) As Long
    End If
    
    'Color values must be individually processed to account for shading, so we need to declare them
    Dim r As Long, g As Long, b As Long, a As Long
    
    'The Voronoi approach we use requires knowledge of the distance to the nearest Voronoi point, and depending on
    ' shading quality, distance to the second-nearest point as well.
    Dim nearestPoint As Long, secondNearestPoint As Long
    
    'Loop through each pixel in the image, calculating nearest Voronoi points as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Use the Voronoi class to find the nearest points to this pixel
        nearestPoint = cVoronoi.getNearestPointIndex(x, y, secondNearestPoint)
        
        'Store the nearest and second-nearest point indices in our master lookup table
        vLookup(x, y) = nearestPoint
        If shadeQuality > SHADE_F1 Then vLookup2(x, y) = secondNearestPoint
        
        'If the user has elected to recolor each cell using the average color for the cell, we need to track
        ' color values.  This is no different from a histogram approach, except in this case, each histogram
        ' bucket corresponds to one Voronoi cell.
        If colorSamplingMethod = 1 Then
        
            'Retrieve RGBA values for this pixel
            r = dstImageData(QuickVal + 2, y)
            g = dstImageData(QuickVal + 1, y)
            b = dstImageData(QuickVal, y)
            If qvDepth = 4 Then a = dstImageData(QuickVal + 3, y)
            
            'Store those RGBA values into their respective lookup "bin"
            rLookup(nearestPoint) = rLookup(nearestPoint) + r
            gLookUp(nearestPoint) = gLookUp(nearestPoint) + g
            bLookup(nearestPoint) = bLookup(nearestPoint) + b
            If qvDepth = 4 Then aLookup(nearestPoint) = aLookup(nearestPoint) + a
            
            'Increment the count of all pixels who share this Voronoi point as their nearest point
            numPixels(nearestPoint) = numPixels(nearestPoint) + 1
            
        End If
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'All lookup tables are now properly initialized.  Depending on the user's color sampling choice, calculate
    ' cell colors now.
    Dim numPixelsCache As Long
    Dim thisPoint As POINTAPI
    
    For x = 0 To numVoronoiPoints
    
        'The user wants a "fast and dirty" approach to coloring.  For each cell, use only the color of the
        ' corresponding Voronoi point pixel of that cell.
        If colorSamplingMethod = 0 Then
        
            'Retrieve the location of this Voronoi point
            thisPoint = cVoronoi.getVoronoiCoordinates(x)
            
            'Validate its bounds
            If thisPoint.x < initX Then thisPoint.x = initX
            If thisPoint.x > finalX Then thisPoint.x = finalX
            
            If thisPoint.y < initX Then thisPoint.y = initY
            If thisPoint.y > finalY Then thisPoint.y = finalY
            
            'Retrieve the color at this Voronoi point's location, and assign it to the lookup arrays
            QuickVal = thisPoint.x * qvDepth
            rLookup(x) = dstImageData(QuickVal + 2, thisPoint.y)
            gLookUp(x) = dstImageData(QuickVal + 1, thisPoint.y)
            bLookup(x) = dstImageData(QuickVal, thisPoint.y)
            If qvDepth = 4 Then aLookup(x) = dstImageData(QuickVal + 3, thisPoint.y)
        
        'The user wants us to find the average color for each cell.  This is effectively just a blur operation;
        ' for each bin in the lookup table, divide the total RGBA values by the number of pixels in that bin.
        Else
        
            numPixelsCache = numPixels(x)
            
            If numPixelsCache > 0 Then
                rLookup(x) = rLookup(x) \ numPixelsCache
                gLookUp(x) = gLookUp(x) \ numPixelsCache
                bLookup(x) = bLookup(x) \ numPixelsCache
                If qvDepth = 4 Then aLookup(x) = aLookup(x) \ numPixelsCache
            End If
            
        End If
    
    Next x
    
    'Our pixel count cache is now unneeded; free it
    Erase numPixels
    
    'Shading requires a number of specialized variables
    Dim shadeAdjustment As Single, shadeThreshold As Single, edgeAdjustment As Single, maxDistance As Single
            
    'Loop through the image, changing colors to match our calculated Voronoi values
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Use the lookup table from step 1 to find the nearest and second-nearest Voronoi point indices for this pixel.
        ' (NOTE: this step could be rewritten to simply re-request a distance calculation from the Voronoi class,
        '        but that would slow the function considerably.)
        nearestPoint = vLookup(x, y)
        If shadeQuality > SHADE_F1 Then secondNearestPoint = vLookup2(x, y)
        
        'Retrieve the RGB values from the relevant Voronoi cell bin
        r = rLookup(nearestPoint)
        g = gLookUp(nearestPoint)
        b = bLookup(nearestPoint)
        If qvDepth = 4 Then a = aLookup(nearestPoint)
        
        'If the user is using a custom edge value, we need to perform a number of extra calculations.  If they are
        ' just doing a generic filter, however, we can greatly shortcut the process.
        If edgeThreshold = 0 Then
        
            If shadeQuality <> NO_SHADE Then
                
                'Retrieve a shade value on the scale [0, 1] from the Voronoi class; it will calculate this
                ' value using the relationship between this point's distance to the nearest Voronoi point,
                ' and the maximum shading value for this cell.
                shadeAdjustment = cVoronoi.getShadingValue(x, y, nearestPoint, secondNearestPoint)
                
                'Modify the RGB values for this pixel by the retrieved shading adjustment
                r = r * shadeAdjustment
                g = g * shadeAdjustment
                b = b * shadeAdjustment
                
            End If
        
        'The user has modified the edge threshold.  Break out your mathbook!
        Else
            
            'We will now proceed to calculate and apply an edge modification on top of the pixel's existing shading value.
            ' Basically, the edge parameter controls an artificial fade for pixels whose distances fall below the edge
            ' threshold value.  Pixels above the threshold value are untouched (meaning they receive only their default
            ' shading adjustment).
            
            'We can shortcut the edge calculation process for the basic, non-shade method.
            If shadeQuality = NO_SHADE Then
                shadeAdjustment = 1
            Else
                shadeAdjustment = cVoronoi.getShadingValue(x, y, nearestPoint, secondNearestPoint)
            End If
            
            'Retrieve the maximum distance for this Voronoi cell, and use that to calculate a cell threshold value.
            maxDistance = cVoronoi.getMaxDistanceForCell(nearestPoint)
            shadeThreshold = (edgeThreshold * maxDistance)
            
            'Different shading methods require different calculations to make the edge algorithm work similarly.
            ' Sort by shade method, and calculate only a relevant edge adjustment value.
            If shadeQuality < SHADE_F2_MINUS_F1 Then
                edgeAdjustment = maxDistance - cVoronoi.getDistance(x, y, nearestPoint)
            Else
                
                If shadeQuality = SHADE_F2_MINUS_F1 Then
                    edgeAdjustment = (cVoronoi.getDistance(x, y, secondNearestPoint) - cVoronoi.getDistance(x, y, nearestPoint))
                Else
                    edgeAdjustment = shadeAdjustment
                End If
                
            End If
            
            'If our calculated adjustment falls below the shading threshold we calculated, this pixel is a candidate for
            ' edge enhancement.
            If edgeAdjustment < shadeThreshold Then
                
                'Check for the special case of 0 shade thresholds; to avoid divide-by-zero errors, we must handle
                ' such pixels separately.
                If shadeThreshold <> 0 Then
                    edgeAdjustment = edgeAdjustment / shadeThreshold
                Else
                    edgeAdjustment = 0
                End If
                
                'To provide a slightly better look, we actually use an n^2 fall-off instead of a linear one
                shadeAdjustment = shadeAdjustment * edgeAdjustment * edgeAdjustment
                
                'To avoid potential overflow errors, make sure our edge parameter only shrinks RGB values.
                ' (This case should never occur, but given the number of parameters at play in this tool,
                '  it doesn't hurt to exert a little extra caution!)
                If shadeAdjustment > 1 Then shadeAdjustment = 1
                
            End If
            
            'With our shade adjustment finalized, we can finally calculate final RGB values for this pixel.  Note
            ' that shading does not currently affect alpha.
            r = r * shadeAdjustment
            g = g * shadeAdjustment
            b = b * shadeAdjustment
            
        End If
        
        'Set the new RGBA values to the image
        dstImageData(QuickVal + 2, y) = r
        dstImageData(QuickVal + 1, y) = g
        dstImageData(QuickVal, y) = b
        If qvDepth = 4 Then dstImageData(QuickVal + 3, y) = a
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal finalX + x
            End If
        End If
    Next x
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'For fun, you can uncomment the code block below to render the calculated Voronoi points onto the image.
'    For x = 0 To numVoronoiPoints
'        thisPoint = cVoronoi.getVoronoiCoordinates(x)
'        GDIPlusDrawCircleToDC workingDIB.getDIBDC, thisPoint.x, thisPoint.y, 2, RGB(255, 0, 255)
'    Next x
        
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub cboColorSampling_Click()
    updatePreview
End Sub

Private Sub cboDistance_Click()
    updatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Stained glass", , buildParams(sltSize, sltTurbulence, cboColorSampling.ListIndex, sltShadeQuality, sltEdge, cboDistance.ListIndex), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltSize.Value = 20
    sltTurbulence.Value = 0.5
    sltShadeQuality.Value = 5
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Request a preview
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully initialized
    cmdBar.markPreviewStatus False
    
    'Provide with user with several color sampling options
    cboColorSampling.Clear
    cboColorSampling.AddItem "fast"
    cboColorSampling.AddItem "accurate"
    cboColorSampling.ListIndex = 0
        
    'Provide three experimental distance functions
    cboDistance.Clear
    cboDistance.AddItem "Cartesian (traditional)"
    cboDistance.AddItem "Manhattan (walking)"
    cboDistance.AddItem "Chebyshev (chessboard)"
    cboDistance.ListIndex = 0
    
    'Calculate a random turbulence seed
    Set cRandom = New pdRandomize
    cRandom.setSeed_AutomaticAndRandom
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then fxStainedGlass sltSize, sltTurbulence, cboColorSampling.ListIndex, sltShadeQuality, sltEdge, cboDistance.ListIndex, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltEdge_Change()
    updatePreview
End Sub

Private Sub sltShadeQuality_Change()
    updatePreview
End Sub

Private Sub sltSize_Change()
    updatePreview
End Sub

Private Sub sltTurbulence_Change()
    updatePreview
End Sub
