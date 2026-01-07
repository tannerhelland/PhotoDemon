VERSION 5.00
Begin VB.Form FormCrystallize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Crystallize"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11655
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
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   Visible         =   0   'False
   Begin PhotoDemon.pdDropDown cboColorSampling 
      Height          =   735
      Left            =   6000
      TabIndex        =   4
      Top             =   2880
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      Caption         =   "color sampling"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5475
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9657
   End
   Begin PhotoDemon.pdSlider sltSize 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "cell size"
      Min             =   3
      Max             =   200
      Value           =   20
      DefaultValue    =   20
   End
   Begin PhotoDemon.pdSlider sltTurbulence 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   3840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1244
      Caption         =   "turbulence"
      Max             =   1
      SigDigits       =   2
   End
   Begin PhotoDemon.pdDropDown cboDistance 
      Height          =   735
      Left            =   6000
      TabIndex        =   5
      Top             =   1920
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      Caption         =   "distance method"
   End
   Begin PhotoDemon.pdDropDown cboPattern 
      Height          =   735
      Left            =   6000
      TabIndex        =   6
      Top             =   960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      Caption         =   "pattern"
   End
End
Attribute VB_Name = "FormCrystallize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Crystallize Effect Interface
'Copyright 2014-2026 by Tanner Helland
'Created: 14/July/14
'Last updated: 09/December/20
'Last update: 2x performance improvements, new "pattern" feature
'
'PhotoDemon's crystallize effect is implemented using Worley Noise...
' (https://en.wikipedia.org/wiki/Worley_noise)
' ...which is basically a special algorithmic approach to Voronoi diagrams...
' (https://en.wikipedia.org/wiki/Voronoi_diagram)
'
'The associated pdVoronoi class does most the heavy lifting for this effect.
' For details on this class and how it works, I recommend referring to the Stained Glass
' effect dialog, which details it in much more detail.
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

'A persistent pdVoronoi object is used to handle all Voronoi-related tasks
Private m_Voronoi As pdVoronoi

'Apply a Crystallize effect to an image
' Inputs:
'  cellSize = size, in pixels, of each initial grid box in the Voronoi array.  Do not make this less than 3.
'  fxTurbulence = how much to distort cell shape, range [0, 1], 0 = perfect grid
'  colorSamplingMethod = how to determine cell color (0 = just use pixel at Voronoi point, 1 = average all pixels in cell)
'  distanceMethod = 0 - Cartesian, 1 - Manhattan, 2 - Chebyshev
Public Sub fxCrystallize(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Crystallizing image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim cellSize As Long, colorSamplingMethod As Long, distanceMethod As Long
    Dim fxTurbulence As Double, fxDensity As Double, fxPatternName As String
    
    With cParams
        cellSize = .GetLong("size", sltSize.Value)
        fxTurbulence = .GetDouble("turbulence", sltTurbulence.Value)
        fxDensity = 1#  'Density is no longer exposed to the user, as it isn't that interesting!
        colorSamplingMethod = .GetLong("colorsampling", cboColorSampling.ListIndex)
        distanceMethod = .GetLong("distance", cboDistance.ListIndex)
        fxPatternName = .GetString("pattern", "square")
    End With
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, , , True
    workingDIB.WrapArrayAroundDIB dstImageData, dstSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'Because this is a two-pass filter, we have to manually change the progress bar maximum to 2 * width
    If (Not toPreview) Then
        SetProgBarMax finalX * 2
    
    'If this is a preview, reduce cell size to better portray how the final image will look
    Else
        cellSize = cellSize * curDIBValues.previewModifier
        If (cellSize < 3) Then cellSize = 3
    End If
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Pass all meaningful input parameters on to the Voronoi class.
    ' IMPORTANTLY, note that parameters need to be set in a certain order.  DO NOT DEVIATE
    ' from the pattern used here!
    If (m_Voronoi Is Nothing) Then Set m_Voronoi = New pdVoronoi
    m_Voronoi.InitPoints cellSize, workingDIB.GetDIBWidth, workingDIB.GetDIBHeight
    m_Voronoi.RandomizePoints fxTurbulence, m_Random.GetSeed()
    m_Voronoi.SetDistanceMode distanceMethod
    m_Voronoi.SetShadingMode vs_NoShade
    m_Voronoi.SetInitialPattern m_Voronoi.GetPatternIDFromName(fxPatternName)
    m_Voronoi.SetDensity fxDensity
    m_Voronoi.FinalizeParameters
    
    'Finally, we will also make two image-sized look-up tables that store the nearest Voronoi point index for
    ' each pixel in the image, and if certain shading types are active, the second-nearest Voronoi point as well.
    ' While this consumes a lot of memory, it makes our second pass through the image (the recoloring pass) much
    ' faster than it would otherwise be.
    Dim vLookup() As Long
    ReDim vLookup(initX To finalX, initY To finalY) As Long
    
    'Size all pixels to match the number of possible Voronoi points; the nearest Voronoi point for each pixel
    ' will be used to determine the relevant point in the lookup tables.
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
    
    'The Voronoi approach we use requires knowledge of the distance to the nearest Voronoi point, and depending on
    ' shading quality, distance to the second-nearest point as well.
    Dim nearestPoint As Long
    
    'Loop through each pixel in the image, calculating nearest Voronoi points as we go
    For x = initX To finalX
        xStride = x * 4
    For y = initY To finalY
        
        'Use the Voronoi class to find the nearest point to this pixel
        nearestPoint = m_Voronoi.GetNearestPointIndex(x, y)
        
        'Store the nearest point index in our central lookup table
        vLookup(x, y) = nearestPoint
        
        'If the user has elected to recolor each cell using the average color for the cell,
        ' we need to track color values.  This is no different from a histogram approach,
        ' except in this case, each histogram bucket corresponds to one Voronoi cell.
        If (colorSamplingMethod = 0) Then
        
            'Retrieve RGBA values for this pixel
            b = dstImageData(xStride, y)
            g = dstImageData(xStride + 1, y)
            r = dstImageData(xStride + 2, y)
            a = dstImageData(xStride + 3, y)
            
            'Store those RGBA values into their respective lookup "bin"
            rLookup(nearestPoint) = rLookup(nearestPoint) + r
            gLookup(nearestPoint) = gLookup(nearestPoint) + g
            bLookup(nearestPoint) = bLookup(nearestPoint) + b
            aLookup(nearestPoint) = aLookup(nearestPoint) + a
            
            'Increment the count of all pixels who share this Voronoi point as their nearest point
            numPixels(nearestPoint) = numPixels(nearestPoint) + 1
            
        End If
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'All lookup tables are now properly initialized.  Depending on the user's color sampling choice, calculate
    ' cell colors now.
    Dim numPixelsCache As Long
    Dim thisPoint As PointFloat
    
    For x = 0 To numVoronoiPoints
        
        Select Case colorSamplingMethod
        
             'accurate
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
            
             'fast
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
            
             'random
             Case 2
                rLookup(x) = m_Random.GetRandomIntRange_WH(0, 255)
                gLookup(x) = m_Random.GetRandomIntRange_WH(0, 255)
                bLookup(x) = m_Random.GetRandomIntRange_WH(0, 255)
                aLookup(x) = 255
                
        End Select
        
    Next x
    
    'Our pixel count cache is now unneeded; free it
    Erase numPixels
                
    'Loop through the image, changing colors to match our calculated Voronoi values
    For x = initX To finalX
        xStride = x * 4
    For y = initY To finalY
        
        'Use the lookup table from step 1 to find the nearest Voronoi point index for this pixel.
        ' (NOTE: this step could be rewritten to simply re-request a distance calculation from the Voronoi class,
        '        but that would slow the function considerably.)
        nearestPoint = vLookup(x, y)
                
        'Retrieve the RGB values from the relevant Voronoi cell bin
        dstImageData(xStride, y) = bLookup(nearestPoint)
        dstImageData(xStride + 1, y) = gLookup(nearestPoint)
        dstImageData(xStride + 2, y) = rLookup(nearestPoint)
        dstImageData(xStride + 3, y) = aLookup(nearestPoint)
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal finalX + x
            End If
        End If
    Next x
    
    workingDIB.UnwrapArrayFromDIB dstImageData
    
    'For fun, you can uncomment the code block below to render the calculated Voronoi points onto the image.
'    For x = 0 To numVoronoiPoints
'        thisPoint = cVoronoi.GetVoronoiCoordinates(x)
'        GDIPlusDrawCircleToDC workingDIB.getDIBDC, thisPoint.x, thisPoint.y, 2, RGB(255, 0, 255)
'    Next x
        
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
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
    Process "Crystallize", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    Set m_Voronoi = New pdVoronoi
    Set m_Random = New pdRandomize
    
    'Disable previews until the dialog is fully initialized
    cmdBar.SetPreviewStatus False
    
    'Initial pattern options
    cboPattern.Clear
    Dim i As Long
    For i = 0 To m_Voronoi.GetPatternCount - 1
        cboPattern.AddItem m_Voronoi.GetPatternUINameFromID(i), i
    Next i
    cboPattern.ListIndex = 0
    
    'Color sampling options
    cboColorSampling.Clear
    cboColorSampling.AddItem "accurate"
    cboColorSampling.AddItem "fast"
    cboColorSampling.AddItem "random"
    cboColorSampling.ListIndex = 0
        
    'Experimental distance functions
    cboDistance.Clear
    cboDistance.AddItem "Cartesian (traditional)"
    cboDistance.AddItem "Manhattan (walking)"
    cboDistance.AddItem "Chebyshev (chessboard)"
    cboDistance.ListIndex = 0
    
    'Calculate a random noise seed
    m_Random.SetSeed_AutomaticAndRandom
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    
    'Request a preview
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the effect preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then fxCrystallize GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
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
        .AddParam "turbulence", sltTurbulence.Value
        .AddParam "colorsampling", cboColorSampling.ListIndex
        .AddParam "distance", cboDistance.ListIndex
        
        'Shape options use string values for future-proofed expansion possibilities
        If (m_Voronoi Is Nothing) Then Set m_Voronoi = New pdVoronoi
        .AddParam "pattern", m_Voronoi.GetPatternNameFromID(cboPattern.ListIndex)
        
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
