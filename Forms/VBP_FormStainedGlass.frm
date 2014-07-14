VERSION 5.00
Begin VB.Form FormStainedGlass 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Stained glass"
   ClientHeight    =   6510
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
   ScaleHeight     =   434
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
      TabIndex        =   11
      Top             =   5160
      Width           =   5775
   End
   Begin VB.ComboBox cboShading 
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
      TabIndex        =   8
      Top             =   4320
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
      TabIndex        =   7
      Top             =   2880
      Width           =   5775
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5760
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltSize 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   690
      Width           =   5895
      _ExtentX        =   10186
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   3
      Max             =   200
      Value           =   50
   End
   Begin PhotoDemon.sliderTextCombo sltTurbulence 
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   1
      SigDigits       =   2
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   100
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "distance technique:"
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
      TabIndex        =   12
      Top             =   4800
      Width           =   2040
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXPERIMENTAL SETTINGS:"
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
      Index           =   4
      Left            =   6000
      TabIndex        =   10
      Top             =   3480
      Width           =   2985
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "shading:"
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
      Index           =   3
      Left            =   6000
      TabIndex        =   9
      Top             =   3960
      Width           =   915
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "color sampling:"
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
      TabIndex        =   6
      Top             =   2520
      Width           =   1650
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "turbulence:"
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
      Index           =   2
      Left            =   6000
      TabIndex        =   4
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cell size:"
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
      Index           =   0
      Left            =   6000
      TabIndex        =   1
      Top             =   360
      Width           =   900
   End
End
Attribute VB_Name = "FormStainedGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Stained Glass Effect Interface
'Copyright ©2013-2014 by Tanner Helland
'Created: 14/July/14
'Last updated: 14/July/14
'Last update: initial build
'
'TODO: explain how Voronoi diagrams are used to generate the effect
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'To make sure the function looks similar in the preview and final image, we cache the random seed used
Private m_RndSeed As Long

'Apply a Stained Glass effect to an image
' Inputs: size of each cell, turbulence of the cells (range [0, 1]), color sampling method (0 = point, 1 = area)
Public Sub fxStainedGlass(ByVal cellSize As Long, ByVal fxTurbulence As Double, ByVal colorSamplingMethod As Long, ByVal shadeMethod As Long, ByVal distanceMethod As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Carving image from stained glass..."
        
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent processed pixels from spreading across the image as we go.)
    'Dim srcDIB As pdDIB
    'Set srcDIB = New pdDIB
    'srcDIB.createFromExistingDIB workingDIB
    
    'Dim srcImageData() As Byte
    'Dim srcSA As SAFEARRAY2D
    'prepSafeArray srcSA, srcDIB
    'CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValDiffuseX As Long, QuickValDiffuseY As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'Because this is a two-pass filter, we have to manually change the progress bar maximum
    If Not toPreview Then
        SetProgBarMax finalX * 2
    Else
        cellSize = cellSize * curDIBValues.previewModifier
        If cellSize < 3 Then cellSize = 3
    End If
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Create a Voronoi class to help us with processing; it's at the heart of this function.
    Dim cVoronoi As pdVoronoi
    Set cVoronoi = New pdVoronoi
    
    cVoronoi.initPoints cellSize, workingDIB.getDIBWidth, workingDIB.getDIBHeight
    cVoronoi.randomizePoints fxTurbulence, m_RndSeed
    cVoronoi.setDistanceMode distanceMethod
    
    'Create several look-up tables; one for each channel (RGBA), and one for the number of points tied to that index.
    ' The last lookup table is necessary to average the contents of each panel to arrive at an ideal color representation.
    Dim rLookup() As Long, gLookUp() As Long, bLookup() As Long, aLookup() As Long
    Dim numPixels() As Long
    
    'Size all pixels to match the number of possible Voronoi points
    Dim numVoronoiPoints As Long
    numVoronoiPoints = cVoronoi.getTotalNumOfVoronoiPoints() - 1
    
    ReDim rLookup(0 To numVoronoiPoints) As Long
    ReDim gLookUp(0 To numVoronoiPoints) As Long
    ReDim bLookup(0 To numVoronoiPoints) As Long
    ReDim aLookup(0 To numVoronoiPoints) As Long
    ReDim numPixels(0 To numVoronoiPoints) As Long
    
    'Color values must be retrieved
    Dim r As Long, g As Long, b As Long, a As Long
    
    'Finally, we will also make a look-up table of nearest Voronoi values at the size of the image; this makes our
    ' second pass through the image much, much faster than it would otherwise be.
    Dim vLookup() As Long
    ReDim vLookup(initX To finalX, initY To finalY) As Long
    
    Dim nearestPoint As Long
    
    'Loop through each pixel in the image, tallying Voronoi values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Use the Voronoi class to find the nearest point
        nearestPoint = cVoronoi.getNearestPointIndex(x, y)
        
        'Populate all relevant look-up tables
        vLookup(x, y) = nearestPoint
        
        If colorSamplingMethod = 1 Then
        
            'Get red, green, and blue values from the array
            r = dstImageData(QuickVal + 2, y)
            g = dstImageData(QuickVal + 1, y)
            b = dstImageData(QuickVal, y)
            If qvDepth = 3 Then a = dstImageData(QuickVal + 3, y)
        
            'Store color and alpha values
            rLookup(nearestPoint) = rLookup(nearestPoint) + r
            gLookUp(nearestPoint) = gLookUp(nearestPoint) + g
            bLookup(nearestPoint) = bLookup(nearestPoint) + b
            If qvDepth = 3 Then aLookup(nearestPoint) = aLookup(nearestPoint) + a
            
            'Update the number of pixels sharing this point as their nearest point
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
    
        If colorSamplingMethod = 0 Then
        
            'Retrieve the location of this Voronoi point
            thisPoint = cVoronoi.getVoronoiCoordinates(x)
            
            If thisPoint.x < initX Then thisPoint.x = initX
            If thisPoint.x > finalX Then thisPoint.x = finalX
            
            If thisPoint.y < initX Then thisPoint.y = initY
            If thisPoint.y > finalY Then thisPoint.y = finalY
            
            'Retrieve the color at this Voronoi point, and assign it to the lookup arrays
            QuickVal = thisPoint.x * qvDepth
            rLookup(x) = dstImageData(QuickVal + 2, thisPoint.y)
            gLookUp(x) = dstImageData(QuickVal + 1, thisPoint.y)
            bLookup(x) = dstImageData(QuickVal, thisPoint.y)
            If qvDepth = 3 Then aLookup(x) = dstImageData(QuickVal + 3, thisPoint.y)
        
        Else
        
            numPixelsCache = numPixels(x)
            
            If numPixelsCache > 0 Then
                rLookup(x) = rLookup(x) \ numPixelsCache
                gLookUp(x) = gLookUp(x) \ numPixelsCache
                bLookup(x) = bLookup(x) \ numPixelsCache
                If qvDepth = 3 Then aLookup(x) = aLookup(x) \ numPixelsCache
            End If
            
        End If
    
    Next x
    
    'Our pixel cache is now unneeded; free it
    Erase numPixels
    
    Dim shadeAdjustment As Double
    
    'Loop through the image, changing colors to match our calculated Voronoi values
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        
        'Use the lookup table to more quickly access the nearest point
        nearestPoint = vLookup(x, y)
        
        'Apply the RGB values from that point, using shading as appropriate
        r = rLookup(nearestPoint)
        g = gLookUp(nearestPoint)
        b = bLookup(nearestPoint)
        If qvDepth = 3 Then a = aLookup(nearestPoint)
        
        If shadeMethod > 0 Then
            
            'Apply shading now, with help from the Voronoi class
            shadeAdjustment = cVoronoi.getShadingValue(x, y, shadeMethod, nearestPoint)
            
            r = r * shadeAdjustment
            g = g * shadeAdjustment
            b = b * shadeAdjustment
            
        End If
        
        dstImageData(QuickVal + 2, y) = r
        dstImageData(QuickVal + 1, y) = g
        dstImageData(QuickVal, y) = b
        If qvDepth = 3 Then dstImageData(QuickVal + 3, y) = a
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal finalX + x
            End If
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    'CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    'Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub cboColorSampling_Click()
    updatePreview
End Sub

Private Sub cboDistance_Click()
    updatePreview
End Sub

Private Sub cboShading_Click()
    updatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Stained glass", , buildParams(sltSize, sltTurbulence, cboColorSampling.ListIndex, cboShading.ListIndex, cboDistance.ListIndex), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltSize.Value = 50
    sltTurbulence.Value = 0.5
End Sub

Private Sub Form_Activate()
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Request a preview
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully initialized
    cmdBar.markPreviewStatus False
    
    'Provide with user with several color sampling options
    cboColorSampling.Clear
    cboColorSampling.AddItem "pixel at center of cell"
    cboColorSampling.AddItem "average of all pixels in cell"
    cboColorSampling.ListIndex = 0
    
    'Provide several experimental shading options
    cboShading.Clear
    cboShading.AddItem "none"
    cboShading.AddItem "pixel distance / max distance for cell"
    cboShading.ListIndex = 0
    
    'Provide three experimental distance functions
    cboDistance.Clear
    cboDistance.AddItem "Cartesian (traditional)"
    cboDistance.AddItem "Manhattan (walking)"
    cboDistance.AddItem "Chebyshev (chessboard)"
    cboDistance.ListIndex = 0
    
    'Calculate a random noise seed
    Rnd -1
    Randomize (-Timer * Now)
    m_RndSeed = Rnd * &HEFFFFFFF
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Redraw the effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then fxStainedGlass sltSize, sltTurbulence, cboColorSampling.ListIndex, cboShading.ListIndex, cboDistance.ListIndex, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub sltSize_Change()
    updatePreview
End Sub

Private Sub sltTurbulence_Change()
    updatePreview
End Sub
