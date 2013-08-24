VERSION 5.00
Begin VB.Form FormSpherize 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Spherize"
   ClientHeight    =   6540
   ClientLeft      =   -15
   ClientTop       =   225
   ClientWidth     =   12105
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   807
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   14
      Top             =   5790
      Width           =   12105
      _ExtentX        =   21352
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
   Begin VB.ComboBox cmbEdges 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   4
      Top             =   3345
      Width           =   5700
   End
   Begin PhotoDemon.fxPreviewCtl fxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   0
      Left            =   6120
      TabIndex        =   6
      Top             =   4200
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   635
      Caption         =   "quality"
      Value           =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartOptionButton OptInterpolate 
      Height          =   330
      Index           =   1
      Left            =   7920
      TabIndex        =   7
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      Caption         =   "speed"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.sliderTextCombo sltAngle 
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   450
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -180
      Max             =   180
      SigDigits       =   1
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
   Begin PhotoDemon.sliderTextCombo sltOffsetY 
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   2370
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
      SigDigits       =   1
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
   Begin PhotoDemon.sliderTextCombo sltOffsetX 
      Height          =   495
      Left            =   6000
      TabIndex        =   10
      Top             =   1410
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -100
      Max             =   100
      SigDigits       =   1
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
   Begin PhotoDemon.smartCheckBox chkRays 
      Height          =   540
      Left            =   6120
      TabIndex        =   12
      Top             =   5040
      Width           =   3705
      _ExtentX        =   5980
      _ExtentY        =   847
      Caption         =   "fill exterior with matching light rays"
      Value           =   1
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
   Begin VB.Label lblOther 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "other options:"
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
      Left            =   6000
      TabIndex        =   13
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Label lblOffset 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "horizontal offset:"
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
      TabIndex        =   11
      Top             =   1080
      Width           =   1800
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "if pixels lie outside the image..."
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
      TabIndex        =   5
      Top             =   2970
      Width           =   3315
   End
   Begin VB.Label lblOffset 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vertical offset:"
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
      TabIndex        =   2
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Label lblInterpolation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "render emphasis:"
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
      Left            =   6000
      TabIndex        =   1
      Top             =   3840
      Width           =   1845
   End
   Begin VB.Label lblAmount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "angle:"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "FormSpherize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image "Spherize" Distortion
'Copyright ©2012-2013 by Tanner Helland
'Created: 05/June/13
'Last updated: 08/June/13
'Last update: performance optimizations; speed was improved ~15-25% with some lookup tables and precalculation tricks
'
'This tool allows the user to map an image around a sphere, with optional rotation and bidirectional offsets.
' Bilinear interpolation (via reverse-mapping) is available to improve mapping quality.
'
'In keeping with the true mathematical nature of spheres, this tool forces x and y mapping to the minimum
' dimension (width or height).  Technically this isn't necessary, as the code works just fine with rectangular
' shapes, but as my lens distort tool already handles rectangular shapes, I'm forcing this one to spheres only.
'
'For a bit of extra fun, the empty space around the sphere can be mapped one of two ways - as light rays "beaming"
' from behind the sphere, or simply erased to white.  Your choice.
'
'The transformation used by this tool is a heavily modified version of basic math originally shared by Paul Bourke.
' You can see Paul's original (and very helpful article) at the following link, good as of 05 June '13:
' http://paulbourke.net/miscellaneous/imagewarp/
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub chkRays_Click()
    updatePreview
End Sub

Private Sub cmbEdges_Click()
    updatePreview
End Sub

'Apply a "swirl" effect to an image
Public Sub SpherizeImage(ByVal sphereAngle As Double, ByVal xOffset As Double, ByVal yOffset As Double, ByVal useRays As Boolean, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    'Reverse the rotationAngle value so that POSITIVE values indicate CLOCKWISE rotation.
    ' Also, convert it to radians.
    sphereAngle = sphereAngle * (PI / 180)

    If toPreview = False Then Message "Wrapping image around sphere..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent diffused pixels from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, edgeHandling, useBilinear, curLayerValues.maxX, curLayerValues.MaxY
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
          
    'Swirling requires some specialized variables
    
    'Polar conversion values
    Dim theta As Double, r As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
        
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = curLayerValues.Width
    tHeight = curLayerValues.Height
    
    Dim minDimension As Long
    Dim minDimVertical As Boolean
    
    If tWidth < tHeight Then
        minDimension = tWidth
        minDimVertical = False
    Else
        minDimension = tHeight
        minDimVertical = True
    End If
        
    Dim halfDimDiff As Double
    halfDimDiff = Abs(tWidth - tHeight) / 2
    
    'Convert offsets to usable amounts
    xOffset = (xOffset / 100) * tWidth
    yOffset = (yOffset / 100) * tHeight
    
    'A slow part of this algorithm is translating all (x, y) coordinates to the range [-1, 1].  We can perform this
    ' step in advance by constructing a look-up table for all possible x values and all possible y values.  This
    ' greatly improves performance.
    Dim xLookup() As Double, yLookup() As Double
    ReDim xLookup(initX To finalX) As Double
    ReDim yLookup(initY To finalY) As Double
    
    For x = initX To finalX
        If minDimVertical Then
            xLookup(x) = (2 * (x - halfDimDiff)) / minDimension - 1
        Else
            xLookup(x) = (2 * x) / minDimension - 1
        End If
    Next x
    
    For y = initY To finalY
        If minDimVertical Then
            yLookup(y) = (2 * y) / minDimension - 1
        Else
            yLookup(y) = (2 * (y - halfDimDiff)) / minDimension - 1
        End If
    Next y
    
    'We can also calculate a few constants in advance
    Dim twoDivByPI As Double
    twoDivByPI = 2 / PI
    
    Dim halfMinDimension As Double
    halfMinDimension = minDimension / 2
            
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Remap the coordinates around a center point of (0, 0), and normalize them to (-1, 1)
        nX = xLookup(x)
        nY = yLookup(y)
        
        'Next, map them to polar coordinates and apply the spherification
        r = Sqr(nX * nX + nY * nY)
        theta = Atan2(nY, nX)
        
        r = Asin(r) * twoDivByPI
        
        'Apply optional rotation
        theta = theta - sphereAngle
                
        'Convert them back to the Cartesian plane
        nX = r * Cos(theta) + 1
        srcX = halfMinDimension * nX + xOffset
        nY = r * Sin(theta) + 1
        srcY = halfMinDimension * nY + yOffset
        
        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
        If useRays Then
            fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
        Else
            If r < 1 Then
                fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
            Else
                fSupport.forcePixels x, y, 255, 255, 255, 0, dstImageData
            End If
        End If
                
    Next y
        If Not toPreview Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
        
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Spherize", , buildParams(sltAngle, sltOffsetX, sltOffsetY, CBool(chkRays.Value), CLng(cmbEdges.ListIndex), OptInterpolate(0).Value)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
        
    'Create the preview
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Don't attempt to preview the image until the dialog is fully initialized
    cmdBar.markPreviewStatus False
    
    'I use a central function to populate the edge handling combo box; this way, I can add new methods and have
    ' them immediately available to all distort functions.
    popDistortEdgeBox cmbEdges, EDGE_WRAP
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub OptInterpolate_Click(Index As Integer)
    updatePreview
End Sub

Private Sub sltAngle_Change()
    updatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then SpherizeImage sltAngle, sltOffsetX, sltOffsetY, CBool(chkRays.Value), CLng(cmbEdges.ListIndex), OptInterpolate(0).Value, True, fxPreview
End Sub

Private Sub sltOffsetX_Change()
    updatePreview
End Sub

Private Sub sltOffsetY_Change()
    updatePreview
End Sub
