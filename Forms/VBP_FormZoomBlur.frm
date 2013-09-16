VERSION 5.00
Begin VB.Form FormZoomBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Zoom Blur"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12030
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
   ScaleWidth      =   802
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
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
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.sliderTextCombo sltDistance 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   3360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      Min             =   -200
      Max             =   200
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
   Begin PhotoDemon.smartOptionButton OptStyle 
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   5
      Top             =   2160
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   635
      Caption         =   "modern"
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
   Begin PhotoDemon.smartOptionButton OptStyle 
      Height          =   360
      Index           =   1
      Left            =   6120
      TabIndex        =   6
      Top             =   2520
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   635
      Caption         =   "traditional"
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
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "style:"
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
      TabIndex        =   4
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "distance:"
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
      TabIndex        =   2
      Top             =   3000
      Width           =   945
   End
End
Attribute VB_Name = "FormZoomBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Zoom Blur Tool
'Copyright ©2012-2013 by Tanner Helland
'Created: 27/August/13
'Last updated: 16/September/13
'Last update: added a "traditional" mode, with optimizations based on aspect ratio (to minimize data loss from
'              repeated polar coord conversions)
'
'Basic zoom blur tool.  For performance reasons, my approach relies heavily on StretchBlt and AlphaBlend.  The
' resulting zoom is of reasonably good quality, and it outperforms similar tools in both GIMP and Paint.NET, so I
' think the implementation is solid.  That said, this function doesn't allow for sub-pixel control, which means that
' even at small levels the blur is quite strong.  I could remedy this by using my internal pan/zoom function instead
' of StretchBlt, but there would be a large performance penalty.  Maybe in the future...
'
'Note that unlike other blur tools, this one performances quite nicely in the IDE (due to its reliance on API
' functions to do the heavy lifting).
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'AlphaBlend is used to render a fast zoom blur estimation
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Because PD now provides two styles of zoom blur, I've added this wrapper function, which calls the appropriate *actual* zoom blur
' function, without the caller having to know details about either implementation.
Public Sub ZoomBlurWrapper(ByVal useModernStyle As Boolean, ByVal zDistance As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)

    If useModernStyle Then
        ZoomBlurModern zDistance, toPreview, dstPic
    Else
        ZoomBlurTraditional zDistance, toPreview, dstPic
    End If

End Sub

'Apply motion blur to an image using a "modern" approach that allows for both in and out zoom
'Inputs: distance of the blur
Public Sub ZoomBlurModern(ByVal zDistance As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying zoom blur..."
    
    'Call prepImageData, which will initialize a workingLayer object for us (with all selection tool masks applied)
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    Dim finalX As Long, finalY As Long
    finalX = workingLayer.getLayerWidth
    finalY = workingLayer.getLayerHeight
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    SetProgBarMax Abs(zDistance)
    progBarCheck = findBestProgBarValue()
    
    'AlphaBlend has two limitations we have to work around: it only stretches using nearest-neighbor interpolation (which
    ' creates very obvious blocking along the lines where the algorithm stretches), and it is incapable of rendering from
    ' a source DC to a matching dest DC.  Thus we will make our own a temporary copy of the image, which we will use to
    ' work around these two issues.
    Dim tmpSrcLayer As pdLayer
    Set tmpSrcLayer = New pdLayer
    tmpSrcLayer.createFromExistingLayer workingLayer
    
    'Next, we set up some AlphaBlend parameters.  If the image already has alpha data, we want to make use of it, although
    ' it will result in a funky final image (I may work around this in the future, but right now I don't care).  If the
    ' image is 24bpp, we'll use halftoning to generate a higher-quality zoom.
    Dim blendVal As Long
    blendVal = 127 * &H10000
    If workingLayer.getLayerColorDepth = 32 Then
        SetStretchBltMode tmpSrcLayer.getLayerDC, STRETCHBLT_COLORONCOLOR
    Else
        SetStretchBltMode tmpSrcLayer.getLayerDC, STRETCHBLT_HALFTONE
    End If
    
    'Finally, we're going to use the image's aspect ratio to try and generate a more realistic zoom blur.  Calculate
    ' that ratio now.
    Dim aspectRatio As Double
    If workingLayer.getLayerWidth > workingLayer.getLayerHeight Then
        aspectRatio = workingLayer.getLayerHeight / workingLayer.getLayerWidth
    Else
        aspectRatio = workingLayer.getLayerWidth / workingLayer.getLayerHeight
    End If
    
    'Zoom distance must be adjusted during a preview, so that the preview accurately represents the finished product.
    If toPreview Then zDistance = zDistance * curLayerValues.previewModifier
    
    'Now comes the actual transform.  We basically just repeat a series of AlphaBlend calls on the image, blending at 50% opacity
    ' as we go.  Ridiculous?  Yes.  Simple?  Yes.  :)
    Dim i As Long, zoomOffset As Double
    
    'Note also that we have to use different functions for +/- distance
    If zDistance > 0 Then
        
        For i = 0 To zDistance Step 1
        
            'Because AlphaBlend can't stretch with anything besides nearest neighbor, we must do our own StretchBlt in advance.
            ' I also modify the blend factors using the aspect ratio calculated above, which applies the blur at roughly the
            ' same aspect ratio as the image itself (e.g. the "zoom" lines extend from the center toward the corners, and not
            ' at 45 degree increments).
            zoomOffset = i / 2
            If workingLayer.getLayerWidth > workingLayer.getLayerHeight Then
                StretchBlt tmpSrcLayer.getLayerDC, 0, 0, finalX, finalY, workingLayer.getLayerDC, zoomOffset, zoomOffset * aspectRatio, finalX - zoomOffset * 2, finalY - ((zoomOffset * aspectRatio) * 2), vbSrcCopy
            Else
                StretchBlt tmpSrcLayer.getLayerDC, 0, 0, finalX, finalY, workingLayer.getLayerDC, zoomOffset * aspectRatio, zoomOffset, finalX - ((zoomOffset * aspectRatio) * 2), finalY - zoomOffset * 2, vbSrcCopy
            End If
            
            'Finally, AlphaBlend the modified layer onto the working layer, then rinse and repeat!
            AlphaBlend workingLayer.getLayerDC, 0, 0, finalX, finalY, tmpSrcLayer.getLayerDC, 0, 0, finalX, finalY, blendVal
            
            If Not toPreview Then
                If (i And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal i
                End If
            End If
        
        Next i
        
    Else
    
        For i = 0 To zDistance Step -1
        
            'Because AlphaBlend can't stretch with anything besides nearest neighbor, we must do our own StretchBlt in advance.
            ' I also modify the blend factors using the aspect ratio calculated above, which applies the blur at roughly the
            ' same aspect ratio as the image itself (e.g. the "zoom" lines extend from the center toward the corners, and not
            ' at 45 degree increments).
            zoomOffset = Abs(i / 2)
            If workingLayer.getLayerWidth > workingLayer.getLayerHeight Then
                StretchBlt tmpSrcLayer.getLayerDC, zoomOffset, zoomOffset * aspectRatio, finalX - zoomOffset * 2, finalY - ((zoomOffset * aspectRatio) * 2), workingLayer.getLayerDC, 0, 0, finalX, finalY, vbSrcCopy
            Else
                StretchBlt tmpSrcLayer.getLayerDC, zoomOffset * aspectRatio, zoomOffset, finalX - ((zoomOffset * aspectRatio) * 2), finalY - zoomOffset * 2, workingLayer.getLayerDC, 0, 0, finalX, finalY, vbSrcCopy
            End If
            
            'Finally, AlphaBlend the modified layer onto the working layer, then rinse and repeat!
            AlphaBlend workingLayer.getLayerDC, 0, 0, finalX, finalY, tmpSrcLayer.getLayerDC, 0, 0, finalX, finalY, blendVal
            
            If Not toPreview Then
                If (Abs(i) And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal Abs(i)
                End If
            End If
        
        Next i
    
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic
    
End Sub

'Apply "traditional" zoom blur to an image
'Inputs: distance of the blur
Public Sub ZoomBlurTraditional(ByVal bDistance As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying zoom blur..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second local array. This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    'By dividing blur distance by 200 (its maximum value), we can use it as a fractional amount to determine the strength of our horizontal blur.
    If toPreview Then
        bDistance = bDistance * curLayerValues.previewModifier
        If bDistance < 1 Then bDistance = 1
    End If
    
    Dim finalX As Long, finalY As Long
    finalX = workingLayer.getLayerWidth
    finalY = workingLayer.getLayerHeight
    
    Dim newProgBarMax As Long
    
    'Zoom blur basically works by converting an image to polar coordinates, applying a horizontal blur, then converting
    ' back to rectangular coordinates.  Even with interpolation, the two coordinate conversion functions result in a loss
    ' of image data, so I've gone to some lengths to try and mitigate this.
    '
    'Polar coordinate conversion basically works by using either X or Y to represent radius, and the other to represent theta
    ' (or the angle of the pixel).  Because images are stored as rectangles, radius tends to preserve more of the original data
    ' than theta, and obviously the larger of X or Y will retain more data by virtue of having more pixels available.
    '
    'Thus, PD checks the image's aspect ratio.  Whichever dimension is bigger is used to determine the type of polar coordinate
    ' conversion used.  This should result in improved quality for both portrait and landscape aspect ratio images.
    
    If finalX > finalY Then
    
        'Because this function actually wraps three functions, calculating the progress bar maximum is a bit convoluted
        newProgBarMax = finalX * 3
    
        'Start by converting the image to polar coordinates, using a specific set of actions to maximize quality
        If CreatePolarCoordLayer(1, 100, EDGE_CLAMP, True, srcLayer, workingLayer, toPreview, newProgBarMax) Then
            
            'Now we can apply the box blur to the temporary layer, using the blur radius supplied by the user
            If CreateVerticalBlurLayer(bDistance, bDistance, workingLayer, srcLayer, toPreview, newProgBarMax, finalX) Then
                
                'Finally, convert back to rectangular coordinates, using the opposite parameters of the first conversion
                CreatePolarCoordLayer 0, 100, EDGE_CLAMP, True, srcLayer, workingLayer, toPreview, newProgBarMax, finalX + finalX
                
            End If
            
        End If
    
    Else
    
        'Because this function actually wraps three functions, calculating the progress bar maximum is a bit convoluted
        newProgBarMax = finalX * 2 + finalY
    
        'Start by converting the image to polar coordinates, using a specific set of actions to maximize quality
        If CreateXSwappedPolarCoordLayer(1, 100, EDGE_CLAMP, True, srcLayer, workingLayer, toPreview, newProgBarMax) Then
            
            'Now we can apply the box blur to the temporary layer, using the blur radius supplied by the user
            If CreateHorizontalBlurLayer(bDistance, bDistance, workingLayer, srcLayer, toPreview, newProgBarMax, finalX) Then
                
                'Finally, convert back to rectangular coordinates, using the opposite parameters of the first conversion
                CreateXSwappedPolarCoordLayer 0, 100, EDGE_CLAMP, True, srcLayer, workingLayer, toPreview, newProgBarMax, finalX + finalY
                
            End If
            
        End If
    
    End If
    
    srcLayer.eraseLayer
    Set srcLayer = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingLayer
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Zoom blur", , buildParams(OptStyle(0), sltDistance)
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()

    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the form is fully initialized
    cmdBar.markPreviewStatus False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Render a new effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ZoomBlurWrapper OptStyle(0), sltDistance, True, fxPreview
End Sub

'Modern style allows for zooming in and out.  Traditional only allows out.
Private Sub OptStyle_Click(Index As Integer)

    If OptStyle(0) Then
        sltDistance.Min = -200
    Else
        If sltDistance.Value < 0 Then sltDistance.Value = Abs(sltDistance.Value)
        sltDistance.Min = 0
    End If
    
    updatePreview

End Sub

Private Sub sltDistance_Change()
    updatePreview
End Sub
