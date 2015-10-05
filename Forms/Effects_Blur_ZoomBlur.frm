VERSION 5.00
Begin VB.Form FormZoomBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Zoom blur"
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
   Begin PhotoDemon.buttonStrip btsStyle 
      Height          =   615
      Left            =   6180
      TabIndex        =   4
      Top             =   2160
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   1085
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
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
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.sliderTextCombo sltDistance 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   3120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "distance"
      Min             =   -200
      Max             =   200
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "style"
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
      TabIndex        =   3
      Top             =   1800
      Width           =   480
   End
End
Attribute VB_Name = "FormZoomBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Zoom Blur Tool
'Copyright 2013-2015 by Tanner Helland
'Created: 27/August/13
'Last updated: 25/September/14
'Last update: improve traditional mode output, and allow negative zoom values
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
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'AlphaBlend is used to render a fast zoom blur estimation
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean

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
    
    'Call prepImageData, which will initialize a workingDIB object for us (with all selection tool masks applied)
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    Dim finalX As Long, finalY As Long
    finalX = workingDIB.getDIBWidth
    finalY = workingDIB.getDIBHeight
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not toPreview Then
        SetProgBarMax Abs(zDistance)
        progBarCheck = findBestProgBarValue()
    End If
    
    'AlphaBlend has two limitations we have to work around: it only stretches using nearest-neighbor interpolation (which
    ' creates very obvious blocking along the lines where the algorithm stretches), and it is incapable of rendering from
    ' a source DC to a matching dest DC.  Thus we will make our own a temporary copy of the image, which we will use to
    ' work around these two issues.
    Dim tmpSrcDIB As pdDIB
    Set tmpSrcDIB = New pdDIB
    tmpSrcDIB.createFromExistingDIB workingDIB
    
    'Next, we set up some AlphaBlend parameters.  If the image already has alpha data, we want to make use of it, although
    ' it will result in a funky final image (I may work around this in the future, but right now I don't care).  If the
    ' image is 24bpp, we'll use halftoning to generate a higher-quality zoom.
    Dim blendVal As Long
    blendVal = 127 * &H10000
    If workingDIB.getDIBColorDepth = 32 Then
        SetStretchBltMode tmpSrcDIB.getDIBDC, STRETCHBLT_COLORONCOLOR
    Else
        SetStretchBltMode tmpSrcDIB.getDIBDC, STRETCHBLT_HALFTONE
    End If
    
    'Finally, we're going to use the image's aspect ratio to try and generate a more realistic zoom blur.  Calculate
    ' that ratio now.
    Dim aspectRatio As Double
    If workingDIB.getDIBWidth > workingDIB.getDIBHeight Then
        aspectRatio = workingDIB.getDIBHeight / workingDIB.getDIBWidth
    Else
        aspectRatio = workingDIB.getDIBWidth / workingDIB.getDIBHeight
    End If
    
    'Zoom distance must be adjusted during a preview, so that the preview accurately represents the finished product.
    If toPreview Then zDistance = zDistance * curDIBValues.previewModifier
    
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
            If workingDIB.getDIBWidth > workingDIB.getDIBHeight Then
                StretchBlt tmpSrcDIB.getDIBDC, 0, 0, finalX, finalY, workingDIB.getDIBDC, zoomOffset, zoomOffset * aspectRatio, finalX - zoomOffset * 2, finalY - ((zoomOffset * aspectRatio) * 2), vbSrcCopy
            Else
                StretchBlt tmpSrcDIB.getDIBDC, 0, 0, finalX, finalY, workingDIB.getDIBDC, zoomOffset * aspectRatio, zoomOffset, finalX - ((zoomOffset * aspectRatio) * 2), finalY - zoomOffset * 2, vbSrcCopy
            End If
            
            'Finally, AlphaBlend the modified DIB onto the working DIB, then rinse and repeat!
            AlphaBlend workingDIB.getDIBDC, 0, 0, finalX, finalY, tmpSrcDIB.getDIBDC, 0, 0, finalX, finalY, blendVal
            
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
            If workingDIB.getDIBWidth > workingDIB.getDIBHeight Then
                StretchBlt tmpSrcDIB.getDIBDC, zoomOffset, zoomOffset * aspectRatio, finalX - zoomOffset * 2, finalY - ((zoomOffset * aspectRatio) * 2), workingDIB.getDIBDC, 0, 0, finalX, finalY, vbSrcCopy
            Else
                StretchBlt tmpSrcDIB.getDIBDC, zoomOffset * aspectRatio, zoomOffset, finalX - ((zoomOffset * aspectRatio) * 2), finalY - zoomOffset * 2, workingDIB.getDIBDC, 0, 0, finalX, finalY, vbSrcCopy
            End If
            
            'Finally, AlphaBlend the modified DIB onto the working DIB, then rinse and repeat!
            AlphaBlend workingDIB.getDIBDC, 0, 0, finalX, finalY, tmpSrcDIB.getDIBDC, 0, 0, finalX, finalY, blendVal
            
            If Not toPreview Then
                If (Abs(i) And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal Abs(i)
                End If
            End If
        
        Next i
    
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
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
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    'By dividing blur distance by 200 (its maximum value), we can use it as a fractional amount to determine the strength of our horizontal blur.
    If toPreview Then bDistance = bDistance * curDIBValues.previewModifier
    
    Dim finalX As Long, finalY As Long
    finalX = workingDIB.getDIBWidth
    finalY = workingDIB.getDIBHeight
    
    Dim newProgBarMax As Long
    
    'Negative and positive zooms are both allowed; these trigger whether we apply a forward or reverse horizontal/vertical blur.
    Dim forwardBlurDistance As Long, backwardBlurDistance As Long
    
    If bDistance < 0 Then
        forwardBlurDistance = Abs(bDistance)
        backwardBlurDistance = 0
    Else
        forwardBlurDistance = 0
        backwardBlurDistance = Abs(bDistance)
    End If
    
    'Zoom blur basically works by converting an image to polar coordinates, applying a horizontal (or vertical) blur,
    ' then converting back to rectangular coordinates.  Even with interpolation, the two coordinate conversion functions
    ' result in a loss of image data, so I've gone to some lengths to try and mitigate this.
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
        If CreatePolarCoordDIB(1, 100, EDGE_CLAMP, True, srcDIB, workingDIB, toPreview, newProgBarMax) Then
            
            'Now we can apply the box blur to the temporary DIB, using the blur radius supplied by the user
            If CreateVerticalBlurDIB(backwardBlurDistance, forwardBlurDistance, workingDIB, srcDIB, toPreview, newProgBarMax, finalX) Then
                
                'Finally, convert back to rectangular coordinates, using the opposite parameters of the first conversion
                CreatePolarCoordDIB 0, 100, EDGE_CLAMP, True, srcDIB, workingDIB, toPreview, newProgBarMax, finalX + finalX
                
            End If
            
        End If
    
    Else
    
        'Because this function actually wraps three functions, calculating the progress bar maximum is a bit convoluted
        newProgBarMax = finalX * 2 + finalY
    
        'Start by converting the image to polar coordinates, using a specific set of actions to maximize quality
        If CreateXSwappedPolarCoordDIB(1, 100, EDGE_CLAMP, True, srcDIB, workingDIB, toPreview, newProgBarMax) Then
            
            'Now we can apply the box blur to the temporary DIB, using the blur radius supplied by the user
            If CreateHorizontalBlurDIB(backwardBlurDistance, forwardBlurDistance, workingDIB, srcDIB, toPreview, newProgBarMax, finalX) Then
                
                'Finally, convert back to rectangular coordinates, using the opposite parameters of the first conversion
                CreateXSwappedPolarCoordDIB 0, 100, EDGE_CLAMP, True, srcDIB, workingDIB, toPreview, newProgBarMax, finalX + finalY
                
            End If
            
        End If
    
    End If
    
    srcDIB.eraseDIB
    Set srcDIB = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    finalizeImageData toPreview, dstPic
    
End Sub

'Modern style allows for zooming in and out.  Traditional only allows out.
Private Sub btsStyle_Click(ByVal buttonIndex As Long)
    updatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Zoom blur", , buildParams((btsStyle.ListIndex = 0), sltDistance), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()

    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the form is fully initialized
    cmdBar.markPreviewStatus False
    
    'Add style options to the button strip
    btsStyle.AddItem "modern", 0
    btsStyle.AddItem "traditional", 1
    btsStyle.ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Render a new effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ZoomBlurWrapper (btsStyle.ListIndex = 0), sltDistance, True, fxPreview
End Sub

Private Sub sltDistance_Change()
    updatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

