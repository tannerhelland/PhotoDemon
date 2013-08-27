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
      Top             =   2640
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
      Top             =   2280
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
'Last updated: 27/August/13
'Last update: initial build
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

'When previewing, we need to modify the strength to be representative of the final filter. This means dividing by the
' original image dimensions in order to establish the right ratio.
Dim iWidth As Long, iHeight As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Apply motion blur to an image
'Inputs: angle of the blur, distance of the blur
Public Sub ZoomBlurFilter(ByVal zDistance As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Applying zoom blur..."
    
    'Call prepImageData, which will initialize a workingLayer object for us (with all selection tool masks applied)
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        If iWidth > iHeight Then
            zDistance = (zDistance / iWidth) * curLayerValues.Width
        Else
            zDistance = (zDistance / iHeight) * curLayerValues.Height
        End If
        If zDistance = 0 Then zDistance = 1
    End If
    
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
    If iWidth > iHeight Then
        aspectRatio = iHeight / iWidth
    Else
        aspectRatio = iWidth / iHeight
    End If
    
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
            If iWidth > iHeight Then
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
            If iWidth > iHeight Then
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

Private Sub cmdBar_OKClick()
    Process "Zoom blur", , buildParams(sltDistance)
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
    
    'Note the current image's width and height, which will be needed to adjust the preview effect
    If pdImages(CurrentImage).selectionActive Then
        iWidth = pdImages(CurrentImage).mainSelection.boundWidth
        iHeight = pdImages(CurrentImage).mainSelection.boundHeight
    Else
        iWidth = pdImages(CurrentImage).Width
        iHeight = pdImages(CurrentImage).Height
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Render a new effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ZoomBlurFilter sltDistance, True, fxPreview
End Sub

Private Sub sltDistance_Change()
    updatePreview
End Sub
