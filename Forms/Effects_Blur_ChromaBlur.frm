VERSION 5.00
Begin VB.Form FormChromaBlur 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Chroma blur"
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
   End
   Begin PhotoDemon.sliderTextCombo sltRadius 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "blur radius"
      Min             =   0.1
      Max             =   200
      SigDigits       =   1
      Value           =   5
   End
   Begin PhotoDemon.smartOptionButton OptQuality 
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   3
      Top             =   3150
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   582
      Caption         =   "good"
      Value           =   -1  'True
   End
   Begin PhotoDemon.smartOptionButton OptQuality 
      Height          =   360
      Index           =   1
      Left            =   6120
      TabIndex        =   4
      Top             =   3570
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   582
      Caption         =   "better"
   End
   Begin PhotoDemon.smartOptionButton OptQuality 
      Height          =   360
      Index           =   2
      Left            =   6120
      TabIndex        =   5
      Top             =   3990
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   582
      Caption         =   "best"
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "quality"
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
      Top             =   2760
      Width           =   705
   End
End
Attribute VB_Name = "FormChromaBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Chroma (Color) Blur Tool
'Copyright 2014-2015 by Tanner Helland
'Created: 11/January/14
'Last updated: 11/January/14
'Last update: initial build
'
'Chroma blur is a useful tool for improving noise in low-quality digital photos (especially image taken with a phone).
' It blurs color data only - not luminance - thus leaving image edges intact while smoothing out regions of mixed
' color.  The results can be fine-tuned more easily than something like a median function.
'
'Despite many optimizations, this function is quite slow in the IDE.  I STRONGLY recommend compiling the project before
' applying any actions at a large radius.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Track the active option button, so we can more easily pass it as a parameter when the user clicks OK
Private qualityIndex As Long

'Selectively blur just the chroma (color) data in an image, but not the luminance.  Very helpful for removing noise,
' particularly in digital photos.
'Inputs: radius of the blur (min 1, no real max - but processing speed obviously drops as the radius increases)
'        quality of the blur (gaussian approximations - fast, lower quality - vs an actual gaussian - slow, excellent quality)
Public Sub ChromaBlurFilter(ByVal gRadius As Double, Optional ByVal gaussQuality As Long = 2, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Blurring chroma (color) data..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, which we need to retrieve luminance
    ' values when merging the blurred color data with the original luminance data.
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    Dim gaussDIB As pdDIB
    Set gaussDIB = New pdDIB
    gaussDIB.createFromExistingDIB workingDIB
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        gRadius = gRadius * curDIBValues.previewModifier
        If gRadius = 0 Then gRadius = 0.1
    End If
    
    Dim blurSuccess As Long
    
    Dim calcProgBarMax As Long, calcProgBarOffset As Long
    
    'The quality parameter we were passed will be used to determine how we blur the image.
    Select Case gaussQuality
    
        '3 iteration box blur
        Case 0
            calcProgBarMax = finalX * 4 + finalY * 3
            calcProgBarOffset = finalX * 3 + finalY * 3
            blurSuccess = CreateApproximateGaussianBlurDIB(gRadius, srcDIB, gaussDIB, 3, toPreview, calcProgBarMax)
        
        '5 iteration box blur
        Case 1
            calcProgBarMax = finalX * 6 + finalY * 5
            calcProgBarOffset = finalX * 5 + finalY * 5
            blurSuccess = CreateApproximateGaussianBlurDIB(gRadius, srcDIB, gaussDIB, 5, toPreview, calcProgBarMax)
        
        'True Gaussian
        Case 2
            calcProgBarMax = finalX + finalY * 2
            calcProgBarOffset = finalY * 2
            blurSuccess = CreateGaussianBlurDIB(gRadius, srcDIB, gaussDIB, toPreview, calcProgBarMax)
            
    End Select
    
    'If the previous blur step was successful (e.g. the user did NOT cancel it), continue with the chroma blur.
    If blurSuccess Then
            
        'Point arrays at three images: the source and gauss DIBs, and the final destination DIB
        Dim dstImageData() As Byte
        prepImageData dstSA, toPreview, dstPic
        CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
        Dim srcImageData() As Byte
        Dim srcSA As SAFEARRAY2D
        prepSafeArray srcSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
            
        Dim GaussImageData() As Byte
        Dim gaussSA As SAFEARRAY2D
        prepSafeArray gaussSA, gaussDIB
        CopyMemory ByVal VarPtrArray(GaussImageData()), VarPtr(gaussSA), 4
                
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim QuickVal As Long, qvDepth As Long
        qvDepth = curDIBValues.BytesPerPixel
        
        'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
        ' based on the size of the area to be processed.
        Dim progBarCheck As Long
        progBarCheck = findBestProgBarValue()
            
        If Not toPreview Then Message "Merging luminance and chroma into final image..."
        
        'More color variables - in this case, sums for each color component
        Dim r As Long, g As Long, b As Long
        Dim h As Double, s As Double, l As Double
        Dim origLuminance As Double
        
        'The final step of the chroma blur function is to merge blurred color data with original luminance data
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
            
            'Retrieve the original image's pixels
            r = srcImageData(QuickVal + 2, y)
            g = srcImageData(QuickVal + 1, y)
            b = srcImageData(QuickVal, y)
            
            'Determine original HSL values
            tRGBToHSL r, g, b, h, s, origLuminance
            
            'Now, retrieve the gaussian pixels
            r = GaussImageData(QuickVal + 2, y)
            g = GaussImageData(QuickVal + 1, y)
            b = GaussImageData(QuickVal, y)
            
            'Determine HSL for the blurred data
            tRGBToHSL r, g, b, h, s, l
            
            'Use the final hue and saturation values but the ORIGINAL luminance value to create a new RGB coordinate
            tHSLToRGB h, s, origLuminance, r, g, b
            
            'Apply the new RGB colors to the image
            dstImageData(QuickVal + 2, y) = r
            dstImageData(QuickVal + 1, y) = g
            dstImageData(QuickVal, y) = b
            If qvDepth = 4 Then dstImageData(QuickVal + 3, y) = srcImageData(QuickVal + 3, y)
            
        Next y
            If Not toPreview Then
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal x + calcProgBarOffset
                End If
            End If
        Next x
            
        'With our work complete, release all arrays
        CopyMemory ByVal VarPtrArray(GaussImageData), 0&, 4
        Erase GaussImageData
        
        gaussDIB.eraseDIB
        Set gaussDIB = Nothing
        
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        Erase srcImageData
        
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
        Erase dstImageData
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Chroma blur", , buildParams(sltRadius, qualityIndex), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 1
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Draw a preview of the effect
    cmdBar.markPreviewStatus True
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully loaded
    cmdBar.markPreviewStatus False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub OptQuality_Click(Index As Integer)
    qualityIndex = Index
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

'Render a new effect preview
Private Sub updatePreview()
    If cmdBar.previewsAllowed Then ChromaBlurFilter sltRadius, qualityIndex, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

