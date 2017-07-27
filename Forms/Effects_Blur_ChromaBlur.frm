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
   Begin PhotoDemon.pdButtonStrip btsQuality 
      Height          =   975
      Left            =   6000
      TabIndex        =   3
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      Caption         =   "mode"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
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
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
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
      DefaultValue    =   5
   End
End
Attribute VB_Name = "FormChromaBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Chroma (Color) Blur Tool
'Copyright 2014-2017 by Tanner Helland
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

'Selectively blur just the chroma (color) data in an image, but not the luminance.  Very helpful for removing noise,
' particularly in digital photos.
'Inputs: radius of the blur (min 1, no real max - but processing speed obviously drops as the radius increases)
'        quality of the blur (gaussian approximations - fast, lower quality - vs an actual gaussian - slow, excellent quality)
Public Sub ChromaBlurFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Blurring chroma (color) data..."
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString effectParams
    
    Dim gRadius As Double, gaussQuality As Long
    gRadius = cParams.GetDouble("radius", 1#)
    gaussQuality = cParams.GetLong("quality", 2)
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    PrepImageData dstSA, toPreview, dstPic
    
    'Create a second local array.  This will contain the a copy of the current image, which we need to retrieve luminance
    ' values when merging the blurred color data with the original luminance data.
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    Dim gaussDIB As pdDIB
    Set gaussDIB = New pdDIB
    gaussDIB.CreateFromExistingDIB workingDIB
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        gRadius = gRadius * curDIBValues.previewModifier
        If (gRadius <= 0.1) Then gRadius = 0.1
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
        Case Else
            calcProgBarMax = finalX * 6 + finalY * 5
            calcProgBarOffset = finalX * 5 + finalY * 5
            blurSuccess = CreateApproximateGaussianBlurDIB(gRadius, srcDIB, gaussDIB, 5, toPreview, calcProgBarMax)
        
    End Select
    
    'If the previous blur step was successful (e.g. the user did NOT cancel it), continue with the chroma blur.
    If blurSuccess Then
            
        'Point arrays at three images: the source and gauss DIBs, and the final destination DIB
        Dim dstImageData() As Byte
        PrepImageData dstSA, toPreview, dstPic
        CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
        Dim srcImageData() As Byte
        Dim srcSA As SAFEARRAY2D
        PrepSafeArray srcSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
            
        Dim GaussImageData() As Byte
        Dim gaussSA As SAFEARRAY2D
        PrepSafeArray gaussSA, gaussDIB
        CopyMemory ByVal VarPtrArray(GaussImageData()), VarPtr(gaussSA), 4
                
        'These values will help us access locations in the array more quickly.
        ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
        Dim quickVal As Long, qvDepth As Long
        qvDepth = curDIBValues.BytesPerPixel
        
        'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
        ' based on the size of the area to be processed.
        Dim progBarCheck As Long
        progBarCheck = FindBestProgBarValue()
            
        If (Not toPreview) Then Message "Merging luminance and chroma into final image..."
        
        'More color variables - in this case, sums for each color component
        Dim r As Long, g As Long, b As Long
        Dim h As Double, s As Double, l As Double
        Dim origLuminance As Double
        
        'The final step of the chroma blur function is to merge blurred color data with original luminance data
        For x = initX To finalX
            quickVal = x * qvDepth
        For y = initY To finalY
            
            'Retrieve the original image's pixels
            b = srcImageData(quickVal, y)
            g = srcImageData(quickVal + 1, y)
            r = srcImageData(quickVal + 2, y)
            
            'Determine original HSL values
            tRGBToHSL r, g, b, h, s, origLuminance
            
            'Now, retrieve the gaussian pixels
            b = GaussImageData(quickVal, y)
            g = GaussImageData(quickVal + 1, y)
            r = GaussImageData(quickVal + 2, y)
            
            'Determine HSL for the blurred data
            tRGBToHSL r, g, b, h, s, l
            
            'Use the final hue and saturation values but the ORIGINAL luminance value to create a new RGB coordinate
            tHSLToRGB h, s, origLuminance, r, g, b
            
            'Apply the new RGB colors to the image
            dstImageData(quickVal, y) = b
            dstImageData(quickVal + 1, y) = g
            dstImageData(quickVal + 2, y) = r
            If (qvDepth = 4) Then dstImageData(quickVal + 3, y) = srcImageData(quickVal + 3, y)
            
        Next y
            If (Not toPreview) Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal x + calcProgBarOffset
                End If
            End If
        Next x
            
        'With our work complete, release all arrays
        CopyMemory ByVal VarPtrArray(GaussImageData), 0&, 4
        Set gaussDIB = Nothing
        
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub btsQuality_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Chroma blur", , GetLocalParamString, UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 1
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully loaded
    cmdBar.MarkPreviewStatus False
    
    btsQuality.AddItem "fast", 0
    btsQuality.AddItem "precise", 1
    btsQuality.ListIndex = 0
        
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Draw a preview of the effect
    cmdBar.MarkPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

'Render a new effect preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ChromaBlurFilter GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.AddParam "radius", sltRadius.Value
    cParams.AddParam "quality", btsQuality.ListIndex
    GetLocalParamString = cParams.GetParamString()
End Function
