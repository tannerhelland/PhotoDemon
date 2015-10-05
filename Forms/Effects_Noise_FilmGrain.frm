VERSION 5.00
Begin VB.Form FormFilmGrain 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Apply Film Grain"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12120
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
   ScaleWidth      =   808
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.sliderTextCombo sltNoise 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "strength"
      Min             =   1
      Max             =   50
      Value           =   10
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
      TabIndex        =   3
      Top             =   3000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "softness"
      Max             =   25
      SigDigits       =   1
      Value           =   5
   End
End
Attribute VB_Name = "FormFilmGrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Add Film Grain Tool
'Copyright 2013-2015 by Tanner Helland
'Created: 31/January/13
'Last updated: 11/January/14
'Last update: convert softness to floating point; minor performance improvements
'
'Tool for simulating film grain. For aesthetic reasons, film grain is restricted to monochromatic noise
' (luminance only) to better mimic traditional film grain.
'
'The separate standalone Gaussian Blur function is used for noise smoothing.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'When previewing, we need to modify the strength to be representative of the final filter. This means dividing by the
' original image dimensions in order to establish the right ratio.
Dim iWidth As Long, iHeight As Long

'Subroutine for adding noise to an image
' Inputs: Amount of noise, monochromatic or not, preview settings
Public Sub AddFilmGrain(ByVal gStrength As Double, ByVal gSoftness As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
        
    If Not toPreview Then Message "Generating film grain texture..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA, toPreview, dstPic
    
    'Create a separate source DIB. This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent adjusted pixel values from spreading across the image as we go.)
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.createFromExistingDIB workingDIB
    
    'Create a DIB to hold the gaussian blur
    Dim gaussDIB As pdDIB
    Set gaussDIB = New pdDIB
    
    'Create a DIB to hold the film grain
    Dim noiseDIB As pdDIB
    Set noiseDIB = New pdDIB
    noiseDIB.createFromExistingDIB workingDIB
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'Point an array at the noise DIB
    Dim dstImageData() As Byte
    prepSafeArray dstSA, noiseDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary. This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not toPreview Then
        SetProgBarMax finalY * 2 + finalX * 2
        progBarCheck = findBestProgBarValue()
    End If
        
    'Noise variables
    Dim nColor As Long
    Dim gStrength2 As Long
    
    'Double the amount of noise we plan on using (so we can add noise above or below the current color value)
    gStrength2 = gStrength * 2
    
    'Although it's slow, we're stuck using random numbers for noise addition. Seed the generator with a pseudo-random value.
    Randomize Timer
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
                    
        'Generate monochromatic noise, e.g. the same amount of noise for each color component, based around RGB(127, 127, 127)
        nColor = 127 + (gStrength2 * Rnd) - gStrength
        
        'Assign that noise to each color component
        dstImageData(QuickVal + 2, y) = nColor
        dstImageData(QuickVal + 1, y) = nColor
        dstImageData(QuickVal, y) = nColor
        
    Next y
        If Not toPreview Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our noise generation complete, point dstImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Next, we need to soften the noise DIB
    If (Not toPreview) And (Not cancelCurrentAction) Then Message "Softening film grain..."
    
    If (gSoftness > 0) And (Not cancelCurrentAction) Then
    
        'If this is a preview, we need to adjust the softening radius to match the size of the preview box
        If toPreview Then
            gSoftness = gSoftness * curDIBValues.previewModifier
            If gSoftness = 0 Then gSoftness = 0.1
        End If
    
        gaussDIB.createFromExistingDIB workingDIB
    
        'Blur the noise texture as required by the user
        CreateGaussianBlurDIB gSoftness, noiseDIB, gaussDIB, toPreview, finalY * 2 + finalX * 2, finalX
        
    Else
        gaussDIB.createFromExistingDIB noiseDIB
    End If
    
    'Delete the original noise DIB to conserve resources
    noiseDIB.eraseDIB
    Set noiseDIB = Nothing
    
    If Not cancelCurrentAction Then
    
        'We now have a softened noise DIB. Next, create three arrays - one pointing at the original image data, one pointing at
        ' the noise data, and one pointing at the destination data.
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
            
        If Not toPreview Then Message "Applying film grain to image..."
        
        Dim r As Long, g As Long, b As Long
        
        'The final step of the smart blur function is to find edges, and replace them with the blurred data as necessary
        For x = initX To finalX
            QuickVal = x * qvDepth
        For y = initY To finalY
            
            'Retrieve the original image's pixels
            r = srcImageData(QuickVal + 2, y)
            g = srcImageData(QuickVal + 1, y)
            b = srcImageData(QuickVal, y)
                    
            'Now, retrieve a noise pixel (we only need one, as each color component will be identical)
            nColor = GaussImageData(QuickVal, y) - 127
                    
            'Add the noise to each color component
            r = r + nColor
            g = g + nColor
            b = b + nColor
            
            If r > 255 Then r = 255
            If r < 0 Then r = 0
            If g > 255 Then g = 255
            If g < 0 Then g = 0
            If b > 255 Then b = 255
            If b < 0 Then b = 0
            
            dstImageData(QuickVal + 2, y) = r
            dstImageData(QuickVal + 1, y) = g
            dstImageData(QuickVal, y) = b
            
        Next y
            If Not toPreview Then
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal finalX + x + finalY + finalY
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
        
        srcDIB.eraseDIB
        Set srcDIB = Nothing
        
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
        Erase dstImageData
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Add film grain", , buildParams(sltNoise.Value, sltRadius.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Render a preview
    updatePreview
    
End Sub

Private Sub Form_Load()
        
    'Note the current image's width and height, which will be needed to adjust the preview effect
    If pdImages(g_CurrentImage).selectionActive Then
        iWidth = pdImages(g_CurrentImage).mainSelection.boundWidth
        iHeight = pdImages(g_CurrentImage).mainSelection.boundHeight
    Else
        iWidth = pdImages(g_CurrentImage).Width
        iHeight = pdImages(g_CurrentImage).Height
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltNoise_Change()
    updatePreview
End Sub

Private Sub sltRadius_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then AddFilmGrain sltNoise, sltRadius, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub
