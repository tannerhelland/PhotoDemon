VERSION 5.00
Begin VB.Form FormIgnite 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Ignite (fire effect)"
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
   Begin PhotoDemon.pdSlider sltIntensity 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "color intensity"
      Min             =   1
      SigDigits       =   1
      Value           =   5
      NotchPosition   =   2
      NotchValueCustom=   5
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
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "flame height"
      Min             =   1
      Max             =   500
      Value           =   50
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdSlider sltOpacity 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   3720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "strength"
      Min             =   1
      Max             =   100
      Value           =   50
      DefaultValue    =   50
   End
End
Attribute VB_Name = "FormIgnite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'"Burn" Fire FX Form
'Copyright 2001-2017 by Tanner Helland
'Created: some time 2001
'Last updated: 09/July/14
'Last update: give tool its own form; overhaul algorithm completely
'
'This fun little tool is a product of my own creation.  It works as follows:
'
' 1) Analyze image edges and create a contour map
' 2) For each pixel in the image, blur it upward at a distance relative to its luminance.  Apply linear decay as
'     each pixel is faded upward.
' 3) Recolor the image from (2) according to the user's intensity value; higher intensity warms the colors more.
' 4) Composite the final result against the original image, using the SCREEN blend mode (so that black pixels
'     are ignored, and bright pixels are maintained).
'
'While a "fire" effect has existed in PD for many years, it didn't receive its own dialog until July '14.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Apply the "burn" fire effect filter
'Input: strength of the filter (min 1, no real max - but above 7 it becomes increasingly blown-out)
Public Sub fxBurn(ByVal fxIntensity As Double, ByVal fxRadius As Long, ByVal fxOpacity As Long, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Lighting image on fire..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    PrepImageData tmpSA, toPreview, dstPic
    
    'Radius needs to be adjusted during previews, to accurately reflect how the final image will appear
    If toPreview Then
        fxRadius = fxRadius * curDIBValues.previewModifier
        If (fxRadius < 1) Then fxRadius = 1
    Else
        SetProgBarMax workingDIB.GetDIBWidth * 3
    End If
    
    'First things first: start by analyzing image edges and generating a white-on-black contour map
    Dim edgeDIB As pdDIB
    Set edgeDIB = New pdDIB
    edgeDIB.CreateFromExistingDIB workingDIB
    Filters_Layers.CreateContourDIB True, workingDIB, edgeDIB, toPreview, workingDIB.GetDIBWidth * 3, 0
    
    'Next, we're going to do two things: blurring the flame upward, while also applying some decay
    ' to the flame.
    PrepSafeArray tmpSA, edgeDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = FindBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Byte
    
    'Build a look-up table of grayscale values (faster than calculating it manually for each pixel)
    Dim grayLookUp(0 To 765) As Byte
    For x = 0 To 765
        grayLookUp(x) = x \ 3
    Next x
    
    Dim distLookUp(0 To 765) As Single
    For x = 0 To 765
        distLookUp(x) = CDbl(x / 382) * fxRadius
    Next x
    
    Dim fDistance As Long, fTargetMin As Long, innerY As Long
    Dim inR As Byte, inG As Byte, inB As Byte
    Dim fadeVal As Double
    
    'Loop through each pixel in the image, applying flame decay as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        b = ImageData(quickVal, y)
        g = ImageData(quickVal + 1, y)
        r = ImageData(quickVal + 2, y)
        
        'Calculate a distance value using our precalculated look-up values.  Basically, this is the max distance
        ' a flame can travel, and it's directly tied to the pixel's luminance (brighter pixels travel further).
        fDistance = distLookUp(r + g + b)
        
        'Calculate an upper bound
        fTargetMin = y - fDistance
        If (fTargetMin < 0) Then
            fTargetMin = 0
            fDistance = y
        End If
        
        'Trace a path upward, blending this value with neighboring pixels as we go
        If (fDistance > 0) Then
        
            For innerY = y To fTargetMin Step -1
                
                inB = ImageData(quickVal, innerY)
                inG = ImageData(quickVal + 1, innerY)
                inR = ImageData(quickVal + 2, innerY)
                
                'Blend this pixel's value with the value at this pixel, using the distance traveled as our blend metric
                fadeVal = (innerY - fTargetMin) / fDistance
                
                ImageData(quickVal, innerY) = BlendColors(inB, b, fadeVal)
                ImageData(quickVal + 1, innerY) = BlendColors(inG, g, fadeVal)
                ImageData(quickVal + 2, innerY) = BlendColors(inR, r, fadeVal)
                
            Next innerY
        
        End If
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If UserPressedESC() Then Exit For
                SetProgBarVal finalX + x
            End If
        End If
    Next x
    
    'Loop through the contour map one final time, recolor pixels to flame-like warm colors
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        b = ImageData(quickVal, y)
        g = ImageData(quickVal + 1, y)
        r = ImageData(quickVal + 2, y)
        
        'Calculate the gray value using the look-up table
        grayVal = grayLookUp(r + g + b)
        
        'Perform the fire conversion
        r = grayVal * fxIntensity
        If (r > 255) Then r = 255
        g = grayVal
        b = grayVal \ fxIntensity
        
        'Assign the new "fire" value to each color channel
        ImageData(quickVal, y) = b
        ImageData(quickVal + 1, y) = g
        ImageData(quickVal + 2, y) = r
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If UserPressedESC() Then Exit For
                SetProgBarVal finalX * 2 + x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Apply premultiplication prior to compositing
    edgeDIB.SetAlphaPremultiplication True
    workingDIB.SetAlphaPremultiplication True
    
    'A pdCompositor class will help us selectively blend the flame results back onto the main image
    Dim cComposite As pdCompositor
    Set cComposite = New pdCompositor
    cComposite.QuickMergeTwoDibsOfEqualSize workingDIB, edgeDIB, BL_SCREEN, fxOpacity
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    FinalizeImageData toPreview, dstPic, True

End Sub

Private Sub cmdBar_OKClick()
    Process "Ignite", , BuildParams(sltIntensity, sltRadius, sltOpacity), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltIntensity.Value = 5
    sltRadius = 50
    sltOpacity.Value = 50
End Sub

Private Sub Form_Load()
    cmdBar.MarkPreviewStatus False
    ApplyThemeAndTranslations Me
    cmdBar.MarkPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

'Update the preview whenever the combination slider/text control has its value changed
Private Sub sltIntensity_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then fxBurn sltIntensity, sltRadius, sltOpacity, True, pdFxPreview
End Sub

Private Sub sltOpacity_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    With cParams
    
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
