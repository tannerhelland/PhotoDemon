VERSION 5.00
Begin VB.Form FormTransparency_FromColor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Make color transparent"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11820
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   788
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11820
      _ExtentX        =   20849
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
      ColorSelection  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltErase 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2640
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1270
      Caption         =   "erase threshold"
      Max             =   199
      Value           =   15
      DefaultValue    =   15
   End
   Begin PhotoDemon.pdSlider sltBlend 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   3600
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1270
      Caption         =   "edge blending"
      Max             =   200
      Value           =   15
      DefaultValue    =   15
   End
   Begin PhotoDemon.pdColorSelector csSource 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   1320
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1931
      Caption         =   "color to erase (right-click preview to select)"
      curColor        =   49152
   End
End
Attribute VB_Name = "FormTransparency_FromColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Make color transparent ("green screen") tool dialog
'Copyright 2013-2026 by Tanner Helland
'Created: 13/August/13
'Last updated: 10/June/16
'Last update: add a LittleCMS path for the algorithm; this improves performance by ~30%
'
'PhotoDemon has long provided the ability to convert a 24bpp image to 32bpp, but the lack of an interface meant it could
' only add a fully opaque alpha channel.  Now the user can select from one of several conversion methods.
'
'This dialog present one of the more interesting conversion methods: a "color to alpha" technique, which allows for
' powerful green-screen capabilities.  A full CieLAB color space transformation is used, and an optional blend parameter
' will antialias and color-correct edges for maximum smoothness.  I don't know of any other software that utilizes this
' dual-threshold approach, and in my own testing, I have found PD to be superior to any other open source package at
' removing complex background colors.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'OK button
Private Sub cmdBar_OKClick()
    Process "Color to alpha", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    csSource.Color = RGB(0, 192, 0)
    sltErase.Value = 15
    sltBlend.Value = 15
End Sub

Private Sub csSource_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()
    cmdBar.SetPreviewStatus False
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The user can select a color from the preview window; this helps green screen calculation immensely
Private Sub pdFxPreview_ColorSelected()
    csSource.Color = pdFxPreview.SelectedColor
    UpdatePreview
End Sub

'Add transparency to an image by making a specified color transparent (chroma-key or "green screen").
' This function uses a high-quality color-matching scheme in the L*a*b* color space.
' LittleCMS is used for transforms, if present.
Public Sub ColorToAlpha(ByVal processParameters As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    
    Dim targetColor As Long
    Dim eraseThreshold As Single, blendThreshold As Single
    
    With cParams
        targetColor = .GetLong("color", vbBlack)
        eraseThreshold = .GetSingle("erase-threshold", 15!)
        blendThreshold = .GetSingle("edge-blending", 30!)
    End With
    
    If (Not toPreview) Then Message "Adding new alpha channel to image..."
    
    'Call prepImageData, which will prepare a temporary copy of the image
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'For this function to work, each pixel needs to be RGBA, 32-bpp
    Dim pxSize As Long
    pxSize = workingDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long, a As Long
    
    'R2/G2/B2 store the RGB values of the color we are attempting to remove
    Dim r2 As Long, g2 As Long, b2 As Long
    r2 = Colors.ExtractRed(targetColor)
    g2 = Colors.ExtractGreen(targetColor)
    b2 = Colors.ExtractBlue(targetColor)
    
    'For maximum quality, we will apply our color comparison in the L*a*b* color space; each scanline will be
    ' transformed to L*a*b* all at once, for performance reasons
    Dim labValues() As Single
    ReDim labValues(0 To finalX * pxSize + pxSize) As Single
    
    'Used with internal LAB functions which require double-precision
    Dim labL As Double, labA As Double, labB As Double
    Dim labL2 As Double, labA2 As Double, labB2 As Double
    
    'Used with LCMS which supports single-level precision
    Dim labL2f As Single, labA2f As Single, labB2f As Single
    
    Dim labTransform As pdLCMSTransform
    Dim useLCMS As Boolean
    useLCMS = PluginManager.IsPluginCurrentlyEnabled(CCP_LittleCMS)
    
    'Calculate the L*a*b* values of the color to be removed
    If useLCMS Then
    
        'If LittleCMS is available, we're going to use it to perform the whole damn L*a*b* transform.
        Set labTransform = New pdLCMSTransform
        labTransform.CreateRGBAToLabTransform , True, INTENT_PERCEPTUAL, 0&
        
        Dim rgbBytes() As Byte
        ReDim rgbBytes(0 To 3) As Byte
        rgbBytes(0) = b2: rgbBytes(1) = g2: rgbBytes(2) = r2
        
        Dim labBytes() As Single
        ReDim labBytes(0 To 3) As Single
        labTransform.ApplyTransformToScanline VarPtr(rgbBytes(0)), VarPtr(labBytes(0)), 1
        
        labL2f = labBytes(0)
        labA2f = labBytes(1)
        labB2f = labBytes(2)
        
    Else
        RGBtoLAB r2, g2, b2, labL2, labA2, labB2
        labL2f = labL2
        labA2f = labA2
        labB2f = labB2
    End If
    
    'The blend threshold is used to "smooth" the edges of the green screen.  Calculate the difference between
    ' the erase and the blend thresholds in advance.
    Dim difThreshold As Single
    blendThreshold = eraseThreshold + blendThreshold
    difThreshold = blendThreshold - eraseThreshold
    If (difThreshold <> 0!) Then difThreshold = 1! / difThreshold
    
    Const ONE_DIV_255 As Double = 1# / 255#
    
    Dim cDistance As Single, cDistanceDenom As Single
    Dim newAlpha As Long
    
    'To improve performance of our horizontal loop, we'll move through bytes an entire pixel at a time
    Dim xStart As Long, xStop As Long
    xStart = initX * pxSize
    xStop = finalX * pxSize
        
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        
        'Wrap an array around the current scanline
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
        
        'Start by pre-calculating all L*a*b* values for this row
        If useLCMS Then
            labTransform.ApplyTransformToScanline VarPtr(imageData(0)), VarPtr(labValues(0)), finalX + 1
        Else
            For x = xStart To xStop Step pxSize
                b = imageData(x)
                g = imageData(x + 1)
                r = imageData(x + 2)
                RGBtoLAB r, g, b, labL, labA, labB
                labValues(x) = labL
                labValues(x + 1) = labA
                labValues(x + 2) = labB
            Next x
        End If
        
        'With all lab values pre-calculated, we can quickly step through each pixel, calculating distances as we go
        For x = xStart To xStop Step pxSize
        
            'Get the source pixel color values
            b = imageData(x)
            g = imageData(x + 1)
            r = imageData(x + 2)
            a = imageData(x + 3)
            
            'Perform a basic distance calculation (not ideal, but faster than a completely correct comparison;
            ' see https://en.wikipedia.org/wiki/Color_difference for a full report)
            If useLCMS Then
                cDistance = PDMath.Distance3D_FastFloat(labValues(x), labValues(x + 1), labValues(x + 2), labL2f, labA2f, labB2f)
            Else
                cDistance = PDMath.DistanceThreeDimensions(labValues(x), labValues(x + 1), labValues(x + 2), labL2, labA2, labB2)
            End If
            
            'If the distance is below the erasure threshold, remove it completely
            If (cDistance < eraseThreshold) Then
                imageData(x + 3) = 0
                
            'If the color is between the erasure and blend threshold, feather it against a partial alpha and
            ' color-correct it to remove any "color fringing" from the removed color.
            ElseIf (cDistance < blendThreshold) Then
                
                'Use a ^2 curve to improve blending response
                cDistance = (blendThreshold - cDistance) * difThreshold
                cDistance = cDistance * cDistance
                
                'Calculate a new alpha value for this pixel, based on its distance from the threshold.  Large
                ' distances from the removed color are made less transparent than small distances.
                newAlpha = 255 - (cDistance * 255)
                
                'Feathering the alpha often isn't enough to fully remove the color fringing caused by the removed
                ' background color, which will have "infected" the core RGB values.  Attempt to correct this by
                ' subtracting the target color from the original color, using the calculated threshold value; this
                ' is the only way I know to approximate the "feathering" caused by light bleeding over object edges.
                If (cDistance > 0.999999) Then cDistance = 0.999999
                cDistanceDenom = 1! / (1! - cDistance)
                r = (r - (r2 * cDistance)) * cDistanceDenom
                g = (g - (g2 * cDistance)) * cDistanceDenom
                b = (b - (b2 * cDistance)) * cDistanceDenom
                
                If (r > 255) Then r = 255
                If (g > 255) Then g = 255
                If (b > 255) Then b = 255
                If (r < 0) Then r = 0
                If (g < 0) Then g = 0
                If (b < 0) Then b = 0
                
                'Assign the new color and alpha values
                imageData(x) = b
                imageData(x + 1) = g
                imageData(x + 2) = r
                imageData(x + 3) = newAlpha * a * ONE_DIV_255
                    
            End If
            
        Next x
        
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
        
    Next y
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub sltBlend_Change()
    UpdatePreview
End Sub

Private Sub sltErase_Change()
    UpdatePreview
End Sub

'Render a new preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ColorToAlpha GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "color", csSource.Color
        .AddParam "erase-threshold", sltErase.Value
        .AddParam "edge-blending", sltBlend.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
