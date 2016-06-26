VERSION 5.00
Begin VB.Form FormTransparency_FromColor 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Make color transparent"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11820
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
   ScaleWidth      =   788
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
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
   Begin PhotoDemon.pdColorSelector colorPicker 
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
'Copyright 2013-2016 by Tanner Helland
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'OK button
Private Sub cmdBar_OKClick()
    Process "Color to alpha", , BuildParams(colorPicker.Color, sltErase.Value, sltBlend.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    colorPicker.Color = RGB(0, 192, 0)
    sltErase.Value = 15
    sltBlend.Value = 15
End Sub

Private Sub colorPicker_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Activate()
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Render a preview of the alpha effect
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The user can select a color from the preview window; this helps green screen calculation immensely
Private Sub pdFxPreview_ColorSelected()
    colorPicker.Color = pdFxPreview.SelectedColor
    UpdatePreview
End Sub

'Convert a DIB from 24bpp to 32bpp, using a high-quality color-matching scheme in the L*a*b* color space.
Public Sub ColorToAlpha(Optional ByVal ConvertColor As Long, Optional ByVal eraseThreshold As Double = 15, Optional ByVal blendThreshold As Double = 30, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Adding new alpha channel to image..."
    
    'Call prepImageData, which will prepare a temporary copy of the image
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    PrepImageData tmpSA, toPreview, dstPic
    
    'Before doing anything else, convert this DIB to 32bpp.
    If (workingDIB.GetDIBColorDepth <> 32) Then workingDIB.ConvertTo32bpp
    
    'Create a local array and point it at the pixel data we want to operate on
    PrepSafeArray tmpSA, workingDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'For this function to work, each pixel needs to be RGBA, 32-bpp
    Dim pxWidth As Long
    pxWidth = workingDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long, a As Long
    
    'R2/G2/B2 store the RGB values of the color we are attempting to remove
    Dim r2 As Long, g2 As Long, b2 As Long
    r2 = ExtractRed(ConvertColor)
    g2 = ExtractGreen(ConvertColor)
    b2 = ExtractBlue(ConvertColor)
    
    'For maximum quality, we will apply our color comparison in the L*a*b* color space; each scanline will be
    ' transformed to L*a*b* all at once, for performance reasons
    Dim labValues() As Single
    ReDim labValues(0 To finalX * pxWidth + pxWidth) As Single
    
    Dim labL As Double, labA As Double, labB As Double
    Dim labL2 As Double, labA2 As Double, labB2 As Double
    
    Dim labLf As Single, labAf As Single, labBf As Single
    Dim labL2f As Single, labA2f As Single, labB2f As Single
    
    Dim labTransform As pdLCMSTransform
    
    'Calculate the L*a*b* values of the color to be removed
    If g_LCMSEnabled Then
    
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
    Dim difThreshold As Double
    blendThreshold = eraseThreshold + blendThreshold
    difThreshold = blendThreshold - eraseThreshold
    
    Dim cDistance As Double
    Dim newAlpha As Long
    
    'To improve performance of our horizontal loop, we'll move through bytes an entire pixel at a time
    Dim xStart As Long, xStop As Long
    xStart = initX * pxWidth
    xStop = finalX * pxWidth
        
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
    
        'Start by pre-calculating all L*a*b* values for this row
        If g_LCMSEnabled Then
            labTransform.ApplyTransformToScanline VarPtr(ImageData(0, y)), VarPtr(labValues(0)), finalX + 1
        Else
            For x = xStart To xStop Step pxWidth
                b = ImageData(x, y)
                g = ImageData(x + 1, y)
                r = ImageData(x + 2, y)
                RGBtoLAB r, g, b, labL, labA, labB
                labValues(x) = labL
                labValues(x + 1) = labA
                labValues(x + 2) = labB
            Next x
        End If
        
        'With all lab values pre-calculated, we can quickly step through each pixel, calculating distances as we go
        For x = xStart To xStop Step pxWidth
        
            'Get the source pixel color values
            b = ImageData(x, y)
            g = ImageData(x + 1, y)
            r = ImageData(x + 2, y)
            a = ImageData(x + 3, y)
            
            'Perform a basic distance calculation (not ideal, but faster than a completely correct comparison;
            ' see http://en.wikipedia.org/wiki/Color_difference for a full report)
            If g_LCMSEnabled Then
                cDistance = Math_Functions.Distance3D_FastFloat(labValues(x), labValues(x + 1), labValues(x + 2), labL2f, labA2f, labB2f)
            Else
                cDistance = Math_Functions.DistanceThreeDimensions(labValues(x), labValues(x + 1), labValues(x + 2), labL2, labA2, labB2)
            End If
            
            'If the distance is below the erasure threshold, remove it completely
            If (cDistance < eraseThreshold) Then
                ImageData(x + 3, y) = 0
                
            'If the color is between the erasure and blend threshold, feather it against a partial alpha and
            ' color-correct it to remove any "color fringing" from the removed color.
            ElseIf (cDistance < blendThreshold) Then
                
                'Use a ^2 curve to improve blending response
                cDistance = ((blendThreshold - cDistance) / difThreshold)
                cDistance = cDistance * cDistance
                
                'Calculate a new alpha value for this pixel, based on its distance from the threshold.  Large
                ' distances from the removed color are made less transparent than small distances.
                newAlpha = 255 - (cDistance * 255)
                
                'Feathering the alpha often isn't enough to fully remove the color fringing caused by the removed
                ' background color, which will have "infected" the core RGB values.  Attempt to correct this by
                ' subtracting the target color from the original color, using the calculated threshold value; this
                ' is the only way I know to approximate the "feathering" caused by light bleeding over object edges.
                If (cDistance = 1) Then cDistance = 0.999999
                r = (r - (r2 * cDistance)) / (1 - cDistance)
                g = (g - (g2 * cDistance)) / (1 - cDistance)
                b = (b - (b2 * cDistance)) / (1 - cDistance)
                
                If (r > 255) Then r = 255
                If (g > 255) Then g = 255
                If (b > 255) Then b = 255
                If (r < 0) Then r = 0
                If (g < 0) Then g = 0
                If (b < 0) Then b = 0
                
                'Assign the new color and alpha values
                ImageData(x, y) = b
                ImageData(x + 1, y) = g
                ImageData(x + 2, y) = r
                ImageData(x + 3, y) = newAlpha * (a / 255)
                    
            End If
            
        Next x
        
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
        
    Next y
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    FinalizeImageData toPreview, dstPic
    
End Sub

Private Sub sltBlend_Change()
    UpdatePreview
End Sub

Private Sub sltErase_Change()
    UpdatePreview
End Sub

'Render a new preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ColorToAlpha colorPicker.Color, sltErase.Value, sltBlend.Value, True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

