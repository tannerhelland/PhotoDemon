VERSION 5.00
Begin VB.Form FormReplaceColor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Replace color"
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
      Top             =   3000
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1270
      Caption         =   "replace threshold"
      Max             =   199
      Value           =   15
      DefaultValue    =   15
   End
   Begin PhotoDemon.pdSlider sltBlend 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   4080
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   1270
      Caption         =   "edge blending"
      Max             =   200
      Value           =   15
      DefaultValue    =   15
   End
   Begin PhotoDemon.pdColorSelector colorOld 
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      Caption         =   "color to replace (right-click preview to select)"
      curColor        =   12582912
   End
   Begin PhotoDemon.pdColorSelector colorNew 
      Height          =   975
      Left            =   6000
      TabIndex        =   5
      Top             =   1680
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      Caption         =   "new color"
      curColor        =   49152
   End
End
Attribute VB_Name = "FormReplaceColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Replace color dialog
'Copyright 2013-2026 by Tanner Helland
'Created: 29/October/13
'Last updated: 27/April/20
'Last update: perf improvements; tool is 2x faster now!
'
'This function uses an algorithm very similar to PhotoDemon's green screen (FormTransparency_FromColor) algorithm.
' Separate sliders are provided for both a replacement threshold, and a blend threshold, to help the user minimize
' harsh edges between the color and its surroundings.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'OK button
Private Sub cmdBar_OKClick()
    Process "Replace color", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    colorOld.Color = RGB(0, 0, 255)
    colorNew.Color = RGB(0, 192, 0)
    sltErase.Value = 15#
    sltBlend.Value = 15#
End Sub

Private Sub colorNew_ColorChanged()
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
    colorOld.Color = pdFxPreview.SelectedColor
    UpdatePreview
End Sub

'Replace one color in an image with another color, with full blending and feathering support
Public Sub ReplaceSelectedColor(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Replacing color..."
    
    Dim oldColor As Long, newColor As Long
    Dim eraseThreshold As Double, blendThreshold As Double
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    With cParams
        oldColor = .GetLong("oldcolor", vbWhite)
        newColor = .GetLong("newcolor", vbBlack)
        eraseThreshold = .GetDouble("erasethreshold", 15#)
        blendThreshold = .GetDouble("blendthreshold", 30#)
    End With
    
    'Call prepImageData, which will prepare a temporary copy of the image
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
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
    Dim r As Long, g As Long, b As Long
    
    'oldR/G/B store the RGB values of the color we are attempting to remove
    Dim oldR As Long, oldG As Long, oldB As Long
    oldR = Colors.ExtractRed(oldColor)
    oldG = Colors.ExtractGreen(oldColor)
    oldB = Colors.ExtractBlue(oldColor)
    
    'newR/G/B store the RGB values of the color we are using to replace the old color
    Dim newR As Long, newG As Long, newB As Long
    newR = Colors.ExtractRed(newColor)
    newG = Colors.ExtractGreen(newColor)
    newB = Colors.ExtractBlue(newColor)
    
    'For maximum quality, we will apply our color comparison in the L*a*b* color space; each scanline will be
    ' transformed to L*a*b* all at once, for performance reasons
    Dim labValues() As Single
    ReDim labValues(0 To finalX * pxWidth + pxWidth) As Single
    
    Dim labL As Double, labA As Double, labB As Double
    Dim labL2 As Double, labA2 As Double, labB2 As Double
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
        rgbBytes(0) = oldB: rgbBytes(1) = oldG: rgbBytes(2) = oldR
        
        Dim labBytes() As Single
        ReDim labBytes(0 To 3) As Single
        labTransform.ApplyTransformToScanline VarPtr(rgbBytes(0)), VarPtr(labBytes(0)), 1
        
        labL2f = labBytes(0)
        labA2f = labBytes(1)
        labB2f = labBytes(2)
        
    Else
        Colors.RGBtoLAB oldR, oldG, oldB, labL2, labA2, labB2
        labL2f = labL2
        labA2f = labA2
        labB2f = labB2
    End If
    
    'The blend threshold is used to "smooth" the edges of replaced color areas.  Calculate the difference between
    ' the erase and the blend thresholds in advance.
    Dim difThreshold As Double
    blendThreshold = eraseThreshold + blendThreshold
    difThreshold = blendThreshold - eraseThreshold
    If (difThreshold <> 0#) Then difThreshold = 1# / difThreshold
    
    Dim cDistance As Double, invCDistance As Double
    
    'To improve performance of our horizontal loop, we'll move through bytes an entire pixel at a time
    Dim xStart As Long, xStop As Long
    xStart = initX * pxWidth
    xStop = finalX * pxWidth
     
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
        
        'Start by pre-calculating all L*a*b* values for this row
        If useLCMS Then
            labTransform.ApplyTransformToScanline VarPtr(imageData(0)), VarPtr(labValues(0)), finalX + 1
        Else
            For x = xStart To xStop Step pxWidth
                b = imageData(x)
                g = imageData(x + 1)
                r = imageData(x + 2)
                Colors.RGBtoLAB r, g, b, labL, labA, labB
                labValues(x) = labL
                labValues(x + 1) = labA
                labValues(x + 2) = labB
            Next x
        End If
        
        'With all lab values pre-calculated, we can quickly step through each pixel, calculating distances as we go
        For x = xStart To xStop Step pxWidth
        
            'Get the source pixel color values
            b = imageData(x)
            g = imageData(x + 1)
            r = imageData(x + 2)
            
            'Perform a basic distance calculation (not ideal, but faster than a completely correct comparison;
            ' see https://en.wikipedia.org/wiki/Color_difference for a full report)
            If useLCMS Then
                cDistance = PDMath.Distance3D_FastFloat(labValues(x), labValues(x + 1), labValues(x + 2), labL2f, labA2f, labB2f)
            Else
                cDistance = PDMath.DistanceThreeDimensions(labValues(x), labValues(x + 1), labValues(x + 2), labL2, labA2, labB2)
            End If
            
            'If the distance is below the erasure threshold, replace it completely
            If (cDistance < eraseThreshold) Then
                imageData(x) = newB
                imageData(x + 1) = newG
                imageData(x + 2) = newR
                
            'If the color is between the replace and blend threshold, feather it against the new color and
            ' color-correct it to remove any "color fringing" from the replaced color.
            ElseIf (cDistance < blendThreshold) Then
                
                'Use a ^2 curve to improve blending response
                cDistance = (blendThreshold - cDistance) * difThreshold
                
                'Feathering the pixel often isn't enough to fully remove the color fringing caused by the replaced
                ' color, which will have "infected" the core RGB values.  Attempt to correct this by subtracting the
                ' target color from the original color, using the calculated threshold value; this is the only way I
                ' know to approximate the "feathering" caused by light bleeding over object edges.
                If (cDistance >= 1#) Then cDistance = 0.999999
                invCDistance = 1# / (1# - cDistance)
                r = (r - (oldR * cDistance)) * invCDistance
                g = (g - (oldG * cDistance)) * invCDistance
                b = (b - (oldB * cDistance)) * invCDistance
                
                If (r > 255) Then r = 255
                If (g > 255) Then g = 255
                If (b > 255) Then b = 255
                If (r < 0) Then r = 0
                If (g < 0) Then g = 0
                If (b < 0) Then b = 0
                
                'Assign the new color and alpha values
                imageData(x) = Colors.BlendColors(b, newB, cDistance)
                imageData(x + 1) = Colors.BlendColors(g, newG, cDistance)
                imageData(x + 2) = Colors.BlendColors(r, newR, cDistance)
                
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
    If cmdBar.PreviewsAllowed Then ReplaceSelectedColor GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "oldcolor", colorOld.Color
        .AddParam "newcolor", colorNew.Color
        .AddParam "erasethreshold", sltErase.Value
        .AddParam "blendthreshold", sltBlend.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
