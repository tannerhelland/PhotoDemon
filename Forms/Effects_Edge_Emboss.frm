VERSION 5.00
Begin VB.Form FormEmbossEngrave 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Emboss"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11535
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
   ScaleWidth      =   769
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdColorSelector csEmboss 
      Height          =   1095
      Left            =   6000
      TabIndex        =   2
      Top             =   3840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1931
      Caption         =   "base color"
      curColor        =   16744576
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
   Begin PhotoDemon.pdSlider sltDistance 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1920
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "thickness"
      Min             =   -10
      SigDigits       =   2
      Value           =   1
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "angle"
      Min             =   -180
      Max             =   180
      SigDigits       =   1
   End
   Begin PhotoDemon.pdSlider sltDepth 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   2880
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "depth"
      Min             =   0.1
      SigDigits       =   2
      Value           =   1
      DefaultValue    =   1
   End
End
Attribute VB_Name = "FormEmbossEngrave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Emboss/Engrave Effect Dialog
'Copyright 2003-2026 by Tanner Helland
'Created: 3/6/03
'Last updated: 21/February/20
'Last update: large performance improvements
'
'This dialog processes a variety of emboss/engrave-style filters.  It's been in PD for a long time,
' but the 6.4 release saw some much-needed improvements in the form of selectable angle, depth,
' and thickness.  Interpolation is used to process all emboss calculations, so the result looks
' very good for any angle and/or depth combination.  Edge handling is also handled much better than
' past versions of the tool.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'OK button
Private Sub cmdBar_OKClick()
    Process "Emboss", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltDepth.Value = 1
    sltDistance.Value = 1
    csEmboss.Color = RGB(127, 127, 127)
End Sub

Private Sub csEmboss_ColorChanged()
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

Private Sub pdFxPreview_ColorSelected()
    csEmboss.Color = pdFxPreview.SelectedColor
End Sub

'Emboss an image
' Inputs: color to emboss to, and whether or not this is a preview (plus the destination picture box if it IS a preview)
Public Sub ApplyEmbossEffect(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)

    If (Not toPreview) Then Message "Embossing image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim eDistance As Double, eAngle As Double, eDepth As Double, eColor As Long
    
    With cParams
        eDistance = .GetDouble("distance", sltDistance.Value)
        eAngle = .GetDouble("angle", sltAngle.Value)
        eDepth = .GetDouble("depth", sltDepth.Value)
        eColor = .GetLong("color", csEmboss.Color)
    End With
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte, dstSA As SafeArray2D, dstSA1D As SafeArray1D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a copy of the current image; we will use it as our source reference.
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'At present, stride is always width * 4 (32-bit RGBA)
    Dim xStride As Long
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.SetDistortParameters pdeo_Clamp, True, curDIBValues.maxX, curDIBValues.maxY
    
    'During previews, adjust the distance parameter to compensate for preview size
    If toPreview Then eDistance = eDistance * curDIBValues.previewModifier
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim tR As Long, tG As Long, tB As Long
    Dim rBase As Long, gBase As Long, bBase As Long

    'Extract the red, green, and blue values from the color we've been passed
    rBase = Colors.ExtractRed(eColor)
    gBase = Colors.ExtractGreen(eColor)
    bBase = Colors.ExtractBlue(eColor)
    
    'Convert the rotation angle to radians
    eAngle = eAngle * (PI / 180#)
    
    'X and Y offsets are hard-coded per the current angle
    Dim xOffset As Double, yOffset As Double
    xOffset = Cos(eAngle) * eDistance
    yOffset = Sin(eAngle) * eDistance
    
    Dim tmpQuad As RGBQuad
    fSupport.AliasTargetDIB srcDIB
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline dstImageData, dstSA1D, y
    For x = initX To finalX
    
        'Retrieve source RGB values
        xStride = x * 4
        b = dstImageData(xStride)
        g = dstImageData(xStride + 1)
        r = dstImageData(xStride + 2)
        
        'Use the filter support class to interpolate and edge-clamp pixels as necessary
        ' on the source pixel at the pre-calculated offset
        tmpQuad = fSupport.GetColorsFromSource(x + xOffset, y + yOffset, x, y)
        tB = tmpQuad.Blue
        tG = tmpQuad.Green
        tR = tmpQuad.Red
            
        'Calculate an emboss value for each color
        r = (r - tR) * eDepth + rBase
        g = (g - tG) * eDepth + gBase
        b = (b - tB) * eDepth + bBase
                
        'Clamp RGB values
        If (r < 0) Then r = 0
        If (r > 255) Then r = 255
        If (g < 0) Then g = 0
        If (g > 255) Then g = 255
        If (b < 0) Then b = 0
        If (b > 255) Then b = 255
        
        dstImageData(xStride) = b
        dstImageData(xStride + 1) = g
        dstImageData(xStride + 2) = r
        
        'Leave alpha as-is for this effect
        
    Next x
        If (Not toPreview) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y
            End If
        End If
    Next y
    
    'Safely deallocate all image arrays
    fSupport.UnaliasTargetDIB
    workingDIB.UnwrapArrayFromDIB dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic
 
End Sub

'Render a new preview
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyEmbossEffect GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltAngle_Change()
    UpdatePreview
End Sub

Private Sub sltDepth_Change()
    UpdatePreview
End Sub

Private Sub sltDistance_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "distance", sltDistance.Value
        .AddParam "angle", sltAngle.Value
        .AddParam "depth", sltDepth.Value
        .AddParam "color", csEmboss.Color
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
