VERSION 5.00
Begin VB.Form FormHSL 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Adjust HSL"
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
   Begin PhotoDemon.pdSlider sltHue 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "hue"
      Min             =   -180
      Max             =   180
      SliderTrackStyle=   4
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
   Begin PhotoDemon.pdSlider sltSaturation 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2520
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "saturation"
      Min             =   -100
      Max             =   100
      SliderTrackStyle=   2
   End
   Begin PhotoDemon.pdSlider sltLuminance 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   3480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1244
      Caption         =   "lightness"
      Min             =   -100
      Max             =   100
      SliderTrackStyle=   2
   End
End
Attribute VB_Name = "FormHSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'HSL Adjustment Form
'Copyright 2012-2026 by Tanner Helland
'Created: 05/October/12
'Last updated: 23/April/20
'Last update: perf improvements; switch to higher-quality HSL transform
'
'Fairly simple and standard HSL adjustment form.  Layout and feature set derived from comparable tools
' in GIMP and Paint.NET.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Colorize an image using a hue defined between -1 and 5
' Input: desired hue, whether to force saturation to 0.5 or maintain the existing value
Public Sub AdjustImageHSL(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Adjusting hue, saturation, and luminance values..."
    
    Dim hModifier As Double, sModifier As Double, lModifier As Double
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    With cParams
        hModifier = .GetDouble("hue", sltHue.Value)
        sModifier = .GetDouble("saturation", sltSaturation.Value)
        lModifier = .GetDouble("value", sltLuminance.Value)
    End With
    
    'Convert the modifiers to be on the same scale as the HSL translation routine
    hModifier = hModifier / 360#
    sModifier = (sModifier + 100#) / 100#
    lModifier = lModifier / 100#
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Double, g As Double, b As Double
    Dim h As Double, s As Double, l As Double
        
    initX = initX * 4
    finalX = finalX * 4
    
    Const ONE_DIV_255 As Double = 1# / 255#
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        'Get the hue and saturation
        Colors.PreciseRGBtoHSL r * ONE_DIV_255, g * ONE_DIV_255, b * ONE_DIV_255, h, s, l
        
        'Apply modifiers
        h = h + hModifier
        If (h > 1#) Then h = h - 1#
        If (h < 0#) Then h = h + 1#
        
        s = s * sModifier
        If (s < 0#) Then s = 0#
        If (s > 1#) Then s = 1#
        
        l = l + lModifier
        If (l < 0#) Then l = 0#
        If (l > 1#) Then l = 1#
        
        'Convert back to RGB using our artificial hue value
        Colors.PreciseHSLtoRGB h, s, l, r, g, b
        
        'Assign the new values to each color channel
        imageData(x) = Int(b * 255 + 0.5)
        imageData(x + 1) = Int(g * 255 + 0.5)
        imageData(x + 2) = Int(r * 255 + 0.5)
        
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

Private Sub cmdBar_OKClick()
    Process "Hue and saturation", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    RedrawSaturationSlider
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

Private Sub sltHue_Change()
    RedrawSaturationSlider
    UpdatePreview
End Sub

Private Sub sltLuminance_Change()
    UpdatePreview
End Sub

Private Sub sltSaturation_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then AdjustImageHSL GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub RedrawSaturationSlider()

    'Update the Saturation background dynamically, to match the hue background!
    Dim r As Long, g As Long, b As Long
    
    ImpreciseHSLtoRGB (sltHue.Value + 180) / 60, 0, 0.5, r, g, b
    sltSaturation.GradientColorLeft = RGB(r, g, b)
    
    ImpreciseHSLtoRGB (sltHue.Value + 180) / 60, 1, 0.5, r, g, b
    sltSaturation.GradientColorRight = RGB(r, g, b)

End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "hue", sltHue.Value
        .AddParam "saturation", sltSaturation.Value
        .AddParam "value", sltLuminance.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
