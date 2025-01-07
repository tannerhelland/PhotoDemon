VERSION 5.00
Begin VB.Form FormColorize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Colorize"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11655
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
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   Begin PhotoDemon.pdButtonStrip btsSaturation 
      Height          =   1095
      Left            =   6000
      TabIndex        =   4
      Top             =   1485
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1931
      Caption         =   "saturation"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5820
      Width           =   11655
      _ExtentX        =   20558
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
   Begin PhotoDemon.pdSlider sldHSL 
      Height          =   705
      Index           =   0
      Left            =   6000
      TabIndex        =   2
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Caption         =   "new color"
      Max             =   359
      SliderTrackStyle=   4
      Value           =   180
      NotchPosition   =   1
      DefaultValue    =   180
   End
   Begin PhotoDemon.pdSlider sldHSL 
      Height          =   705
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   2685
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Max             =   100
      SliderTrackStyle=   2
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdButtonStrip btsLuminance 
      Height          =   1095
      Left            =   6000
      TabIndex        =   5
      Top             =   3480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1931
      Caption         =   "luminance"
   End
   Begin PhotoDemon.pdSlider sldHSL 
      Height          =   705
      Index           =   2
      Left            =   6000
      TabIndex        =   6
      Top             =   4680
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1244
      Max             =   100
      SliderTrackStyle=   2
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
End
Attribute VB_Name = "FormColorize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Colorize Form
'Copyright 2006-2025 by Tanner Helland
'Created: 12/January/07
'Last updated: 19/April/20
'Last update: perf improvements; allow user to control saturation and luminance (if not preserving)
'
'This dialog has slowly morphed over the years, and now it bears a lot of similarity to
' the HSL adjustment dialog.  The difference here is that values can be forced to a specific
' constant value, instead of merely scaling them proportionally.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Sub btsLuminance_Click(ByVal buttonIndex As Long)
    UpdatePreview
    ReflowInterface
End Sub

Private Sub btsSaturation_Click(ByVal buttonIndex As Long)
    UpdatePreview
    ReflowInterface
End Sub

Private Sub cmdBar_OKClick()
    Process "Colorize", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

'Colorize an image using a hue defined between 0 and 359
' Input: desired hue, whether to force saturation to 0.5 or maintain the existing value
Public Sub ColorizeImage(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Colorizing image..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim hToUse As Double
    hToUse = cParams.GetDouble("hue", sldHSL(0).Value, True)
    
    Dim sToUse As Double, maintainSaturation As Boolean
    Dim lToUse As Double, maintainLuminance As Boolean
    
    'Old param strings used slightly different IDs
    If cParams.DoesParamExist("preservesaturation", True) Then
        maintainSaturation = cParams.GetBool("preservesaturation", True, True)
        sToUse = 0.5
    Else
        maintainSaturation = cParams.GetBool("preserve-saturation", True)
        sToUse = cParams.GetDouble("saturation", 0.5, True) / 100#
    End If
    
    maintainLuminance = cParams.GetBool("preserve-luminance", True, True)
    lToUse = cParams.GetDouble("luminance", 0.5, True) / 100#
    
    'Convert HSL values to safe ranges
    hToUse = hToUse / 360#
    If (hToUse < 0#) Then hToUse = 0#
    If (hToUse > 1#) Then hToUse = 1#
    
    If (sToUse < 0#) Then sToUse = 0#
    If (sToUse > 1#) Then sToUse = 1#
    
    If (lToUse < 0#) Then lToUse = 0#
    If (lToUse > 1#) Then lToUse = 1#
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left * 4
    initY = curDIBValues.Top
    finalX = curDIBValues.Right * 4
    finalY = curDIBValues.Bottom
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not toPreview) Then ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, l As Double
    Dim rFloat As Double, gFloat As Double, bFloat As Double
    
    Const ONE_DIV_255 As Double = 1# / 255#
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX Step 4
        
        'Get the source pixel color values
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
        
        bFloat = b * ONE_DIV_255
        gFloat = g * ONE_DIV_255
        rFloat = r * ONE_DIV_255
        
        'Get the hue and saturation
        Colors.PreciseRGBtoHSL rFloat, gFloat, bFloat, h, s, l
        
        'Convert back to RGB using our artificial hue value
        If (Not maintainSaturation) Then s = sToUse
        If (Not maintainLuminance) Then l = lToUse
        Colors.PreciseHSLtoRGB hToUse, s, l, rFloat, gFloat, bFloat
        
        'Assign the new values to each color channel
        imageData(x) = bFloat * 255#
        imageData(x + 1) = gFloat * 255#
        imageData(x + 2) = rFloat * 255#
        
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

'Reset the hue bar to the center position
Private Sub cmdBar_ResetClick()
    sldHSL(0).Value = 180#
End Sub

Private Sub Form_Load()

    cmdBar.SetPreviewStatus False
    
    btsSaturation.AddItem "preserve", 0
    btsSaturation.AddItem "custom", 1
    btsSaturation.ListIndex = 0
    
    btsLuminance.AddItem "preserve", 0
    btsLuminance.AddItem "custom", 1
    btsLuminance.ListIndex = 0
    
    ReflowInterface
    ApplyThemeAndTranslations Me, True, True
    
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ColorizeImage GetLocalParamString(), True, pdFxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub RedrawSaturationSlider()

    'Update the Saturation background dynamically, to match the hue background!
    Dim r As Double, g As Double, b As Double
    
    Colors.PreciseHSLtoRGB sldHSL(0).Value / 360#, 0#, 0.5, r, g, b
    sldHSL(1).GradientColorLeft = RGB(r, g, b)
    
    Colors.PreciseHSLtoRGB sldHSL(0).Value / 360#, 1#, 0.5, r, g, b
    sldHSL(1).GradientColorRight = RGB(r, g, b)

End Sub

Private Sub ReflowInterface()

    Dim yOffset As Long
    yOffset = btsSaturation.GetTop + btsSaturation.GetHeight + Interface.FixDPI(8)
    
    If (btsSaturation.ListIndex = 1) Then
        sldHSL(1).SetTop yOffset
        yOffset = yOffset + sldHSL(1).GetHeight + Interface.FixDPI(8)
    End If
    sldHSL(1).Visible = (btsSaturation.ListIndex = 1)
    If sldHSL(1).Visible Then RedrawSaturationSlider
    
    yOffset = yOffset + Interface.FixDPI(4)
    btsLuminance.SetTop yOffset
    yOffset = yOffset + btsLuminance.GetHeight + Interface.FixDPI(8)
    sldHSL(2).SetTop yOffset
    sldHSL(2).Visible = (btsLuminance.ListIndex = 1)
    
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "hue", sldHSL(0).Value
        .AddParam "preserve-saturation", (btsSaturation.ListIndex = 0)
        .AddParam "saturation", sldHSL(1).Value
        .AddParam "preserve-luminance", (btsLuminance.ListIndex = 0)
        .AddParam "luminance", sldHSL(2).Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Sub sldHSL_Change(Index As Integer)
    UpdatePreview
End Sub
