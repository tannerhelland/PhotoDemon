VERSION 5.00
Begin VB.Form FormHSL 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Adjust Hue / Saturation / Lightness"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12090
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
   ScaleWidth      =   806
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.sliderTextCombo sltHue 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "hue"
      Min             =   -180
      Max             =   180
      SliderTrackStyle=   4
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
   Begin PhotoDemon.sliderTextCombo sltSaturation 
      Height          =   720
      Left            =   6000
      TabIndex        =   3
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "saturation"
      Min             =   -100
      Max             =   100
      SliderTrackStyle=   2
   End
   Begin PhotoDemon.sliderTextCombo sltLuminance 
      Height          =   720
      Left            =   6000
      TabIndex        =   4
      Top             =   3480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
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
'Copyright 2012-2015 by Tanner Helland
'Created: 05/October/12
'Last updated: 26/April/13
'Last update: simplify code by relying on new slider/text custom control
'
'Fairly simple and standard HSL adjustment form.  Layout and feature set derived from comparable tools
' in GIMP and Paint.NET.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Colorize an image using a hue defined between -1 and 5
' Input: desired hue, whether to force saturation to 0.5 or maintain the existing value
Public Sub AdjustImageHSL(ByVal hModifier As Double, ByVal sModifier As Double, ByVal lModifier As Double, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As fxPreviewCtl)
    
    If Not toPreview Then Message "Adjusting hue, saturation, and luminance values..."
    
    'Convert the modifiers to be on the same scale as the HSL translation routine
    hModifier = hModifier / 60
    sModifier = (sModifier + 100) / 100
    lModifier = lModifier / 100
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, toPreview, dstPic
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, l As Double
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Get the hue and saturation
        tRGBToHSL r, g, b, h, s, l
        
        'Apply the modifiers
        h = h + hModifier
        If h > 5 Then h = h - 6
        If h < -1 Then h = h + 6
        
        s = s * sModifier
        If s < 0 Then s = 0
        If s > 1 Then s = 1
        
        l = l + lModifier
        If l < 0 Then l = 0
        If l > 1 Then l = 1
        
        'Convert back to RGB using our artificial hue value
        tHSLToRGB h, s, l, r, g, b
        
        'Assign the new values to each color channel
        ImageData(QuickVal + 2, y) = r
        ImageData(QuickVal + 1, y) = g
        ImageData(QuickVal, y) = b
        
    Next y
        If toPreview = False Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData toPreview, dstPic
    
End Sub

Private Sub cmdBar_OKClick()
    Process "Hue and saturation", , buildParams(sltHue.Value, sltSaturation.Value, sltLuminance.Value), UNDO_LAYER
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    redrawSaturationSlider
    updatePreview
End Sub

Private Sub Form_Activate()
        
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'Display the previewed effect in the neighboring window
    updatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltHue_Change()
    redrawSaturationSlider
    updatePreview
End Sub

Private Sub sltLuminance_Change()
    updatePreview
End Sub

Private Sub sltSaturation_Change()
    updatePreview
End Sub

Private Sub updatePreview()
    If cmdBar.previewsAllowed Then AdjustImageHSL sltHue.Value, sltSaturation.Value, sltLuminance.Value, True, fxPreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub fxPreview_ViewportChanged()
    updatePreview
End Sub

Private Sub redrawSaturationSlider()

    'Update the Saturation background dynamically, to match the hue background!
    Dim r As Long, g As Long, b As Long
    
    tHSLToRGB (sltHue.Value + 180) / 60, 0, 0.5, r, g, b
    sltSaturation.GradientColorLeft = RGB(r, g, b)
    
    tHSLToRGB (sltHue.Value + 180) / 60, 1, 0.5, r, g, b
    sltSaturation.GradientColorRight = RGB(r, g, b)

End Sub
