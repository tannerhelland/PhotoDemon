VERSION 5.00
Begin VB.Form FormFilmNoir 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Film noir"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11700
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
   ScaleWidth      =   780
   Begin PhotoDemon.pdSlider sltShadow 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   480
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   1244
      Caption         =   "shadow cut-off"
      Max             =   100
      SigDigits       =   1
      Value           =   50
      DefaultValue    =   50
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
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltContrast 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2400
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   1244
      Caption         =   "contrast boost"
      Max             =   100
      SigDigits       =   1
      Value           =   50
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdSlider sltHighlight 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   1440
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   1244
      Caption         =   "highlight cut-off"
      Max             =   100
      SigDigits       =   1
      Value           =   50
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdSlider sltMidpoint 
      Height          =   705
      Left            =   6000
      TabIndex        =   5
      Top             =   3360
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   1244
      Caption         =   "contrast midpoint"
      Max             =   100
      SigDigits       =   1
      Value           =   50
      DefaultValue    =   50
   End
   Begin PhotoDemon.pdSlider sltGrain 
      Height          =   705
      Left            =   6000
      TabIndex        =   6
      Top             =   4320
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   1244
      Caption         =   "film grain"
      Max             =   100
      SigDigits       =   1
   End
End
Attribute VB_Name = "FormFilmNoir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Film Noir Effect Interface
'Copyright 2013-2026 by Tanner Helland
'Created: some time 2013
'Last updated: 04/October/15
'Last update: rewrite the old "one-click" filter from scratch, and completely rethink the algorithm while I'm at it.
'
'As usual, if you're not familiar with the film noir genre, Wikipedia is a good place to start:
'
'https://en.wikipedia.org/wiki/Film_noir#Visual_style
'
'Classic, sci-fi, neo noir - it's all fantastically fun.  PD's effect is a throwback to the classic, high-contrast style
' of 40's noir (The Maltese Falcon, Double Indemnity, etc.), and the entire effect operates in HDR space to try
' and minimize contrast loss.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a noir-inspired filter to an image.
Public Sub fxFilmNoir(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim shadowStrength As Double, contrastStrength As Double, luminancePoint As Double, highlightStrength As Double, artificialGrain As Double
    shadowStrength = cParams.GetDouble("shadow")
    contrastStrength = cParams.GetDouble("contrast")
    luminancePoint = cParams.GetDouble("midpoint")
    highlightStrength = cParams.GetDouble("highlight")
    artificialGrain = cParams.GetDouble("grain")
    
    If (Not toPreview) Then Message "Asking Sam Spade for help..."
    
    'Shadow and highlight strength are on the range 0-100.  Invert highlight so it's on the range [155, 255]
    highlightStrength = 255 - highlightStrength
    
    'Given the distance between shadow and highlight values, determine the remap value on a [0, 100] scale.
    Dim contrastRange As Double
    contrastRange = (highlightStrength - shadowStrength)
    contrastRange = (255 / contrastRange)
    
    'The luminance midpoint is on the range [0, 100], with a default of 50.  Remap it to [0, 255].
    luminancePoint = luminancePoint * 2.55
    
    'Film grain is on the range [0, 100].  Remap it to [-50, 50], and prep a randomizer as necessary.
    artificialGrain = artificialGrain - 50
    
    Dim cRandomize As pdRandomize
    If (artificialGrain > 0) Then
        Set cRandomize = New pdRandomize
        cRandomize.SetSeed_AutomaticAndRandom
    End If
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D
    EffectPrep.PrepImageData tmpSA, toPreview, dstPic
    workingDIB.WrapArrayAroundDIB imageData, tmpSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color and grayscale variables
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Double, grayByte As Long
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        xStride = x * 4
    For y = initY To finalY
    
        'Get the source pixel color values
        b = imageData(xStride, y)
        g = imageData(xStride + 1, y)
        r = imageData(xStride + 2, y)
        
        'Starting by convert the pixel to its grayscale equivalent
        grayVal = (218 * r + 732 * g + 74 * b) \ 1024
        
        'Gray is now a floating-point value on the scale 0-255.  We leave it as a floating-point, because we're about to
        ' do a bunch of detailed contrast work, and we don't want to lost resolution.
        
        'First thing we want to do is remap the gray value to ignore the shadow and highlight cutoffs specified by
        ' the caller.  A standard gradient formula works wonders here.
        grayVal = (grayVal - shadowStrength) * contrastRange
        
        'Our gray value may now lie outside the desired [0, 100] range - and that's okay.  We won't crop it until the
        ' last possible moment, to try and retain as much image data as possible.
        
        'Remap contrast based on the luminance point supplied by the caller
        grayVal = grayVal + (((grayVal - luminancePoint) * contrastStrength) / 100)
        
        'We now have a contrast-corrected gray value.  If the user wants noise applied, do so now.
        If (artificialGrain > 0) Then
            grayVal = grayVal + (artificialGrain * cRandomize.GetRandomFloat_VB)
        End If
        
        'Copy it to an integer and clamp.
        grayByte = grayVal
        If (grayByte < 0) Then grayByte = 0
        If (grayByte > 255) Then grayByte = 255
        
        'Assign that gray value to each color channel
        imageData(xStride, y) = grayByte
        imageData(xStride + 1, y) = grayByte
        imageData(xStride + 2, y) = grayByte
        
    Next y
        If (Not toPreview) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic

End Sub

Private Sub cmdBar_OKClick()
    Process "Film noir", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltShadow.Value = 50
    sltMidpoint.Value = 50
    sltContrast.Value = 50
    sltHighlight.Value = 50
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

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then fxFilmNoir GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub sltContrast_Change()
    UpdatePreview
End Sub

Private Sub sltGrain_Change()
    UpdatePreview
End Sub

Private Sub sltHighlight_Change()
    UpdatePreview
End Sub

Private Sub sltMidpoint_Change()
    UpdatePreview
End Sub

Private Sub sltShadow_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("shadow", sltShadow.Value, "contrast", sltContrast.Value, "midpoint", sltMidpoint.Value, "highlight", sltHighlight.Value, "grain", sltGrain.Value)
End Function
