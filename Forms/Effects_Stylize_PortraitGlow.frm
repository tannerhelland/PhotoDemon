VERSION 5.00
Begin VB.Form FormPortraitGlow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Portrait glow"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11580
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
   ScaleWidth      =   772
   Begin PhotoDemon.pdButtonStrip btsStyle 
      Height          =   1095
      Left            =   6000
      TabIndex        =   5
      Top             =   840
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   1931
      Caption         =   "style"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11580
      _ExtentX        =   20426
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
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   2040
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   1244
      Caption         =   "glow radius"
      Min             =   1
      Max             =   100
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdSlider sltBoost 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   2880
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   1244
      Caption         =   "exposure boost"
      Max             =   200
   End
   Begin PhotoDemon.pdSlider sltStrength 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   3720
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   1244
      Caption         =   "strength"
      Max             =   100
      Value           =   100
      DefaultValue    =   100
   End
End
Attribute VB_Name = "FormPortraitGlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Portrait glow (sometimes called "soft glow") image effect
'Copyright 2015-2026 by Tanner Helland
'Created: 20/Dec/15
'Last updated: 20/Dec/15
'Last update: initial build
'
'Basic portrait glow function.  This effect is easily achieved manually, using a duplicate layer + gaussian blur +
' Screen blend mode, but a one-shot menu has been requested by multiple users, so there's probably some merit to it.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Apply a "portrait glow" effect to an image
Public Sub ApplyPortraitGlow(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Parse out the parameter list
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim glowRadius As Double, glowBoost As Double, glowOpacity As Double, glowStyle As Long
    glowRadius = cParams.GetDouble("radius", 1#)
    glowBoost = cParams.GetDouble("exposure", 0#)
    glowOpacity = cParams.GetDouble("strength", 100#)
    glowStyle = cParams.GetLong("style", 0&)
    
    'Change the exposure boost to a 1-based measurement, where 1 = no change
    glowBoost = 1# + (glowBoost / 100)
    
    If (Not toPreview) Then Message "Applying petroleum jelly to camera lens..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic
    
    'Create a copy of the image.  "Portrait glow" requires a blurred image copy as part of the effect, and we maintain
    ' that copy separate from the original (as the two must be blended as the final step of the filter).
    Dim blurDIB As pdDIB
    Set blurDIB = New pdDIB
    blurDIB.CreateFromExistingDIB workingDIB
    
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    Dim progBarMax As Long, progBarOffset As Long
    If toPreview Then
        glowRadius = glowRadius * curDIBValues.previewModifier
        
    'If this is not a preview, initialize the main program progress bar
    Else
        
        progBarMax = finalY * 3 + finalX * 4
        SetProgBarMax progBarMax
        
        Dim progBarCheck As Long
        progBarCheck = ProgressBars.FindBestProgBarValue()
        
    End If
    
    'Failsafe verification for the glow radius
    If (glowRadius < 1) Then glowRadius = 1
    
    'Start by creating the blurred DIB
    If CreateApproximateGaussianBlurDIB(glowRadius, workingDIB, blurDIB, 3, toPreview, progBarMax) Then
        
        progBarOffset = finalY * 3 + finalX * 3
        
        'Now that we have a gaussian DIB created in blurDIB, we can apply any subsequent exposure adjustments.
        If (glowBoost > 0) Then
            
            Dim cLUT As pdFilterLUT
            Set cLUT = New pdFilterLUT
            
            Dim exposureLookup() As Byte
            cLUT.FillLUT_BrightnessMultiplicative exposureLookup, glowBoost
            cLUT.ApplyLUTToAllColorChannels blurDIB, exposureLookup
            
        End If
        
        'With the blur and post-application exposure adjustments applied, we can now merge down the blur+glow layer.
        
        'Start by fixing premultiplication status
        blurDIB.SetAlphaPremultiplication True
        workingDIB.SetAlphaPremultiplication True
        
        'A pdCompositor class will help us blend the images together
        Dim cComposite As pdCompositor
        Set cComposite = New pdCompositor
        
        'Composite our invert+blur image against the base layer (workingDIB) using the COLOR DODGE blend mode;
        ' this will emphasize areas where the layers differ, while ignoring areas where they're the same.
        Dim dstBlendMode As PD_BlendMode
        Select Case glowStyle
            Case 0
                dstBlendMode = BM_Screen
            Case 1
                dstBlendMode = BM_Overlay
            Case 2
                dstBlendMode = BM_SoftLight
        End Select
        
        cComposite.QuickMergeTwoDibsOfEqualSize workingDIB, blurDIB, dstBlendMode, glowOpacity
        
        'Release our temporary DIB
        blurDIB.EraseDIB
        
    End If
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
    
End Sub

Private Sub btsStyle_Click(ByVal buttonIndex As Long)
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Portrait glow", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sltRadius.Value = 5
    sltStrength.Value = 100
End Sub

Private Sub Form_Load()
    
    'Disable previews until the dialog is fully loaded
    cmdBar.SetPreviewStatus False
    
    btsStyle.AddItem "classic", 0
    btsStyle.AddItem "modern", 1
    btsStyle.AddItem "subtle", 2
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltBoost_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltStrength_Change()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyPortraitGlow GetLocalParamString, True, pdFxPreview
End Sub

Private Function GetLocalParamString() As String
    GetLocalParamString = BuildParamList("style", btsStyle.ListIndex, "radius", sltRadius.Value, "exposure", sltBoost.Value, "strength", sltStrength.Value)
End Function
