VERSION 5.00
Begin VB.Form FormGaussianBlur 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Gaussian blur"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
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
   ScaleWidth      =   802
   Begin PhotoDemon.pdSlider sldIterations 
      Height          =   855
      Left            =   6000
      TabIndex        =   5
      Top             =   4560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1508
      Caption         =   "iterations"
      Min             =   2
      Max             =   4
      Value           =   3
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   3
   End
   Begin PhotoDemon.pdDropDown ddCustom 
      Height          =   975
      Left            =   6000
      TabIndex        =   4
      Top             =   3480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      Caption         =   "algorithms"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   465
      Left            =   6000
      TabIndex        =   2
      Top             =   1740
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   820
      Min             =   0.1
      Max             =   1000
      SigDigits       =   1
      ScaleStyle      =   1
      Value           =   1
      DefaultValue    =   1
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
   Begin PhotoDemon.pdButtonStrip btsQuality 
      Height          =   1080
      Left            =   6000
      TabIndex        =   3
      Top             =   2280
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   1905
      Caption         =   "mode"
   End
   Begin PhotoDemon.pdDropDown ddRadius 
      Height          =   420
      Left            =   6000
      TabIndex        =   6
      Top             =   1260
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   741
      FontSize        =   12
   End
End
Attribute VB_Name = "FormGaussianBlur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gaussian Blur Tool
'Copyright 2010-2026 by Tanner Helland
'Created: 01/July/10
'Last updated: 07/November/19
'Last update: switched to an all-new Deriche implementation for "high quality" gaussian
'
'Not much to say here.  Gaussian blur is a standard function in an image editor!
'
'Like most programs, PhotoDemon does not attempt to calculate the gaussian precisely.
' (I previously provided a function for this, but even a separable implementation is
' incredibly slow on large images.)  Instead, it uses several different approximation
' methods with varying performance/quality trade-offs.  Currently available methods include...
' - iterative box blurs (a la Photoshop)
' - heat equation estimation (Alvarez-Mazorra)
' - recursive IIR (Deriche)
'
'These functions were written with help from outside researchers; please see the linked
' gaussian functions for details.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Convolve an image using a gaussian kernel (separable implementation!)
'Input: radius of the blur (min 1, no real max - but the scroll bar is maxed at 200 presently)
Public Sub GaussianBlurFilter(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
        
    If (Not toPreview) Then Message "Applying gaussian blur..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim gRadius As Double, gaussQuality As String, gaussAlgo As String, gaussIterations As Long
    
    With cParams
        gRadius = .GetDouble("radius", sltRadius.Value, True)
        gaussQuality = .GetString("quality", "fast", True)
        gaussAlgo = .GetString("algorithm", "box", True)
        gaussIterations = .GetLong("iterations", 3, True)
    End With
    
    'Legacy handling of old numeric quality indicators
    If TextSupport.IsNumberLocaleUnaware(gaussQuality) Then gaussQuality = GetQualityAsString(TextSupport.CDblCustom(gaussQuality))
    
    'If "fast" or "precise" mode is used, populate the equivalent alog
    If (gaussQuality = "fast") Then
        gaussAlgo = "box"
        gaussIterations = 3
    End If
    
    If (gaussQuality = "precise") Then
        gaussAlgo = "deriche"
        gaussIterations = 3
    End If
    
    'Validate settings
    If (gaussIterations < 2) Then gaussIterations = 2
    If (gaussIterations > 4) Then gaussIterations = 4
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, , , True
    
    'If this is a preview, we need to adjust the kernel radius to match the size of the preview box
    If toPreview Then
        gRadius = gRadius * curDIBValues.previewModifier
        If (gRadius <= 0.1) Then gRadius = 0.1
    End If
    
    'I almost always recommend quality over speed for PD tools, but in this case, the fast option is SO much faster,
    ' and the results so indistinguishable (3% different according to the Central Limit Theorem:
    ' https://www.khanacademy.org/math/probability/statistics-inferential/sampling_distribution/v/central-limit-theorem?playlist=Statistics
    ' ), that I recommend the faster methods instead.
    Select Case gaussAlgo
    
        'Iterative box blurs
        Case "box"
            Filters_Layers.CreateApproximateGaussianBlurDIB gRadius, workingDIB, workingDIB, gaussIterations, toPreview
            
        'Alvarez-Mazorra anisotropic diffusion
        Case "am"
            Filters_Area.GaussianBlur_AM workingDIB, gRadius, gaussIterations, toPreview
        
        'Deriche IIR
        Case "deriche"
            Filters_Area.GaussianBlur_Deriche workingDIB, gRadius, gaussIterations, toPreview
        
        'Failsafe for unsupported quality IDs
        Case Else
            Filters_Layers.CreateApproximateGaussianBlurDIB gRadius, workingDIB, workingDIB, gaussIterations, toPreview
            
    End Select
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering using the data inside workingDIB
    EffectPrep.FinalizeImageData toPreview, dstPic, True
            
End Sub

Private Sub btsQuality_Click(ByVal buttonIndex As Long)
    updateUI
    UpdatePreview
End Sub

'OK button
Private Sub cmdBar_OKClick()
    Process "Gaussian blur", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub ddCustom_Click()
    UpdatePreview
End Sub

Private Sub ddRadius_Click()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Populate radius options
    ddRadius.SetAutomaticRedraws False
    ddRadius.AddItem "radius", 0
    ddRadius.AddItem "radius (Photoshop equivalent)", 1
    
    Dim stdDevText As String
    If PDMain.IsProgramRunning() Then stdDevText = g_Language.TranslateMessage("standard deviation (%1)", ChrW$(&H3C3))
    ddRadius.AddItem stdDevText, 2
    ddRadius.ListIndex = 0
    ddRadius.SetAutomaticRedraws True
    
    'Populate the quality selector
    btsQuality.AddItem "fast", 0
    btsQuality.AddItem "precise", 1
    btsQuality.AddItem "custom", 2
    btsQuality.ListIndex = 0
    
    'Custom algorithms for the "custom" quality setting
    ddCustom.SetAutomaticRedraws False
    ddCustom.AddItem "iterative box blur", 0
    ddCustom.AddItem "anisotropic diffusion (Alvarez-Mazorra)", 1
    ddCustom.AddItem "IIR (Deriche)", 2
    ddCustom.ListIndex = 0
    ddCustom.SetAutomaticRedraws True
    updateUI
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sldIterations_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then GaussianBlurFilter GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub updateUI()
    ddCustom.Visible = (btsQuality.ListIndex = 2)
    sldIterations.Visible = (btsQuality.ListIndex = 2)
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        
        If (ddRadius.ListIndex = 0) Then
            .AddParam "radius", sltRadius.Value
        
        'At present, Photoshop radius matching simply doubles the current radius; for some
        ' reason, PS appears to use "diameter" instead of "radius" even though their UI says
        ' radius... sigh.  (It's also possible they use sigma instead of radius - or they
        ' use sigma and radius interchangeably?  Who knows...)
        ElseIf (ddRadius.ListIndex = 1) Then
            .AddParam "radius", sltRadius.Value * 2
        
        'Some software (GIMP) uses standard deviation as their measurement; we silently
        ' convert this to an equivalent radius.  (Note, however, that new versions of GIMP
        ' deliberately blur in a linear RGB space, so their results will *not* match PD's,
        ' even at an equivalent sigma.)
        Else
            Const LOG_255_BASE_10 As Double = 2.40654018043395
            .AddParam "radius", CDbl((sltRadius.Value * Sqr(2# * LOG_255_BASE_10)) - 1#)
        End If
        
        'Quality is passed as a string to improve forward-compatibility
        .AddParam "quality", GetQualityAsString(btsQuality.ListIndex)
        
        'We always pass custom settings, even if they won't necessarily be used
        ' with the current "quality" setting
        .AddParam "algorithm", GetAlgoAsString(ddCustom.ListIndex)
        .AddParam "iterations", sldIterations.Value
        
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Function GetQualityAsString(ByVal srcIndex As Long) As String
    Select Case srcIndex
        Case 0
            GetQualityAsString = "fast"
        Case 1
            GetQualityAsString = "precise"
        Case 2
            GetQualityAsString = "custom"
    End Select
End Function

Private Function GetAlgoAsString(ByVal srcIndex As Long) As String
    Select Case srcIndex
        Case 0
            GetAlgoAsString = "box"
        Case 1
            GetAlgoAsString = "am"
        Case 2
            GetAlgoAsString = "deriche"
    End Select
End Function
