VERSION 5.00
Begin VB.Form FormFragment 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Fragment"
   ClientHeight    =   6525
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   802
   Begin PhotoDemon.pdSlider sldOpacity 
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   4320
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      Min             =   1
      Max             =   100
      Value           =   50
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdButtonStrip btsOpacity 
      Height          =   975
      Left            =   6000
      TabIndex        =   5
      Top             =   3240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1720
      Caption         =   "opacity"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   12030
      _ExtentX        =   21220
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
      DisableZoomPan  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltDistance 
      Height          =   705
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1270
      Caption         =   "distance"
      Max             =   100
      SigDigits       =   2
      Value           =   5
      DefaultValue    =   5
   End
   Begin PhotoDemon.pdSlider sltFragments 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1270
      Caption         =   "number of fragments"
      Min             =   2
      Max             =   25
      Value           =   4
      DefaultValue    =   4
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   2400
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1270
      Caption         =   "angle"
      Max             =   360
      SigDigits       =   1
   End
End
Attribute VB_Name = "FormFragment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Fragment Filter Dialog
'Copyright 2017-2026 by Tanner Helland
'Created: 01/August/17
'Last updated: 01/August/17
'Last update: complete rewrite using new, original algorithm.  (Performance increase is ~20x over the old method,
'             so a pretty great improvement!)
'
'The fragment tool superimposes multiple copies of an image over itself, at specified angle and distance offsets.
' pd2D is used to greatly improve rendering performance.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Sub Fragment(ByVal effectParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
   
    If (Not toPreview) Then Message "Calculating image fragments..."
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString effectParams
    
    Dim numOfFragments As Long, fragmentDistance As Double, startingAngle As Double
    Dim opacityAutomatic As Boolean, opacityManual As Double
        
    With cParams
        numOfFragments = .GetLong("count", sltFragments.Value)
        fragmentDistance = .GetDouble("distance", sltDistance.Value)
        startingAngle = .GetDouble("angle", sltAngle.Value)
        opacityAutomatic = .GetBool("opacityauto", False)
        opacityManual = .GetDouble("opacitymanual", sldOpacity.Value)
    End With
    
    'Request a working copy of the current layer
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, , , True
    
    'Make a copy of said working data
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    srcDIB.CreateFromExistingDIB workingDIB
    
    'If this is a preview, we need to adjust the distance values to match the size of the preview box
    If (curDIBValues.Width < curDIBValues.Height) Then
        fragmentDistance = fragmentDistance * 0.005 * curDIBValues.Width
    Else
        fragmentDistance = fragmentDistance * 0.005 * curDIBValues.Height
    End If
    If (fragmentDistance < 0.1) Then fragmentDistance = 0.1
    
    'To keep processing quick, only update the progress bar when absolutely necessary. This function calculates that value
    ' based on the size of the area to be processed.
    If (Not toPreview) Then ProgressBars.SetProgBarMax numOfFragments
    
    'Calculate the step between angles.  (This value is only needed if numOfFragments > 1, obviously.)
    Dim angleStep As Double
    angleStep = 360# / CDbl(numOfFragments)
    
    'Calculate a similar step for opacity.
    Dim opacityLevel As Double
    If opacityAutomatic Then
    
        'Opacity scales non-linearly with the number of fragments.  (These constants were arrived at by
        ' trial-and-error, frankly; a straight division results in alpha values too low, so we artificially
        ' boost the opacity, while not allowing it to drop to the point where everything becomes a
        ' muddled gray.)
        opacityLevel = 100# / (numOfFragments + 1) * 1.667
        If (opacityLevel > 100#) Then opacityLevel = 100#
        If (opacityLevel < 7.5) Then opacityLevel = 7.5
        
    Else
        opacityLevel = opacityManual
    End If
    
    'Wrap pd2D surfaces around the source and destination images.
    Dim srcSurface As pd2DSurface, dstSurface As pd2DSurface
    Set srcSurface = New pd2DSurface: Set dstSurface = New pd2DSurface
    srcSurface.WrapSurfaceAroundPDDIB srcDIB
    dstSurface.WrapSurfaceAroundPDDIB workingDIB
    dstSurface.SetSurfaceResizeQuality P2_RQ_Bilinear
    
    'Prep a transformer.  (It will calculate rotations for us.)
    Dim cTransform As pd2DTransform
    Set cTransform = New pd2DTransform
    cTransform.ApplyTranslation_Polar startingAngle, fragmentDistance, True
    
    Dim topLeft As PointFloat
    
    'Starting at the user's specified initial angle, superimpose copies of the image at the specified intervals
    Dim i As Long
    For i = 1 To numOfFragments
        
        'Calculate a top-left point for this copy
        topLeft.x = 0#
        topLeft.y = 0#
        cTransform.ApplyTransformToPointF topLeft
        
        'Draw a superimposed, semi-transparent copy
        PD2D.DrawSurfaceF dstSurface, topLeft.x, topLeft.y, srcSurface, opacityLevel
        
        'Advance the rotation
        cTransform.Reset
        cTransform.ApplyTranslation_Polar startingAngle + (angleStep * i), fragmentDistance, True
        
        'Provide feedback (when not in preview mode, obviously)
        If (Not toPreview) Then ProgressBars.SetProgBarVal i
        
    Next i
    
    'Free our pd2D surfaces
    Set srcSurface = Nothing
    Set dstSurface = Nothing
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
        
End Sub

Private Sub btsOpacity_Click(ByVal buttonIndex As Long)
    ToggleOpacitySliderVisibility
    UpdatePreview
End Sub

Private Sub ToggleOpacitySliderVisibility()
    sldOpacity.Visible = (btsOpacity.ListIndex = 1)
End Sub

Private Sub cmdBar_OKClick()
    Process "Fragment", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()

    'Disable previews until the dialog has been fully initialized
    cmdBar.SetPreviewStatus False
    
    btsOpacity.AddItem "auto", 0
    btsOpacity.AddItem "manual", 1
    btsOpacity.ListIndex = 0
    ToggleOpacitySliderVisibility
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.Fragment GetLocalParamString(), True, pdFxPreview
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sldOpacity_Change()
    UpdatePreview
End Sub

Private Sub sltAngle_Change()
    UpdatePreview
End Sub

Private Sub sltDistance_Change()
    UpdatePreview
End Sub

Private Sub sltFragments_Change()
    UpdatePreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "count", sltFragments.Value
        .AddParam "distance", sltDistance.Value
        .AddParam "angle", sltAngle.Value
        .AddParam "opacityauto", (btsOpacity.ListIndex = 0)
        .AddParam "opacitymanual", sldOpacity.Value
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function
