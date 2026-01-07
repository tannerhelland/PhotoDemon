VERSION 5.00
Begin VB.Form FormVignette 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Vignetting"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12090
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
   ScaleHeight     =   443
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   806
   Begin PhotoDemon.pdButtonStrip btsShape 
      Height          =   975
      Left            =   6000
      TabIndex        =   8
      Top             =   3960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      Caption         =   "shape"
   End
   Begin PhotoDemon.pdSlider sltXCenter 
      Height          =   405
      Left            =   6000
      TabIndex        =   2
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   4
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5895
      Width           =   12090
      _ExtentX        =   21325
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
      PointSelection  =   -1  'True
   End
   Begin PhotoDemon.pdSlider sltRadius 
      Height          =   705
      Left            =   6000
      TabIndex        =   3
      Top             =   1440
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1270
      Caption         =   "radius"
      Min             =   1
      Max             =   100
      Value           =   50
      NotchPosition   =   2
      NotchValueCustom=   50
   End
   Begin PhotoDemon.pdSlider sltFeathering 
      Height          =   705
      Left            =   6000
      TabIndex        =   4
      Top             =   2160
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1244
      Caption         =   "softness"
      Min             =   1
      Max             =   100
      Value           =   1
      DefaultValue    =   1
   End
   Begin PhotoDemon.pdSlider sltTransparency 
      Height          =   705
      Left            =   9000
      TabIndex        =   5
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "opacity"
      Min             =   1
      Max             =   100
      Value           =   100
      DefaultValue    =   100
   End
   Begin PhotoDemon.pdColorSelector csVignette 
      Height          =   810
      Left            =   6000
      TabIndex        =   6
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1429
      Caption         =   "color"
      curColor        =   0
   End
   Begin PhotoDemon.pdSlider sltYCenter 
      Height          =   405
      Left            =   9000
      TabIndex        =   7
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Max             =   1
      SigDigits       =   4
      Value           =   0.5
      NotchPosition   =   2
      NotchValueCustom=   0.5
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   435
      Index           =   0
      Left            =   6120
      Top             =   1050
      Width           =   5655
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   2
      Caption         =   "you can also set a center position by clicking the preview window"
      FontSize        =   9
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   6000
      Top             =   120
      Width           =   5925
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "center position (x, y)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdSlider sltAspectRatio 
      Height          =   705
      Left            =   6000
      TabIndex        =   9
      Top             =   5010
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "aspect ratio"
      Min             =   0.2
      Max             =   4
      SigDigits       =   3
      ScaleStyle      =   1
      ScaleExponent   =   5
      Value           =   1
      NotchPosition   =   2
      NotchValueCustom=   1
   End
   Begin PhotoDemon.pdSlider sltAngle 
      Height          =   705
      Left            =   9000
      TabIndex        =   10
      Top             =   5010
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1244
      Caption         =   "angle"
      Max             =   360
      SigDigits       =   1
   End
End
Attribute VB_Name = "FormVignette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Vignette tool
'Copyright 2013-2026 by Tanner Helland
'Created: 31/January/13
'Last updated: 27/February/17
'Last update: large performance improvements; added "custom shape" mode
'
'This tool allows the user to apply vignetting to an image.  Many options are available, and all should be
' self-explanatory.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To improve performance, we cache a temporary DIB at form level; the DIB is freed with the form.
Private m_OverlayDIB As pdDIB

'Apply vignetting to an image
Public Sub ApplyVignette(ByVal vignetteParams As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Applying vignetting..."
    
    'Parse out individual parameters from the incoming XML packet.  (Note that not all modes will use all settings.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString vignetteParams
    
    Dim maxRadius As Double, vFeathering As Double, vTransparency As Double, vMode As Long
    Dim vColor As Long, centerPosX As Double, centerPosY As Double, vAspectRatio As Double, vAngle As Double
    
    With cParams
        maxRadius = .GetDouble("radius", 50#)
        vFeathering = .GetDouble("softness", 0#)
        vTransparency = .GetDouble("strength", 100#)
        vMode = .GetLong("shape", 0)
        centerPosX = .GetDouble("centerx", 0.5)
        centerPosY = .GetDouble("centery", 0.5)
        vColor = .GetLong("color", vbBlack)
        vAspectRatio = .GetDouble("aspectratio", 1#)
        vAngle = .GetDouble("angle", 0#)
    End With
    
    'Prep a working copy of the source image, and note that we leave the color data premultiplied.
    ' (We're only going to be blending atop the source, so we don't need to un-premultiply it.)
    Dim dstImageData() As Long
    Dim dstSA As SafeArray2D
    EffectPrep.PrepImageData dstSA, toPreview, dstPic, , , True
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Prep an overlay at the same size as the underlying image.  (We're going to render the vignette to
    ' this overlay, then merge down onto the existing image at the very end.)
    If (m_OverlayDIB Is Nothing) Then Set m_OverlayDIB = New pdDIB
    If (m_OverlayDIB.GetDIBWidth = workingDIB.GetDIBWidth) And (m_OverlayDIB.GetDIBHeight = workingDIB.GetDIBHeight) Then
        m_OverlayDIB.ResetDIB 0
    Else
        m_OverlayDIB.CreateBlank workingDIB.GetDIBWidth, workingDIB.GetDIBHeight, 32, 0, 0
        m_OverlayDIB.SetInitialAlphaPremultiplicationState True
    End If
    
    m_OverlayDIB.WrapLongArrayAroundDIB dstImageData, dstSA
    
    'Calculate the center of the image, in absolute pixels
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * centerPosX
    midX = midX + initX
    midY = CDbl(finalY - initY) * centerPosY
    midY = midY + initY
            
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    Dim nX2 As Double, nY2 As Double
            
    'Radius is based off the smaller of the two dimensions - width or height.  (This is used in the "circle" mode.)
    Dim tWidth As Long, tHeight As Long
    tWidth = curDIBValues.Width
    tHeight = curDIBValues.Height
    
    Dim sRadiusW As Double, sRadiusH As Double
    Dim sRadiusW2 As Double, sRadiusH2 As Double
    Dim minDimension As Double
    
    If (vMode = 0) Or (vMode = 1) Then
        sRadiusW = tWidth * (maxRadius / 100#)
        sRadiusH = tHeight * (maxRadius / 100#)
    Else
        If (tWidth < tHeight) Then minDimension = tWidth Else minDimension = tHeight
        sRadiusW = minDimension * (maxRadius / 100) * vAspectRatio
        sRadiusH = minDimension * (maxRadius / 100) * (1# / vAspectRatio)
    End If
    
    sRadiusW2 = sRadiusW * sRadiusW
    sRadiusH2 = sRadiusH * sRadiusH
    
    'Adjust the vignetting to be a proportion of the image's maximum radius.  This ensures accurate correlations
    ' between the preview and the final result.
    Dim vFeathering2 As Double
    
    If (vMode = 0) Then
        vFeathering2 = (vFeathering / 100) * (sRadiusW * sRadiusH)
    ElseIf (vMode = 1) Then
        If (sRadiusW < sRadiusH) Then minDimension = sRadiusW Else minDimension = sRadiusH
        vFeathering2 = (vFeathering / 100) * (minDimension * minDimension)
    ElseIf (vMode = 2) Then
        vFeathering = 1# - (vFeathering / 100)
        If (vFeathering = 1#) Then vFeathering = 0.99999
    End If
    
    'Calculate the smaller of the two radii.  (Used for "circular" mode.)
    Dim sRadiusCircular As Double, sRadiusMax As Double, sRadiusMin As Double
    If (sRadiusW < sRadiusH) Then sRadiusCircular = sRadiusW2 Else sRadiusCircular = sRadiusH2
    sRadiusMin = sRadiusCircular - vFeathering2
    
    'In "custom aspect ratio" mode, we want to cache a few other relevant values
    Dim vCos As Double, vSin As Double
    vAngle = vAngle * (PI / 180#)
    vCos = Cos(vAngle)
    vSin = Sin(vAngle)
    
    Dim tmpH As Double, tmpV As Double
    
    Dim blendVal As Double
    
    'Build a lookup table of vignette values.  Because we're just applying the vignette to a standalone layer,
    ' we can treat the vignette as a constant color scaled from transparent to opaque.  This makes it *very*
    ' fast to apply.
    Dim vLookup() As Long
    ReDim vLookup(0 To 255) As Long
    Dim tmpQuad As RGBQuad
    
    'Extract the RGB values of the vignetting color
    Dim newR As Byte, newG As Byte, newB As Byte
    newR = Colors.ExtractRed(vColor)
    newG = Colors.ExtractGreen(vColor)
    newB = Colors.ExtractBlue(vColor)
    
    For x = 0 To 255
        With tmpQuad
            .Alpha = x
            blendVal = CSng(x / 255)
            .Red = Int(blendVal * CSng(newR))
            .Green = Int(blendVal * CSng(newG))
            .Blue = Int(blendVal * CSng(newB))
        End With
        CopyMemoryStrict VarPtr(vLookup(x)), VarPtr(tmpQuad), 4&
    Next x
    
    'And that's it!  Loop through each pixel in the image, converting values as we go.
    For x = initX To finalX
    For y = initY To finalY
    
        'Remap the coordinates around a center point of (0, 0)
        nX = x - midX
        nY = y - midY
        nX2 = nX * nX
        nY2 = nY * nY
        
        'Based on the current render mode, figure out if this pixel lies...
        ' 1) Outside the vignette (so it's forced to the vignette color)
        ' 2) Inside the vignette (so it's left alone)
        ' 3) Somewhere between (1) and (2) (so it's feathered, per the caller's parameters)
        
        'Fit to image (elliptical)
        If (vMode = 0) Then
            
            sRadiusMax = sRadiusH2 - ((sRadiusH2 * nX2) / sRadiusW2)
            
            'Outside
            If (nY2 > sRadiusMax) Then
                dstImageData(x, y) = vLookup(255)
            
            'Inside
            Else
                
                sRadiusMin = sRadiusMax - vFeathering2
                
                'Feathered
                If (nY2 >= sRadiusMin) Then
                    blendVal = (nY2 - sRadiusMin) / vFeathering2
                    dstImageData(x, y) = vLookup(blendVal * 255)
                End If
                    
            End If
                
        'Circular
        ElseIf (vMode = 1) Then
        
            'Outside
            If ((nX2 + nY2) > sRadiusCircular) Then
                dstImageData(x, y) = vLookup(255)
                
            'Inside
            Else
                
                'Feathered
                If ((nX2 + nY2) >= sRadiusMin) Then
                    blendVal = (nX2 + nY2 - sRadiusMin) / vFeathering2
                    dstImageData(x, y) = vLookup(blendVal * 255)
                End If
                
            End If
                
        'Custom
        Else
        
            tmpH = vCos * nX + vSin * nY
            tmpV = vSin * nX - vCos * nY
            sRadiusMax = (tmpH * tmpH / sRadiusW2) + (tmpV * tmpV / sRadiusH2)
            
            'Outside
            If (sRadiusMax > 1#) Then
                dstImageData(x, y) = vLookup(255)
            Else
            
                'Feathered
                If (sRadiusMax >= vFeathering) Then
                    blendVal = 1# - (1 - sRadiusMax) / (1 - vFeathering)
                    dstImageData(x, y) = vLookup(blendVal * 255)
                End If
            
            End If
        
        End If
                        
    Next y
        If (Not toPreview) Then
            If ((x And progBarCheck) = 0) Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x
            End If
        End If
    Next x
    
    'Merge the final result onto the working layer
    Dim tmpCompositor As pdCompositor
    Set tmpCompositor = New pdCompositor
    tmpCompositor.QuickMergeTwoDibsOfEqualSize workingDIB, m_OverlayDIB, , vTransparency
    
    m_OverlayDIB.UnwrapLongArrayFromDIB dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData toPreview, dstPic, True
        
End Sub

Private Sub btsShape_Click(ByVal buttonIndex As Long)
    UpdateShapeVisibility
    UpdatePreview
End Sub

Private Sub UpdateShapeVisibility()
    sltAspectRatio.Visible = (btsShape.ListIndex = 2)
    sltAngle.Visible = sltAspectRatio.Visible
End Sub

Private Sub cmdBar_OKClick()
    Process "Vignetting", , GetFunctionParams(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    csVignette.Color = RGB(0, 0, 0)
End Sub

Private Sub csVignette_ColorChanged()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    btsShape.AddItem "fit to image", 0
    btsShape.AddItem "circular", 1
    btsShape.AddItem "custom", 2
    btsShape.ListIndex = 0
    UpdateShapeVisibility
    
    ApplyThemeAndTranslations Me, True, True
    cmdBar.SetPreviewStatus True
    UpdatePreview
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub pdFxPreview_ColorSelected()
    csVignette.Color = pdFxPreview.SelectedColor
    UpdatePreview
End Sub

'The user can right-click the preview area to select a new center point
Private Sub pdFxPreview_PointSelected(xRatio As Double, yRatio As Double)
    cmdBar.SetPreviewStatus False
    sltXCenter.Value = xRatio
    sltYCenter.Value = yRatio
    cmdBar.SetPreviewStatus True
    UpdatePreview
End Sub

Private Sub sltAngle_Change()
    UpdatePreview
End Sub

Private Sub sltAspectRatio_Change()
    UpdatePreview
End Sub

Private Sub sltFeathering_Change()
    UpdatePreview
End Sub

Private Sub sltRadius_Change()
    UpdatePreview
End Sub

Private Sub sltTransparency_Change()
    UpdatePreview
End Sub

'Redraw the on-screen preview of the transformed image
Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then ApplyVignette GetFunctionParams(), True, pdFxPreview
End Sub

Private Function GetFunctionParams() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "radius", sltRadius.Value
        .AddParam "softness", sltFeathering.Value
        .AddParam "strength", sltTransparency.Value
        .AddParam "shape", btsShape.ListIndex
        .AddParam "centerx", sltXCenter.Value
        .AddParam "centery", sltYCenter.Value
        .AddParam "color", csVignette.Color
        .AddParam "aspectratio", sltAspectRatio.Value
        .AddParam "angle", sltAngle.Value
    End With
    
    GetFunctionParams = cParams.GetParamString()

End Function

'If the user changes the position and/or zoom of the preview viewport, the entire preview must be redrawn.
Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

Private Sub sltXCenter_Change()
    UpdatePreview
End Sub

Private Sub sltYCenter_Change()
    UpdatePreview
End Sub
