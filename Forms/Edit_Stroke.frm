VERSION 5.00
Begin VB.Form FormStroke 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Stroke"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11790
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
   ScaleWidth      =   786
   Begin PhotoDemon.pdPenSelector penSelector 
      Height          =   1815
      Left            =   6000
      TabIndex        =   5
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3201
      Caption         =   "pen"
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
   Begin PhotoDemon.pdCommandBar cmdBar 
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdSlider sldOpacity 
      Height          =   720
      Left            =   6000
      TabIndex        =   2
      Top             =   2280
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1270
      Caption         =   "opacity"
      CaptionPadding  =   2
      Max             =   100
      SigDigits       =   1
      Value           =   100
      NotchValueCustom=   25
      DefaultValue    =   100
   End
   Begin PhotoDemon.pdDropDown cboBlendMode 
      Height          =   735
      Left            =   6000
      TabIndex        =   3
      Top             =   3120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      Caption         =   "blend mode"
   End
   Begin PhotoDemon.pdDropDown cboAlphaMode 
      Height          =   735
      Left            =   6000
      TabIndex        =   4
      Top             =   3960
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      Caption         =   "alpha mode"
   End
   Begin PhotoDemon.pdDropDown cboLayerSize 
      Height          =   735
      Left            =   6000
      TabIndex        =   6
      Top             =   4800
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      Caption         =   "layer size"
   End
End
Attribute VB_Name = "FormStroke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Selection Stroke Dialog
'Copyright 2022-2026 by Tanner Helland
'Created: 04/May/22
'Last updated: 05/May/22
'Last update: wrap up initial build
'
'This UI provides a way for users to stroke the current selection (or layer) boundary with an arbitrary pen.
'
'Note also that the user can ask this function to resize the active layer to match the selection size
' (or the union of the current layer and selection size).  If no selection is active, this function will
' always stroke the border of the underlying layer.  Importantly, *this will still result in a layer size change*,
' as the stroke will increase the dimensions of the underlying layer.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Edit > Stroke works differently in different tools.
' 1) Photoshop updates layer boundaries to the union of the existing layer and the stroke.
' 2) GIMP leaves layer boundaries unchanged, and simply crops the stroke to fit the current layer.
'
'PhotoDemon supports both modes.  (Obviously, this is only relevant if a selection is active;
' if one is *not* active, we'll stroke the entire layer border as-is, with no boundary changes.)
Private Enum PD_StrokeBoundary
    sb_UseLayer = 0
    sb_UseUnion = 1
End Enum

#If False Then
    Private Const sb_UseLayer = 0, sb_UseUnion = 1
#End If

'Because this function can change layer size (depending on the user's choices), we need to handle previews
' in a non-standard way.  At Form_Load, we'll retrieve a null-padded copy of the current layer and work from there.
Private m_CachedLayer As pdDIB

'At preview time, the active selection - if any - gets scaled down to preview size.  We then scan *that* copy
' to produce a (much simpler) stroke path.  This is a workaround until I can solve some issues with
' pd2DPath.GetPathPoints() which relies on GDI+ subpath enumeration (and I can't figure out why it doesn't
' work sometimes).
Private m_SelMaskSmall As pdDIB

'The actual effect will be rendered into this DIB, and we'll try to not change its size unless absolutely necessary.
Private m_EffectDIB As pdDIB

'Copy of the current selection region in pdPath format, resized to match the current working DIB size (in preview mode).
Private m_SelectionPath As pd2DPath

'During previews, we'll cache a "minified" version of the current selection path
Private m_SelectionPathMini As pd2DPath

'Original boundary rects of the target layer, selection, and union of the two - IN IMAGE COORDINATES
Private m_LayerRect As RectF, m_SelectionRect As RectF, m_UnionRect As RectF

'Find the outline boundary of an image and trace it with a variable pen
Public Sub ApplyStrokeEffect(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    If (Not toPreview) Then Message "Stroking outline..."
    
    'This function behaves differently when a selection is active
    Dim useSelectionData As Boolean
    useSelectionData = PDImages.GetActiveImage.IsSelectionActive()
    
    'Parse out the parameter list.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim strokePenParams As String, strokeOpacity As Single, strokeBlendMode As PD_BlendMode, strokeAlphaMode As PD_AlphaMode
    Dim strokeBoundary As PD_StrokeBoundary
    
    With cParams
        strokePenParams = .GetString("stroke-pen")
        strokeOpacity = .GetSingle("brush-opacity")
        strokeBlendMode = .GetLong("blend-mode")
        strokeAlphaMode = .GetLong("alpha-mode")
        strokeBoundary = .GetLong("boundary")
    End With
    
    'Failsafe sanity check only
    If (m_CachedLayer Is Nothing) Then CacheActiveLayerAsNullPadded
    If (m_CachedLayer Is Nothing) Then Exit Sub
    
    'Because this function may change layer size, we need to handle preview behavior manually.
    Dim dstDIB As pdDIB, tmpSA As SafeArray2D
    If toPreview Then
        EffectPrep.PreviewNonStandardImage tmpSA, m_CachedLayer, dstPic, False
        Set dstDIB = workingDIB
    
    'In non-preview "apply the effect" mode, just point at our cached null-padded layer copy
    Else
        Set dstDIB = m_CachedLayer
        If dstDIB.GetAlphaPremultiplication Then dstDIB.SetAlphaPremultiplication False
    End If
    
    'Initialize a progress bar (for non-previews only)
    Dim xMax As Long, yMax As Long
    xMax = dstDIB.GetDIBWidth - 1
    yMax = dstDIB.GetDIBHeight - 1
    
    Dim progBarCheck As Long
    If (Not toPreview) Then
        ProgressBars.SetProgBarMax yMax
        progBarCheck = ProgressBars.FindBestProgBarValue
    End If
    
    'If we haven't already, retrieve a copy of the active selection boundary.
    ' (And if a selection is *not* active, create a path of the active *layer* boundary.)
    Dim outlinePath As pd2DPath
    Set outlinePath = New pd2DPath
    
    If useSelectionData Then
        
        'During preview mode, we must retrieve a copy of the current selection mask, but padded by 1-px on
        ' all sides (for edge detection purposes) and shrunk to the size of the current preview window.
        If toPreview Then
            
            If (m_SelMaskSmall Is Nothing) Then Set m_SelMaskSmall = New pdDIB
            If (m_SelMaskSmall.GetDIBWidth <> dstDIB.GetDIBWidth) Or (m_SelMaskSmall.GetDIBHeight <> dstDIB.GetDIBHeight) Then
                
                m_SelMaskSmall.CreateBlank dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, 32, 0, 0
                GDI_Plus.GDIPlus_StretchBlt m_SelMaskSmall, 0, 0, m_SelMaskSmall.GetDIBWidth, m_SelMaskSmall.GetDIBHeight, PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, 0, 0, PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, 1!, GP_IM_Bilinear, dstCopyIsOkay:=True
                
                'We now need to generate an outline path for this small version of the image.
                ' (This is faster than trying to scale-down the "full size" selection boundary path.)
                
                'Convert the shrunk mask to a 1-byte-per-pixel array with fixed 0/255 values
                Dim maskBytes() As Byte
                If (Not DIBs.GetSingleChannel_2D(m_SelMaskSmall, maskBytes, 0)) Then
                    PDDebug.LogAction "Unexpected channel error"
                    Exit Sub
                End If
                
                Filters_ByteArray.ThresholdByteArray maskBytes, m_SelMaskSmall.GetDIBWidth, m_SelMaskSmall.GetDIBHeight, 1, False
                
                'Convert *that* array to an edge-safe one.
                Dim cEdges As pdEdgeDetector
                Set cEdges = New pdEdgeDetector
                
                Dim safeBytes() As Byte
                cEdges.MakeArrayEdgeSafe maskBytes, safeBytes, dstDIB.GetDIBWidth - 1, dstDIB.GetDIBHeight - 1
                
                'Retrieve a full outline path!
                Set m_SelectionPathMini = Nothing
                If (Not cEdges.FindAllEdges(m_SelectionPathMini, safeBytes, 1, 1, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, -1, -1)) Then
                    PDDebug.LogAction "WARNING: edge detector failed to find preview selection outline"
                End If
                
            End If
            
            'With a "minified" selection boundary generated, we can now clone the path into the local path object
            If (outlinePath Is Nothing) Then Set outlinePath = New pd2DPath
            outlinePath.ResetPath
            outlinePath.CloneExistingPath m_SelectionPathMini
            
        'In non-preview "apply mode", retrieve the current selection boundary as-is.
        Else
            If (m_SelectionPath Is Nothing) Then Set m_SelectionPath = PDImages.GetActiveImage.MainSelection.GetSelectionBoundaryPath()
            outlinePath.CloneExistingPath m_SelectionPath
            Set m_SelMaskSmall = Nothing
        End If
        
        'Because the path generated by the selection was guaranteed to come from a marching-squares implementation,
        ' we want to simplify it prior to rendering to improve antialiasing quality.
        PDMath.SimplifyPathFromMarchingSquares outlinePath
        
    Else
        Dim tmpLayerBoundsRect As RectF
        PDImages.GetActiveImage.GetActiveLayer.GetLayerBoundaryRect tmpLayerBoundsRect
        outlinePath.AddRectangle_RectF tmpLayerBoundsRect
    End If
    
    'Create a temporary stroke DIB at the same size as the destination DIB.  (We need a separate DIB so we
    ' can apply custom blend- and/or alpha-mode settings.)
    If (m_EffectDIB Is Nothing) Then Set m_EffectDIB = New pdDIB
    If (m_EffectDIB.GetDIBWidth <> dstDIB.GetDIBWidth) Or (m_EffectDIB.GetDIBHeight <> dstDIB.GetDIBHeight) Then
        m_EffectDIB.CreateBlank dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, 32, 0, 0
    Else
        m_EffectDIB.ResetDIB 0
    End If
    
    m_EffectDIB.SetInitialAlphaPremultiplicationState True
    
    'In preview mode, all boundary rects all need to be converted to preview sizes - and so does the
    ' stroke width, to ensure it actually looks how it will in the final render.
    Dim previewConversion As Double
    If toPreview Then
        
        previewConversion = workingDIB.GetDIBWidth / m_CachedLayer.GetDIBWidth
        
        'Because we generated an outline path manually from a shrunk copy of the selection mask,
        ' our outline path is already OK *if* it's based on the active selection.
        If (Not useSelectionData) Then
            Dim tmpTransform As pd2DTransform
            Set tmpTransform = New pd2DTransform
            tmpTransform.ApplyScaling previewConversion, previewConversion
            outlinePath.ApplyTransformation tmpTransform
        End If
        
    Else
        previewConversion = 1#
    End If
    
    Dim previewLayerRect As RectF
    With previewLayerRect
        .Left = m_LayerRect.Left * previewConversion
        .Top = m_LayerRect.Top * previewConversion
        .Width = m_LayerRect.Width * previewConversion
        .Height = m_LayerRect.Height * previewConversion
    End With
    
    'The boundary rect we'll use for this operation depends on the user.
    ' (They can choose to use the original layer or selection rect, or the union of the two.)
    Dim cropToLayerArea As Boolean
    If useSelectionData Then
        cropToLayerArea = (strokeBoundary = sb_UseLayer)
    Else
        cropToLayerArea = False
    End If
    
    Dim finalBoundsRect As RectL, srcRectF As RectF
    If cropToLayerArea Then
        srcRectF = previewLayerRect
    Else
        With srcRectF
            .Left = 0
            .Top = 0
            .Width = dstDIB.GetDIBWidth
            .Height = dstDIB.GetDIBHeight
        End With
    End If
    
    finalBoundsRect.Left = Int(srcRectF.Left)
    finalBoundsRect.Top = Int(srcRectF.Top)
    finalBoundsRect.Right = finalBoundsRect.Left + Int(PDMath.Frac(srcRectF.Left) + srcRectF.Width + 0.5) - 1
    finalBoundsRect.Bottom = finalBoundsRect.Top + Int(PDMath.Frac(srcRectF.Top) + srcRectF.Height + 0.5) - 1
    
    'Sanity check for OOB layers
    If (finalBoundsRect.Left < 0) Then finalBoundsRect.Left = 0
    If (finalBoundsRect.Top < 0) Then finalBoundsRect.Top = 0
    If (finalBoundsRect.Right >= dstDIB.GetDIBWidth) Then finalBoundsRect.Right = dstDIB.GetDIBWidth - 1
    If (finalBoundsRect.Bottom >= dstDIB.GetDIBHeight) Then finalBoundsRect.Bottom = dstDIB.GetDIBHeight - 1
    
    'Apply the stroke to the effect DIB
    Dim cPen As pd2DPen
    Set cPen = New pd2DPen
    cPen.SetPenPropertiesFromXML strokePenParams
    
    'In preview mode, reduce the width of the pen proportionally
    If toPreview Then
        Dim curWidth As Single
        curWidth = cPen.GetPenWidth()
        curWidth = curWidth * previewConversion
        If (curWidth < 1!) Then curWidth = 1!
        cPen.SetPenWidth curWidth
    End If
    
    'TODO: expose antialiasing and pixel offset as settings??
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundPDDIB m_EffectDIB
    cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
    cSurface.SetSurfacePixelOffset P2_PO_Half
    
    'NOTE: this line is the biggest performance hit on this function.  Simplifying the underlying path
    ' may help; it's on my TODO list.
    PD2D.DrawPath cSurface, cPen, outlinePath
    Set cPen = Nothing: Set cSurface = Nothing
    
    'Unpremultiply the stroked surface
    m_EffectDIB.SetAlphaPremultiplication False
    
    'm_EffectDIB now contains a copy of the stroke effect.  We need to blend this onto the base layer,
    ' but if the user wants the result cropped to the current layer's size, we need to manually crop
    ' the stroke effect to match!
    If (useSelectionData And cropToLayerArea) Then
    
        Dim x As Long, y As Long, xOffset As Long
        
        Dim pxDst() As Byte, saDst As SafeArray1D, ptrDst As Long, strideDst As Long
        m_EffectDIB.WrapArrayAroundScanline pxDst, saDst, 0
        ptrDst = saDst.pvData
        strideDst = saDst.cElements
        
        For y = 0 To yMax
            
            'Update array pointers to point at the current line in both the source and destination images
            saDst.pvData = ptrDst + strideDst * y
            
            'Ignore any lines outside the boundary rect
            If (y >= finalBoundsRect.Top) And (y <= finalBoundsRect.Bottom) Then
                
                'Remove any OOB pixels
                For x = 0 To xMax
                    xOffset = x * 4
                    If (x < finalBoundsRect.Left) Or (x > finalBoundsRect.Right) Then pxDst(xOffset + 3) = 0
                Next x
                
            Else
                'Make this whole line invisible
                For x = 0 To xMax
                    xOffset = x * 4
                    pxDst(xOffset + 3) = 0
                Next x
            End If
            
            If (Not toPreview) Then
                If ((y And progBarCheck) = 0) Then ProgressBars.SetProgBarVal y
            End If
            
        Next y
        
        'Free unsafe array wrappers!
        m_EffectDIB.UnwrapArrayFromDIB pxDst
        
    End If
    
    'Max out the progress bar while we finalize the last few items
    If (Not toPreview) Then ProgressBars.SetProgBarVal ProgressBars.GetProgBarMax
    
    'Premultiply any changed surfaces
    If (Not dstDIB.GetAlphaPremultiplication) Then dstDIB.SetAlphaPremultiplication True
    m_EffectDIB.SetAlphaPremultiplication True
    
    'Use a compositor to blend the finished result onto the destination DIB
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    cCompositor.QuickMergeTwoDibsOfEqualSize dstDIB, m_EffectDIB, strokeBlendMode, strokeOpacity, strokeAlphaMode, IIf(strokeAlphaMode = AM_Inherit, AM_Inherit, AM_Normal)
    
    'In preview mode, pass control to finalizeImageData, which will handle the rest of the rendering
    If toPreview Then
        EffectPrep.FinalizeNonstandardPreview dstPic, True
    
    'In apply mode, set the new DIB as the target layer DIB, then ask it to auto-crop.
    Else
        
        'Reset any affine transform data in the target DIB, then assign new left/top offsets and backing surface
        With PDImages.GetActiveImage.GetActiveLayer
            .ResetAffineTransformProperties
            .SetLayerOffsetX finalBoundsRect.Left
            .SetLayerOffsetY finalBoundsRect.Top
            .SetLayerDIB dstDIB
            .CropNullPaddedLayer
        End With
        
        ProgressBars.ReleaseProgressBar
        
        'Notify the parent image of the change, then redraw the primary viewport before exiting
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
    End If
    
End Sub

Private Sub cboAlphaMode_Click()
    UpdatePreview
End Sub

Private Sub cboBlendMode_Click()
    UpdatePreview
End Sub

Private Sub cboLayerSize_Click()
    UpdatePreview
End Sub

Private Sub cmdBar_OKClick()
    Process "Stroke", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    Interface.PopulateAlphaModeDropDown cboAlphaMode, AM_Normal
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Normal
    
    cboLayerSize.SetAutomaticRedraws False
    cboLayerSize.AddItem "do not change", 0
    cboLayerSize.AddItem "expand to fit stroke (if necessary)", 1
    cboLayerSize.ListIndex = 0
    cboLayerSize.SetAutomaticRedraws True, True
    
    'Do not display the "layer size" option if no selection is active.
    ' (We'll simply stroke the layer boundary in that case, and this *always* requires an expansion.)
    cboLayerSize.Visible = PDImages.GetActiveImage.IsSelectionActive()
    
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

Private Sub CacheActiveLayerAsNullPadded()
    
    'Because this function may change layer size, we need to handle preview caching in a non-standard way.
    ' Manually retrieve a null-padded copy of the active layer and store it locally.
    If (m_CachedLayer Is Nothing) Then PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB_NullPadded m_CachedLayer, PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
    
    'Retrieve - in image coordinates - the boundary rect of the target layer and active selection.
    PDImages.GetActiveImage.GetActiveLayer.GetLayerBoundaryRect m_LayerRect
    
    '(If a selection is not active, simply mirror the layer selection.)
    If PDImages.GetActiveImage.IsSelectionActive() Then
        m_SelectionRect = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect
    Else
        m_SelectionRect = m_LayerRect
    End If
    
    'Pre-calculate the union of the two as well
    PDMath.UnionRectF m_UnionRect, m_LayerRect, m_SelectionRect, False
    
End Sub

Private Sub UpdatePreview()
    If cmdBar.PreviewsAllowed Then Me.ApplyStrokeEffect GetLocalParamString(), True, pdFxPreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "stroke-pen", penSelector.Pen
        .AddParam "brush-opacity", sldOpacity.Value
        .AddParam "blend-mode", cboBlendMode.ListIndex
        .AddParam "alpha-mode", cboAlphaMode.ListIndex
        .AddParam "boundary", cboLayerSize.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Sub penSelector_PenChanged(ByVal isFinalChange As Boolean)
    If isFinalChange Then UpdatePreview
End Sub

Private Sub sldOpacity_Change()
    UpdatePreview
End Sub
