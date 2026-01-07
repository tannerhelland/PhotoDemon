VERSION 5.00
Begin VB.Form FormFill 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Fill"
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
   Begin PhotoDemon.pdBrushSelector brshSelector 
      Height          =   1815
      Left            =   6000
      TabIndex        =   2
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3201
      Caption         =   "brush"
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
      TabIndex        =   3
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
      TabIndex        =   4
      Top             =   3120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1296
      Caption         =   "blend mode"
   End
   Begin PhotoDemon.pdDropDown cboAlphaMode 
      Height          =   735
      Left            =   6000
      TabIndex        =   5
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
Attribute VB_Name = "FormFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Selection Fill Dialog
'Copyright 2022-2026 by Tanner Helland
'Created: 04/May/22
'Last updated: 05/May/22
'Last update: wrap up initial build
'
'Users have always been able to fill selections via the paint bucket tool, but this UI provides a shortcut
' method for filling the entire selection or active layer without messing with paint bucket settings.
'
'Note also that the user can ask this function to resize the active layer to match the selection size
' (or the union of the current layer and selection size).  If no selection is active, this function will
' always fill the entire underlying layer, leaving its size unchanged.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Edit > Fill works differently in different tools.
' 1) Photoshop updates layer boundaries to the union of the existing layer and the active selection.
' 2) GIMP leaves the layer boundaries unchanged.  (Paint.NET does too, implicitly, since all layers are image-sized.)
'
'PhotoDemon supports both modes and also adds a 3rd one - cropping the layer to the selection during the fill.
' (Obviously, these settings are only relevant if a selection is active; if one is *not* active, we'll fill the
' entire layer as-is, with no boundary changes.)
Private Enum PD_FillBoundary
    fb_UseLayer = 0
    fb_UseSelection = 1
    fb_UseUnion = 2
End Enum

#If False Then
    Private Const fb_UseLayer = 0, fb_UseSelection = 1, fb_UseUnion = 2
#End If

'Because this function can change layer size (depending on the user's choices), we need to handle previews
' in a non-standard way.  At Form_Load, we'll retrieve a null-padded copy of the current layer and work from there.
Private m_CachedLayer As pdDIB

'The actual effect will be rendered into this DIB, and we'll try to not change its size unless absolutely necessary.
Private m_EffectDIB As pdDIB

'Copy of the current selection mask, resized to match the current working DIB size (in preview mode).
Private m_SelectionMask As pdDIB

'Original boundary rects of the target layer, selection, and union of the two - IN IMAGE COORDINATES
Private m_LayerRect As RectF, m_SelectionRect As RectF, m_UnionRect As RectF

'Find the outline boundary of an image and paint it with a variable pen stroke
Public Sub ApplyFillEffect(ByVal parameterList As String, Optional ByVal toPreview As Boolean = False, Optional ByRef dstPic As pdFxPreviewCtl)
    
    'Failsafe sanity check only
    If (m_CachedLayer Is Nothing) Then CacheActiveLayerAsNullPadded
    If (m_CachedLayer Is Nothing) Then Exit Sub
    
    If (Not toPreview) Then Message "Applying fill..."
    
    'This function behaves differently when a selection is active
    Dim useSelectionData As Boolean
    useSelectionData = PDImages.GetActiveImage.IsSelectionActive()
    
    'Parse out the parameter list.
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString parameterList
    
    Dim fillBrushParams As String, fillOpacity As Single, fillBlendMode As PD_BlendMode, fillAlphaMode As PD_AlphaMode
    Dim fillBoundary As PD_FillBoundary
    
    With cParams
        fillBrushParams = .GetString("fill-brush")
        fillOpacity = .GetSingle("brush-opacity")
        fillBlendMode = .GetLong("blend-mode")
        fillAlphaMode = .GetLong("alpha-mode")
        fillBoundary = .GetLong("boundary")
    End With
    
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
    
    'If we haven't already, retrieve a copy of the active selection mask at the same size as the target DIB
    If useSelectionData Then
        
        If (m_SelectionMask Is Nothing) Then Set m_SelectionMask = New pdDIB
        
        If toPreview Then
            If (m_SelectionMask.GetDIBWidth <> dstDIB.GetDIBWidth) Or (m_SelectionMask.GetDIBHeight <> dstDIB.GetDIBHeight) Then
                m_SelectionMask.CreateBlank dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, 32, 0, 0
                GDI_Plus.GDIPlus_StretchBlt m_SelectionMask, 0, 0, m_SelectionMask.GetDIBWidth, m_SelectionMask.GetDIBHeight, PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, 0, 0, PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, 1!, GP_IM_Bilinear, dstCopyIsOkay:=True
            End If
        Else
            m_SelectionMask.CreateFromExistingDIB PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB
        End If
        
    End If
    
    'Create a temporary fill DIB at the same size as the destination DIB.  (Because the destination DIB
    ' has been auto-scaled to the preview control UI - in preview mode - and regardless of mode is based
    ' on a null-padded copy of the current layer, we can simply use the selection mask copy as-is to
    ' blend the fill onto this effect DIB.)
    If (m_EffectDIB Is Nothing) Then Set m_EffectDIB = New pdDIB
    If (m_EffectDIB.GetDIBWidth <> dstDIB.GetDIBWidth) Or (m_EffectDIB.GetDIBHeight <> dstDIB.GetDIBHeight) Then
        m_EffectDIB.CreateBlank dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, 32, 0, 0
    Else
        m_EffectDIB.ResetDIB 0
    End If
    
    m_EffectDIB.SetInitialAlphaPremultiplicationState True
    
    'In preview mode, all boundary rects all need to be converted to preview sizes
    Dim previewConversion As Double
    If toPreview Then
        previewConversion = workingDIB.GetDIBWidth / m_CachedLayer.GetDIBWidth
    Else
        previewConversion = 1#
    End If
    
    Dim previewLayerRect As RectF, previewSelectionRect As RectF, previewUnionRect As RectF
    With previewLayerRect
        .Left = m_LayerRect.Left * previewConversion
        .Top = m_LayerRect.Top * previewConversion
        .Width = m_LayerRect.Width * previewConversion
        .Height = m_LayerRect.Height * previewConversion
    End With
    
    With previewSelectionRect
        .Left = m_SelectionRect.Left * previewConversion
        .Top = m_SelectionRect.Top * previewConversion
        .Width = m_SelectionRect.Width * previewConversion
        .Height = m_SelectionRect.Height * previewConversion
    End With
    
    With previewUnionRect
        .Left = m_UnionRect.Left * previewConversion
        .Top = m_UnionRect.Top * previewConversion
        .Width = m_UnionRect.Width * previewConversion
        .Height = m_UnionRect.Height * previewConversion
    End With
    
    'The boundary rect we'll use for this operation depends on the user.  (They can choose to use
    ' the original layer or selection rect, or the union of the two.)
    Dim cropToSelectionArea As Boolean: cropToSelectionArea = False
    Dim finalBoundsRect As RectL, srcRectF As RectF
    If (fillBoundary = fb_UseLayer) Then
        srcRectF = previewLayerRect
    ElseIf (fillBoundary = fb_UseSelection) Then
        If useSelectionData Then
            cropToSelectionArea = True
            srcRectF = previewSelectionRect
        Else
            srcRectF = previewLayerRect
        End If
    ElseIf (fillBoundary = fb_UseUnion) Then
        srcRectF = previewUnionRect
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
    
    'Apply the fill to the effect DIB
    Dim cBrush As pd2DBrush
    Set cBrush = New pd2DBrush
    cBrush.SetBrushPropertiesFromXML fillBrushParams
    
    Dim brshBounds As RectF
    If useSelectionData Then
        brshBounds = previewSelectionRect
    Else
        brshBounds = previewLayerRect
    End If
    cBrush.SetBoundaryRect brshBounds
    
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundPDDIB m_EffectDIB
    
    PD2D.FillRectangleI_FromRectF cSurface, cBrush, brshBounds
    Set cBrush = Nothing: Set cSurface = Nothing
    
    'Unpremultiply the filled surface
    m_EffectDIB.SetAlphaPremultiplication False
    
    'm_EffectDIB now has a copy of the fill effect, in a rectangular boundary either the size of
    ' the current layer (if no selection is active) or the size of the current selection.
    '
    'If a selection is active, we need to mask the filled region against the selection mask.
    ' During this step, we will also mask the *layer* image against the user's choice of new boundaries,
    ' if any.
    Dim x As Long, y As Long, xOffset As Long, selValue As Byte
    
    Dim pxWorking() As Byte, saWorking As SafeArray1D, ptrWorking As Long, strideWorking As Long
    dstDIB.WrapArrayAroundScanline pxWorking, saWorking, 0
    ptrWorking = saWorking.pvData
    strideWorking = saWorking.cElements
    
    Dim pxDst() As Byte, saDst As SafeArray1D, ptrDst As Long, strideDst As Long
    m_EffectDIB.WrapArrayAroundScanline pxDst, saDst, 0
    ptrDst = saDst.pvData
    strideDst = saDst.cElements
    
    Dim pxMask() As Byte, saMask As SafeArray1D, ptrMask As Long, strideMask As Long
    If useSelectionData Then
        m_SelectionMask.WrapArrayAroundScanline pxMask, saMask, 0
        ptrMask = saMask.pvData
        strideMask = saMask.cElements
    End If
    
    Dim blendAmount As Double
    Const ONE_DIV_255 As Double = 1# / 255#
    
    For y = 0 To yMax
        
        'Update array pointers to point at the current line in both the source and destination images
        saWorking.pvData = ptrWorking + strideWorking * y
        saDst.pvData = ptrDst + strideDst * y
        If useSelectionData Then saMask.pvData = ptrMask + strideMask * y
        
        'Ignore any lines outside the boundary rect
        If (y >= finalBoundsRect.Top) And (y <= finalBoundsRect.Bottom) Then
            
            For x = 0 To xMax
                
                xOffset = x * 4
                
                If (x >= finalBoundsRect.Left) And (x <= finalBoundsRect.Right) Then
                    
                    If useSelectionData Then
                        
                        selValue = pxMask(x * 4)
                        blendAmount = selValue * ONE_DIV_255
                        
                        'Deal with the filled DIB
                        If (selValue > 0) Then
                            
                            'Reduce the opacity of this pixel proportionately
                            If (selValue < 255) Then
                                pxDst(xOffset + 3) = Int(pxDst(xOffset + 3) * blendAmount)
                                
                            'No need for an Else, because fully masked pixels are left as-is
                            End If
                        
                        'Unmasked pixels are fully removed
                        Else
                            pxDst(xOffset) = 0
                            pxDst(xOffset + 1) = 0
                            pxDst(xOffset + 2) = 0
                            pxDst(xOffset + 3) = 0
                        End If
                        
                        'This branch is only necessary if "use selection bounds" is selected
                        If cropToSelectionArea Then
                            If (selValue < 255) Then
                                pxWorking(xOffset + 3) = Int(pxWorking(xOffset + 3) * blendAmount)
                            End If
                        End If
                        
                    End If
                    
                'OOB pixels are fully removed
                Else
                    pxWorking(xOffset + 3) = 0
                    pxDst(xOffset + 3) = 0
                End If
                
            Next x
            
        Else
            'Make this whole line invisible
            For x = 0 To xMax
                xOffset = x * 4
                pxWorking(xOffset + 3) = 0
                pxDst(xOffset + 3) = 0
            Next x
        End If
        
        If (Not toPreview) Then
            If ((y And progBarCheck) = 0) Then ProgressBars.SetProgBarVal y
        End If
        
    Next y
    
    'Free unsafe array wrappers!
    dstDIB.UnwrapArrayFromDIB pxWorking
    m_EffectDIB.UnwrapArrayFromDIB pxDst
    If useSelectionData Then m_SelectionMask.UnwrapArrayFromDIB pxMask
    
    'Max out the progress bar while we finalize the last few items
    If (Not toPreview) Then ProgressBars.SetProgBarVal ProgressBars.GetProgBarMax
    
    'Premultiply any changed surfaces
    If (Not dstDIB.GetAlphaPremultiplication) Then dstDIB.SetAlphaPremultiplication True
    m_EffectDIB.SetAlphaPremultiplication True
    
    'Use a compositor to blend the finished result onto a temporary copy of workingDIB=
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    cCompositor.QuickMergeTwoDibsOfEqualSize dstDIB, m_EffectDIB, fillBlendMode, fillOpacity, fillAlphaMode, IIf(fillAlphaMode = AM_Inherit, AM_Inherit, AM_Normal)
    
    'In preview mode, pass control to finalizeImageData, which will handle the rest of the rendering
    If toPreview Then
        EffectPrep.FinalizeNonstandardPreview dstPic, True
    
    'In apply mode, set the new DIB as the target layer DIB, manually crop out the target region,
    ' then assign that DIB as the backing surface of the current layer.
    Else
        
        'Crop out the relevant region of the image
        Dim finalWidth As Long, finalHeight As Long
        finalWidth = (finalBoundsRect.Right - finalBoundsRect.Left) + 1
        finalHeight = (finalBoundsRect.Bottom - finalBoundsRect.Top) + 1
        
        Dim cropDIB As pdDIB
        Set cropDIB = New pdDIB
        cropDIB.CreateBlank finalWidth, finalHeight, 32, 0, 0
        GDI.BitBltWrapper cropDIB.GetDIBDC, 0, 0, finalWidth, finalHeight, dstDIB.GetDIBDC, finalBoundsRect.Left, finalBoundsRect.Top, vbSrcCopy
        cropDIB.SetInitialAlphaPremultiplicationState True
        
        'Reset any affine transform data in the target DIB, then assign new left/top offsets and backing surface
        With PDImages.GetActiveImage.GetActiveLayer
            .ResetAffineTransformProperties
            .SetLayerOffsetX finalBoundsRect.Left
            .SetLayerOffsetY finalBoundsRect.Top
            .SetLayerDIB cropDIB
        End With
        
        ProgressBars.ReleaseProgressBar
        
        'Notify the parent image of the change, then redraw the primary viewport before exiting
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
    End If

End Sub

Private Sub brshSelector_BrushChanged()
    UpdatePreview
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
    Process "Fill", , GetLocalParamString(), UNDO_Layer
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub Form_Load()
    
    cmdBar.SetPreviewStatus False
    
    'Cache the active layer as a null-padded layer, and generate boundary rects for all relevant items
    CacheActiveLayerAsNullPadded
    
    Interface.PopulateAlphaModeDropDown cboAlphaMode, AM_Normal
    Interface.PopulateBlendModeDropDown cboBlendMode, BM_Normal
    
    cboLayerSize.SetAutomaticRedraws False
    cboLayerSize.AddItem "do not change", 0
    cboLayerSize.AddItem "use selection size", 1
    cboLayerSize.AddItem "use combined size (union)", 2
    cboLayerSize.ListIndex = 0
    cboLayerSize.SetAutomaticRedraws True, True
    
    'Do not display the "layer size" option if no selection is active.  (We'll simply fill the full layer
    ' as-is in that case.)
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
    
    'Because this function may change layer size, we need to handle it in a non-standard way
    PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB_NullPadded m_CachedLayer, PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
    
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
    If cmdBar.PreviewsAllowed Then Me.ApplyFillEffect GetLocalParamString(), True, pdFxPreview
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
        .AddParam "fill-brush", brshSelector.Brush
        .AddParam "brush-opacity", sldOpacity.Value
        .AddParam "blend-mode", cboBlendMode.ListIndex
        .AddParam "alpha-mode", cboAlphaMode.ListIndex
        .AddParam "boundary", cboLayerSize.ListIndex
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

Private Sub sldOpacity_Change()
    UpdatePreview
End Sub
