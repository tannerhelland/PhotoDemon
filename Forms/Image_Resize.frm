VERSION 5.00
Begin VB.Form FormResize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   " Resize image"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14430
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
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   962
   Begin PhotoDemon.pdCheckBox chkPreview 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      Caption         =   "preview changes"
   End
   Begin PhotoDemon.pdPictureBoxInteractive picPreview 
      Height          =   4920
      Left            =   120
      Top             =   240
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   8678
   End
   Begin PhotoDemon.pdLabel lblLanczos 
      Height          =   375
      Left            =   10200
      Top             =   4185
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "radius:"
   End
   Begin PhotoDemon.pdSlider sldLanczos 
      Height          =   375
      Left            =   11400
      TabIndex        =   5
      Top             =   4125
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Min             =   2
      Value           =   3
      NotchPosition   =   2
      NotchValueCustom=   3
   End
   Begin PhotoDemon.pdCheckBox chkEstimate 
      Height          =   375
      Left            =   6495
      TabIndex        =   4
      Top             =   4170
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   661
      Caption         =   "optimize for speed"
   End
   Begin PhotoDemon.pdDropDown cmbFit 
      Height          =   855
      Left            =   6360
      TabIndex        =   2
      Top             =   4710
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1508
      Caption         =   "when changing aspect ratio, fit image to new size by"
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5730
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdResize ucResize 
      Height          =   2850
      Left            =   5760
      TabIndex        =   1
      Top             =   240
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5027
   End
   Begin PhotoDemon.pdDropDown cboResample 
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   3240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1296
      Caption         =   "resampling"
   End
End
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Size Handler
'Copyright 2001-2026 by Tanner Helland
'Created: 12/December/01
'Last updated: 09/April/22
'Last update: ensure at least 1x1 source pixels exist when generating a preview (this solves an issue where
'             previews could disappear when resizing from e.g. 1x1 to 5000x5000)
'
'Standard image resize dialog.  A number of resample algorithms are provided, with some being provided
' by the 3rd-party FreeImage library.  PD also supports three different modes of "fitting" the resized
' image into the new size - standard (which stretches the image as necessary), inclusive (which preserves
' aspect ratio and fits the larger image dimension completely within the new boundaries, with empty
' borders as necessary), and exclusive (which preserves aspect ratio and fits the smaller image dimension
' completely within the new boundaries, cropping the other image dimension as necessary).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Internal flag to use/not use GDI+ resize functions.  GDI+ is significantly faster than our internal resampler,
' but its algorithms are nonstandard and have some weird quirks vs theoretically "correct" implementations
' (see https://photosauce.net/blog/post/image-scaling-with-gdi-part-4-examining-the-interpolationmode-values).
' My current feeling is that the performance vs quality trade-offs are worth it, and the GDI+ resamplers should
' be available to users (for the three resampling modes where they apply, and only if fast "approximation" mode
' is enabled).  Thus I prefer this constant set to TRUE in production code.
Private Const ALLOW_GDIPLUS_RESIZE As Boolean = True

'Internal flag to report performance.  Very helpful while debugging; not helpful in production.
Private Const REPORT_PREVIEW_PERF As Boolean = False

'PhotoDemon's resampling options have expanded over the years.  Unfortunately, old versions of the app
' stored resampling settings using 0-based integers.  I have since switched to a version-agnostic string system,
' but these old enums are kept around so PD can map old macros to modern resampling constants.
Private Enum PD_ResampleOld_V1
    ResizeNormal = 0
    ResizeBilinear = 1
    ResizeBspline = 2
    ResizeBicubicMitchell = 3
    ResizeBicubicCatmull = 4
    ResizeSincLanczos = 5
End Enum

#If False Then
    Private Const ResizeNormal = 0, ResizeBilinear = 1, ResizeBspline = 2, ResizeBicubicMitchell = 3, ResizeBicubicCatmull = 4, ResizeSincLanczos = 5
#End If

Private Enum PD_ResampleOld_V2
    pdrc_Automatic = 0
    pdrc_NearestNeighbor = 1
    pdrc_BilinearNormal = 2
    pdrc_BilinearShrink = 3
    pdrc_BicubicNormal = 4
    pdrc_BicubicShrink = 5
    pdrc_Mitchell = 6
    pdrc_CatmullRom = 7
    pdrc_Sinc = 8
End Enum

#If False Then
    Private Const pdrc_Automatic = 0, pdrc_NearestNeighbor = 1, pdrc_BilinearNormal = 2, pdrc_BilinearShrink = 3, pdrc_BicubicNormal = 4, pdrc_BicubicShrink = 5, pdrc_Mitchell = 6, pdrc_CatmullRom = 7, pdrc_Sinc = 8
#End If

'This dialog can be used to resize the full image, or a single layer.  The requested target will be stored here,
' and can be externally accessed by the ResizeTarget property.
Private m_ResizeTarget As PD_ActionTarget

'This dialog now supports live previews, which means we need a track a bunch of preview-related settings.

'Original, untouched source DIB (full image or layer, depending on dialog mode)
Private m_SrcComposite As pdDIB

'Cropped region of the source image pertaining to the current resize area.  May be larger or smaller
' than the preview DIB, depending on resize settings (e.g. up- or downsampling).
Private m_PreviewSrc As pdDIB

'Current preview, resized to the same size as the interactive preview box on the left and containing
' a preview of what a region of the resized image will look like.
Private m_PreviewDst As pdDIB

'Current x/y offset - IN SOURCE IMAGE/LAYER COORDINATES - of the preview box.  (Also, values necessary
' for tracking state changes to those values during mouse-drag.)
Private m_PreviewSrcX As Long, m_PreviewSrcY As Long
Private m_OrigPreviewSrcX As Long, m_OrigPreviewSrcY As Long
Private m_OrigMouseDownX As Long, m_OrigMouseDownY As Long

'Current source width/height of the image/layer (depends on m_ResizeTarget, above)
Private m_SrcImageWidth As Long, m_SrcImageHeight As Long

'Source and destination rectangle of the current preview, in their respective coordinate spaces.
Private m_DstRectF As RectF, m_SrcRectF As RectF

'Mouse is in the preview area
Private m_MouseInPreview As Boolean

'Forcibly suspend previews until the dialog is ready
Private m_PreviewsAllowed As Boolean

Public Property Let ResizeTarget(newTarget As PD_ActionTarget)
    m_ResizeTarget = newTarget
    If (m_ResizeTarget = pdat_Image) Then
        m_SrcImageWidth = PDImages.GetActiveImage.Width
        m_SrcImageHeight = PDImages.GetActiveImage.Height
        PDImages.GetActiveImage.GetCompositedImage m_SrcComposite, True
    Else
        m_SrcImageWidth = PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False)
        m_SrcImageHeight = PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)
        If PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True) Then
            PDImages.GetActiveImage.GetActiveLayer.GetAffineTransformedDIB m_SrcComposite, 0, 0
        Else
            Set m_SrcComposite = New pdDIB
            m_SrcComposite.CreateFromExistingDIB PDImages.GetActiveImage.GetActiveLayer.GetLayerDIB
        End If
    End If
End Property

Private Sub cboResample_Click()
    sldLanczos.Visible = (cboResample.ListIndex = rf_Lanczos)
    lblLanczos.Visible = (cboResample.ListIndex = rf_Lanczos)
    UpdatePreview
End Sub

Private Sub chkEstimate_Click()
    UpdatePreview
End Sub

Private Sub chkPreview_Click()
    UpdatePreview True
End Sub

Private Sub cmbFit_Click()
    UpdatePreview   'Is this necessary?
End Sub

'OK button
Private Sub cmdBar_OKClick()
    
    'The Undo method used varies if we are resizing the entire image (which requires undo data for all
    ' layers in the image) vs resizing a single layer.
    Select Case m_ResizeTarget
    
        Case pdat_Image
            Process "Resize image", , GetLocalParamString(), UNDO_Image
        
        Case pdat_SingleLayer
            Process "Resize layer", , GetLocalParamString(), UNDO_Layer
    
    End Select
    
End Sub

'I'm not sure that randomize serves any purpose on this dialog, but as I don't have a way to hide that button at
' present, simply randomize the width/height to +/- the current image's width/height divided by two.
Private Sub cmdBar_RandomizeClick()
    ucResize.AspectRatioLock = False
    ucResize.ResizeWidthInPixels = (PDImages.GetActiveImage.Width / 2) + (Rnd * PDImages.GetActiveImage.Width)
    ucResize.ResizeHeightInPixels = (PDImages.GetActiveImage.Height / 2) + (Rnd * PDImages.GetActiveImage.Height)
End Sub

Private Sub cmdBar_ResetClick()
    
    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.UnitOfMeasurement = mu_Pixels
    
    Select Case m_ResizeTarget
    
        Case pdat_Image
            ucResize.SetInitialDimensions PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, PDImages.GetActiveImage.GetDPI
            
        Case pdat_SingleLayer
            ucResize.SetInitialDimensions PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False), PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False), PDImages.GetActiveImage.GetDPI
        
    End Select
    
    ucResize.AspectRatioLock = True
    cboResample.ListIndex = 0
    
    'Stretch to new aspect ratio by default
    cmbFit.ListIndex = 0
    
    'It's possible that none of the above changes trigger a preview redraw, so request a manual one "just in case"
    UpdatePreview True
    
End Sub

Private Sub Form_Activate()

    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.UnitOfMeasurement = mu_Pixels
    
    'Set the dialog caption to match the current resize operation (resize image or resize single layer),
    ' and also set the width/height text boxes to match.
    If (m_ResizeTarget = pdat_Image) Then
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, g_Language.TranslateMessage("Resize image")
        ucResize.SetInitialDimensions PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, PDImages.GetActiveImage.GetDPI
    ElseIf (m_ResizeTarget = pdat_SingleLayer) Then
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, g_Language.TranslateMessage("Resize layer")
        ucResize.SetInitialDimensions PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False), PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False), PDImages.GetActiveImage.GetDPI
    End If
    
    ucResize.AspectRatioLock = True
    
    m_PreviewsAllowed = True
    UpdatePreview True
    
End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()
    
    'Populate the dropdowns with all available resampling algorithms.
    cboResample.SetAutomaticRedraws False
    cboResample.Clear
    
    Dim i As PD_ResamplingFilter
    For i = 0 To rf_Max - 1
        cboResample.AddItem Resampling.GetResamplerNameUI(i), i, (i = rf_Automatic) Or (i = rf_Box) Or (i = rf_Hermite) Or (i = rf_QuadraticBSpline) Or (i = rf_Mitchell)
    Next i
    
    cboResample.ListIndex = 0
    cboResample.SetAutomaticRedraws True, True
    
    'Populate the "fit" options
    cmbFit.SetAutomaticRedraws False
    cmbFit.Clear
    cmbFit.AddItem "stretching to new size  (default)", 0
    cmbFit.AddItem "fitting inclusively, with transparent borders as necessary", 1
    cmbFit.AddItem "fitting exclusively, and cropping as necessary", 2
    cmbFit.ListIndex = 0
    cmbFit.SetAutomaticRedraws True, True
    
    'Automatically set the width and height text boxes to match the image's current dimensions.  (Note that we must
    ' do this again in the Activate step, as the last-used settings will automatically override these values.  However,
    ' if we do not also provide these values here, the resize control may attempt to set parameters while having
    ' a width/height/resolution of 0, which will cause divide-by-zero errors.)
    Select Case m_ResizeTarget
    
        Case pdat_Image
            ucResize.SetInitialDimensions PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, PDImages.GetActiveImage.GetDPI
            
        Case pdat_SingleLayer
            ucResize.SetInitialDimensions PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False), PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False), PDImages.GetActiveImage.GetDPI
        
    End Select
    
    'Add some tooltips
    chkEstimate.AssignTooltip "Some image resampling filters support approximate calculations.  These improve performance at a minor penalty to quality."
    cboResample.AssignTooltip "Resampling affects the quality of a resized image.  For a good summary of resampling techniques, visit the Image Resampling article on Wikipedia."
    cmbFit.AssignTooltip "When changing an image's aspect ratio, undesirable stretching may occur.  PhotoDemon can avoid this by using empty borders or cropping instead."
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True, picPreview.hWnd
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Resize an image using our own internal algorithms.  Slower, but better quality.
Private Function InternalImageResize(ByRef dstDIB As pdDIB, ByRef srcDIB As pdDIB, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal interpolationMethod As PD_ResamplingFilter, ByVal allowApproximation As Boolean, ByVal displayProgress As Boolean) As Boolean
    
    PDDebug.LogAction "Using internal resampler for this operation."
    
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    'Unpremultiply alpha prior to resampling
    If srcDIB.GetAlphaPremultiplication() Then srcDIB.SetAlphaPremultiplication False
    
    'Resize the destination DIB in preparation for the resize
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    If (dstDIB.GetDIBWidth <> dstWidth) Or (dstDIB.GetDIBHeight <> dstHeight) Then
        dstDIB.CreateBlank dstWidth, dstHeight, 32, 0, 0
    Else
        dstDIB.ResetDIB 0
    End If
    
    'Hand off the image to PD's internal resampler, and if approximation is allowed, use the integer-based transform
    ' for a nice performance boost.
    If allowApproximation Then
        InternalImageResize = Resampling.ResampleImageI(dstDIB, srcDIB, dstWidth, dstHeight, interpolationMethod, displayProgress)
    Else
        InternalImageResize = Resampling.ResampleImage(dstDIB, srcDIB, dstWidth, dstHeight, interpolationMethod, displayProgress)
    End If
    
    'Premultiply the resulting image
    dstDIB.SetAlphaPremultiplication True, True
    
End Function

'Resize an image using any one of several resampling algorithms.  (Some algorithms are provided by FreeImage.)
Public Sub ResizeImage(ByVal resizeParams As String)
        
    'Parse incoming parameters into type-appropriate vars
    Dim imgWidth As Double, imgHeight As Double, imgDPI As Double
    Dim resampleMethod As PD_ResamplingFilter, allowApproximation As Boolean, lanczosLobes As Long
    Dim fitMethod As PD_ResizeFit, imgResizeUnit As PD_MeasurementUnit, thingToResize As PD_ActionTarget
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString resizeParams
    
    With cParams
        imgWidth = .GetDouble("width")
        imgHeight = .GetDouble("height")
        imgResizeUnit = .GetLong("unit", mu_Pixels)
        imgDPI = .GetDouble("ppi", 96)
        fitMethod = .GetLong("fit", ResizeFitStretch)
        thingToResize = .GetLong("target", pdat_Image)
    End With
    
    'Use a separate function to retrieve resampling method and approximation permission
    ' (legacy items may need to be dealt with)
    resampleMethod = GetResampleMethod(cParams, allowApproximation, lanczosLobes)
    Resampling.SetLanczosRadius lanczosLobes
    
    'Depending on the requested fitting technique, we may have to resize the image to a slightly different size
    ' than the one requested.  Before doing anything else, calculate that new size.
    Dim srcWidth As Long, srcHeight As Long
    
    Select Case thingToResize
        Case pdat_Image
            srcWidth = PDImages.GetActiveImage.Width
            srcHeight = PDImages.GetActiveImage.Height
        Case pdat_SingleLayer
            srcWidth = PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False)
            srcHeight = PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)
    End Select
    
    'In past versions of the software, we could assume the passed measurements were always in pixels,
    ' but that is no longer the case!  Using the supplied "unit of measurement", convert the passed
    ' width and height values to pixel measurements.
    imgWidth = ConvertOtherUnitToPixels(imgResizeUnit, imgWidth, imgDPI, srcWidth)
    imgHeight = ConvertOtherUnitToPixels(imgResizeUnit, imgHeight, imgDPI, srcHeight)
    
    'Finally, use the above values to determine a "fit" width/height (necessary if stretch-to-fit
    ' is *not* enabled).
    Dim fitWidth As Long, fitHeight As Long
    Select Case fitMethod
    
        'Stretch-to-fit.  Default behavior, and no size changes are required.
        Case ResizeFitStretch
            fitWidth = imgWidth
            fitHeight = imgHeight
        
        'Fit inclusively.  Fit the image's largest dimension.  No cropping will occur, but blank space may be present.
        Case ResizeFitInclusive
            PDMath.ConvertAspectRatio srcWidth, srcHeight, imgWidth, imgHeight, fitWidth, fitHeight
            
        'Fit exclusively.  Fit the image's smallest dimension.  Cropping will occur, but no blank space will be present.
        Case ResizeFitExclusive
            PDMath.ConvertAspectRatio srcWidth, srcHeight, imgWidth, imgHeight, fitWidth, fitHeight, False
        
    End Select
    
    'If the image contains an active selection, automatically deactivate it.
    ' (In the future, perhaps we could resize it, depending on the selection type.)
    If PDImages.GetActiveImage.IsSelectionActive And (thingToResize = pdat_Image) Then
        PDImages.GetActiveImage.SetSelectionActive False
        PDImages.GetActiveImage.MainSelection.LockRelease
    End If

    'Because most resize methods require a temporary DIB, create one here
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'If the user has requested "automatic" resample mode, convert that to an actual resample mode.
    ' (In the future, it would be nice to develop a smarter heuristic for this step - but for now,
    '  use bicubic resampling, and auto-select the shrink-optimized variant based on the
    '  dimensions used for the resize.)
    If (resampleMethod = rf_Automatic) Then resampleMethod = ConvertAutoResample(srcWidth, srcHeight, fitWidth, fitHeight, allowApproximation)
    
    'If we use an outside library for the resize (e.g. FreeImage, GDI+), we won't receive progress reports
    ' during the resize.  For full-image resizing this may not be a problem (as we can use layer count as
    ' a surrogate), but for a single-layer resize, we don't have a good fallback.
    If (thingToResize = pdat_Image) Then
        ProgressBars.SetProgBarMax PDImages.GetActiveImage.GetNumOfLayers
        Message "Resizing image..."
    Else
        ProgressBars.SetProgBarMax 1
        Message "Resizing layer..."
    End If
        
    Dim srcAspect As Double, dstAspect As Double
    Dim dstX As Long, dstY As Long
    
    'It is now time to iterate through all layers, resizing as we go.  Note that PD's approach to multi-layer
    ' operations allows us to use the same resize code for each layer, because layers smaller than the image
    ' will be automatically padded to the image's full size.
    Dim tmpLayerRef As pdLayer
    
    'If we are resizing the entire image, we must handle all layers in turn.  Otherwise, we can handle just
    ' the active layer.  Set loop boundaries accordingly.
    Dim firstLayerIndex As Long, lastLayerIndex As Long
    
    Select Case thingToResize
        Case pdat_Image
            firstLayerIndex = 0
            lastLayerIndex = PDImages.GetActiveImage.GetNumOfLayers - 1
        Case pdat_SingleLayer
            firstLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex
            lastLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex
    End Select
    
    Dim i As Long
    For i = firstLayerIndex To lastLayerIndex
        
        'When resizing the full image, report progress on a layer-by-layer basis
        If (thingToResize = pdat_Image) Then ProgressBars.SetProgBarVal i
        
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(i)
        
        'Null-pad the layer
        If (thingToResize = pdat_Image) Then tmpLayerRef.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, False
        
        'Call the appropriate external function, based on the user's resize selection.  Each function will
        ' place a resized version of tmpLayerRef.GetLayerDIB into tmpDIB.
        
        Select Case resampleMethod
            
            'For nearest-neighbor scaling, GDI can be used for great performance
            Case rf_Box
            
                'Copy the current DIB into this temporary DIB at the new size.  (StretchBlt is used
                ' for a fast resize.)
                If ALLOW_GDIPLUS_RESIZE And allowApproximation Then
                    tmpDIB.CreateFromExistingDIB tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, GP_IM_NearestNeighbor
                
                'For a slower, but more mathematically accurate approach, you could also use our internal scaler
                Else
                    InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_Box, allowApproximation, (firstLayerIndex = lastLayerIndex)
                End If
                
            'Bilinear and bicubic resampling can use GDI+ (preferentially), our internal resampler, or the FreeImage library
            Case rf_BilinearTriangle
                If ALLOW_GDIPLUS_RESIZE And allowApproximation Then
                    If (tmpDIB.GetDIBWidth <> fitWidth) Or (tmpDIB.GetDIBHeight <> fitHeight) Then tmpDIB.CreateBlank fitWidth, fitHeight, 32, 0 Else tmpDIB.ResetDIB 0
                    GDIPlusResizeDIB tmpDIB, 0, 0, fitWidth, fitHeight, tmpLayerRef.GetLayerDIB, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), GP_IM_HighQualityBilinear, P2_PO_Half
                Else
                    InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_BilinearTriangle, allowApproximation, (firstLayerIndex = lastLayerIndex)
                End If
            
            Case rf_CubicBSpline
                If ALLOW_GDIPLUS_RESIZE And allowApproximation Then
                    If (tmpDIB.GetDIBWidth <> fitWidth) Or (tmpDIB.GetDIBHeight <> fitHeight) Then tmpDIB.CreateBlank fitWidth, fitHeight, 32, 0 Else tmpDIB.ResetDIB 0
                    GDIPlusResizeDIB tmpDIB, 0, 0, fitWidth, fitHeight, tmpLayerRef.GetLayerDIB, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), GP_IM_HighQualityBicubic, P2_PO_Half
                Else
                    InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_CubicBSpline, allowApproximation, (firstLayerIndex = lastLayerIndex)
                End If
            
            'All remaining methods rely on our own internal resampling engine
            Case rf_Mitchell
                InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_Mitchell, allowApproximation, (firstLayerIndex = lastLayerIndex)
                
            Case rf_CatmullRom
                InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_CatmullRom, allowApproximation, (firstLayerIndex = lastLayerIndex)
                
            Case rf_Lanczos
                InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_Lanczos, allowApproximation, (firstLayerIndex = lastLayerIndex)
                
            Case rf_Cosine
                InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_Cosine, allowApproximation, (firstLayerIndex = lastLayerIndex)
                
            Case rf_Hermite
                InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_Hermite, allowApproximation, (firstLayerIndex = lastLayerIndex)
            
            Case rf_Bell
                InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_Bell, allowApproximation, (firstLayerIndex = lastLayerIndex)
            
            Case rf_Quadratic
                InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_Quadratic, allowApproximation, (firstLayerIndex = lastLayerIndex)
            
            Case rf_QuadraticBSpline
                InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_QuadraticBSpline, allowApproximation, (firstLayerIndex = lastLayerIndex)
                
            Case rf_CubicConvolution
                InternalImageResize tmpDIB, tmpLayerRef.GetLayerDIB, fitWidth, fitHeight, rf_CubicConvolution, allowApproximation, (firstLayerIndex = lastLayerIndex)
                
            'This failsafe should never be triggered
            Case Else
                PDDebug.LogAction "WARNING: FormResize.ResizeImage encountered an unknown resize filter: " & resampleMethod
                If (tmpDIB.GetDIBWidth <> fitWidth) Or (tmpDIB.GetDIBHeight <> fitHeight) Then tmpDIB.CreateBlank fitWidth, fitHeight, 32, 0 Else tmpDIB.ResetDIB 0
                GDIPlusResizeDIB tmpDIB, 0, 0, fitWidth, fitHeight, tmpLayerRef.GetLayerDIB, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), GP_IM_HighQualityBicubic
            
        End Select
        
        'tmpDIB now holds a copy of the resized image.
        
        'Calculate the aspect ratio of the temporary DIB and the target size.  If the user has selected
        ' a resize mode other than "fit exactly", we still need to do a bit of extra trimming.
        srcAspect = srcWidth / srcHeight
        dstAspect = imgWidth / imgHeight
        
        'We now want to copy the resized image into the current image using the technique requested by the user.
        Select Case fitMethod
        
            'Stretch-to-fit.  This is default resize behavior in all image editing software
            Case ResizeFitStretch
        
                'Very simple - just copy the resized image back into the main DIB
                tmpLayerRef.GetLayerDIB.CreateFromExistingDIB tmpDIB
        
            'Fit inclusively.  This fits the image's largest dimension into the destination image, which can leave
            ' blank space - that space is filled by the background color parameter passed in (or transparency,
            ' in the case of 32bpp images).
            Case ResizeFitInclusive
            
                'Resize the main DIB (destructively!) to fit the new dimensions
                tmpLayerRef.GetLayerDIB.CreateBlank imgWidth, imgHeight, 32, 0
                
                'BitBlt the old image, centered, onto the new DIB
                If (srcAspect > dstAspect) Then
                    dstY = Int((imgHeight - fitHeight) / 2# + 0.5)
                    dstX = 0
                Else
                    dstX = Int((imgWidth - fitWidth) / 2# + 0.5)
                    dstY = 0
                End If
                
                GDI.BitBltWrapper tmpLayerRef.GetLayerDIB.GetDIBDC, dstX, dstY, fitWidth, fitHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
                tmpLayerRef.GetLayerDIB.SetInitialAlphaPremultiplicationState tmpDIB.GetAlphaPremultiplication
            
            'Fit exclusively.  This fits the image's smallest dimension into the destination image, which means no
            ' blank space - but parts of the image may get cropped out.
            Case ResizeFitExclusive
            
                'Resize the main DIB (destructively!) to fit the new dimensions
                tmpLayerRef.GetLayerDIB.CreateBlank imgWidth, imgHeight, 32, 0
            
                'BitBlt the old image, centered, onto the new DIB
                If (srcAspect < dstAspect) Then
                    dstY = Int((imgHeight - fitHeight) / 2# + 0.5)
                    dstX = 0
                Else
                    dstX = Int((imgWidth - fitWidth) / 2# + 0.5)
                    dstY = 0
                End If
                
                GDI.BitBltWrapper tmpLayerRef.GetLayerDIB.GetDIBDC, dstX, dstY, fitWidth, fitHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
                tmpLayerRef.GetLayerDIB.SetInitialAlphaPremultiplicationState tmpDIB.GetAlphaPremultiplication
                
        End Select
        
        'With the layer now successfully resized, we can remove any null-padding that may still exist.
        ' (Note that we skip this step when resizing a single layer only.)
        If (thingToResize = pdat_Image) Then tmpLayerRef.CropNullPaddedLayer
        
        'Notify the parent image of the change
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
        
    'Move on to the next layer
    Next i
    
    'We are finished with the temporary DIB, so release it, along with any other temporary buffers we allocated
    Set tmpDIB = Nothing
    Resampling.FreeBuffers
    
    'Update the main image's size and DPI values
    If (thingToResize = pdat_Image) Then
        PDImages.GetActiveImage.UpdateSize False, imgWidth, imgHeight
        PDImages.GetActiveImage.SetDPI imgDPI, imgDPI
        Interface.DisplaySize PDImages.GetActiveImage()
        Tools.NotifyImageSizeChanged
    End If
        
    'Fit the new image on-screen and redraw its viewport
    Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    'Release the progress bar
    ProgressBars.SetProgBarVal 0
    ReleaseProgressBar
    
    Message "Finished."
    
End Sub

Private Function GetLocalParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    With cParams
    
        'In July 2021, the parameters for this tool were modified.  String values are now used
        ' to define resampling algorithms, which should successfully future-proof this tool
        ' (finally lol)
        .SetParamVersion 3#
        .AddParam "width", ucResize.ResizeWidth
        .AddParam "height", ucResize.ResizeHeight
        .AddParam "unit", ucResize.UnitOfMeasurement
        .AddParam "ppi", ucResize.ResizeDPIAsPPI
        .AddParam "resample", Resampling.GetResamplerName(cboResample.ListIndex)
        .AddParam "approximations-ok", chkEstimate.Value
        .AddParam "lanczos-lobes", sldLanczos.Value
        .AddParam "fit", cmbFit.ListIndex
        .AddParam "target", m_ResizeTarget
        
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

'This separate function is necessary for retrieving the preferred resampling method from
' PD's XML-based serializer (which stores macro commands, among other details).  Old versions
' of the serializer used 0-based integers to define resampling methods; we now use string IDs
' as they're much easier to make backward- and forward-compatible against new features.
Private Function GetResampleMethod(ByRef cParams As pdSerialize, ByRef allowApproximation As Boolean, ByRef lanczosLobes As Long) As PD_ResamplingFilter

    'Current parameter strings (v3) use string IDs
    If (cParams.GetParamVersion() >= 3#) Then
        GetResampleMethod = Resampling.GetResamplerID(cParams.GetString("resample", "auto", True))
        allowApproximation = cParams.GetBool("approximations-ok", True, True)
        lanczosLobes = cParams.GetLong("lanczos-lobes", 3, True)
        
    'Legacy implementations follow
    Else
        
        'In legacy implementations, the only available Lanczos lobe count was 3
        lanczosLobes = 3
        
        'In July 2018, the resample options for this tool were updated to allow for downsample-optimized
        ' filters (provided by GDI+).
        If (cParams.GetParamVersion() = 2#) Then
            
            Dim rsMethodV2 As PD_ResampleOld_V2
            rsMethodV2 = cParams.GetLong("algorithm", pdrc_Automatic)
            
            'Translate the old enum into a modern one
            Select Case rsMethodV2
                Case pdrc_Automatic
                    GetResampleMethod = rf_Automatic
                    allowApproximation = True
                Case pdrc_NearestNeighbor
                    GetResampleMethod = rf_Box
                    allowApproximation = True
                Case pdrc_BilinearNormal
                    GetResampleMethod = rf_BilinearTriangle
                    allowApproximation = False
                Case pdrc_BilinearShrink
                    GetResampleMethod = rf_BilinearTriangle
                    allowApproximation = True
                Case pdrc_BicubicNormal
                    GetResampleMethod = rf_CubicBSpline
                    allowApproximation = False
                Case pdrc_BicubicShrink
                    GetResampleMethod = rf_CubicBSpline
                    allowApproximation = True
                Case pdrc_Mitchell
                    GetResampleMethod = rf_Mitchell
                    allowApproximation = True
                Case pdrc_CatmullRom
                    GetResampleMethod = rf_CatmullRom
                    allowApproximation = True
                Case pdrc_Sinc
                    GetResampleMethod = rf_Lanczos
                    allowApproximation = True
            End Select
            
        'v1 parameters use a different enum
        Else
            
            Dim rsMethodV1 As PD_ResampleOld_V1
            rsMethodV1 = cParams.GetLong("algorithm", ResizeNormal)
            
            Select Case rsMethodV1
                Case ResizeNormal
                    GetResampleMethod = rf_Box
                    allowApproximation = True
                Case ResizeBilinear
                    GetResampleMethod = rf_BilinearTriangle
                    allowApproximation = False
                Case ResizeBspline
                    GetResampleMethod = rf_CubicBSpline
                    allowApproximation = False
                Case ResizeBicubicMitchell
                    GetResampleMethod = rf_Mitchell
                    allowApproximation = True
                Case ResizeBicubicCatmull
                    GetResampleMethod = rf_CatmullRom
                    allowApproximation = True
                Case ResizeSincLanczos
                    GetResampleMethod = rf_Lanczos
                    allowApproximation = True
                Case Else
                    GetResampleMethod = rf_Automatic
                    allowApproximation = True
            End Select
            
        End If
    
    End If

End Function

'Rendering the preview box is easy - just paint the current preview cache to the target DC!
Private Sub picPreview_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    
    'The user can disable live previews.  This is helpful on older/performance-constrained PCs.
    If (Not m_PreviewDst Is Nothing) Then
    
        'Start by painting a checkerboard background on the destination
        Dim xOffset As Long, yOffset As Long
        xOffset = (picPreview.GetWidth - m_PreviewDst.GetDIBWidth) \ 2
        yOffset = (picPreview.GetHeight - m_PreviewDst.GetDIBHeight) \ 2
        GDI_Plus.GDIPlusFillDIBRect_Pattern Nothing, xOffset, yOffset, m_PreviewDst.GetDIBWidth, m_PreviewDst.GetDIBHeight, g_CheckerboardPattern, targetDC, True, True
        
        'Then paint the preview atop it
        m_PreviewDst.AlphaBlendToDC targetDC, dstX:=xOffset, dstY:=yOffset
        
    End If
    
    'Always finish by drawing a border around the control
    Dim borderColor As Long
    If m_MouseInPreview And chkPreview.Value Then
        borderColor = g_Themer.GetGenericUIColor(UI_Accent, True, False, True)
    Else
        borderColor = g_Themer.GetGenericUIColor(UI_GrayDefault)
    End If
    
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundDC targetDC
    cSurface.SetSurfaceAntialiasing P2_AA_None
    
    Dim cPen As pd2DPen
    Set cPen = New pd2DPen
    cPen.SetPenWidth 1!
    cPen.SetPenColor borderColor
    cPen.SetPenLineJoin P2_LJ_Miter
    
    PD2D.DrawRectangleI cSurface, cPen, 0, 0, picPreview.GetWidth - 1, picPreview.GetHeight - 1
    
    Set cSurface = Nothing: Set cPen = Nothing
End Sub

Private Sub picPreview_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    If (Button And pdLeftButton) Then
        m_OrigPreviewSrcX = m_PreviewSrcX
        m_OrigPreviewSrcY = m_PreviewSrcY
        m_OrigMouseDownX = x
        m_OrigMouseDownY = y
    End If
End Sub

Private Sub picPreview_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInPreview = True
    picPreview.RequestCursor IDC_HAND
    picPreview.RequestRedraw True
End Sub

Private Sub picPreview_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInPreview = False
    picPreview.RequestRedraw True
End Sub

Private Sub picPreview_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    If (Button And pdLeftButton) Then
        
        Dim xScale As Double, yScale As Double
        If (m_SrcImageWidth > 0#) Then
            xScale = m_SrcImageWidth / ucResize.ResizeWidthInPixels
        Else
            xScale = 1#
        End If
        If (m_SrcImageHeight > 0#) Then
            yScale = m_SrcImageHeight / ucResize.ResizeHeightInPixels
        Else
            yScale = 1#
        End If
        
        m_PreviewSrcX = m_OrigPreviewSrcX - (x - m_OrigMouseDownX) * xScale
        m_PreviewSrcY = m_OrigPreviewSrcY - (y - m_OrigMouseDownY) * yScale
        
        'Fix boundary conditions
        If (m_PreviewSrcX < 0) Then m_PreviewSrcX = 0
        If (m_PreviewSrcY < 0) Then m_PreviewSrcY = 0
        If (m_PreviewSrcX + (picPreview.GetWidth * xScale) > m_SrcImageWidth) Then m_PreviewSrcX = m_SrcImageWidth - (picPreview.GetWidth * xScale)
        If (m_PreviewSrcY + (picPreview.GetHeight * yScale) > m_SrcImageHeight) Then m_PreviewSrcY = m_SrcImageHeight - (picPreview.GetHeight * yScale)
        
        UpdatePreview True
        
    End If
End Sub

Private Sub picPreview_WindowResizeDetected()
    UpdatePreview
End Sub

Private Sub sldLanczos_Change()
    UpdatePreview
End Sub

Private Sub ucResize_Change(ByVal newWidthPixels As Double, ByVal newHeightPixels As Double)
    UpdatePreview
End Sub

'If the user has requested "automatic" resample mode, convert that to an actual resample mode.
' (In the future, it would be nice to develop a smarter heuristic for this step - but for now,
'  use standard resampling rules from Photoshop.)
Private Function ConvertAutoResample(ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal allowApproximation As Boolean) As PD_ResamplingFilter

    If (dstWidth < srcWidth) Or (dstHeight < srcHeight) Then
        If allowApproximation Then
            ConvertAutoResample = rf_CubicBSpline
        Else
            ConvertAutoResample = rf_Lanczos
        End If
    Else
        ConvertAutoResample = rf_Mitchell
    End If
    
End Function

Private Sub UpdatePreview(Optional ByVal forcePreviewNow As Boolean = False)
    
    'Failsafe checks
    If (m_SrcComposite Is Nothing) Then Exit Sub
    If ((Not cmdBar.PreviewsAllowed) Or (Not m_PreviewsAllowed)) And (Not forcePreviewNow) Then Exit Sub
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
        
    'The user can disable live previews.  This is helpful on older PCs.
    If chkPreview.Value Then
        
        'Cache the current source and destination rectangles.
        ' (If they haven't changed since the last preview, we don't need to regenerate them.)
        Dim lastSrcRectF As RectF
        lastSrcRectF = m_SrcRectF
        
        'Next, we need to calculate the rectangle in the source image we need to "snip" out a preview source.
        
        'Calculate the the source image's width/height.  If the source image is small (or the destination image
        ' is small due to current resizing), we'll need to position the preview image accordingly.
        Dim srcWidth As Long, srcHeight As Long
        srcWidth = m_SrcComposite.GetDIBWidth
        srcHeight = m_SrcComposite.GetDIBHeight
        
        'Next, figure out an x/y scale of the current image vs the resized coordinates
        Dim xScale As Double, yScale As Double
        xScale = ucResize.ResizeWidthInPixels / srcWidth
        yScale = ucResize.ResizeHeightInPixels / srcHeight
        
        'Failsafe check
        If (xScale <= 0#) Or (yScale <= 0#) Then
            Debug.Print "bad scale; can't preview"
            Exit Sub
        End If
        
        'Working backward from the destination image (whose size is fixed, since it's tied to the preview box),
        ' calculate a corresponding rectangle in the source image that matches the current scale factor.
        
        'Start by filling the destination rect with dummy values
        With m_DstRectF
            .Left = 0
            .Top = 0
            .Width = picPreview.GetWidth
            .Height = picPreview.GetHeight
        End With
        
        'Now, work backward - using the current scale factor - to find a corresponding region in the source image.
        With m_SrcRectF
            .Left = m_PreviewSrcX
            .Top = m_PreviewSrcY
            .Width = m_DstRectF.Width / xScale
            .Height = m_DstRectF.Height / yScale
        End With
        
        'We now have a corresponding source rectangle.  We now need to modify the source rectangle to account
        ' for things like "larger than the source image" and other boundary conditions.
        With m_SrcRectF
            
            'Check OOB on right edge
            If ((.Left + .Width) > srcWidth) Then
                
                'The current rectangle lies outside the image.  Attempt to bring it in-bounds by
                ' moving the rectangle left.
                .Left = srcWidth - .Width
                
            End If
            
            'If moving the rectangle left pushes it out of bounds, we have no choice but to shrink the
            ' width entirely, to the smallest functional width we can afford.
            If (.Left < 0) Then
                .Left = 0
                .Width = srcWidth
                m_PreviewSrcX = 0
            End If
            
            'Repeat the above steps for the bottom boundary
            If ((.Top + .Height) > srcHeight) Then
                
                'The current rectangle lies outside the image.  Attempt to bring it in-bounds by
                ' moving the rectangle up.
                .Top = srcHeight - .Height
                
            End If
            
            'If moving the rectangle up pushes it out of bounds, we have no choice but to shrink the
            ' height entirely, to the smallest functional height we can afford.
            If (.Top < 0) Then
                .Top = 0
                .Height = srcHeight
                m_PreviewSrcY = 0
            End If
            
        End With
        
        'Ensure at least one pixel is being used from the source!  (This is necessary to ensure a preview
        ' if resizing from e.g. 1x1 to 5000x5000.)
        If (m_SrcRectF.Width < 1!) Then m_SrcRectF.Width = 1!
        If (m_SrcRectF.Height < 1!) Then m_SrcRectF.Height = 1!
        
        'Left and top boundaries for the source rectangle are now properly cropped, as are width/height.
        ' This means we have the source rectangle we want to use.
        
        'We now need to reverse-calculate a corresponding destination width/height for this rectangle.
        With m_DstRectF
            .Width = m_SrcRectF.Width * xScale
            .Height = m_SrcRectF.Height * yScale
            
            'Left and Top will need to be dealt with in the future; for now, zero them for testing
            .Left = 0
            .Top = 0
            
        End With
        
        'Crop out the source region into a standalone DIB
        If (m_PreviewSrc Is Nothing) Then Set m_PreviewSrc = New pdDIB
        If (Not PDMath.AreRectFsEqual(lastSrcRectF, m_SrcRectF)) Then
            m_PreviewSrc.CreateBlank m_SrcRectF.Width, m_SrcRectF.Height, 32, 0, 0
            GDI.BitBltWrapper m_PreviewSrc.GetDIBDC, 0, 0, m_SrcRectF.Width, m_SrcRectF.Height, m_SrcComposite.GetDIBDC, m_SrcRectF.Left, m_SrcRectF.Top, vbSrcCopy
            m_PreviewSrc.SetInitialAlphaPremultiplicationState True
        End If
        
        'Failsafe check during initialization
        If (m_PreviewSrc.GetDIBWidth = 0) Or (m_PreviewSrc.GetDIBDC = 0) Then Exit Sub
        
        If REPORT_PREVIEW_PERF Then PDDebug.LogAction "Prep: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
        
        'Create the destination DIB at the required size
        If (m_PreviewDst Is Nothing) Then Set m_PreviewDst = New pdDIB
        m_PreviewDst.CreateBlank m_DstRectF.Width, m_DstRectF.Height, 32, 0, 0
        
        If REPORT_PREVIEW_PERF Then PDDebug.LogAction "Dest: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
        
        'If the x/y scales are 1.0, it means the image is not being resized.  Simply copy it as-is into place.
        If (xScale = 1#) And (yScale = 1#) Then
            GDI.BitBltWrapper m_PreviewDst.GetDIBDC, 0, 0, m_DstRectF.Width, m_DstRectF.Height, m_PreviewSrc.GetDIBDC, 0, 0, vbSrcCopy
            m_PreviewDst.SetInitialAlphaPremultiplicationState True
            
        'Ifi the x/y scales are *not* 1.0, it means the image is being resized.  We need to generate a preview.
        Else
            
            'Determine which resampling strategy to preview
            Dim resampleMethod As PD_ResamplingFilter
            resampleMethod = cboResample.ListIndex
            If (resampleMethod = rf_Automatic) Then resampleMethod = ConvertAutoResample(m_SrcRectF.Width, m_SrcRectF.Height, m_DstRectF.Width, m_DstRectF.Height, chkEstimate.Value)
            
            Dim allowApproximation As Boolean
            allowApproximation = chkEstimate.Value
            
            'Perform the resize.  Note that - as best we can - we try to mimic different resize algorithms
            ' as they'll appear in the "official" resize, but some engines - like FreeImage are ignored in
            ' favor of our internal functions during a preview.
            Resampling.SetLanczosRadius sldLanczos.Value
            
            Select Case resampleMethod
                Case rf_Box
                    If allowApproximation Then
                        If ALLOW_GDIPLUS_RESIZE Then
                            m_PreviewDst.CreateFromExistingDIB m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, GP_IM_NearestNeighbor
                        Else
                            Resampling.ResampleImageI m_PreviewDst, m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, resampleMethod, False
                        End If
                    Else
                        'If GDI+ is active, it will handle "fast" mode.  For "pure" mode, let's use our
                        ' faster-but-just-about-identical integer mode if we're competing with GDI+.
                        If ALLOW_GDIPLUS_RESIZE Then
                            Resampling.ResampleImageI m_PreviewDst, m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, resampleMethod, False
                        Else
                            Resampling.ResampleImage m_PreviewDst, m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, resampleMethod, False
                        End If
                    End If
                Case rf_BilinearTriangle
                    If allowApproximation Then
                        If ALLOW_GDIPLUS_RESIZE Then
                            GDI_Plus.GDIPlusResizeDIB m_PreviewDst, 0, 0, m_DstRectF.Width, m_DstRectF.Height, m_PreviewSrc, 0, 0, m_SrcRectF.Width, m_SrcRectF.Height, GP_IM_HighQualityBilinear, P2_PO_Half
                        Else
                            Resampling.ResampleImageI m_PreviewDst, m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, resampleMethod, False
                        End If
                    Else
                        If ALLOW_GDIPLUS_RESIZE Then
                            Resampling.ResampleImageI m_PreviewDst, m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, resampleMethod, False
                        Else
                            Resampling.ResampleImage m_PreviewDst, m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, resampleMethod, False
                        End If
                    End If
                Case rf_CubicBSpline
                    If allowApproximation Then
                        If ALLOW_GDIPLUS_RESIZE Then
                            GDI_Plus.GDIPlusResizeDIB m_PreviewDst, 0, 0, m_DstRectF.Width, m_DstRectF.Height, m_PreviewSrc, 0, 0, m_SrcRectF.Width, m_SrcRectF.Height, GP_IM_HighQualityBicubic, P2_PO_Half
                        Else
                            Resampling.ResampleImageI m_PreviewDst, m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, resampleMethod, False
                        End If
                    Else
                        If ALLOW_GDIPLUS_RESIZE Then
                            Resampling.ResampleImageI m_PreviewDst, m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, resampleMethod, False
                        Else
                            Resampling.ResampleImage m_PreviewDst, m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, resampleMethod, False
                        End If
                    End If
                
                'Anything else uses our internal resampler for preview
                Case Else
                    If allowApproximation Then
                        Resampling.ResampleImageI m_PreviewDst, m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, resampleMethod, False
                    Else
                        Resampling.ResampleImage m_PreviewDst, m_PreviewSrc, m_DstRectF.Width, m_DstRectF.Height, resampleMethod, False
                    End If
                
            End Select
            
            'Ensure correct premultiplication by un-premultiplying, then re-premultiplying alpha.
            ' (This is necessary because some operators with sharpening tendencies - like Lanczos -
            ' will weight color values from transparent pixels as part of the algorithm, so we want
            ' to resample in the premultiplied color space to minimize absorption of color from
            ' transparent regions.)
            m_PreviewDst.SetAlphaPremultiplication False, True
            m_PreviewDst.SetAlphaPremultiplication True
            
        End If
        
    '/end "show previews" checkbox
    End If
    
    'Request a redraw
    picPreview.RequestRedraw forcePreviewNow
        
    If REPORT_PREVIEW_PERF Then PDDebug.LogAction "Final: " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub
