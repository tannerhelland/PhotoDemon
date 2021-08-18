VERSION 5.00
Begin VB.Form FormResize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resize image"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9630
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
   ScaleHeight     =   495
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCheckBox chkEstimate 
      Height          =   375
      Left            =   825
      TabIndex        =   5
      Top             =   4440
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      Caption         =   "use approximations to improve performance"
   End
   Begin PhotoDemon.pdDropDown cmbFit 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   5520
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   635
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6675
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdResize ucResize 
      Height          =   2850
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5027
   End
   Begin PhotoDemon.pdColorSelector csBackground 
      Height          =   495
      Left            =   840
      TabIndex        =   4
      ToolTipText     =   "Click to change the color used for empty borders"
      Top             =   6000
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   873
   End
   Begin PhotoDemon.pdDropDown cboResample 
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   3960
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   635
   End
   Begin PhotoDemon.pdLabel lblFit 
      Height          =   315
      Left            =   480
      Top             =   5040
      Width           =   8685
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "when changing aspect ratio, fit image to new size by"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblResample 
      Height          =   315
      Left            =   480
      Top             =   3480
      Width           =   8820
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "resize quality"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Size Handler
'Copyright 2001-2021 by Tanner Helland
'Created: 12/December/01
'Last updated: 16/August/21
'Last update: attempt a new custom-built resize engine, specific to PhotoDemon
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

'Internal flag to use/not use GDI+ resize functions.  GDI+ is much faster than our internal PD resampler,
' but its algorithms are nonstandard and have some weird quirks vs theoretically "perfect" implementations
' (see https://photosauce.net/blog/post/image-scaling-with-gdi-part-4-examining-the-interpolationmode-values).
' This value should be set to TRUE in production code.
Private Const USE_GDIPLUS_RESIZE As Boolean = True

'Internal flag to use/not use the 3rd-party FreeImage library's resize functions (if available).
' FreeImage's functions are faster than our internal PD resampler, but they require a lot of extra memory due
' to FreeImage needing its own container for pixel data.  (If FreeImage is *not* available, this value is ignored.)
' This value should be set to TRUE in production code.
Private Const USE_FREEIMAGE_RESIZE As Boolean = True

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

Public Property Let ResizeTarget(newTarget As PD_ActionTarget)
    m_ResizeTarget = newTarget
End Property

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
    
    'Make borders fill with black by default
    csBackground.Color = RGB(0, 0, 0)
    
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

End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()
    
    'Populate the dropdowns with all available resampling algorithms.
    cboResample.SetAutomaticRedraws False
    cboResample.Clear
    
    Dim i As PD_ResamplingFilter
    For i = 0 To rf_Max - 1
        cboResample.AddItem Resampling.GetResamplerNameUI(i)
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
    ApplyThemeAndTranslations Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Resize an image using the FreeImage library.  Very fast.
Private Sub FreeImageResize(ByRef dstDIB As pdDIB, ByRef srcDIB As pdDIB, ByVal iWidth As Long, ByVal iHeight As Long, ByVal interpolationMethod As FREE_IMAGE_FILTER)
    
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    
    'Double-check that FreeImage exists
    If ImageFormats.IsFreeImageEnabled() Then
        
        'If srcDIB.GetAlphaPremultiplication Then srcDIB.SetAlphaPremultiplication False
        
        'Convert the current image to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = Plugin_FreeImage.GetFIHandleFromPDDib_NoCopy(srcDIB, False)
        
        'Use that handle to request an image resize
        If (fi_DIB <> 0) Then
            
            Dim returnDIB As Long
            returnDIB = FreeImage_RescaleByPixel(fi_DIB, iWidth, iHeight, True, interpolationMethod)
            
            'Resize the destination DIB in preparation for the transfer
            If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
            If (dstDIB.GetDIBWidth <> iWidth) Or (dstDIB.GetDIBHeight <> iHeight) Then
                dstDIB.CreateBlank iWidth, iHeight, srcDIB.GetDIBColorDepth
            Else
                dstDIB.ResetDIB 0
            End If
            
            'Copy the bits from the FreeImage DIB to our DIB
            Plugin_FreeImage.PaintFIDibToPDDib dstDIB, returnDIB, 0, 0, iWidth, iHeight
            
            'With the transfer complete, release the FreeImage DIB and unload the library
            If (returnDIB <> 0) Then FreeImage_UnloadEx returnDIB
            
        End If
        
        'If the original image is 32bpp, mark correct premultiplication state now
        'If (srcDIB.GetDIBColorDepth = 32) Then dstDIB.SetInitialAlphaPremultiplicationState True
        
        'We now need to do something weird - because certain interpolation methods can cause
        ' "ringing" artifacts that don't obey alpha premultiplication rules (e.g. the alpha data
        ' is no longer guaranteed to be in sync with RGB values), we need to un-premultiply the
        ' current results, then re-premultiply them again.
        dstDIB.SetAlphaPremultiplication False, True
        dstDIB.SetAlphaPremultiplication True, True
        
    End If
    
End Sub

'Resize an image using our own internal algorithms.  Slower, but better quality.
Private Function InternalImageResize(ByRef dstDIB As pdDIB, ByRef srcDIB As pdDIB, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal interpolationMethod As PD_ResamplingFilter, ByVal displayProgress As Boolean) As Boolean
    
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
    
    'Hand off the image to PD's internal resampler
    InternalImageResize = Resampling.ResampleImage(dstDIB, srcDIB, dstWidth, dstHeight, interpolationMethod, displayProgress)
    
    'Premultiply the resulting image
    dstDIB.SetAlphaPremultiplication True, True
    
End Function

'Resize an image using any one of several resampling algorithms.  (Some algorithms are provided by FreeImage.)
Public Sub ResizeImage(ByVal resizeParams As String)
        
    'Parse incoming parameters into type-appropriate vars
    Dim imgWidth As Double, imgHeight As Double, imgDPI As Double
    Dim resampleMethod As PD_ResamplingFilter, allowApproximation As Boolean
    Dim fitMethod As PD_ResizeFit, newBackColor As Long
    Dim imgResizeUnit As PD_MeasurementUnit
    Dim thingToResize As PD_ActionTarget
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString resizeParams
    
    With cParams
        imgWidth = .GetDouble("width")
        imgHeight = .GetDouble("height")
        imgResizeUnit = .GetLong("unit", mu_Pixels)
        imgDPI = .GetDouble("ppi", 96)
        fitMethod = .GetLong("fit", ResizeFitStretch)
        newBackColor = .GetLong("fillcolor", vbWhite)
        thingToResize = .GetLong("target", pdat_Image)
    End With
    
    'Use a separate function to retrieve resampling method and approximation permission
    ' (legacy items may need to be dealt with)
    resampleMethod = GetResampleMethod(cParams, allowApproximation)
    
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
    If (resampleMethod = rf_Automatic) Then
        If (fitWidth < srcWidth) Or (fitHeight < srcHeight) Then
            If allowApproximation Then
                resampleMethod = rf_CubicBSpline
            Else
                resampleMethod = rf_Mitchell
            End If
        Else
            resampleMethod = rf_Lanczos3
        End If
    End If
    
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
        ' place a resized version of tmpLayerRef.layerDIB into tmpDIB.
        
        Select Case resampleMethod
            
            'For nearest-neighbor scaling, GDI can be used for great performance
            Case rf_Box
            
                'Copy the current DIB into this temporary DIB at the new size.  (StretchBlt is used
                ' for a fast resize.)
                If USE_GDIPLUS_RESIZE And allowApproximation Then
                    tmpDIB.CreateFromExistingDIB tmpLayerRef.layerDIB, fitWidth, fitHeight, GP_IM_NearestNeighbor
                
                'For a slower, but more mathematically accurate approach, you could also use our internal scaler
                Else
                    InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_Box, (firstLayerIndex = lastLayerIndex)
                End If
                
            'Bilinear resampling can use GDI+ (preferentially), our internal resampler, or the FreeImage library
            Case rf_BilinearTriangle
                If USE_GDIPLUS_RESIZE And allowApproximation Then
                    If (tmpDIB.GetDIBWidth <> fitWidth) Or (tmpDIB.GetDIBHeight <> fitHeight) Then tmpDIB.CreateBlank fitWidth, fitHeight, 32, 0 Else tmpDIB.ResetDIB 0
                    GDIPlusResizeDIB tmpDIB, 0, 0, fitWidth, fitHeight, tmpLayerRef.layerDIB, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), GP_IM_HighQualityBilinear, P2_PO_Half
                Else
                    InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_BilinearTriangle, (firstLayerIndex = lastLayerIndex)
                End If
            
            Case rf_Hermite
                InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_Hermite, (firstLayerIndex = lastLayerIndex)
            
            Case rf_Bell
                InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_Bell, (firstLayerIndex = lastLayerIndex)
                
            'Bicubic sampling can use GDI+ (preferentially), our internal resampler, or the FreeImage library (currently disabled)
            Case rf_CubicBSpline
                If USE_GDIPLUS_RESIZE And allowApproximation Then
                    If (tmpDIB.GetDIBWidth <> fitWidth) Or (tmpDIB.GetDIBHeight <> fitHeight) Then tmpDIB.CreateBlank fitWidth, fitHeight, 32, 0 Else tmpDIB.ResetDIB 0
                    GDIPlusResizeDIB tmpDIB, 0, 0, fitWidth, fitHeight, tmpLayerRef.layerDIB, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), GP_IM_HighQualityBicubic, P2_PO_Half
                Else
                    InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_CubicBSpline, (firstLayerIndex = lastLayerIndex)
                End If
            
            Case rf_Lanczos3
                If (USE_FREEIMAGE_RESIZE And ImageFormats.IsFreeImageEnabled) Then
                    FreeImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, FILTER_LANCZOS3
                Else
                    InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_Lanczos3, (firstLayerIndex = lastLayerIndex)
                End If
            
            'All subsequent methods use either our internal resampler, or the FreeImage library (when applicable)
            Case rf_Mitchell
                If (USE_FREEIMAGE_RESIZE And ImageFormats.IsFreeImageEnabled) Then
                    FreeImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, FILTER_BICUBIC
                Else
                    InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_Mitchell, (firstLayerIndex = lastLayerIndex)
                End If
            
            Case rf_Cosine
                InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_Cosine, (firstLayerIndex = lastLayerIndex)
            
            Case rf_CatmullRom
                If (USE_FREEIMAGE_RESIZE And ImageFormats.IsFreeImageEnabled) Then
                    FreeImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, FILTER_CATMULLROM
                Else
                    InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_CatmullRom, (firstLayerIndex = lastLayerIndex)
                End If
            
            Case rf_Quadratic
                InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_Quadratic, (firstLayerIndex = lastLayerIndex)
            
            Case rf_QuadraticBSpline
                InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_QuadraticBSpline, (firstLayerIndex = lastLayerIndex)
                
            Case rf_CubicConvolution
                InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_CubicConvolution, (firstLayerIndex = lastLayerIndex)
                
            Case rf_Lanczos8
                InternalImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, rf_Lanczos8, (firstLayerIndex = lastLayerIndex)
            
            'This failsafe should never be triggered
            Case Else
                PDDebug.LogAction "WARNING: FormResize.ResizeImage encountered an unknown resize filter: " & resampleMethod
                If (tmpDIB.GetDIBWidth <> fitWidth) Or (tmpDIB.GetDIBHeight <> fitHeight) Then tmpDIB.CreateBlank fitWidth, fitHeight, 32, 0 Else tmpDIB.ResetDIB 0
                GDIPlusResizeDIB tmpDIB, 0, 0, fitWidth, fitHeight, tmpLayerRef.layerDIB, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), GP_IM_Bicubic
            
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
                tmpLayerRef.layerDIB.CreateFromExistingDIB tmpDIB
        
            'Fit inclusively.  This fits the image's largest dimension into the destination image, which can leave
            ' blank space - that space is filled by the background color parameter passed in (or transparency,
            ' in the case of 32bpp images).
            Case ResizeFitInclusive
            
                'Resize the main DIB (destructively!) to fit the new dimensions
                tmpLayerRef.layerDIB.CreateBlank imgWidth, imgHeight, 32, 0
                
                'BitBlt the old image, centered, onto the new DIB
                If (srcAspect > dstAspect) Then
                    dstY = Int((imgHeight - fitHeight) / 2# + 0.5)
                    dstX = 0
                Else
                    dstX = Int((imgWidth - fitWidth) / 2# + 0.5)
                    dstY = 0
                End If
                
                GDI.BitBltWrapper tmpLayerRef.layerDIB.GetDIBDC, dstX, dstY, fitWidth, fitHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
                tmpLayerRef.layerDIB.SetInitialAlphaPremultiplicationState tmpDIB.GetAlphaPremultiplication
            
            'Fit exclusively.  This fits the image's smallest dimension into the destination image, which means no
            ' blank space - but parts of the image may get cropped out.
            Case ResizeFitExclusive
            
                'Resize the main DIB (destructively!) to fit the new dimensions
                tmpLayerRef.layerDIB.CreateBlank imgWidth, imgHeight, 32, 0
            
                'BitBlt the old image, centered, onto the new DIB
                If (srcAspect < dstAspect) Then
                    dstY = Int((imgHeight - fitHeight) / 2# + 0.5)
                    dstX = 0
                Else
                    dstX = Int((imgWidth - fitWidth) / 2# + 0.5)
                    dstY = 0
                End If
                
                GDI.BitBltWrapper tmpLayerRef.layerDIB.GetDIBDC, dstX, dstY, fitWidth, fitHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
                tmpLayerRef.layerDIB.SetInitialAlphaPremultiplicationState tmpDIB.GetAlphaPremultiplication
                
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
        DisplaySize PDImages.GetActiveImage()
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
        .AddParam "fit", cmbFit.ListIndex
        .AddParam "fillcolor", csBackground.Color
        .AddParam "target", m_ResizeTarget
        
    End With
    
    GetLocalParamString = cParams.GetParamString()
    
End Function

'This separate function is necessary for retrieving the preferred resampling method from
' PD's XML-based serializer (which stores macro commands, among other details).  Old versions
' of the serializer used 0-based integers to define resampling methods; we now use string IDs
' as they're much easier to make backward- and forward-compatible against new features.
Private Function GetResampleMethod(ByRef cParams As pdSerialize, ByRef allowApproximation As Boolean) As PD_ResamplingFilter

    'Current parameter strings (v3) use string IDs
    If (cParams.GetParamVersion() >= 3#) Then
        GetResampleMethod = Resampling.GetResamplerID(cParams.GetString("resample", "auto", True))
        allowApproximation = cParams.GetBool("approximations-ok", True, True)
        
    'Legacy implementations follow
    Else
    
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
                    GetResampleMethod = rf_Lanczos3
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
                    GetResampleMethod = rf_Lanczos3
                    allowApproximation = True
                Case Else
                    GetResampleMethod = rf_Automatic
                    allowApproximation = True
            End Select
            
        End If
    
    End If

End Function
