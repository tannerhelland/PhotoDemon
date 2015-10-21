VERSION 5.00
Begin VB.Form FormResize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resize image"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   9630
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
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdComboBox cmbFit 
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   5640
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   635
   End
   Begin PhotoDemon.pdComboBox cboResampleFriendly 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   3960
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   635
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6750
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.smartResize ucResize 
      Height          =   2850
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5027
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartCheckBox chkNames 
      Height          =   300
      Left            =   840
      TabIndex        =   2
      Top             =   4440
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   582
      Caption         =   "show technical names"
   End
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   495
      Left            =   840
      TabIndex        =   5
      ToolTipText     =   "Click to change the color used for empty borders"
      Top             =   6120
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   873
   End
   Begin PhotoDemon.pdComboBox cboResampleTechnical 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   3960
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   635
   End
   Begin VB.Label lblFit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "when changing aspect ratio, fit image to new size by"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Top             =   5160
      Width           =   5565
   End
   Begin VB.Label lblResample 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "resize quality"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   3480
      Width           =   1380
   End
End
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Size Handler
'Copyright 2001-2015 by Tanner Helland
'Created: 6/12/01
'Last updated: 09/May/14
'Last update: allow resizing of the entire image, or a single layer
'
'Handles all image-size related functions.  Currently supports nearest-neighbor and halftone resampling
' (via the API; not 100% accurate but faster than doing it manually), bilinear resampling via pure VB, and
' a number of more advanced resampling techniques via FreeImage.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The list of available resampling algorithms changes based on the presence of FreeImage, and the use of
' "friendly" vs "technical" names.  As a result, we have to track them dynamically using a custom type.
Private Type resampleAlgorithm
    Name As String
    ProgramID As Long
End Type

Dim resampleTypes() As resampleAlgorithm
Dim numResamples() As Long

Private Enum ResampleNameType
    rsFriendly = 0
    rsTechnical = 1
End Enum

#If False Then
    Const rsFriendly = 0, rsTechnical = 1
#End If

'This dialog can be used to resize the full image, or a single layer.  The requested target will be stored here,
' and can be externally accessed by the ResizeTarget property.
Private m_ResizeTarget As PD_ACTION_TARGET

Public Property Let ResizeTarget(newTarget As PD_ACTION_TARGET)
    m_ResizeTarget = newTarget
End Property

'Whenever the user toggles technical and friendly resample options, this sub is called.  It will translate between
' friendly and technical choices, as well as displaying the proper combo box.
Private Sub switchResampleOption()

    'Technical names
    If CBool(chkNames) Then
    
        'Show a descriptive label
        lblResample.Caption = g_Language.TranslateMessage("resampling algorithm:")
    
        'Show the proper combo box
        cboResampleTechnical.Visible = True
        cboResampleFriendly.Visible = False
        
    'Friendly names are selected
    Else
    
        'Show a descriptive label
        lblResample.Caption = g_Language.TranslateMessage("resampling quality:")
        
        'Show the proper combo box
        cboResampleFriendly.Visible = True
        cboResampleTechnical.Visible = False
        
    End If
    
End Sub

'Used by refillResampleBoxes, below, to keep track of what resample algorithms we have available
Private Sub addResample(ByVal rName As String, ByVal rID As Long, ByVal rCategory As ResampleNameType)
    resampleTypes(rCategory, numResamples(rCategory)).Name = rName
    resampleTypes(rCategory, numResamples(rCategory)).ProgramID = rID
    numResamples(rCategory) = numResamples(rCategory) + 1
End Sub

'Display all available resample algorithms in the combo box (contingent on the "show technical names" check box as well)
Private Sub refillResampleBoxes()

    'Resample Types stores resample data for two combo boxes: one that displays "friendly" names (0),
    ' and one that displays "technical" ones (1).  The numResamples() array stores the number of
    ' resample algorithms available as "friendly" entries (0) and "technical" entries (1).
    ReDim resampleTypes(0 To 1, 0 To 20) As resampleAlgorithm
    ReDim numResamples(0 To 1) As Long
    
    'Start with the "friendly" names options.  If FreeImage is available, we will map the friendly
    ' names to more advanced resample algorithms.  Without it, we are stuck with standard algorithms.
    If g_ImageFormats.FreeImageEnabled Then
        addResample g_Language.TranslateMessage("best for photographs"), RESIZE_LANCZOS, rsFriendly
        addResample g_Language.TranslateMessage("best for text and illustrations"), RESIZE_BICUBIC_MITCHELL, rsFriendly
        addResample g_Language.TranslateMessage("fastest"), RESIZE_NORMAL, rsFriendly
    Else
        addResample g_Language.TranslateMessage("best for photographs"), RESIZE_BSPLINE, rsFriendly
        addResample g_Language.TranslateMessage("best for text and illustrations"), RESIZE_BILINEAR, rsFriendly
        addResample g_Language.TranslateMessage("fastest"), RESIZE_NORMAL, rsFriendly
    End If
    
    'Next, populate the "technical" names options.  This list should expose every algorithm we have
    ' access to.  Again, if FreeImage is available, far more options exist.
    addResample g_Language.TranslateMessage("Nearest Neighbor"), RESIZE_NORMAL, rsTechnical
    addResample g_Language.TranslateMessage("Halftone"), RESIZE_HALFTONE, rsTechnical
    addResample g_Language.TranslateMessage("Bilinear"), RESIZE_BILINEAR, rsTechnical
    
    'If FreeImage is not enabled, only a single bicubic option will be listed; but if FreeImage IS enabled,
    ' we can list multiple bicubic variants.
    If Not g_ImageFormats.FreeImageEnabled Then
        addResample g_Language.TranslateMessage("Bicubic"), RESIZE_BSPLINE, rsTechnical
    Else
        addResample g_Language.TranslateMessage("B-Spline"), RESIZE_BSPLINE, rsTechnical
        addResample g_Language.TranslateMessage("Bicubic (Mitchell and Netravali)"), RESIZE_BICUBIC_MITCHELL, rsTechnical
        addResample g_Language.TranslateMessage("Bicubic (Catmull-Rom)"), RESIZE_BICUBIC_CATMULL, rsTechnical
        addResample g_Language.TranslateMessage("Sinc (Lanczos 3-lobe)"), RESIZE_LANCZOS, rsTechnical
    End If
    
    'Populate the Friendly combo box with friendly names, and the Technical box with technical ones.
    Dim i As Long
    
    cboResampleFriendly.Clear
    For i = 0 To numResamples(rsFriendly) - 1
        cboResampleFriendly.AddItem " " & resampleTypes(rsFriendly, i).Name, i
    Next i
    
    cboResampleTechnical.Clear
    For i = 0 To numResamples(rsTechnical) - 1
        cboResampleTechnical.AddItem " " & resampleTypes(rsTechnical, i).Name, i
    Next i
    
    'Intelligently select default values for the user.
    
    'Technical drop-down:
    
        'FreeImage enabled; select Bicubic (Catmull-Rom)
        If g_ImageFormats.FreeImageEnabled Then
            cboResampleTechnical.ListIndex = 5
        
        'FreeImage not enabled; select plain Bicubic
        Else
            cboResampleTechnical.ListIndex = 3
        End If
        
    'Friendly drop-down:
    
        'Always select "best for photos"
        cboResampleFriendly.ListIndex = 0
    
End Sub

'New to v6.0, PhotoDemon displays friendly resample names by default.  The user can toggle these off at their liking.
Private Sub chkNames_Click()
    switchResampleOption
End Sub

Private Sub cmbFit_Click()
    
    'NOTE: as of April '14 (version 6.4), layers support dictates that all images are treated as 32bpp by default.
    '      As such, there's no longer a reason to display a color selector, because transparency will always be
    '      available (regardless of actual layer contents).
    
    'Hide the color picker as necessary
    'If (cmbFit.ListIndex = 1) And (pdImages(g_CurrentImage).getCompositeImageColorDepth <> 32) Then
    '    colorPicker.Visible = True
    'Else
    '    colorPicker.Visible = False
    'End If
    
End Sub

Private Sub cmdBar_ExtraValidations()
    If Not ucResize.IsValid(True) Then cmdBar.validationFailed
End Sub

'OK button
Private Sub cmdBar_OKClick()
    
    'Retrieve the resample type selected by the user, which will vary depending on whether they used
    ' "technical" names or "friendly" ones.
    Dim resampleAlgorithm As Long
    If CBool(chkNames) Then
        resampleAlgorithm = resampleTypes(rsTechnical, cboResampleTechnical.ListIndex).ProgramID
    Else
        resampleAlgorithm = resampleTypes(rsFriendly, cboResampleFriendly.ListIndex).ProgramID
    End If
    
    Select Case m_ResizeTarget
    
        Case PD_AT_WHOLEIMAGE
            Process "Resize image", , buildParams(ucResize.imgWidth, ucResize.imgHeight, resampleAlgorithm, cmbFit.ListIndex, colorPicker.Color, ucResize.unitOfMeasurement, ucResize.imgDPIAsPPI, m_ResizeTarget), UNDO_IMAGE
        
        Case PD_AT_SINGLELAYER
            Process "Resize layer", , buildParams(ucResize.imgWidth, ucResize.imgHeight, resampleAlgorithm, cmbFit.ListIndex, colorPicker.Color, ucResize.unitOfMeasurement, ucResize.imgDPIAsPPI, m_ResizeTarget), UNDO_LAYER
    
    End Select
    
End Sub

'I'm not sure that randomize serves any purpose on this dialog, but as I don't have a way to hide that button at
' present, simply randomize the width/height to +/- the current image's width/height divided by two.
Private Sub cmdBar_RandomizeClick()

    ucResize.lockAspectRatio = False
    ucResize.imgWidthInPixels = (pdImages(g_CurrentImage).Width / 2) + (Rnd * pdImages(g_CurrentImage).Width)
    ucResize.imgHeightInPixels = (pdImages(g_CurrentImage).Height / 2) + (Rnd * pdImages(g_CurrentImage).Height)

End Sub

Private Sub cmdBar_ResetClick()
    
    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.unitOfMeasurement = MU_PIXELS
    
    Select Case m_ResizeTarget
    
        Case PD_AT_WHOLEIMAGE
            ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).getDPI
            
        Case PD_AT_SINGLELAYER
            ucResize.setInitialDimensions pdImages(g_CurrentImage).getActiveLayer.getLayerWidth(False), pdImages(g_CurrentImage).getActiveLayer.getLayerHeight(False), pdImages(g_CurrentImage).getDPI
        
    End Select
    
    ucResize.lockAspectRatio = True
    
    'Use friendly resample names by default
    cboResampleTechnical.ListIndex = 0
    cboResampleFriendly.ListIndex = 0
    chkNames.Value = vbUnchecked
    
    'Stretch to new aspect ratio by default
    cmbFit.ListIndex = 0
    
    'Make borders fill with black by default
    colorPicker.Color = RGB(0, 0, 0)
    
End Sub

Private Sub Form_Activate()

    'Set the dialog caption to match the current resize operation (resize image or resize single layer)
    Select Case m_ResizeTarget
        
        Case PD_AT_WHOLEIMAGE
            Me.Caption = g_Language.TranslateMessage("Resize image")
        
        Case PD_AT_SINGLELAYER
            Me.Caption = g_Language.TranslateMessage("Resize layer")
        
    End Select

    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.unitOfMeasurement = MU_PIXELS
    
    Select Case m_ResizeTarget
        
        Case PD_AT_WHOLEIMAGE
            ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).getDPI
            
        Case PD_AT_SINGLELAYER
            ucResize.setInitialDimensions pdImages(g_CurrentImage).getActiveLayer.getLayerWidth(False), pdImages(g_CurrentImage).getActiveLayer.getLayerHeight(False), pdImages(g_CurrentImage).getDPI
        
    End Select
    
    ucResize.lockAspectRatio = True

End Sub

'Certain actions are done at LOAD time instead of ACTIVATE time to minimize visible flickering
Private Sub Form_Load()
    
    'Populate the dropdowns with all available resampling algorithms.  (Availability depends on FreeImage.)
    refillResampleBoxes
    
    'Populate the "fit" options
    cmbFit.Clear
    cmbFit.AddItem " stretching to new size  (default)", 0
    If pdImages(g_CurrentImage).getCompositeImageColorDepth = 32 Then
        cmbFit.AddItem " fitting inclusively, with transparent borders as necessary", 1
    Else
        cmbFit.AddItem " fitting inclusively, with colored borders as necessary", 1
    End If
    cmbFit.AddItem " fitting exclusively, and cropping as necessary", 2
    cmbFit.ListIndex = 0
    
    'Automatically set the width and height text boxes to match the image's current dimensions.  (Note that we must
    ' do this again in the Activate step, as the last-used settings will automatically override these values.  However,
    ' if we do not also provide these values here, the resize control may attempt to set parameters while having
    ' a width/height/resolution of 0, which will cause divide-by-zero errors.)
    Select Case m_ResizeTarget
    
        Case PD_AT_WHOLEIMAGE
            ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).getDPI
            
        Case PD_AT_SINGLELAYER
            ucResize.setInitialDimensions pdImages(g_CurrentImage).getActiveLayer.getLayerWidth(False), pdImages(g_CurrentImage).getActiveLayer.getLayerHeight(False), pdImages(g_CurrentImage).getDPI
        
    End Select
    
    'Add some tooltips
    cboResampleFriendly.AssignTooltip "Resampling affects the quality of a resized image.  For a good summary of resampling techniques, visit the Image Resampling article on Wikipedia."
    cboResampleTechnical.AssignTooltip "Resampling affects the quality of a resized image.  For a good summary of resampling techniques, visit the Image Resampling article on Wikipedia."
    chkNames.AssignTooltip "By default, descriptive names are used in place of technical ones.  Advanced users can toggle this option to expose more resampling techniques."
    cmbFit.AssignTooltip "When changing an image's aspect ratio, undesirable stretching may occur.  PhotoDemon can avoid this by using empty borders or cropping instead."
    
    'Apply translations and visual themes
    MakeFormPretty Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Resize an image using the FreeImage library.  Very fast.
Private Sub FreeImageResize(ByRef dstDIB As pdDIB, ByRef srcDIB As pdDIB, ByVal iWidth As Long, iHeight As Long, ByVal interpolationMethod As Long)
    
    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
        
        'If the original image is 32bpp, remove premultiplication now
        If srcDIB.getDIBColorDepth = 32 Then srcDIB.setAlphaPremultiplication
        
        'Convert the current image to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(srcDIB.getDIBDC)
        
        'Use that handle to request an image resize
        If fi_DIB <> 0 Then
            
            Dim returnDIB As Long
            returnDIB = FreeImage_RescaleByPixel(fi_DIB, iWidth, iHeight, True, interpolationMethod)
            
            'Resize the destination DIB in preparation for the transfer
            dstDIB.createBlank iWidth, iHeight, srcDIB.getDIBColorDepth
            
            'Copy the bits from the FreeImage DIB to our DIB
            SetDIBitsToDevice dstDIB.getDIBDC, 0, 0, iWidth, iHeight, 0, 0, 0, iHeight, ByVal FreeImage_GetBits(returnDIB), ByVal FreeImage_GetInfo(returnDIB), 0&
     
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            
        End If
        
        'If the original image is 32bpp, add back in premultiplication now
        If srcDIB.getDIBColorDepth = 32 Then dstDIB.setAlphaPremultiplication True
        
    End If
    
End Sub

'Resize an image using any one of several resampling algorithms.  (Some algorithms are provided by FreeImage.)
Public Sub ResizeImage(ByVal iWidth As Double, ByVal iHeight As Double, ByVal resampleMethod As Long, ByVal fitMethod As Long, Optional ByVal newBackColor As Long = vbWhite, Optional ByVal unitOfMeasurement As MeasurementUnit = MU_PIXELS, Optional ByVal iDPI As Long, Optional ByVal thingToResize As PD_ACTION_TARGET = PD_AT_WHOLEIMAGE)
    
    'Depending on the requested fitting technique, we may have to resize the image to a slightly different size
    ' than the one requested.  Before doing anything else, calculate that new size.
    Dim fitWidth As Long, fitHeight As Long
    
    Dim srcWidth As Long, srcHeight As Long
    Select Case thingToResize
    
        Case PD_AT_WHOLEIMAGE
            srcWidth = pdImages(g_CurrentImage).Width
            srcHeight = pdImages(g_CurrentImage).Height
        
        Case PD_AT_SINGLELAYER
            srcWidth = pdImages(g_CurrentImage).getActiveLayer.getLayerWidth(False)
            srcHeight = pdImages(g_CurrentImage).getActiveLayer.getLayerHeight(False)
        
    End Select
    
    'In past versions of the software, we could assume the passed measurements were always in pixels,
    ' but that is no longer the case!  Using the supplied "unit of measurement", convert the passed
    ' width and height values to pixel measurements.
    iWidth = convertOtherUnitToPixels(unitOfMeasurement, iWidth, iDPI, srcWidth)
    iHeight = convertOtherUnitToPixels(unitOfMeasurement, iHeight, iDPI, srcHeight)
    
    Select Case fitMethod
    
        'Stretch-to-fit.  Default behavior, and no size changes are required.
        Case 0
            fitWidth = iWidth
            fitHeight = iHeight
        
        'Fit inclusively.  Fit the image's largest dimension.  No cropping will occur, but blank space may be present.
        Case 1
            
            'We have an existing function for this purpose.  (It's used when rendering preview images, for example.)
            convertAspectRatio srcWidth, srcHeight, iWidth, iHeight, fitWidth, fitHeight
            
        'Fit exclusively.  Fit the image's smallest dimension.  Cropping will occur, but no blank space will be present.
        Case 2
        
            convertAspectRatio srcWidth, srcHeight, iWidth, iHeight, fitWidth, fitHeight, False
        
    End Select
    
    'If the image contains an active selection, automatically deactivate it
    If pdImages(g_CurrentImage).selectionActive And (thingToResize = PD_AT_WHOLEIMAGE) Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    'Because most resize methods require a temporary DIB, create one here
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Because we will likely use outside libraries for the resize (FreeImage, GDI+), we won't be able to track
    ' detailed progress of the actions.  Instead, let the user know when a layer has been resized by using
    ' the number of layers as our progress guide.
    If (thingToResize = PD_AT_WHOLEIMAGE) Then
        SetProgBarMax pdImages(g_CurrentImage).getNumOfLayers
        Message "Resizing image..."
    Else
        SetProgBarMax 1
        Message "Resizing layer..."
    End If
        
    Dim srcAspect As Double, dstAspect As Double
    Dim dstX As Long, dstY As Long
    
    'It is now time to iterate through all layers, resizing as we go.  Note that PD's approach to multi-layer
    ' operations allows us to use the same resize code for each layer, because layers smaller than the image
    ' will be automatically padded to the image's full size.
    Dim tmpLayerRef As pdLayer
    
    'If we are resizing the entire image, we must handle all layers in turn.  Otherwise, we can handle just
    ' the active layer.
    Dim lInit As Long, lFinal As Long
    
    Select Case thingToResize
    
        Case PD_AT_WHOLEIMAGE
            lInit = 0
            lFinal = pdImages(g_CurrentImage).getNumOfLayers - 1
        
        Case PD_AT_SINGLELAYER
            lInit = pdImages(g_CurrentImage).getActiveLayerIndex
            lFinal = pdImages(g_CurrentImage).getActiveLayerIndex
    
    End Select
    
    Dim i As Long
    For i = lInit To lFinal
    
        If thingToResize = PD_AT_WHOLEIMAGE Then SetProgBarVal i
        
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).getLayerByIndex(i)
        
        'Null-pad the layer
        If thingToResize = PD_AT_WHOLEIMAGE Then tmpLayerRef.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, False
        
        'Call the appropriate external function, based on the user's resize selection.  Each function will
        ' place a resized version of tmpLayerRef.layerDIB into tmpDIB.
        Select Case resampleMethod
    
            'Nearest neighbor...
            Case RESIZE_NORMAL
                
                'Copy the current DIB into this temporary DIB at the new size
                tmpDIB.createFromExistingDIB tmpLayerRef.layerDIB, fitWidth, fitHeight, False
                
            'Halftone resampling... I'm not sure what to actually call it, but since it's based off the
            ' StretchBlt mode Microsoft calls "halftone," I'm sticking with that.  Basically we get cheap
            ' supersampling when shrinking an image, and nearest-neighbor when enlarging.  This method
            ' is extraordinarily fast for batch shrinking of images, while providing reasonably good
            ' results.
            Case RESIZE_HALFTONE
                
                'Copy the current DIB into this temporary DIB at the new size
                tmpDIB.createFromExistingDIB tmpLayerRef.layerDIB, fitWidth, fitHeight, True
            
            'Bilinear sampling
            Case RESIZE_BILINEAR
            
                'If FreeImage is enabled, use their bilinear filter.
                If g_ImageFormats.FreeImageEnabled Then
                
                    FreeImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, FILTER_BILINEAR
                
                'If FreeImage is not enabled, use GDI+ instead.
                Else
                
                    tmpDIB.createBlank fitWidth, fitHeight, 32, 0
                    GDIPlusResizeDIB tmpDIB, 0, 0, fitWidth, fitHeight, tmpLayerRef.layerDIB, 0, 0, tmpLayerRef.getLayerWidth(False), tmpLayerRef.getLayerHeight(False), InterpolationModeHighQualityBilinear
                    
                End If
            
            Case RESIZE_BSPLINE
            
                'If FreeImage is enabled, use their bilinear filter.
                If g_ImageFormats.FreeImageEnabled Then
                
                    FreeImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, FILTER_BSPLINE
                
                'If FreeImage is not enabled, use GDI+ instead.
                Else
                
                    tmpDIB.createBlank fitWidth, fitHeight, 32, 0
                    GDIPlusResizeDIB tmpDIB, 0, 0, fitWidth, fitHeight, tmpLayerRef.layerDIB, 0, 0, tmpLayerRef.getLayerWidth(False), tmpLayerRef.getLayerHeight(False), InterpolationModeHighQualityBicubic
                    
                End If
                
            Case RESIZE_BICUBIC_MITCHELL
                FreeImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, FILTER_BICUBIC
                
            Case RESIZE_BICUBIC_CATMULL
                FreeImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, FILTER_CATMULLROM
            
            Case RESIZE_LANCZOS
                FreeImageResize tmpDIB, tmpLayerRef.layerDIB, fitWidth, fitHeight, FILTER_LANCZOS3
                
        End Select
        
        'tmpDIB now holds a copy of the resized image.
        
        'Calculate the aspect ratio of the temporary DIB and the target size.  If the user has selected
        ' a resize mode other than "fit exactly", we still need to do a bit of extra trimming.
        srcAspect = srcWidth / srcHeight
        dstAspect = iWidth / iHeight
        
        'We now want to copy the resized image into the current image using the technique requested by the user.
        Select Case fitMethod
        
            'Stretch-to-fit.  This is default resize behavior in all image editing software
            Case 0
        
                'Very simple - just copy the resized image back into the main DIB
                tmpLayerRef.layerDIB.createFromExistingDIB tmpDIB
        
            'Fit inclusively.  This fits the image's largest dimension into the destination image, which can leave
            ' blank space - that space is filled by the background color parameter passed in (or transparency,
            ' in the case of 32bpp images).
            Case 1
            
                'Resize the main DIB (destructively!) to fit the new dimensions
                tmpLayerRef.layerDIB.createBlank iWidth, iHeight, 32, 0
            
                'BitBlt the old image, centered, onto the new DIB
                If srcAspect > dstAspect Then
                    dstY = CLng((iHeight - fitHeight) / 2)
                    dstX = 0
                Else
                    dstX = CLng((iWidth - fitWidth) / 2)
                    dstY = 0
                End If
                
                BitBlt tmpLayerRef.layerDIB.getDIBDC, dstX, dstY, fitWidth, fitHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
            
            'Fit exclusively.  This fits the image's smallest dimension into the destination image, which means no
            ' blank space - but parts of the image may get cropped out.
            Case 2
            
                'Resize the main DIB (destructively!) to fit the new dimensions
                tmpLayerRef.layerDIB.createBlank iWidth, iHeight, 32, 0
            
                'BitBlt the old image, centered, onto the new DIB
                If srcAspect < dstAspect Then
                    dstY = CLng((iHeight - fitHeight) / 2)
                    dstX = 0
                Else
                    dstX = CLng((iWidth - fitWidth) / 2)
                    dstY = 0
                End If
                
                BitBlt tmpLayerRef.layerDIB.getDIBDC, dstX, dstY, fitWidth, fitHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
            
        End Select
        
        'With the layer now successfully resized, we can remove any null-padding that may still exist
        If thingToResize = PD_AT_WHOLEIMAGE Then tmpLayerRef.cropNullPaddedLayer
        
        'Notify the parent image of the change
        pdImages(g_CurrentImage).notifyImageChanged UNDO_LAYER, i
        
    'Move on to the next layer
    Next i
    
    'We are finished with the temporary DIB, so release it
    Set tmpDIB = Nothing
    
    'Update the main image's size and DPI values
    If thingToResize = PD_AT_WHOLEIMAGE Then
        pdImages(g_CurrentImage).updateSize False, iWidth, iHeight
        pdImages(g_CurrentImage).setDPI iDPI, iDPI
        DisplaySize pdImages(g_CurrentImage)
    End If
        
    'Fit the new image on-screen and redraw its viewport
    Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    'Release the progress bar
    SetProgBarVal 0
    releaseProgressBar
    
    Message "Finished."
    
End Sub

