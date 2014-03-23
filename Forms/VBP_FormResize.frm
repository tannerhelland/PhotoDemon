VERSION 5.00
Begin VB.Form FormResize 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Resize Image"
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
   Begin VB.ComboBox cboResampleFriendly 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.ComboBox cmbFit 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5640
      Width           =   7935
   End
   Begin VB.ComboBox cboResampleTechnical 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3960
      Width           =   7935
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.smartResize ucResize 
      Height          =   2850
      Left            =   480
      TabIndex        =   1
      Top             =   480
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
      Height          =   480
      Left            =   840
      TabIndex        =   3
      Top             =   4440
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   847
      Caption         =   "show technical names"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PhotoDemon.colorSelector colorPicker 
      Height          =   495
      Left            =   840
      TabIndex        =   7
      ToolTipText     =   "Click to change the color used for empty borders"
      Top             =   6120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   873
   End
   Begin VB.Label lblSize 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "new size:"
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
      TabIndex        =   8
      Top             =   120
      Width           =   990
   End
   Begin VB.Label lblFit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "when changing aspect ratio, fit image to new size by:"
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
      TabIndex        =   5
      Top             =   5160
      Width           =   5655
   End
   Begin VB.Label lblResample 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "resize quality:"
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
      Top             =   3480
      Width           =   1470
   End
End
Attribute VB_Name = "FormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Size Handler
'Copyright ©2001-2014 by Tanner Helland
'Created: 6/12/01
'Last updated: 09/February/14
'Last update: further redesigns to account for "SmartResize" user control.
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
Dim lastSelectedResample As Long

Private Enum ResampleNameType
    rsFriendly = 0
    rsTechnical = 1
End Enum

#If False Then
    Const rsFriendly = 0, rsTechnical = 1
#End If

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Whenever the user toggles technical and friendly resample options, this sub is called.  It will translate between
' friendly and technical choices, as well as displaying the proper combo box.
Private Sub switchResampleOption()

    Dim i As Long

    'Technical names
    If CBool(chkNames) Then
    
        'Show a descriptive label
        lblResample.Caption = g_Language.TranslateMessage("resampling algorithm:")
    
        'Show the proper combo box
        cboResampleTechnical.Visible = True
        cboResampleFriendly.Visible = False
        
        'Find the list entry that corresponds to the current "friendly name" option
'        For i = 0 To numResamples(rsTechnical) - 1
'            If resampleTypes(rsTechnical, i).ProgramID = resampleTypes(rsFriendly, cboResampleFriendly.ListIndex).ProgramID Then
'                cboResampleTechnical.ListIndex = i
'                Exit For
'            End If
'        Next i
    
    'Friendly names are selected
    Else
    
        'Show a descriptive label
        lblResample.Caption = g_Language.TranslateMessage("resampling quality:")
        
        'Show the proper combo box
        cboResampleFriendly.Visible = True
        cboResampleTechnical.Visible = False
        
'        'Find the list entry that corresponds to the current "technical name" option.  If one does not exist,
'        ' select "Best for photographs"
'        Dim entryFound As Boolean
'        entryFound = False
'
'        For i = 0 To numResamples(rsFriendly) - 1
'            If resampleTypes(rsFriendly, i).ProgramID = resampleTypes(rsTechnical, cboResampleTechnical.ListIndex).ProgramID Then
'                entryFound = True
'                cboResampleFriendly.ListIndex = i
'                Exit For
'            End If
'        Next i
'
'        'No friendly entry was found that matches the user's selected technical entry.  Select "best for photographs".
'        If Not entryFound Then cboResampleFriendly.ListIndex = 0
        
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
        addResample g_Language.TranslateMessage("best for photographs"), RESIZE_BILINEAR, rsFriendly
        addResample g_Language.TranslateMessage("best for text and illustrations"), RESIZE_HALFTONE, rsFriendly
        addResample g_Language.TranslateMessage("fastest"), RESIZE_NORMAL, rsFriendly
    End If
    
    'Next, populate the "technical" names options.  This list should expose every algorithm we have
    ' access to.  Again, if FreeImage is available, far more options exist.
    addResample g_Language.TranslateMessage("Nearest Neighbor"), RESIZE_NORMAL, rsTechnical
    addResample g_Language.TranslateMessage("Halftone"), RESIZE_HALFTONE, rsTechnical
    addResample g_Language.TranslateMessage("Bilinear"), RESIZE_BILINEAR, rsTechnical
    
    'If the FreeImage library is available, add additional resize options to the combo box
    If g_ImageFormats.FreeImageEnabled Then
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
        
        'FreeImage not enabled; select Bilinear
        Else
            cboResampleTechnical.ListIndex = 2
        End If
        
    'Friendly drop-down:
    
        'Always select "best for photos"
        cboResampleFriendly.ListIndex = 0
    
End Sub

'New to v6.0, PhotoDemon gives the user friendly resample names by default.  They can toggle these off at their liking.
Private Sub chkNames_Click()
    switchResampleOption
End Sub

Private Sub cmbFit_Click()
    
    'Hide the color picker as necessary
    If (cmbFit.ListIndex = 1) And (pdImages(g_CurrentImage).getCompositeImageColorDepth <> 32) Then
        colorPicker.Visible = True
    Else
        colorPicker.Visible = False
    End If
    
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
    
    Process "Resize", , buildParams(ucResize.imgWidth, ucResize.imgHeight, resampleAlgorithm, cmbFit.ListIndex, colorPicker.Color, ucResize.unitOfMeasurement, ucResize.imgDPIAsPPI)
    
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
    ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).getDPI
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

    'Automatically set the width and height text boxes to match the image's current dimensions
    ucResize.unitOfMeasurement = MU_PIXELS
    ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).getDPI
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
    ucResize.setInitialDimensions pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).getDPI
    
    'Add some tooltips
    cboResampleFriendly.ToolTipText = g_Language.TranslateMessage("Resampling affects the quality of a resized image.  For a good summary of resampling techniques, visit the Image Resampling article on Wikipedia.")
    cboResampleTechnical.ToolTipText = g_Language.TranslateMessage("Resampling affects the quality of a resized image.  For a good summary of resampling techniques, visit the Image Resampling article on Wikipedia.")
    chkNames.ToolTipText = g_Language.TranslateMessage("By default, descriptive names are used in place of technical ones.  Advanced users can toggle this option to expose more resampling techniques.")
    cmbFit.ToolTipText = g_Language.TranslateMessage("When changing an image's aspect ratio, undesirable stretching may occur.  PhotoDemon can avoid this by using empty borders or cropping instead.")
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Resize an image using the FreeImage library.  Very fast.
Private Sub FreeImageResize(ByRef dstDIB As pdDIB, ByRef srcDIB As pdDIB, ByVal iWidth As Long, iHeight As Long, ByVal interpolationMethod As Long)
    
    'Double-check that FreeImage exists
    If g_ImageFormats.FreeImageEnabled Then
        
        Message "Resampling image using the FreeImage plugin..."
        
        'If the original image is 32bpp, remove premultiplication now
        If srcDIB.getDIBColorDepth = 32 Then srcDIB.fixPremultipliedAlpha
        
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
        If srcDIB.getDIBColorDepth = 32 Then dstDIB.fixPremultipliedAlpha True
        
    End If
    
End Sub

'Resize an image using any one of several resampling algorithms.  (Some algorithms are provided by FreeImage.)
Public Sub ResizeImage(ByVal iWidth As Double, ByVal iHeight As Double, ByVal resampleMethod As Long, ByVal fitMethod As Long, Optional ByVal newBackColor As Long = vbWhite, Optional ByVal unitOfMeasurement As MeasurementUnit = MU_PIXELS, Optional ByVal iDPI As Long)

    'TODO!  - Add GDI+ as a resize option
    'TODO!  Rework this function to work with layers

    'Depending on the requested fitting technique, we may have to resize the image to a slightly different size
    ' than the one requested.  Before doing anything else, calculate that new size.
    Dim fitWidth As Long, fitHeight As Long
    
    Dim srcWidth As Long, srcHeight As Long
    srcWidth = pdImages(g_CurrentImage).Width
    srcHeight = pdImages(g_CurrentImage).Height
    
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
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    'Because most resize methods require a temporary DIB, create one here
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB

    Select Case resampleMethod

        'Nearest neighbor...
        Case RESIZE_NORMAL
        
            Message "Resizing image..."
            
            'Copy the current DIB into this temporary DIB at the new size
            'tmpDIB.createFromExistingDIB pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, False
            
        'Halftone resampling... I'm not sure what to actually call it, but since it's based off the
        ' StretchBlt mode Microsoft calls "halftone," I'm sticking with that
        Case RESIZE_HALFTONE
            
            Message "Resizing image..."
            
            'Copy the current DIB into this temporary DIB at the new size
            'tmpDIB.createFromExistingDIB pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, True
        
        'True bilinear sampling
        Case RESIZE_BILINEAR
        
            'If FreeImage is enabled, use their bilinear filter.  Similar results, much faster.
            If g_ImageFormats.FreeImageEnabled Then
            
                'FreeImageResize tmpDIB, pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, FILTER_BILINEAR
            
            'If FreeImage is not enabled, we have to do the resample ourselves.
            Else
            
                'TODO!  Probably kill this function, as it's going to be a lot of work to rework it for layers.
                '        Instead, rely on GDI+ for bilinear resizes when FreeImage is not available.
            
                Message "Resampling image..."
        
                'Create a local array and point it at the pixel data of the current image
                Dim srcImageData() As Byte
                Dim srcSA As SAFEARRAY2D
                prepImageData srcSA
                CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
                'Resize the temporary DIB to the target size, and point a second local array at it
                'tmpDIB.createBlank fitWidth, fitHeight, pdImages(g_CurrentImage).mainDIB.getDIBColorDepth
                
                Dim dstImageData() As Byte
                Dim dstSA As SAFEARRAY2D
                
                prepSafeArray dstSA, tmpDIB
                CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
                
                'These values will help us access locations in the array more quickly.
                ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
                Dim qvDepth As Long
                qvDepth = tmpDIB.getDIBColorDepth \ 8
                
                'Create a filter support class, which will aid with edge handling and interpolation
                Dim fSupport As pdFilterSupport
                Set fSupport = New pdFilterSupport
                fSupport.setDistortParameters qvDepth, EDGE_CLAMP, True, curDIBValues.maxX, curDIBValues.MaxY
    
                'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
                ' based on the size of the area to be processed.
                Dim progBarCheck As Long
                SetProgBarMax iWidth
                progBarCheck = findBestProgBarValue()
            
                'Resampling requires many variables
                
                'Scaled ratios between the old x and y values and the new ones
                Dim xScale As Double, yScale As Double
                xScale = (srcWidth - 1) / fitWidth
                yScale = (srcHeight - 1) / fitHeight
                            
                'Coordinate variables for source and destination
                Dim x As Long, y As Long
                Dim srcX As Double, srcY As Double
                            
                For x = 0 To fitWidth - 1
                    
                    'Generate the x calculation variables
                    srcX = x * xScale
                    
                    'Draw each pixel in the new image
                    For y = 0 To fitHeight - 1
                        
                        'Generate the y calculation variables
                        srcY = y * yScale
                        
                        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
                        fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
                                            
                    Next y
                
                    If (x And progBarCheck) = 0 Then SetProgBarVal x
                    
                Next x
                            
                'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
                CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
                Erase srcImageData
                
                CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
                Erase dstImageData
                
                SetProgBarVal 0
                releaseProgressBar
                
            End If
        
        Case RESIZE_BSPLINE
            'FreeImageResize tmpDIB, pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, FILTER_BSPLINE
            
        Case RESIZE_BICUBIC_MITCHELL
            'FreeImageResize tmpDIB, pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, FILTER_BICUBIC
            
        Case RESIZE_BICUBIC_CATMULL
            'FreeImageResize tmpDIB, pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, FILTER_CATMULLROM
        
        Case RESIZE_LANCZOS
            'FreeImageResize tmpDIB, pdImages(g_CurrentImage).mainDIB, fitWidth, fitHeight, FILTER_LANCZOS3
            
    End Select
    
    'The temporary DIB now holds a copy of the resized image.
    
    'Calculate the aspect ratio of this DIB and the target picture box
    Dim srcAspect As Double, dstAspect As Double
    srcAspect = pdImages(g_CurrentImage).Width / pdImages(g_CurrentImage).Height
    dstAspect = iWidth / iHeight
    
    Dim dstX As Long, dstY As Long
    
    'We now want to copy the resized image into the current image using the technique requested by the user.
    Select Case fitMethod
    
        'Stretch-to-fit.  This is default resize behavior in all image editing software
        Case 0
    
            'Very simple - just copy the resized image back into the main DIB
            'pdImages(g_CurrentImage).mainDIB.createFromExistingDIB tmpDIB
    
        'Fit inclusively.  This fits the image's largest dimension into the destination image, which can leave
        ' blank space - that space is filled by the background color parameter passed in.
        Case 1
        
            'Resize the main DIB (destructively!) to fit the new dimensions
            'pdImages(g_CurrentImage).mainDIB.createBlank iWidth, iHeight, pdImages(g_CurrentImage).mainDIB.getDIBColorDepth, newBackColor
        
            'BitBlt the old image, centered, onto the new DIB
            If srcAspect > dstAspect Then
                dstY = CLng((iHeight - fitHeight) / 2)
                dstX = 0
            Else
                dstX = CLng((iWidth - fitWidth) / 2)
                dstY = 0
            End If
            
            'BitBlt pdImages(g_CurrentImage).mainDIB.getDIBDC, dstX, dstY, fitWidth, fitHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
        
        'Fit exclusively.  This fits the image's smallest dimension into the destination image, which means no
        ' blank space - but parts of the image may get cropped out.
        Case 2
        
            'Resize the main DIB (destructively!) to fit the new dimensions
            'pdImages(g_CurrentImage).mainDIB.createBlank iWidth, iHeight, pdImages(g_CurrentImage).mainDIB.getDIBColorDepth, newBackColor
        
            'BitBlt the old image, centered, onto the new DIB
            If srcAspect < dstAspect Then
                dstY = CLng((iHeight - fitHeight) / 2)
                dstX = 0
            Else
                dstX = CLng((iWidth - fitWidth) / 2)
                dstY = 0
            End If
            
            'BitBlt pdImages(g_CurrentImage).mainDIB.getDIBDC, dstX, dstY, fitWidth, fitHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
        
    End Select
    
    'We are finished with the temporary DIB, so release it
    Set tmpDIB = Nothing
    
    'Update the main image's size and DPI values
    pdImages(g_CurrentImage).updateSize
    pdImages(g_CurrentImage).setDPI iDPI, iDPI
    DisplaySize pdImages(g_CurrentImage)
    
    'Fit the new image on-screen and redraw its viewport
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Image resize"
    
    Message "Finished."
    
End Sub

