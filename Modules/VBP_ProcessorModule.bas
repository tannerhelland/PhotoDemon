Attribute VB_Name = "Processor"
'***************************************************************************
'Program Sub-Processor and Error Handler
'Copyright ©2000-2013 by Tanner Helland
'Created: 4/15/01
'Last updated: 13/August/12
'Last update: built GetNameOfProcess for returning human-readable descriptions of processes
'
'Module for controlling calls to the various program functions.  Any action the program takes has to pass
' through here.  Why go to all that extra work?  A couple of reasons:
' 1) a central error handler that works for every sub throughout the program (due to recursive error handling)
' 2) PhotoDemon can run macros by simply tracking the values that pass through this routine
' 3) PhotoDemon can control code flow by delaying requests that pass through here (for example,
'    if the program is busy applying a filter, we can wait to process subsequent calls)
' 4) miscellaneous semantic benefits
'
'Due to the nature of this routine, very little of interest happens here - this is primarily a router
' for various functions, so the majority of the routine is a huge Case Select statement.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'GROUP IDENTIFIERS: Specify the broader group that an option is within
    '...may be added later, depending on preview options (such as if an effect browser is
    'created...)
'END GROUP IDENTIFIERS

'SUBIDENTIFIERS: Specify specific actions within a group

    'Main functions (not used for image editing); numbers 1-99
    '-File I/O
    Public Const FileOpen As Long = 1
    Public Const FileSave As Long = 2
    Public Const FileSaveAs As Long = 3
    '-Screen Capture
    Public Const capScreen As Long = 10
    'Clipboard constants:
    Public Const cCopy As Long = 20
    Public Const cPaste As Long = 21
    Public Const cEmpty As Long = 22
    'Undo
    Public Const Undo As Long = 30
    Public Const Redo As Long = 31
    'Macro conversion
    Public Const MacroStartRecording As Long = 40
    Public Const MacroStopRecording As Long = 41
    Public Const MacroPlayRecording As Long = 42
    'Scanning
    Public Const SelectScanner As Long = 50
    Public Const ScanImage As Long = 51
    
    'Histogram functions; numbers 100-199
    Public Const ViewHistogram As Long = 100
    Public Const StretchHistogram As Long = 101
    Public Const Equalize As Long = 102
    Public Const WhiteBalance As Long = 104
    'Note: 103 is empty (formerly EqualizeLuminance, which is now handled as part of Equalize)
    
    'Black/White conversion; numbers 200-299
    Public Const BWMaster As Long = 200 'Added 9/2012 - this is a single BW conversion routine to rule them all
    Public Const RemoveBW As Long = 201
    
    'Grayscale conversion; numbers 300-399
    Public Const Desaturate As Long = 300
    Public Const GrayScale As Long = 301
    Public Const GrayscaleAverage As Long = 302
    Public Const GrayscaleCustom As Long = 303
    Public Const GrayscaleCustomDither As Long = 304
    Public Const GrayscaleDecompose As Long = 305
    Public Const GrayscaleSingleChannel As Long = 306
    
    'Area filters; numbers 400-499
    '-Blur
    Public Const Blur As Long = 400
    Public Const BlurMore As Long = 401
    Public Const Soften As Long = 402
    Public Const SoftenMore As Long = 403
    '-Sharpen
    Public Const Sharpen As Long = 404
    Public Const SharpenMore As Long = 405
    Public Const Unsharp As Long = 406
    '-Diffuse
    Public Const Diffuse As Long = 407
    Public Const DiffuseMore As Long = 408
    Public Const CustomDiffuse As Long = 409
    '-Mosaic
    Public Const Mosaic As Long = 410
    '-Rank
    '413-414 have been deprecated
    Public Const MinimumRank As Long = 411
    Public Const MaximumRank As Long = 412
    '-Grid Blurring
    Public Const GridBlur As Long = 415
    '-Gaussian Blur
    Public Const GaussianBlur As Long = 416
    'Smart Blur
    Public Const SmartBlur As Long = 417
    'Box Blur
    Public Const BoxBlur As Long = 418
    
    'Edge filters; numbers 500-599
    '-Emboss
    Public Const EmbossToColor As Long = 500
    '-Engrave
    Public Const EngraveToColor As Long = 501
    '-Pencil
    Public Const Pencil As Long = 504
    '-Relief
    Public Const Relief As Long = 505
    '-Find Edges
    Public Const PrewittHorizontal As Long = 506
    Public Const PrewittVertical As Long = 507
    Public Const SobelHorizontal As Long = 508
    Public Const SobelVertical As Long = 509
    Public Const Laplacian As Long = 510
    Public Const SmoothContour As Long = 511
    Public Const HiliteEdge As Long = 512
    Public Const PhotoDemonEdgeLinear = 513
    Public Const PhotoDemonEdgeCubic = 514
    '-Edge enhance
    Public Const EdgeEnhance As Long = 515
    '-Trace contour
    Public Const Contour As Long = 516
    
    'Color operations; numbers 600-699
    '-Rechanneling
    Public Const Rechannel As Long = 600
    Public Const RechannelGreen As Long = 601   'This is here for legacy reasons only
    Public Const RechannelRed As Long = 602     'This is here for legacy reasons only
    '-Shifting
    Public Const ColorShiftLeft As Long = 603
    Public Const ColorShiftRight As Long = 604
    '-Intensity
    Public Const BrightnessAndContrast As Long = 605
    Public Const GammaCorrection As Long = 606
    '-Invert/Negative
    Public Const Invert As Long = 607
    Public Const InvertHue As Long = 608
    Public Const Negative As Long = 609
    Public Const CompoundInvert As Long = 617
    '-AutoEnhance
    Public Const AutoEnhance As Long = 610
    Public Const AutoHighlights As Long = 611
    Public Const AutoMidtones As Long = 612
    Public Const AutoShadows As Long = 613
    'Image levels
    Public Const ImageLevels As Long = 614
    'Colorize
    Public Const Colorize As Long = 615
    'Reduce image colors
    Public Const ReduceColors As Long = 616
    'Temperature
    Public Const AdjustTemperature As Long = 618
    'HSL Adjustment
    Public Const AdjustHSL As Long = 619
    'Color balance
    Public Const AdjustColorBalance As Long = 620
    'Shadow / midtone / highlight adjustments
    Public Const ShadowHighlight As Long = 621
    'NOTE: 621 is the max value for this section (ShadowHighlight)
    
    'Coordinate filters/transformations; numbers 700-799
    '-Resize
    Public Const ImageSize As Long = 700
    '-Orientation
    Public Const Flip As Long = 701
    Public Const Mirror As Long = 702
    '-Rotation
    Public Const Rotate90Clockwise As Long = 703
    Public Const Rotate180 As Long = 704
    Public Const Rotate270Clockwise As Long = 705
    Public Const FreeRotate As Long = 706
    '-Isometric
    Public Const Isometric As Long = 707
    '-Tiling
    Public Const Tile As Long = 708
    '-Crop to Selection
    Public Const CropToSelection As Long = 709
    '-Image Mode (it's a kind of transformation, right?)
    Public Const ChangeImageMode24 As Long = 710
    Public Const ChangeImageMode32 As Long = 711
    'This also includes DISTORT filters
    '-Distort: Swirl
    Public Const DistortSwirl As Long = 712
    '-Distort: Lens distortion/correction
    Public Const DistortLens As Long = 713
    Public Const DistortLensFix As Long = 714
    '-Distort: Water ripple
    Public Const DistortRipple As Long = 715
    '-Distort: Pinch and whirl
    Public Const DistortPinchAndWhirl As Long = 716
    '-Distort: Waves
    Public Const DistortWaves As Long = 717
    '-Distort: Etched glass
    Public Const DistortFiguredGlass As Long = 718
    '-Distort: Kaleidoscope
    Public Const DistortKaleidoscope As Long = 719
    '-Distort: Polar conversion
    Public Const ConvertPolar As Long = 720
    'Autocrop
    Public Const Autocrop As Long = 721
    '-Distort: Shear
    Public Const DistortShear As Long = 722
    '-Distort: Squish (formerly Fixed Perspective)
    Public Const DistortSquish As Long = 723
    '-Distort: Perspective (free)
    Public Const FreePerspective As Long = 724
    '-Distort: Pan and zoom (Ken Burns effect)
    Public Const DistortPanAndZoom As Long = 725
    
    'Other filters; numbers 800-899
    Public Const FilmNoir As Long = 801
    '-Compound invert
    '800-802 used to be specific CompoundInvert values; this is superceded by passing the values to CompoundInvert, which has been moved with the other Inverts
    '-Fade
    Public Const Fade As Long = 803
    '804-806 used to be specific Fade values; these have been superceded by passing the values to Fade
    Public Const UnFade As Long = 807
    '-Natural
    Public Const Atmospheric As Long = 808
    Public Const Frozen As Long = 809
    Public Const Lava As Long = 810
    Public Const Burn As Long = 811
    'Public Const Ocean As Long = 812
    Public Const Water As Long = 813
    Public Const Steel As Long = 814
    Public Const FogEffect As Long = 828
    Public Const Rainbow As Long = 829
    '-Custom filters
    Public Const CustomFilter As Long = 817
    '-Miscellaneous
    Public Const Dream As Long = 815
    Public Const Alien As Long = 816
    Public Const Antique As Long = 818
    Public Const BlackLight As Long = 819
    Public Const Posterize As Long = 820
    Public Const Radioactive As Long = 821
    Public Const Solarize As Long = 822
    Public Const Twins As Long = 823
    Public Const Synthesize As Long = 824
    Public Const Noise As Long = 825
    Public Const Sepia As Long = 826
    Public Const CountColors As Long = 827
    Public Const Vibrate As Long = 830
    Public Const Despeckle As Long = 831
    Public Const CustomDespeckle As Long = 832
    Public Const HeatMap As Long = 833
    Public Const ComicBook As Long = 840
    Public Const FilmGrain As Long = 841
    Public Const Vignetting As Long = 842
    Public Const Median As Long = 843
    Public Const ModernArt As Long = 844
    
    'Relative processes
    Public Const LastCommand As Long = 900
    Public Const FadeLastEffect As Long = 901
    
    'Other filters end at 844

    'On-Canvas Tools; numbers 1000-2000
    
    'Selections
    Public Const SelectionCreate As Long = 1000
    Public Const SelectionClear As Long = 1001
    
    'Reserved bytes; 2000 and up
    
'END SUBIDENTIFIERS (~130? currently)

'Data type for tracking processor calls - used for macros
'2012 model: MOST CURRENT
Public Type ProcessCall
    MainType As Long
    pOPCODE As Variant
    pOPCODE2 As Variant
    pOPCODE3 As Variant
    pOPCODE4 As Variant
    pOPCODE5 As Variant
    pOPCODE6 As Variant
    pOPCODE7 As Variant
    pOPCODE8 As Variant
    pOPCODE9 As Variant
    LoadForm As Boolean
    RecordAction As Boolean
End Type

'Array of processor calls - tracks what is going on
Public Calls() As ProcessCall

'Tracks the current array position
Public CurrentCall As Long

'Last filter call
Public LastFilterCall As ProcessCall

'Track processing (i.e. whether or not the software processor is busy right now
Public Processing As Boolean

'Processing time (to enable this, see the top constant in the Public_Constants module)
Private m_ProcessingTime As Double

'PhotoDemon's software processor.  Almost every action the program takes is routed through this method.  This is what
' allows us to record and playback macros, among other things.  (See comment at top of page for more details.)
Public Sub Process(ByVal pType As Long, Optional pOPCODE As Variant = 0, Optional pOPCODE2 As Variant = 0, Optional pOPCODE3 As Variant = 0, Optional pOPCODE4 As Variant = 0, Optional pOPCODE5 As Variant = 0, Optional pOPCODE6 As Variant = 0, Optional pOPCODE7 As Variant = 0, Optional pOPCODE8 As Variant = 0, Optional pOPCODE9 As Variant = 0, Optional LoadForm As Boolean = False, Optional RecordAction As Boolean = True)

    'Main error handler for the entire program is initialized by this line
    On Error GoTo MainErrHandler
    
    'If desired, this line can be used to artificially raise errors (to test the error handler)
    'Err.Raise 339
    
    'Mark the software processor as busy
    Processing = True
        
    FormMain.Enabled = False
    
    'Set the mouse cursor to an hourglass and lock the main form (to prevent additional input)
    If LoadForm = False Then
        Screen.MousePointer = vbHourglass
    Else
        setArrowCursor FormMain.ActiveForm
    End If
        
    'If we are to perform the last command, simply replace all the method parameters using data from the
    ' LastFilterCall object, then let the routine carry on as usual
    If pType = LastCommand Then
        pType = LastFilterCall.MainType
        pOPCODE = LastFilterCall.pOPCODE
        pOPCODE2 = LastFilterCall.pOPCODE2
        pOPCODE3 = LastFilterCall.pOPCODE3
        pOPCODE4 = LastFilterCall.pOPCODE4
        pOPCODE5 = LastFilterCall.pOPCODE5
        pOPCODE6 = LastFilterCall.pOPCODE6
        pOPCODE7 = LastFilterCall.pOPCODE7
        pOPCODE8 = LastFilterCall.pOPCODE8
        pOPCODE9 = LastFilterCall.pOPCODE9
        LoadForm = LastFilterCall.LoadForm
    End If
    
    'If the macro recorder is running and this option is recordable, store it in our array of
    'processor calls
    If (MacroStatus = MacroSTART) And (RecordAction = True) Then
        'Tracker variable (remembers where we are at in the array)
        CurrentCall = CurrentCall + 1
        
        'Copy the current function variables into the array
        ReDim Preserve Calls(0 To CurrentCall) As ProcessCall
        With Calls(CurrentCall)
            .MainType = pType
            .pOPCODE = pOPCODE
            .pOPCODE2 = pOPCODE2
            .pOPCODE3 = pOPCODE3
            .pOPCODE4 = pOPCODE4
            .pOPCODE5 = pOPCODE5
            .pOPCODE6 = pOPCODE6
            .pOPCODE7 = pOPCODE7
            .pOPCODE8 = pOPCODE8
            .pOPCODE9 = pOPCODE9
            .LoadForm = LoadForm
            .RecordAction = RecordAction
        End With
    End If
    
    
    'SUB HANDLER/PROCESSOR
    'From this point on, all we do is check the pType variable (the first variable passed
    ' to this subroutine) and depending on what it is, we call the appropriate subroutine.
    ' Very simple and very fast.
    
    'I have also subdivided the "Select Case" statements into groups of 100, just as I do
    ' above in the declarations part.  This is purely organizational.
    
    'Process types 0-99.  Main functions.  These are never recorded as part of macros.
    If pType > 0 And pType <= 99 Then
        Select Case pType
            Case FileOpen
                MenuOpen
            Case FileSave
                MenuSave CurrentImage
            Case FileSaveAs
                MenuSaveAs CurrentImage
            Case capScreen
                CaptureScreen
            Case cCopy
                ClipboardCopy
            Case cPaste
                ClipboardPaste
            Case cEmpty
                ClipboardEmpty
            Case Undo
                RestoreImage
                'Also, redraw the current child form icon
                CreateCustomFormIcon FormMain.ActiveForm
            Case Redo
                RedoImageRestore
                'Also, redraw the current child form icon
                CreateCustomFormIcon FormMain.ActiveForm
            Case MacroStartRecording
                StartMacro
            Case MacroStopRecording
                StopMacro
            Case MacroPlayRecording
                PlayMacro
            Case SelectScanner
                Twain32SelectScanner
            Case ScanImage
                Twain32Scan
        End Select
    End If
    
    'NON-IDENTIFIER CODE
    'Get image data and build the undo for any action that changes the image buffer
    
    'First, make sure that the current command is a filter or image-changing event
    If pType >= 101 Then
        
        'Temporarily disable drag-and-drop operations for the main form
        g_AllowDragAndDrop = False
        FormMain.OLEDropMode = 0
        
        'Only save an "undo" image if we are NOT loading a form for user input, and if
        'we ARE allowed to record this action, and if it's not counting colors (useless),
        ' and if we're not performing a batch conversion (saves a lot of time to not generate undo files!)
        If MacroStatus <> MacroBATCH Then
            If LoadForm <> True And RecordAction <> False And pType <> CountColors Then CreateUndoFile pType
        End If
        
        'Save this information in the LastFilterCall variable (to be used if the user clicks on
        ' Edit -> Redo Last Command.
        FormMain.MnuRepeatLast.Enabled = True
        LastFilterCall.MainType = pType
        LastFilterCall.pOPCODE = pOPCODE
        LastFilterCall.pOPCODE2 = pOPCODE2
        LastFilterCall.pOPCODE3 = pOPCODE3
        LastFilterCall.pOPCODE4 = pOPCODE4
        LastFilterCall.pOPCODE5 = pOPCODE5
        LastFilterCall.pOPCODE6 = pOPCODE6
        LastFilterCall.pOPCODE7 = pOPCODE7
        LastFilterCall.pOPCODE8 = pOPCODE8
        LastFilterCall.pOPCODE9 = pOPCODE9
        LastFilterCall.LoadForm = LoadForm
        
        'If the user wants us to time how long this action takes, mark the current time now
        If Not LoadForm Then
            If DISPLAY_TIMINGS Then m_ProcessingTime = Timer
        End If
        
    End If
    
    'Histogram functions
    Select Case pType
        Case ViewHistogram
            FormHistogram.Show 0, FormMain
        Case StretchHistogram
            FormHistogram.StretchHistogram
        Case Equalize
            If LoadForm Then
                FormEqualize.Show vbModal, FormMain
            Else
                FormEqualize.EqualizeHistogram pOPCODE, pOPCODE2, pOPCODE3, pOPCODE4
            End If
        Case WhiteBalance
            If LoadForm Then
                FormWhiteBalance.Show vbModal, FormMain
            Else
                FormWhiteBalance.AutoWhiteBalance pOPCODE
            End If
        
    'Black/White conversion
    'NOTE: as of PhotoDemon v5.0 all black/white conversions are being rebuilt in a single master function (masterBlackWhiteConversion).
    ' For sake of compatibility with old macros, I need to make sure old processor values are rerouted through the new master function.
        Case BWMaster
            If LoadForm Then
                FormBlackAndWhite.Show vbModal, FormMain
            Else
                FormBlackAndWhite.masterBlackWhiteConversion pOPCODE, pOPCODE2, pOPCODE3, pOPCODE4
            End If
        Case RemoveBW
            If LoadForm Then
                FormMonoToColor.Show vbModal, FormMain
            Else
                FormMonoToColor.ConvertMonoToColor pOPCODE
            End If
        
    'Grayscale conversion
        Case Desaturate
            FormGrayscale.MenuDesaturate
        Case GrayScale
            If LoadForm Then
                FormGrayscale.Show vbModal, FormMain
            Else
                FormGrayscale.MenuGrayscale
            End If
        Case GrayscaleAverage
            FormGrayscale.MenuGrayscaleAverage
        Case GrayscaleCustom
            FormGrayscale.fGrayscaleCustom pOPCODE
        Case GrayscaleCustomDither
            FormGrayscale.fGrayscaleCustomDither pOPCODE
        Case GrayscaleDecompose
            FormGrayscale.MenuDecompose pOPCODE
        Case GrayscaleSingleChannel
            FormGrayscale.MenuGrayscaleSingleChannel pOPCODE
        
    'Area filters
        Case Blur
            FilterBlur
        Case BlurMore
            FilterBlurMore
        Case Soften
            FilterSoften
        Case SoftenMore
            FilterSoftenMore
        Case Sharpen
            FilterSharpen
        Case SharpenMore
            FilterSharpenMore
        Case Unsharp
            If LoadForm Then
                FormUnsharpMask.Show vbModal, FormMain
            Else
                FormUnsharpMask.UnsharpMask CDbl(pOPCODE), CDbl(pOPCODE2), CLng(pOPCODE3)
            End If
        Case Diffuse
            FormDiffuse.DiffuseCustom 5, 5, False
        Case DiffuseMore
            FormDiffuse.DiffuseCustom 25, 25, False
        Case CustomDiffuse
            If LoadForm Then
                FormDiffuse.Show vbModal, FormMain
            Else
                FormDiffuse.DiffuseCustom pOPCODE, pOPCODE2, pOPCODE3
            End If
        Case Mosaic
            If LoadForm Then
                FormMosaic.Show vbModal, FormMain
            Else
                FormMosaic.MosaicFilter CInt(pOPCODE), CInt(pOPCODE2)
            End If
        Case MaximumRank
            If LoadForm Then
                FormMedian.showMedianDialog 100
            Else
                FormMedian.ApplyMedianFilter CLng(pOPCODE), CDbl(pOPCODE2)
            End If
        Case MinimumRank
            If LoadForm Then
                FormMedian.showMedianDialog 1
            Else
                FormMedian.ApplyMedianFilter CLng(pOPCODE), CDbl(pOPCODE2)
            End If
        Case GridBlur
            FilterGridBlur
        Case GaussianBlur
            If LoadForm Then
                FormGaussianBlur.Show vbModal, FormMain
            Else
                FormGaussianBlur.GaussianBlurFilter CDbl(pOPCODE)
            End If
        Case SmartBlur
            If LoadForm Then
                FormSmartBlur.Show vbModal, FormMain
            Else
                FormSmartBlur.SmartBlurFilter CDbl(pOPCODE), CByte(pOPCODE2), CBool(pOPCODE3)
            End If
        Case BoxBlur
            If LoadForm Then
                FormBoxBlur.Show vbModal, FormMain
            Else
                FormBoxBlur.BoxBlurFilter CLng(pOPCODE), CLng(pOPCODE2)
            End If
        
    'Edge filters
        Case EmbossToColor
            If LoadForm Then
                FormEmbossEngrave.Show vbModal, FormMain
            Else
                FormEmbossEngrave.FilterEmbossColor CLng(pOPCODE)
            End If
        Case EngraveToColor
            FormEmbossEngrave.FilterEngraveColor CLng(pOPCODE)
        Case Pencil
            FilterPencil
        Case Relief
            FilterRelief
        Case SmoothContour
            FormFindEdges.FilterSmoothContour pOPCODE
        Case PrewittHorizontal
            FormFindEdges.FilterPrewittHorizontal pOPCODE
        Case PrewittVertical
            FormFindEdges.FilterPrewittVertical pOPCODE
        Case SobelHorizontal
            FormFindEdges.FilterSobelHorizontal pOPCODE
        Case SobelVertical
            FormFindEdges.FilterSobelVertical pOPCODE
        Case Laplacian
            If LoadForm Then
                FormFindEdges.Show vbModal, FormMain
            Else
                FormFindEdges.FilterLaplacian pOPCODE
            End If
        Case HiliteEdge
            FormFindEdges.FilterHilite pOPCODE
        Case PhotoDemonEdgeLinear
            FormFindEdges.PhotoDemonLinearEdgeDetection pOPCODE
        Case PhotoDemonEdgeCubic
            FormFindEdges.PhotoDemonCubicEdgeDetection pOPCODE
        Case EdgeEnhance
            FilterEdgeEnhance
        Case Contour
            If LoadForm Then
                FormContour.Show vbModal, FormMain
            Else
                FormContour.TraceContour pOPCODE, CBool(pOPCODE2), CBool(pOPCODE3)
            End If
        
    'Color operations
        Case Rechannel
            If LoadForm Then
                FormRechannel.Show vbModal, FormMain
            Else
                FormRechannel.RechannelImage CLng(pOPCODE)
            End If
        Case ColorShiftLeft
            MenuCShift pOPCODE
        Case ColorShiftRight
            MenuCShift pOPCODE
        Case BrightnessAndContrast
            If LoadForm Then
                FormBrightnessContrast.Show vbModal, FormMain
            Else
                FormBrightnessContrast.BrightnessContrast CInt(pOPCODE), CSng(pOPCODE2), CBool(pOPCODE3)
            End If
        Case GammaCorrection
            If LoadForm Then
                FormGamma.Show vbModal, FormMain
            Else
                FormGamma.GammaCorrect CSng(pOPCODE), CSng(pOPCODE2), CSng(pOPCODE3)
            End If
        Case Invert
            MenuInvert
        Case CompoundInvert
            MenuCompoundInvert CLng(pOPCODE)
        Case Negative
            MenuNegative
        Case InvertHue
            MenuInvertHue
        Case AutoEnhance
            MenuAutoEnhanceContrast
        Case AutoHighlights
            MenuAutoEnhanceHighlights
        Case AutoMidtones
            MenuAutoEnhanceMidtones
        Case AutoShadows
            MenuAutoEnhanceShadows
        Case ImageLevels
            If LoadForm Then
                FormLevels.Show vbModal, FormMain
            Else
                FormLevels.MapImageLevels pOPCODE, pOPCODE2, pOPCODE3, pOPCODE4, pOPCODE5
            End If
        Case Colorize
            If LoadForm Then
                FormColorize.Show vbModal, FormMain
            Else
                FormColorize.ColorizeImage pOPCODE, pOPCODE2
            End If
        Case ReduceColors
            If LoadForm Then
                FormReduceColors.Show vbModal, FormMain
            Else
                If pOPCODE = REDUCECOLORS_AUTO Then
                    FormReduceColors.ReduceImageColors_Auto pOPCODE2
                ElseIf pOPCODE = REDUCECOLORS_MANUAL Then
                    FormReduceColors.ReduceImageColors_BitRGB pOPCODE2, pOPCODE3, pOPCODE4, pOPCODE5
                ElseIf pOPCODE = REDUCECOLORS_MANUAL_ERRORDIFFUSION Then
                    FormReduceColors.ReduceImageColors_BitRGB_ErrorDif pOPCODE2, pOPCODE3, pOPCODE4, pOPCODE5
                Else
                    pdMsgBox "Unsupported color reduction method.", vbCritical + vbOKOnly + vbApplicationModal, "Color reduction error"
                End If
            End If
        Case AdjustTemperature
            If LoadForm Then
                FormColorTemp.Show vbModal, FormMain
            Else
                FormColorTemp.ApplyTemperatureToImage pOPCODE, pOPCODE2, pOPCODE3
            End If
        Case AdjustHSL
            If LoadForm Then
                FormHSL.Show vbModal, FormMain
            Else
                FormHSL.AdjustImageHSL pOPCODE, pOPCODE2, pOPCODE3
            End If
        Case AdjustColorBalance
            If LoadForm Then
                FormColorBalance.Show vbModal, FormMain
            Else
                FormColorBalance.ApplyColorBalance CLng(pOPCODE), CLng(pOPCODE2), CLng(pOPCODE3), CBool(pOPCODE4)
            End If
        Case ShadowHighlight
            If LoadForm Then
                FormShadowHighlight.Show vbModal, FormMain
            Else
                FormShadowHighlight.ApplyShadowHighlight CDbl(pOPCODE), CDbl(pOPCODE2), CLng(pOPCODE3)
            End If
    
    'Coordinate filters/transformations
        Case Flip
            MenuFlip
        Case FreeRotate
            If LoadForm Then
                FormRotate.Show vbModal, FormMain
            Else
                FormRotate.RotateArbitrary pOPCODE, pOPCODE2
            End If
        Case Mirror
            MenuMirror
        Case Rotate90Clockwise
            MenuRotate90Clockwise
        Case Rotate180
            MenuRotate180
        Case Rotate270Clockwise
            MenuRotate270Clockwise
        Case Isometric
            FilterIsometric
        Case ImageSize
            If LoadForm Then
                FormResize.Show vbModal, FormMain
            Else
                FormResize.ResizeImage CLng(pOPCODE), CLng(pOPCODE2), CByte(pOPCODE3)
            End If
        Case Tile
            If LoadForm Then
                FormTile.Show vbModal, FormMain
            Else
                FormTile.GenerateTile CByte(pOPCODE), CLng(pOPCODE2), CLng(pOPCODE3)
            End If
        Case CropToSelection
            MenuCropToSelection
        Case ChangeImageMode24
            ConvertImageColorDepth 24
        Case ChangeImageMode32
            ConvertImageColorDepth 32
        Case DistortSwirl
            If LoadForm Then
                FormSwirl.Show vbModal, FormMain
            Else
                FormSwirl.SwirlImage CDbl(pOPCODE), CDbl(pOPCODE2), CLng(pOPCODE3), CBool(pOPCODE4)
            End If
        Case DistortLens
            If LoadForm Then
                FormLens.Show vbModal, FormMain
            Else
                FormLens.ApplyLensDistortion CDbl(pOPCODE), CDbl(pOPCODE2), CBool(pOPCODE3)
            End If
        Case DistortLensFix
            If LoadForm Then
                FormLensCorrect.Show vbModal, FormMain
            Else
                FormLensCorrect.ApplyLensCorrection CDbl(pOPCODE), CDbl(pOPCODE2), CDbl(pOPCODE3), CLng(pOPCODE4), CBool(pOPCODE5)
            End If
        Case DistortRipple
            If LoadForm Then
                FormRipple.Show vbModal, FormMain
            Else
                FormRipple.RippleImage CDbl(pOPCODE), CDbl(pOPCODE2), CDbl(pOPCODE3), CDbl(pOPCODE4), CLng(pOPCODE5), CBool(pOPCODE6)
            End If
        Case DistortPinchAndWhirl
            If LoadForm Then
                FormPinch.Show vbModal, FormMain
            Else
                FormPinch.PinchImage CDbl(pOPCODE), CDbl(pOPCODE2), CDbl(pOPCODE3), CLng(pOPCODE4), CBool(pOPCODE5)
            End If
        Case DistortWaves
            If LoadForm Then
                FormWaves.Show vbModal, FormMain
            Else
                FormWaves.WaveImage CDbl(pOPCODE), CDbl(pOPCODE2), CDbl(pOPCODE3), CDbl(pOPCODE4), CLng(pOPCODE5), CBool(pOPCODE6)
            End If
        Case DistortFiguredGlass
            If LoadForm Then
                FormFiguredGlass.Show vbModal, FormMain
            Else
                FormFiguredGlass.FiguredGlassFX CDbl(pOPCODE), CDbl(pOPCODE2), CLng(pOPCODE3), CBool(pOPCODE4)
            End If
        Case DistortKaleidoscope
            If LoadForm Then
                FormKaleidoscope.Show vbModal, FormMain
            Else
                FormKaleidoscope.KaleidoscopeImage CDbl(pOPCODE), CDbl(pOPCODE2), CDbl(pOPCODE3), CDbl(pOPCODE4), CBool(pOPCODE5)
            End If
        Case ConvertPolar
            If LoadForm Then
                FormPolar.Show vbModal, FormMain
            Else
                FormPolar.ConvertToPolar CLng(pOPCODE), CDbl(pOPCODE2), CLng(pOPCODE3), CBool(pOPCODE4)
            End If
        Case Autocrop
            AutocropImage
        Case DistortShear
            If LoadForm Then
                FormShear.Show vbModal, FormMain
            Else
                FormShear.ShearImage CDbl(pOPCODE), CDbl(pOPCODE2), CLng(pOPCODE3), CBool(pOPCODE4)
            End If
        Case DistortSquish
            If LoadForm Then
                FormSquish.Show vbModal, FormMain
            Else
                FormSquish.SquishImage CDbl(pOPCODE), CDbl(pOPCODE2), CLng(pOPCODE3), CBool(pOPCODE4)
            End If
        Case FreePerspective
            If LoadForm Then
                FormTruePerspective.Show vbModal, FormMain
            Else
                FormTruePerspective.PerspectiveImage CStr(pOPCODE), CLng(pOPCODE2), CBool(pOPCODE3)
            End If
        Case DistortPanAndZoom
            If LoadForm Then
                FormPanAndZoom.Show vbModal, FormMain
            Else
                FormPanAndZoom.PanAndZoomFilter CDbl(pOPCODE), CDbl(pOPCODE2), CDbl(pOPCODE3), CLng(pOPCODE4), CBool(pOPCODE5)
            End If
            
            
    'Other filters
        Case Antique
            MenuAntique
        Case Atmospheric
            MenuAtmospheric
        Case BlackLight
            If LoadForm Then
                FormBlackLight.Show vbModal, FormMain
            Else
                FormBlackLight.fxBlackLight pOPCODE
            End If
        Case Dream
            MenuDream
        Case Posterize
            If LoadForm Then
                FormPosterize.Show vbModal, FormMain
            Else
                FormPosterize.PosterizeImage CByte(pOPCODE)
            End If
        Case Radioactive
            MenuRadioactive
        Case Solarize
            If LoadForm Then
                FormSolarize.Show vbModal, FormMain
            Else
                FormSolarize.SolarizeImage CByte(pOPCODE)
            End If
        Case Twins
            If LoadForm Then
                FormTwins.Show vbModal, FormMain
            Else
                FormTwins.GenerateTwins CByte(pOPCODE)
            End If
        Case Fade
            If LoadForm Then
                FormFade.Show vbModal, FormMain
            Else
                FormFade.FadeImage CSng(pOPCODE)
            End If
        Case UnFade
            FormFade.UnfadeImage
        Case Alien
            MenuAlien
        Case Synthesize
            MenuSynthesize
        Case Water
            MenuWater
        Case Noise
            If LoadForm Then
                FormNoise.Show vbModal, FormMain
            Else
                FormNoise.AddNoise CInt(pOPCODE), CByte(pOPCODE2)
            End If
        Case Frozen
            MenuFrozen
        Case Lava
            MenuLava
        Case CustomFilter
            If LoadForm Then
                FormCustomFilter.Show vbModal, FormMain
            Else
                DoFilter , , pOPCODE
            End If
        Case Burn
            MenuBurn
        Case Steel
            MenuSteel
        Case FogEffect
            MenuFogEffect
        Case CountColors
            MenuCountColors
        Case Rainbow
            MenuRainbow
        Case Vibrate
            MenuVibrate
        Case Despeckle
            FormDespeckle.QuickDespeckle
        Case CustomDespeckle
            If LoadForm Then
                FormDespeckle.Show vbModal, FormMain
            Else
                FormDespeckle.Despeckle pOPCODE
            End If
        Case Sepia
            MenuSepia
        Case HeatMap
            MenuHeatMap
        Case ComicBook
            MenuComicBook
        Case FilmGrain
            If LoadForm Then
                FormFilmGrain.Show vbModal, FormMain
            Else
                FormFilmGrain.AddFilmGrain CLng(pOPCODE), CLng(pOPCODE2)
            End If
        Case FilmNoir
            MenuFilmNoir
        Case Vignetting
            If LoadForm Then
                FormVignette.Show vbModal, FormMain
            Else
                FormVignette.ApplyVignette CDbl(pOPCODE), CDbl(pOPCODE2), CDbl(pOPCODE3), CBool(pOPCODE4), CLng(pOPCODE5)
            End If
        Case Median
            If LoadForm Then
                FormMedian.showMedianDialog 50
            Else
                FormMedian.ApplyMedianFilter CLng(pOPCODE), CDbl(pOPCODE2)
            End If
        Case ModernArt
            If LoadForm Then
                FormModernArt.Show vbModal, FormMain
            Else
                FormModernArt.ApplyModernArt CLng(pOPCODE)
            End If
        
    End Select
    
    'If the user wants us to time this action, display the results now
    If (Not LoadForm) And (pType >= 100) Then
        If DISPLAY_TIMINGS Then Message "Time taken: " & Timer - m_ProcessingTime & " seconds"
    End If
    
    'Finally, check to see if the user wants us to Fade the last effect applied to the image...
    If pType = FadeLastEffect Then MenuFadeLastEffect
    
    'Restore the mouse pointer to its default value; if we are running a batch conversion, however, leave it busy
    ' The batch routine will handle restoring the cursor to normal.
    If MacroStatus <> MacroBATCH Then Screen.MousePointer = vbDefault
    
    'If the histogram form is visible and images are loaded, redraw the histogram
    If FormHistogram.Visible Then
        If NumOfWindows > 0 Then
            FormHistogram.TallyHistogramValues
            FormHistogram.DrawHistogram
        Else
            'If the histogram is visible but no images are open, unload the histogram
            Unload FormHistogram
        End If
    End If
    
    'If the image is potentially being changed and we are not performing a batch conversion (disabled to save speed!),
    ' redraw the active MDI child form icon.
    If (pType >= 101) And (MacroStatus <> MacroBATCH) And (LoadForm <> True) And (RecordAction <> False) And (pType <> CountColors) Then CreateCustomFormIcon FormMain.ActiveForm
    
    'Mark the processor as no longer busy and unlock the main form
    FormMain.Enabled = True
    
    'If a filter or tool was just used, return focus to the active form
    If (pType >= 101) And (MacroStatus <> MacroBATCH) And (LoadForm <> True) Then
        If NumOfWindows > 0 Then FormMain.ActiveForm.SetFocus
    End If
        
    'Also, re-enable drag and drop operations
    If pType >= 101 Then
        g_AllowDragAndDrop = True
        FormMain.OLEDropMode = 1
    End If
    
    Processing = False
    
    Exit Sub


'MAIN PHOTODEMON ERROR HANDLER STARTS HERE

MainErrHandler:

    'Reset the mouse pointer and access to the main form
    Screen.MousePointer = vbDefault
    FormMain.Enabled = True

    'We'll use this string to hold additional error data
    Dim AddInfo As String
    
    'This variable stores the message box type
    Dim mType As VbMsgBoxStyle
    
    'Tracks the user input from the message box
    Dim msgReturn As VbMsgBoxResult
    
    'Ignore errors that aren't actually errors
    If Err.Number = 0 Then Exit Sub
    
    'Object was unloaded before it could be shown - this is intentional, so ignore the error
    If Err.Number = 364 Then Exit Sub
        
    'Out of memory error
    If Err.Number = 480 Or Err.Number = 7 Then
        AddInfo = g_Language.TranslateMessage("There is not enough memory available to continue this operation.  Please free up system memory (RAM) by shutting down unneeded programs - especially your web browser, if it is open - then try the action again.")
        Message "Out of memory.  Function cancelled."
        mType = vbExclamation + vbOKOnly + vbApplicationModal
    
    'Invalid picture error
    ElseIf Err.Number = 481 Then
        AddInfo = g_Language.TranslateMessage("Unfortunately, this image file appears to be invalid.  This can happen if a file does not contain image data, or if it contains image data in an unsupported format." & vbCrLf & vbCrLf & "- If you downloaded this image from the Internet, the download may have terminated prematurely.  Please try downloading the image again." & vbCrLf & vbCrLf & "- If this image file came from a digital camera, scanner, or other image editing program, it's possible that PhotoDemon simply doesn't understand this particular file format.  Please save the image in a generic format (such as JPEG or PNG), then reload it.")
        Message "Invalid image.  Image load cancelled."
        mType = vbExclamation + vbOKOnly + vbApplicationModal
    
        'Since we know about this error, there's no need to display the extended box.  Display a smaller one, then exit.
        pdMsgBox AddInfo, mType, "Invalid image file"
        
        'On an invalid picture load, there will be a blank form that needs to be dealt with.
        pdImages(CurrentImage).deactivateImage
        Unload FormMain.ActiveForm
        Exit Sub
    
    'File not found error
    ElseIf Err.Number = 53 Then
        AddInfo = g_Language.TranslateMessage("The specified file could not be located.  If it was located on removable media, please re-insert the proper floppy disk, CD, or portable drive.  If the file is not located on portable media, make sure that:" & vbCrLf & "1) the file hasn't been deleted, and..." & vbCrLf & "2) the file location provided to PhotoDemon is correct.")
        Message "File not found."
        mType = vbExclamation + vbOKOnly + vbApplicationModal
        
    'Unknown error
    Else
        AddInfo = g_Language.TranslateMessage("PhotoDemon cannot locate additional information for this error.  That probably means this error is a bug, and it needs to be fixed!" & vbCrLf & vbCrLf & "Would you like to submit a bug report?  (It takes less than one minute, and it helps everyone who uses the software.)")
        mType = vbCritical + vbYesNo + vbApplicationModal
        Message "Unknown error."
    End If
    
    'Create the message box to return the error information
    msgReturn = pdMsgBox("PhotoDemon has experienced an error.  Details on the problem include:" & vbCrLf & vbCrLf & "Error number %1" & vbCrLf & "Description: %2" & vbCrLf & vbCrLf & "%3", mType, "PhotoDemon Error Handler", Err.Number, Err.Description, AddInfo)
    
    'If the message box return value is "Yes", the user has opted to file a bug report.
    If msgReturn = vbYes Then
    
        'GitHub requires a login for submitting Issues; check for that first
        Dim secondaryReturn As VbMsgBoxResult
    
        secondaryReturn = pdMsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, PhotoDemon needs you to answer one more question." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbQuestion + vbApplicationModal + vbYesNo, "Thanks for making PhotoDemon better")
    
        'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the tannerhelland.com contact form
        If secondaryReturn = vbYes Then
            'Shell a browser window with the GitHub issue report form
            OpenURL "https://github.com/tannerhelland/PhotoDemon/issues/new"
            
            'Display one final message box with additional instructions
            pdMsgBox "PhotoDemon has automatically opened a GitHub bug report webpage for you.  In the Title box, please enter the following error number with a short description of the problem: " & vbCrLf & "%1" & vbCrLf & vbCrLf & "Any additional details you can provide in the large text box, including the steps that led up to this error, will help it get fixed as quickly as possible." & vbCrLf & vbCrLf & "When finished, click the Submit New Issue button.  Thank you!", vbInformation + vbApplicationModal + vbOKOnly, "GitHub bug report instructions", Err.Number
            
        Else
            'Shell a browser window with the tannerhelland.com PhotoDemon contact form
            OpenURL "http://www.tannerhelland.com/photodemon-contact/"
            
            'Display one final message box with additional instructions
            pdMsgBox "PhotoDemon has automatically opened a bug report webpage for you.  In the Additional Details box, please describe the steps that led to this error." & vbCrLf & vbCrLf & "In the bottom box of that page, please enter the following error number: " & vbCrLf & "%1" & vbCrLf & vbCrLf & "When finished, click the Submit button.  Thank you!", vbInformation + vbApplicationModal + vbOKOnly, "Bug report instructions", Err.Number
            
        End If
    
    End If
        
End Sub

'Return a string with a human-readable name of a given process ID.
Public Function GetNameOfProcess(ByVal processID As Long) As String

    Select Case processID
    
        'Main functions (not used for image editing); numbers 1-99
        Case FileOpen
            GetNameOfProcess = "Open"
        Case FileSave
            GetNameOfProcess = "Save"
        Case FileSaveAs
            GetNameOfProcess = "Save As"
        Case capScreen
            GetNameOfProcess = "Screen Capture"
        Case cCopy
            GetNameOfProcess = "Copy"
        Case cPaste
            GetNameOfProcess = "Paste"
        Case cEmpty
            GetNameOfProcess = "Empty Clipboard"
        Case Undo
            GetNameOfProcess = "Undo"
        Case Redo
            GetNameOfProcess = "Redo"
        Case MacroStartRecording
            GetNameOfProcess = "Start Macro Recording"
        Case MacroStopRecording
            GetNameOfProcess = "Stop Macro Recording"
        Case MacroPlayRecording
            GetNameOfProcess = "Play Macro"
        Case SelectScanner
            GetNameOfProcess = "Select Scanner or Camera"
        Case ScanImage
            GetNameOfProcess = "Scan Image"
            
        'Histogram functions; numbers 100-199
        Case ViewHistogram
            GetNameOfProcess = "Display Histogram"
        Case StretchHistogram
            GetNameOfProcess = "Stretch Histogram"
        Case Equalize
            GetNameOfProcess = "Equalize"
        Case WhiteBalance
            GetNameOfProcess = "White Balance"
            
        'Black/White conversion; numbers 200-299
        Case BWMaster
            GetNameOfProcess = "Color to Monochrome"
        Case RemoveBW
            GetNameOfProcess = "Monochrome to Grayscale"
            
        'Grayscale conversion; numbers 300-399
        Case Desaturate
            GetNameOfProcess = "Desaturate"
        Case GrayScale
            GetNameOfProcess = "Grayscale (ITU Standard)"
        Case GrayscaleAverage
            GetNameOfProcess = "Grayscale (Average)"
        Case GrayscaleCustom
            GetNameOfProcess = "Grayscale (Custom # of Colors)"
        Case GrayscaleCustomDither
            GetNameOfProcess = "Grayscale (Custom Dither)"
        Case GrayscaleDecompose
            GetNameOfProcess = "Grayscale (Decomposition)"
        Case GrayscaleSingleChannel
            GetNameOfProcess = "Grayscale (Single Channel)"
        
        'Area filters; numbers 400-499
        Case Blur
            GetNameOfProcess = "Blur"
        Case BlurMore
            GetNameOfProcess = "Blur More"
        Case Soften
            GetNameOfProcess = "Soften"
        Case SoftenMore
            GetNameOfProcess = "Soften More"
        Case Sharpen
            GetNameOfProcess = "Sharpen"
        Case SharpenMore
            GetNameOfProcess = "Sharpen More"
        Case Unsharp
            GetNameOfProcess = "Unsharp"
        Case Diffuse
            GetNameOfProcess = "Diffuse"
        Case DiffuseMore
            GetNameOfProcess = "Diffuse More"
        Case CustomDiffuse
            GetNameOfProcess = "Custom Diffuse"
        Case Mosaic
            GetNameOfProcess = "Mosaic"
        Case MaximumRank
            GetNameOfProcess = "Dilate (maximum rank)"
        Case MinimumRank
            GetNameOfProcess = "Erode (minimum rank)"
        Case GridBlur
            GetNameOfProcess = "Grid Blur"
        Case GaussianBlur
            GetNameOfProcess = "Gaussian Blur"
        Case SmartBlur
            GetNameOfProcess = "Smart Blur"
        Case BoxBlur
            GetNameOfProcess = "Box Blur"
        
        'Edge filters; numbers 500-599
        Case EmbossToColor
            GetNameOfProcess = "Emboss"
        Case EngraveToColor
            GetNameOfProcess = "Engrave"
        Case Pencil
            GetNameOfProcess = "Pencil Drawing"
        Case Relief
            GetNameOfProcess = "Relief"
        Case PrewittHorizontal
            GetNameOfProcess = "Find Edges (Prewitt Horizontal)"
        Case PrewittVertical
            GetNameOfProcess = "Find Edges (Prewitt Vertical)"
        Case SobelHorizontal
            GetNameOfProcess = "Find Edges (Sobel Horizontal)"
        Case SobelVertical
            GetNameOfProcess = "Find Edges (Sobel Vertical)"
        Case Laplacian
            GetNameOfProcess = "Find Edges (Laplacian)"
        Case SmoothContour
            GetNameOfProcess = "Artistic Contour"
        Case HiliteEdge
            GetNameOfProcess = "Find Edges (Hilite)"
        Case PhotoDemonEdgeLinear
            GetNameOfProcess = "Find Edges (PhotoDemon Linear)"
        Case PhotoDemonEdgeCubic
            GetNameOfProcess = "Find Edges (PhotoDemon Cubic)"
        Case EdgeEnhance
            GetNameOfProcess = "Edge Enhance"
        Case Contour
            GetNameOfProcess = "Trace Contour"
            
        'Color operations; numbers 600-699
        Case Rechannel
            GetNameOfProcess = "Rechannel"
        'Rechannel Green and Red are only included for legacy reasons
        Case RechannelGreen
            GetNameOfProcess = "Rechannel (Green)"
        Case RechannelRed
            GetNameOfProcess = "Rechannel (Red)"
        '-------
        Case ColorShiftLeft
            GetNameOfProcess = "Shift Colors (Left)"
        Case ColorShiftRight
            GetNameOfProcess = "Shift Colors (Right)"
        Case BrightnessAndContrast
            GetNameOfProcess = "Brightness/Contrast"
        Case GammaCorrection
            GetNameOfProcess = "Gamma Correction"
        Case Invert
            GetNameOfProcess = "Invert Colors"
        Case InvertHue
            GetNameOfProcess = "Invert Hue"
        Case Negative
            GetNameOfProcess = "Film Negative"
        Case CompoundInvert
            GetNameOfProcess = "Compound Invert"
        Case AutoEnhance
            GetNameOfProcess = "Auto-Enhance Contrast"
        Case AutoHighlights
            GetNameOfProcess = "Auto-Enhance Highlights"
        Case AutoMidtones
            GetNameOfProcess = "Auto-Enhance Midtones"
        Case AutoShadows
            GetNameOfProcess = "Auto-Enhance Shadows"
        Case ImageLevels
            GetNameOfProcess = "Image Levels"
        Case Colorize
            GetNameOfProcess = "Colorize"
        Case ReduceColors
            GetNameOfProcess = "Reduce Colors"
        Case AdjustTemperature
            GetNameOfProcess = "Color Temperature"
        Case AdjustHSL
            GetNameOfProcess = "Hue/Saturation/Lightness"
        Case AdjustColorBalance
            GetNameOfProcess = "Color Balance"
        Case ShadowHighlight
            GetNameOfProcess = "Shadow/Highlight"
            
        'Coordinate filters/transformations; numbers 700-799
        Case ImageSize
            GetNameOfProcess = "Resize"
        Case Flip
            GetNameOfProcess = "Flip"
        Case Mirror
            GetNameOfProcess = "Mirror"
        Case Rotate90Clockwise
            GetNameOfProcess = "Rotate 90° Clockwise"
        Case Rotate180
            GetNameOfProcess = "Rotate 180°"
        Case Rotate270Clockwise
            GetNameOfProcess = "Rotate 90° Counter-Clockwise"
        Case FreeRotate
            GetNameOfProcess = "Arbitrary Rotation"
        Case Isometric
            GetNameOfProcess = "Isometric Conversion"
        Case Tile
            GetNameOfProcess = "Tile Image"
        Case CropToSelection
            GetNameOfProcess = "Crop"
        Case ChangeImageMode24
            GetNameOfProcess = "Convert to Photo Mode (RGB, 24bpp)"
        Case ChangeImageMode32
            GetNameOfProcess = "Convert to Web Mode (RGBA, 32bpp)"
        Case DistortSwirl
            GetNameOfProcess = "Swirl"
        Case DistortLens
            GetNameOfProcess = "Apply lens distortion"
        Case DistortLensFix
            GetNameOfProcess = "Correct lens distortion"
        Case DistortRipple
            GetNameOfProcess = "Ripple"
        Case DistortPinchAndWhirl
            GetNameOfProcess = "Pinch and whirl"
        Case DistortWaves
            GetNameOfProcess = "Waves"
        Case DistortFiguredGlass
            GetNameOfProcess = "Figured glass"
        Case DistortKaleidoscope
            GetNameOfProcess = "Kaleidoscope"
        Case ConvertPolar
            GetNameOfProcess = "Polar conversion"
        Case Autocrop
            GetNameOfProcess = "Autocrop image"
        Case DistortShear
            GetNameOfProcess = "Shear"
        Case DistortSquish
            GetNameOfProcess = "Squish"
        Case FreePerspective
            GetNameOfProcess = "Perspective (free)"
        Case DistortPanAndZoom
            GetNameOfProcess = "Pan and Zoom"
            
        'Miscellaneous filters; numbers 800-899
        Case Fade
            GetNameOfProcess = "Fade"
        Case UnFade
            GetNameOfProcess = "Unfade"
        Case Atmospheric
            GetNameOfProcess = "Atmosphere"
        Case Frozen
            GetNameOfProcess = "Freeze"
        Case Lava
            GetNameOfProcess = "Lava"
        Case Burn
            GetNameOfProcess = "Burn"
        Case Water
            GetNameOfProcess = "Water"
        Case Steel
            GetNameOfProcess = "Steel"
        Case Dream
            GetNameOfProcess = "Dream"
        Case Alien
            GetNameOfProcess = "Alien"
        Case CustomFilter
            GetNameOfProcess = "Custom Filter"
        Case Antique
            GetNameOfProcess = "Antique"
        Case BlackLight
            GetNameOfProcess = "Blacklight"
        Case Posterize
            GetNameOfProcess = "Posterize"
        Case Radioactive
            GetNameOfProcess = "Radioactive"
        Case Solarize
            GetNameOfProcess = "Solarize"
        Case Twins
            GetNameOfProcess = "Generate Twins"
        Case Synthesize
            GetNameOfProcess = "Synthesize"
        Case Noise
            GetNameOfProcess = "Add Noise"
        Case CountColors
            GetNameOfProcess = "Count Image Colors"
        Case FogEffect
            GetNameOfProcess = "Fog"
        Case Rainbow
            GetNameOfProcess = "Rainbow"
        Case Vibrate
            GetNameOfProcess = "Vibrate"
        Case Despeckle
            GetNameOfProcess = "Despeckle"
        Case CustomDespeckle
            GetNameOfProcess = "Custom Despeckle"
        Case ComicBook
            GetNameOfProcess = "Comic book"
        Case Sepia
            GetNameOfProcess = "Sepia"
        Case HeatMap
            GetNameOfProcess = "Thermograph (Heat Map)"
        Case FilmGrain
            GetNameOfProcess = "Film Grain"
        
        Case LastCommand
            GetNameOfProcess = "Repeat Last Action"
        Case FadeLastEffect
            GetNameOfProcess = "Fade last effect"
            
        Case SelectionCreate
            GetNameOfProcess = "Create New Selection"
        Case SelectionClear
            GetNameOfProcess = "Clear Active Selection"
        Case Vignetting
            GetNameOfProcess = "Vignetting"
        Case Median
            GetNameOfProcess = "Median filter"
        Case ModernArt
            GetNameOfProcess = "Modern art"
            
        'This "Else" statement should never trigger, but if it does, return an empty string
        Case Else
            GetNameOfProcess = ""
            
    End Select
    
    GetNameOfProcess = g_Language.TranslateMessage(GetNameOfProcess)
    
End Function
