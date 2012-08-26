Attribute VB_Name = "Processor"
'***************************************************************************
'Program Sub-Processor and Error Handler
'Copyright ©2000-2012 by Tanner Helland
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
    Public Const cScreen As Long = 10
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
    Public Const EqualizeLuminance As Long = 103
    Public Const WhiteBalance As Long = 104
    
    'Black/White conversion; numbers 200-299
    Public Const BWImpressionist As Long = 200
    Public Const BWNearestColor As Long = 201
    Public Const BWNearestColor2 As Long = 202
    Public Const BWOrderedDither As Long = 203
    Public Const BWDiffusionDither As Long = 204
    Public Const Threshold As Long = 205
    Public Const ComicBook As Long = 206
    Public Const BWEnhancedDither As Long = 207
    Public Const BWFloydSteinberg As Long = 208
    
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
    Public Const RankMaximum As Long = 411
    Public Const RankMinimum As Long = 412
    Public Const RankExtreme As Long = 413
    Public Const CustomRank As Long = 414
    '-Grid Blurring
    Public Const GridBlur As Long = 415
    '-Gaussian Blur
    Public Const GaussianBlur As Long = 416
    Public Const GaussianBlurMore As Long = 417
    '-Antialias
    Public Const Antialias As Long = 418
    
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
    
    'Color operations; numbers 600-699
    '-Rechanneling
    Public Const RechannelBlue As Long = 600
    Public Const RechannelGreen As Long = 601
    Public Const RechannelRed As Long = 602
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
    'Tiling
    Public Const Tile As Long = 708
    
    'Other filters; numbers 800-899
    '-Compound invert
    Public Const DarkCompoundInvert As Long = 800
    Public Const LightCompoundInvert As Long = 801
    Public Const MediumCompoundInvert As Long = 802
    '-Fade
    Public Const Fade As Long = 803
    '804-806 used to be specific Fade values; these have been superceded by passing the values to Fade
    Public Const Unfade As Long = 807
    '-Natural
    Public Const Atmospheric As Long = 808
    Public Const Frozen As Long = 809
    Public Const Lava As Long = 810
    Public Const Burn As Long = 811
    Public Const Ocean As Long = 812
    Public Const Water As Long = 813
    Public Const Steel As Long = 814
    Public Const FogEffect As Long = 828
    Public Const Rainbow As Long = 829
    '-Custom filters
    Public Const CustomFilter As Long = 817
    '-Miscellaneous
    Public Const Antique As Long = 818
    Public Const BlackLight As Long = 819
    Public Const Posterize As Long = 820
    Public Const Radioactive As Long = 821
    Public Const Solarize As Long = 822
    Public Const Twins As Long = 823
    Public Const Synthesize As Long = 824
    Public Const Noise As Long = 825
    'Public Const ??? As Long = 826   'This value is free for use - the filter originally here has been removed
    Public Const CountColors As Long = 827
    Public Const Dream As Long = 815
    Public Const Alien As Long = 816
    Public Const Vibrate As Long = 830
    Public Const Despeckle As Long = 831
    Public Const CustomDespeckle As Long = 832
    Public Const Animate As Long = 840
    
    'Relative processes
    Public Const LastCommand As Long = 900
    Public Const FadeLastEffect As Long = 901
    
    'Other filters end at 840

    'Reserved bytes; 1000 and up
    
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

'PhotoDemon's software processor.  Almost every action the program takes is routed through this method.  This is what
' allows us to record and playback macros, among other things.  (See comment at top of page for more details.)
Public Sub Process(ByVal pType As Long, Optional pOPCODE As Variant = 0, Optional pOPCODE2 As Variant = 0, Optional pOPCODE3 As Variant = 0, Optional pOPCODE4 As Variant = 0, Optional pOPCODE5 As Variant = 0, Optional pOPCODE6 As Variant = 0, Optional pOPCODE7 As Variant = 0, Optional pOPCODE8 As Variant = 0, Optional pOPCODE9 As Variant = 0, Optional LoadForm As Boolean = False, Optional RecordAction As Boolean = True)

    'Main error handler for the entire program is initialized by this line
    On Error GoTo MainErrHandler
    
    'If desired, this line can be used to artificially raise errors (to test the error handler)
    'Err.Raise 339
    
    'Mark the software processor as busy
    Processing = True
    
    'Set the mouse cursor to an hourglass
    FormMain.MousePointer = vbHourglass
    
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
            Case cScreen
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
    
        'Get the image data (to get image size and information)
        GetImageData
        
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
        
    End If
    
    'Histogram functions
    If pType >= 100 And pType <= 199 Then
        Select Case pType
            Case ViewHistogram
                FormHistogram.Show 0
            Case StretchHistogram
                FormHistogram.StretchHistogram
            Case Equalize
                FormHistogram.EqualizeHistogram pOPCODE, pOPCODE2, pOPCODE3
            Case EqualizeLuminance
                FormHistogram.EqualizeLuminance
            Case WhiteBalance
                If LoadForm = True Then
                    FormWhiteBalance.Show 1, FormMain
                Else
                    FormWhiteBalance.AutoWhiteBalance pOPCODE
                End If
        End Select
    End If
    
    'Black/White conversion
    If pType >= 200 And pType <= 299 Then
        Select Case pType
            Case BWImpressionist
                If LoadForm = True Then
                    FormBlackAndWhite.Show 1, FormMain
                Else
                    MenuBWImpressionist
                End If
            Case BWNearestColor
                MenuBWNearestColor
            Case BWNearestColor2
                MenuBWNearestColor2
            Case BWOrderedDither
                MenuBWOrderedDither
            Case BWDiffusionDither
                MenuBWDiffusionDither
            Case Threshold
                MenuThreshold pOPCODE
            Case ComicBook
                MenuComicBook
            Case BWEnhancedDither
                MenuBWEnhancedDither
            Case BWFloydSteinberg
                MenuBWFloydSteinberg
        End Select
    End If
    
    'Grayscale conversion
    If pType >= 300 And pType <= 399 Then
        Select Case pType
            Case Desaturate
                FormGrayscale.MenuDesaturate
            Case GrayScale
                If LoadForm = True Then
                    FormGrayscale.Show 1, FormMain
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
        End Select
    End If
    
    'Area filters
    If pType >= 400 And pType <= 499 Then
        Select Case pType
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
                FilterUnsharp
            Case Diffuse
                FormDiffuse.Diffuse
            Case DiffuseMore
                FormDiffuse.DiffuseMore
            Case CustomDiffuse
                If LoadForm = True Then
                    FormDiffuse.Show 1, FormMain
                Else
                    FormDiffuse.DiffuseCustom pOPCODE, pOPCODE2
                End If
            Case Mosaic
                If LoadForm = True Then
                    FormMosaic.Show 1, FormMain
                Else
                    FormMosaic.MosaicFilter CInt(pOPCODE), CInt(pOPCODE2)
                End If
            Case RankMaximum
                FormRank.rMaximize
            Case RankMinimum
                FormRank.rMinimize
            Case RankExtreme
                FormRank.rExtreme
            Case CustomRank
                If LoadForm = True Then
                    FormRank.Show 1, FormMain
                Else
                    FormRank.CustomRankFilter CInt(pOPCODE), CByte(pOPCODE2)
                End If
            Case GridBlur
                FilterGridBlur
            Case Antialias
                FilterAntialias
            Case GaussianBlur
                FilterGaussianBlur
            Case GaussianBlurMore
                FilterGaussianBlurMore
        End Select
    End If
    
    'Edge filters
    If pType >= 500 And pType <= 599 Then
        Select Case pType
            Case EmbossToColor
                If LoadForm = True Then
                    FormEmbossEngrave.Show 1, FormMain
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
                FormFindEdges.FilterSmoothContour
            Case PrewittHorizontal
                FormFindEdges.FilterPrewittHorizontal
            Case PrewittVertical
                FormFindEdges.FilterPrewittVertical
            Case SobelHorizontal
               FormFindEdges.FilterSobelHorizontal
            Case SobelVertical
                FormFindEdges.FilterSobelVertical
            Case Laplacian
                If LoadForm = True Then
                    FormFindEdges.Show 1, FormMain
                Else
                    FormFindEdges.FilterLaplacian
                End If
            Case HiliteEdge
                FormFindEdges.FilterHilite
            Case PhotoDemonEdgeLinear
                FormFindEdges.PhotoDemonLinearEdgeDetection
            Case PhotoDemonEdgeCubic
                FormFindEdges.PhotoDemonCubicEdgeDetection
            Case EdgeEnhance
                FilterEdgeEnhance
        End Select
    End If
    
    'Color operations
    If pType >= 600 And pType <= 699 Then
        Select Case pType
            Case RechannelBlue
                MenuRechannel pOPCODE
            Case RechannelGreen
                MenuRechannel pOPCODE
            Case RechannelRed
                MenuRechannel pOPCODE
            Case ColorShiftLeft
                MenuCShift pOPCODE
            Case ColorShiftRight
                MenuCShift pOPCODE
            Case BrightnessAndContrast
                If LoadForm = True Then
                    FormBrightnessContrast.Show 1, FormMain
                Else
                    FormBrightnessContrast.BrightnessContrast CInt(pOPCODE), CSng(pOPCODE2), CBool(pOPCODE3)
                End If
            Case GammaCorrection
                If LoadForm = True Then
                    FormGamma.Show 1, FormMain
                Else
                    FormGamma.GammaCorrect CSng(pOPCODE), CByte(pOPCODE2)
                End If
            Case Invert
                MenuInvert
            Case AutoEnhance
                MenuAutoEnhanceContrast
            Case AutoHighlights
                MenuAutoEnhanceHighlights
            Case AutoMidtones
                MenuAutoEnhanceMidtones
            Case AutoShadows
                MenuAutoEnhanceShadows
            Case Negative
                MenuNegative
            Case InvertHue
                MenuInvertHue
            Case ImageLevels
                If LoadForm = True Then
                    FormImageLevels.Show 1, FormMain
                Else
                    FormImageLevels.MapImageLevels pOPCODE, pOPCODE2, pOPCODE3, pOPCODE4, pOPCODE5
                End If
            Case Colorize
                If LoadForm = True Then
                    FormColorize.Show 1, FormMain
                Else
                    FormColorize.ColorizeImage pOPCODE
                End If
            Case ReduceColors
                If LoadForm = True Then
                    FormReduceColors.Show 1, FormMain
                Else
                    If pOPCODE = REDUCECOLORS_AUTO Then
                        FormReduceColors.ReduceImageColors_Auto pOPCODE2
                    ElseIf pOPCODE = REDUCECOLORS_MANUAL Then
                        FormReduceColors.ReduceImageColors_BitRGB pOPCODE2, pOPCODE3, pOPCODE4, pOPCODE5
                    ElseIf pOPCODE = REDUCECOLORS_MANUAL_ERRORDIFFUSION Then
                        FormReduceColors.ReduceImageColors_BitRGB_ErrorDif pOPCODE2, pOPCODE3, pOPCODE4, pOPCODE5
                    Else
                        MsgBox "Unsupported color reduction method."
                    End If
                End If
        End Select
    End If
    
    'Coordinate filters/transformations
    If pType >= 700 And pType <= 799 Then
        Select Case pType
            Case Flip
                MenuFlip
            'Case FreeRotate
            '    FormRotate.Visible = LoadForm
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
                If LoadForm = True Then
                    FormResize.Show 1, FormMain
                Else
                    FormResize.ResizeImage CLng(pOPCODE), CLng(pOPCODE2), CByte(pOPCODE3)
                End If
            Case Tile
                If LoadForm = True Then
                    FormTile.Show 1, FormMain
                Else
                    FormTile.GenerateTile CByte(pOPCODE), CLng(pOPCODE2), CLng(pOPCODE3)
                End If
        End Select
    End If
    
    'Other filters
    If pType >= 800 And pType <= 899 Then
        Select Case pType
            Case Antique
                MenuAntique
            Case Atmospheric
                MenuAtmospheric
            Case BlackLight
                If LoadForm = True Then
                    FormBlackLight.Show 1, FormMain
                Else
                    FormBlackLight.fxBlackLight pOPCODE
                End If
            Case DarkCompoundInvert
                MenuCompoundInvert pOPCODE
            Case Dream
                MenuDream
            Case LightCompoundInvert
                MenuCompoundInvert pOPCODE
            Case MediumCompoundInvert
                MenuCompoundInvert pOPCODE
            Case Posterize
                If LoadForm = True Then
                    FormPosterize.Show 1, FormMain
                Else
                    FormPosterize.PosterizeImage CByte(pOPCODE)
                End If
            Case Radioactive
                MenuRadioactive
            Case Solarize
                If LoadForm = True Then
                    FormSolarize.Show 1, FormMain
                Else
                    FormSolarize.SolarizeImage CByte(pOPCODE)
                End If
            Case Twins
                If LoadForm = True Then
                    FormTwins.Show 1, FormMain
                Else
                    FormTwins.GenerateTwins CByte(pOPCODE)
                End If
            Case Fade
                If LoadForm = True Then
                    FormFade.Show 1, FormMain
                Else
                    FormFade.FadeImage CInt(pOPCODE)
                End If
            Case Unfade
                FormFade.UnfadeImage
            Case Alien
                MenuAlien
            Case Synthesize
                MenuSynthesize
            Case Water
                MenuWater
            Case Noise
                If LoadForm = True Then
                    FormNoise.Show 1, FormMain
                Else
                    FormNoise.AddNoise CInt(pOPCODE), CByte(pOPCODE2)
                End If
            Case Frozen
                MenuFrozen
            Case Lava
                MenuLava
            Case CustomFilter
                If LoadForm = True Then
                    FormCustomFilter.Show 1, FormMain
                Else
                    FormCustomFilter.DoCustomFilterFromFile pOPCODE
                End If
            Case Burn
                MenuBurn
            Case Ocean
                MenuOcean
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
                If LoadForm = True Then
                    FormDespeckle.Show 1, FormMain
                Else
                    FormDespeckle.Despeckle pOPCODE
                End If
            Case Animate
                MenuAnimate
        
        End Select
    End If
    
    'Finally, check to see if the user wants us to fade the last effect applied to the image...
    If pType = FadeLastEffect Then MenuFadeLastEffect
    
    'Restore the mouse pointer to its default value; if we are running a batch conversion, however, leave it busy
    ' The batch routine will handle restoring the cursor to normal.
    If MacroStatus <> MacroBATCH Then FormMain.MousePointer = vbDefault
    
    'If the histogram form is visible and images are loaded, redraw the histogram
    If FormHistogram.Visible = True Then
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
    
    'Mark the processor as no longer busy
    Processing = False
    
    Exit Sub


'MAIN PHOTODEMON ERROR HANDLER STARTS HERE

MainErrHandler:

    'Reset the mouse pointer
    FormMain.MousePointer = vbDefault

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
        AddInfo = "There is not enough memory available to continue this operation.  Please free up system memory (RAM) by shutting down unneeded programs - especially your web browser, if it is open - then try the action again."
        Message "Out of memory.  Function cancelled."
        mType = vbCritical + vbOKOnly + vbApplicationModal
    
    'Invalid picture error
    ElseIf Err.Number = 481 Then
        AddInfo = "You have attempted to load an invalid picture.  This can happen if a file does not contain image data, or if it contains image data in an unsupported format." & vbCrLf & vbCrLf & "- If you downloaded this image from the Internet, the download may have terminated prematurely.  Please try downloading the image again." & vbCrLf & vbCrLf & "- If this image file came from a digital camera, scanner, or other image editing program, it's possible that " & PROGRAMNAME & " simply doesn't understand this particular file format.  Please save the image in a generic format (such as bitmap or JPEG), then reload it."
        Message "Invalid picture.  Image load cancelled."
        mType = vbCritical + vbOKOnly + vbApplicationModal
    
    'File not found error
    ElseIf Err.Number = 53 Then
        AddInfo = "The specified file could not be located.  If it was located on removable media, please re-insert the proper floppy disk, CD, or portable drive.  If the file is not located on portable media, make sure that:" & vbCrLf & "1) the file hasn't been deleted, and..." & "2) the file location provided to " & PROGRAMNAME & " is correct."
        Message "File not found."
        mType = vbCritical + vbOKOnly + vbApplicationModal
        
    'Unknown error
    Else
        AddInfo = PROGRAMNAME & " cannot locate additional information for this error.  That probably means this error is a bug, and it needs to be fixed!" & vbCrLf & vbCrLf & "Would you like to submit a bug report?  (It takes less than one minute, and it helps everyone who uses " & PROGRAMNAME & ".)"
        mType = vbCritical + vbYesNo + vbApplicationModal
        Message "Unknown error."
    End If
    
    'Create the message box to return the error information
    msgReturn = MsgBox(PROGRAMNAME & " has experienced an error.  Details on the problem include:" & vbCrLf & vbCrLf & _
    "Error number " & Err.Number & vbCrLf & _
    "Description: " & Err.Description & vbCrLf & vbCrLf & _
    AddInfo, mType, PROGRAMNAME & " Error Handler: #" & Err.Number)
    
    'If the message box return value is "Yes", the user has opted to file a bug report.
    If msgReturn = vbYes Then
    
        'GitHub requires a login for submitting Issues; check for that first
        Dim secondaryReturn As VbMsgBoxResult
    
        secondaryReturn = MsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, PhotoDemon needs you to answer one more question." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbQuestion + vbApplicationModal + vbYesNo, "Thanks for making " & PROGRAMNAME & " better")
    
        'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the tannerhelland.com contact form
        If secondaryReturn = vbYes Then
            'Shell a browser window with the GitHub issue report form
            ShellExecute FormMain.HWnd, "Open", "https://github.com/tannerhelland/PhotoDemon/issues/new", "", 0, SW_SHOWNORMAL
            
            'Display one final message box with additional instructions
            MsgBox "PhotoDemon has automatically opened a GitHub bug report webpage for you.  In the ""Title"" box, please enter the following error number with a short description of the problem: " & vbCrLf & Err.Number & vbCrLf & vbCrLf & "Any additional details you can provide in the large text box, including the steps that led up to this error, will help it get fixed as quickly as possible." & vbCrLf & vbCrLf & "When finished, click the ""Submit new issue"" button.  Thank you so much for your help!", vbInformation + vbApplicationModal + vbOKOnly, "GitHub bug report instructions"
            
        Else
            'Shell a browser window with the tannerhelland.com PhotoDemon contact form
            ShellExecute FormMain.HWnd, "Open", "http://www.tannerhelland.com/photodemon-contact/", "", 0, SW_SHOWNORMAL
            
            'Display one final message box with additional instructions
            MsgBox "PhotoDemon has automatically opened a bug report webpage for you.  In the ""Additional details"" box, please describe the steps that led up to this error." & vbCrLf & vbCrLf & "In the bottom box of that page, please enter the following error number: " & vbCrLf & Err.Number & vbCrLf & vbCrLf & "When finished, click the ""Submit"" button.  Thank you so much for your help!", vbInformation + vbApplicationModal + vbOKOnly, "Bug report instructions"
            
        End If
    
    End If
        
    'If an invalid picture was loaded, unload the active form (since it will just be empty and pictureless)
    If Err.Number = 481 Then Unload FormMain.ActiveForm

End Sub

'Return a string with a human-readable name of a given process ID.
Public Function GetNameOfProcess(ByVal processID As Long) As String

    Select Case processID
    
        'Main functions (not used for image editing); numbers 1-99
        Case 1
            GetNameOfProcess = "Open"
        Case 2
            GetNameOfProcess = "Save"
        Case 3
            GetNameOfProcess = "Save As"
        Case 10
            GetNameOfProcess = "Screen Capture"
        Case 20
            GetNameOfProcess = "Copy"
        Case 21
            GetNameOfProcess = "Paste"
        Case 22
            GetNameOfProcess = "Empty Clipboard"
        Case 30
            GetNameOfProcess = "Undo"
        Case 31
            GetNameOfProcess = "Redo"
        Case 40
            GetNameOfProcess = "Start Macro Recording"
        Case 41
            GetNameOfProcess = "Stop Macro Recording"
        Case 42
            GetNameOfProcess = "Play Macro"
        Case 50
            GetNameOfProcess = "Select Scanner or Camera"
        Case 51
            GetNameOfProcess = "Scan Image"
            
        'Histogram functions; numbers 100-199
        Case 100
            GetNameOfProcess = "Display Histogram"
        Case 101
            GetNameOfProcess = "Stretch Histogram"
        Case 102
            GetNameOfProcess = "Equalize"
        Case 103
            GetNameOfProcess = "Equalize Luminance"
        Case 104
            GetNameOfProcess = "White Balance"
            
        'Black/White conversion; numbers 200-299
        Case 200
            GetNameOfProcess = "Black and White (Impressionist)"
        Case 201
            GetNameOfProcess = "Black and White (Nearest Color)"
        Case 202
            GetNameOfProcess = "Black and White (Component Color)"
        Case 203
            GetNameOfProcess = "Black and White (Ordered Dither)"
        Case 204
            GetNameOfProcess = "Black and White (Diffusion Dither)"
        Case 205
            GetNameOfProcess = "Black and White (Threshold)"
        Case 206
            GetNameOfProcess = "Comic Book"
        Case 207
            GetNameOfProcess = "Black and White (Santos Enhanced)"
        Case 208
            GetNameOfProcess = "Black and White (Floyd-Steinberg)"
            
        'Grayscale conversion; numbers 300-399
        Case 300
            GetNameOfProcess = "Desaturate"
        Case 301
            GetNameOfProcess = "Grayscale (ITU Standard)"
        Case 302
            GetNameOfProcess = "Grayscale (Average)"
        Case 303
            GetNameOfProcess = "Grayscale (Custom # of Colors)"
        Case 304
            GetNameOfProcess = "Grayscale (Custom Dither)"
        Case 305
            GetNameOfProcess = "Grayscale (Decomposition)"
        Case 306
            GetNameOfProcess = "Grayscale (Single Channel)"
        
        'Area filters; numbers 400-499
        Case 400
            GetNameOfProcess = "Blur"
        Case 401
            GetNameOfProcess = "Blur More"
        Case 402
            GetNameOfProcess = "Soften"
        Case 403
            GetNameOfProcess = "Soften More"
        Case 404
            GetNameOfProcess = "Sharpen"
        Case 405
            GetNameOfProcess = "Sharpen More"
        Case 406
            GetNameOfProcess = "Unsharp"
        Case 407
            GetNameOfProcess = "Diffuse"
        Case 408
            GetNameOfProcess = "Diffuse More"
        Case 409
            GetNameOfProcess = "Custom Diffuse"
        Case 410
            GetNameOfProcess = "Mosaic"
        Case 411
            GetNameOfProcess = "Dilate"
        Case 412
            GetNameOfProcess = "Erode"
        Case 413
            GetNameOfProcess = "Extreme Rank"
        Case 414
            GetNameOfProcess = "Custom Rank"
        Case 415
            GetNameOfProcess = "Grid Blur"
        Case 416
            GetNameOfProcess = "Gaussian Blur"
        Case 417
            GetNameOfProcess = "Gaussian Blur More"
        Case 418
            GetNameOfProcess = "Antialias"
    
        'Edge filters; numbers 500-599
        Case 500
            GetNameOfProcess = "Emboss"
        Case 501
            GetNameOfProcess = "Engrave"
        Case 504
            GetNameOfProcess = "Pencil Drawing"
        Case 505
            GetNameOfProcess = "Relief"
        Case 506
            GetNameOfProcess = "Find Edges (Prewitt Horizontal)"
        Case 507
            GetNameOfProcess = "Find Edges (Prewitt Vertical)"
        Case 508
            GetNameOfProcess = "Find Edges (Sobel Horizontal)"
        Case 509
            GetNameOfProcess = "Find Edges (Sobel Vertical)"
        Case 510
            GetNameOfProcess = "Find Edges (Laplacian)"
        Case 511
            GetNameOfProcess = "Artistic Contour"
        Case 512
            GetNameOfProcess = "Find Edges (Hilite)"
        Case 513
            GetNameOfProcess = "Find Edges (PhotoDemon Linear)"
        Case 514
            GetNameOfProcess = "Find Edges (PhotoDemon Cubic)"
        Case 515
            GetNameOfProcess = "Edge Enhance"
            
        'Color operations; numbers 600-699
        Case 600
            GetNameOfProcess = "Rechannel (Blue)"
        Case 601
            GetNameOfProcess = "Rechannel (Green)"
        Case 602
            GetNameOfProcess = "Rechannel (Red)"
        Case 603
            GetNameOfProcess = "Shift Colors (Left)"
        Case 604
            GetNameOfProcess = "Shift Colors (Right)"
        Case 605
            GetNameOfProcess = "Brightness/Contrast"
        Case 606
            GetNameOfProcess = "Gamma Correction"
        Case 607
            GetNameOfProcess = "Invert Colors"
        Case 608
            GetNameOfProcess = "Invert Hue"
        Case 609
            GetNameOfProcess = "Film Negative"
        Case 610
            GetNameOfProcess = "Auto-Enhance Contrast"
        Case 611
            GetNameOfProcess = "Auto-Enhance Highlights"
        Case 612
            GetNameOfProcess = "Auto-Enhance Midtones"
        Case 613
            GetNameOfProcess = "Auto-Enhance Shadows"
        Case 614
            GetNameOfProcess = "Image Levels"
        Case 615
            GetNameOfProcess = "Colorize"
        Case 616
            GetNameOfProcess = "Reduce Colors"
            
        'Coordinate filters/transformations; numbers 700-799
        Case 700
            GetNameOfProcess = "Resize"
        Case 701
            GetNameOfProcess = "Flip"
        Case 702
            GetNameOfProcess = "Mirror"
        Case 703
            GetNameOfProcess = "Rotate 90 Clockwise"
        Case 704
            GetNameOfProcess = "Rotate 180"
        Case 705
            GetNameOfProcess = "Rotate 90 Counter-Clockwise"
        Case 706
            GetNameOfProcess = "Free Rotation"
        Case 707
            GetNameOfProcess = "Isometric Conversion"
        Case 708
            GetNameOfProcess = "Tile Image"
            
        'Miscellaneous filters; numbers 800-899
        Case 800
            GetNameOfProcess = "Compound Invert (Dark)"
        Case 801
            GetNameOfProcess = "Compound Invert (Light)"
        Case 802
            GetNameOfProcess = "Compound Invert (Moderate)"
        Case 803
            GetNameOfProcess = "Fade"
        Case 807
            GetNameOfProcess = "Unfade"
        Case 808
            GetNameOfProcess = "Atmosphere"
        Case 809
            GetNameOfProcess = "Freeze"
        Case 810
            GetNameOfProcess = "Lava"
        Case 811
            GetNameOfProcess = "Burn"
        Case 812
            GetNameOfProcess = "Ocean"
        Case 813
            GetNameOfProcess = "Water"
        Case 814
            GetNameOfProcess = "Steel"
        Case 815
            GetNameOfProcess = "Dream"
        Case 816
            GetNameOfProcess = "Alien"
        Case 817
            GetNameOfProcess = "Custom Filter"
        Case 818
            GetNameOfProcess = "Sepia/Antique"
        Case 819
            GetNameOfProcess = "Blacklight"
        Case 820
            GetNameOfProcess = "Posterize"
        Case 821
            GetNameOfProcess = "Radioactive"
        Case 822
            GetNameOfProcess = "Solarize"
        Case 823
            GetNameOfProcess = "Generate Twins"
        Case 824
            GetNameOfProcess = "Synthesize"
        Case 825
            GetNameOfProcess = "Add Noise"
        Case 827
            GetNameOfProcess = "Count Image Colors"
        Case 828
            GetNameOfProcess = "Fog"
        Case 829
            GetNameOfProcess = "Rainbow"
        Case 830
            GetNameOfProcess = "Vibrate"
        Case 831
            GetNameOfProcess = "Despeckle"
        Case 832
            GetNameOfProcess = "Custom Despeckle"
        Case 840
            GetNameOfProcess = "Animate"
        
        Case 900
            GetNameOfProcess = "Repeat Last Action"
        Case 901
            GetNameOfProcess = "Fade last effect"
            
        'This "Else" statement should never trigger, but if it does, return an empty string
        Case Else
            GetNameOfProcess = ""
            
    End Select
    
End Function
