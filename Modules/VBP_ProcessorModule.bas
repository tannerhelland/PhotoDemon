Attribute VB_Name = "Processor"
'***************************************************************************
'Program Sub-Processor and Error Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 20/June/12
'Last update: changed the way RepeatLastAction is handled to prevent collisions when the user repeats
'             their previous action but NOT by clicking RepeatLastAction (basically, when they do the
'             exact same filter with identical parameters twice in a row).
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
    Public Const GrayscaleDitherCustom As Long = 304
    
    'Area filters; numbers 400-499
    '-Blur
    Public Const Antialias As Long = 416
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
    
    'Other filters; numbers 800-899
    '-Compound invert
    Public Const DarkCompoundInvert As Long = 800
    Public Const LightCompoundInvert As Long = 801
    Public Const MediumCompoundInvert As Long = 802
    '-Fade
    Public Const Fade As Long = 803
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
    Public Const Tile As Long = 823
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
' allows us to record amd playback macros, among other things.  (See comment at top of page for more details.)
Public Sub Process(ByVal PType As Long, Optional pOPCODE As Variant = 0, Optional pOPCODE2 As Variant = 0, Optional pOPCODE3 As Variant = 0, Optional pOPCODE4 As Variant = 0, Optional pOPCODE5 As Variant = 0, Optional pOPCODE6 As Variant = 0, Optional pOPCODE7 As Variant = 0, Optional pOPCODE8 As Variant = 0, Optional pOPCODE9 As Variant = 0, Optional LoadForm As Boolean = False, Optional RecordAction As Boolean = True)

    'Main error handler for the entire program is initialized by this line
    On Error GoTo MainErrHandler
    
    'This line is used for raising errors to test the error handler
    'Err.Raise 339
    
    'Mark the software processor as busy
    Processing = True
    
    'Set the mouse cursor to an hourglass
    FormMain.MousePointer = vbHourglass
    
    'If we are to perform the last command, simply replace all the method parameters using data from the
    ' LastFilterCall object, then let the routine carry on as usual
    If PType = LastCommand Then
        PType = LastFilterCall.MainType
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
            .MainType = PType
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
    'From this point on, all we do is check the PType variable (the constant that is the first
    'variable passed to this subroutine) and depending on what it is, we call the appropriate
    'subroutine.  Very simple and very fast to do.
    
    'I have also subdivided the "Select Case" statements up by groups of 100, just as I do
    'above in the declarations part.
    
    'Main functions.  These are never recorded by macros.
    If PType > 0 And PType <= 99 Then
        Select Case PType
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
            Case Redo
                RedoImageRestore
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
    If PType >= 101 Then
    
        'Get the image data (to get image size and information)
        GetImageData
        
        'Only save an "undo" image if we are NOT loading a form for user input, and if
        'we ARE allowed to record this action, and if it's not counting colors (useless),
        ' and if we're not performing a batch conversion (saves a lot of time to not generate undo files!)
        If MacroStatus <> MacroBATCH Then
            If LoadForm <> True And RecordAction <> False And PType <> CountColors Then BuildImageRestore
        End If
        
        'Save this information in the LastFilterCall variable (to be used if the user clicks on
        ' Edit -> Redo Last Command.
        FormMain.MnuRepeatLast.Enabled = True
        LastFilterCall.MainType = PType
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
    If PType >= 100 And PType <= 199 Then
        Select Case PType
            Case ViewHistogram
                FormHistogram.Show 0, FormMain
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
    If PType >= 200 And PType <= 299 Then
        Select Case PType
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
    If PType >= 300 And PType <= 399 Then
        Select Case PType
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
            Case GrayscaleDitherCustom
                FormGrayscale.fGrayscaleCustomDither pOPCODE
        End Select
    End If
    
    'Area filters
    If PType >= 400 And PType <= 499 Then
        Select Case PType
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
    If PType >= 500 And PType <= 599 Then
        Select Case PType
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
    If PType >= 600 And PType <= 699 Then
        Select Case PType
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
    If PType >= 700 And PType <= 799 Then
        Select Case PType
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
        End Select
    End If
    
    'Other filters
    If PType >= 800 And PType <= 899 Then
        Select Case PType
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
            Case Tile
                If LoadForm = True Then
                    FormTile.Show 1, FormMain
                Else
                    FormTile.GenerateTwins CByte(pOPCODE)
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
    If PType = FadeLastEffect Then MenuFadeLastEffect
    
    'Restore the mouse pointer to its default value; if we are running a batch conversion, however, leave it busy
    ' The batch routine will handle restoring the cursor to normal.
    If MacroStatus <> MacroBATCH Then FormMain.MousePointer = vbDefault
    
    'If the histogram form is visible, redraw the histogram
    If FormHistogram.Visible = True Then
        FormHistogram.TallyHistogramValues
        FormHistogram.DrawHistogram
    End If
    
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
    Dim MsgReturn As VbMsgBoxResult
    
    'Ignore errors that aren't actually errors
    If Err.Number = 0 Then Exit Sub
    
    'Object was unloaded before it could be shown - this is intentional, so ignore the error
    If Err.Number = 364 Then Exit Sub
        
    'Out of memory error
    If Err.Number = 480 Or Err.Number = 7 Then
        AddInfo = "There is not enough memory available to continue this operation.  Please free up system memory (RAM) and try again."
        Message "Out of memory.  Function cancelled."
        mType = vbCritical + vbOKOnly
    
    'Invalid picture error
    ElseIf Err.Number = 481 Then
        AddInfo = "You have attempted to load an invalid picture.  This can happen if a file does not contain image data, or if it contains image data in an unsupported format." & vbCrLf & vbCrLf & "- If you downloaded this image from the Internet, the download may have terminated prematurely.  Please try downloading the image again." & vbCrLf & vbCrLf & "- If this image file came from a digital camera, scanner, or other image editing program, it's possible that " & PROGRAMNAME & " simply doesn't understand this particular file format.  Please save the image in a generic format (such as bitmap or JPEG), then reload it."
        Message "Invalid picture.  Image load cancelled."
        mType = vbCritical + vbOKOnly
    
    'File not found error
    ElseIf Err.Number = 53 Then
        AddInfo = "The specified file could not be located.  If it was located on removable media, please re-insert the proper floppy disk, CD, or portable drive.  If the file is not located on portable media, make sure that:" & vbCrLf & "1) the file hasn't been deleted, and..." & "2) the file location provided to " & PROGRAMNAME & " is correct."
        Message "File not found."
        mType = vbCritical + vbOKOnly
        
    'Unknown error
    Else
        AddInfo = PROGRAMNAME & " cannot locate additional information for this error.  If it persists, please contact the e-mail address below."
        mType = vbCritical + vbOKOnly
        Message "Unknown error."
    End If
    
    'Create the message box to return the error information
    MsgReturn = MsgBox(PROGRAMNAME & " has experienced an error.  Details on the problem include:" & vbCrLf & vbCrLf & _
    "Error number " & Err.Number & vbCrLf & _
    "Description: " & Err.Description & vbCrLf & vbCrLf & _
    AddInfo & vbCrLf & vbCrLf & _
    "Sorry for the inconvenience!" & vbCrLf & vbCrLf & _
    "- Tanner Helland, " & PROGRAMNAME & " Developer" & vbCrLf & _
    "www.tannerhelland.com/contact", mType, PROGRAMNAME & " Error Handler: #" & Err.Number)
    
    'If an invalid picture was loaded, unload the active form (since it will just be empty and pictureless)
    If Err.Number = 481 Then Unload FormMain.ActiveForm

End Sub
