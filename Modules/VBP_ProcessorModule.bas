Attribute VB_Name = "Processor"
'***************************************************************************
'Program Sub-Processor and Error Handler
'Copyright ©2001-2014 by Tanner Helland
'Created: 4/15/01
'Last updated: 22/August/13
'Last update: finish organizing processor calls by category.  They now match their menu order, and I will try to keep it that way.
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
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit
Option Compare Text

'Data type for tracking processor calls - used for macros (NOTE: this is the 2014 model; older models are no longer supported.)
Public Type ProcessCall
    Id As String
    Dialog As Boolean
    Parameters As String
    MakeUndo As Long
    Tool As Long
    Recorded As Boolean
End Type

'During macro recording, all requests to the processor are stored in this array.
Public Processes() As ProcessCall

'How many processor requests we currently have stored.
Public ProcessCount As Long

'Full processor information of the previous request (used to provide a "Repeat Last Action" feature)
Public LastProcess As ProcessCall

'Track processing (e.g. whether or not the software processor is busy right now)
Public Processing As Boolean

'Elapsed time of this processor request (to enable this, see the top constant in the Public_Constants module)
Private m_ProcessingTime As Double

'PhotoDemon's software processor.  (Almost) every action the program takes is first routed through this method.  This processor is what
' makes recording and playing back macros possible, as well as a host of other features.  (See comment at top of page for more details.)
'
'INPUTS (asterisks denote optional parameters):
' - processID: a string identifying the action to be performed, e.g. "Blur"
' - *showDialog: some functions can be run with or without a dialog; for example, "Blur", "True" will display a blur settings dialog,
'                while "Blur", "False" will actually apply the blur.  If showDialog is true, no Undo will be created for the action.
' - *processParameters: all parameters for this function, concatenated into a single string.  The processor will automatically parse out
'                       individual parameters as necessary.
' - *createUndo: ID describing what kind of Undo entry to create for this action.  0 prevents Undo creation, while values > 0 correspond
'                to a specific type of Undo.  (1 = image undo, 2 = selection undo - these values are needed because undoing a selection
'                requires completely different code vs undoing an image filter.)  This value is set to 1 by default, but some functions
'                - like "Count image colors" - may explicitly specify that no Undo is necessary.  NOTE: if showDialog is TRUE, this value
'                will automatically be set to 0, which means "DO NOT CREATE UNDO", because we never create Undo data when showing a
'                dialog (as the user may cancel it).
' - *relevantTool: some Process calls are initiated by a particular tool (for example, "create selection" will be called by one of the
'                  selection tools).  This parameter can contain the relevant tool for a given action.  If Undo is used to return to a
'                  previous state, the relevant tool can automatically be selected, making it much easier for the user to make changes
'                  to an action using the proper tool.
' - *recordAction: are macros allowed to record this action?  Actions are assumed to be recordable.  However, some PhotoDemon functions
'                  are actually several actions strung together; when these are used, subsequent actions are marked as "not recordable"
'                  to prevent them from being executed twice.
Public Sub Process(ByVal processID As String, Optional showDialog As Boolean = False, Optional processParameters As String = "", Optional createUndo As Long = 1, Optional relevantTool As Long = -1, Optional recordAction As Boolean = True)

    'Main error handler for the entire program is initialized by this line
    On Error GoTo MainErrHandler
    
    'Mark the software processor as busy
    Processing = True
        
    'Disable the main form to prevent the user from clicking additional menus or tools while this action is processing
    FormMain.Enabled = False
    
    'If we need to display an additional dialog, restore the default mouse cursor.  Otherwise, set the cursor to busy.
    If showDialog Then
        If g_OpenImageCount > 0 Then
            If Not (pdImages(g_CurrentImage).containingForm Is Nothing) Then setArrowCursor pdImages(g_CurrentImage).containingForm
        End If
    Else
        Screen.MousePointer = vbHourglass
    End If
        
    'If we are to perform the last command, simply replace all the method parameters using data from the
    ' LastEffectsCall object, then let the routine carry on as usual
    If processID = "Repeat last action" Then
        processID = LastProcess.Id
        showDialog = LastProcess.Dialog
        processParameters = LastProcess.Parameters
        createUndo = LastProcess.MakeUndo
        relevantTool = LastProcess.Tool
        recordAction = LastProcess.Recorded
    End If
    
    'If the macro recorder is running and this action is marked as recordable, store it in our array of processor calls
    If (MacroStatus = MacroSTART) And recordAction Then
    
        'First things first: if the current action is NOT selection-related, but the last one was, make a backup of all selection settings.
        If (createUndo <> 2) And (LastProcess.MakeUndo = 2) And (Not (LastProcess.Id = "Finalize selection for macro")) Then
            Process "Finalize selection for macro", False, pdImages(g_CurrentImage).mainSelection.getSelectionParamString, 2, g_CurrentTool, True
        End If
    
        'Increase the process count
        ProcessCount = ProcessCount + 1
        
        'Copy the current process's information into the tracking array
        ReDim Preserve Processes(0 To ProcessCount) As ProcessCall
        
        With Processes(ProcessCount)
            .Id = processID
            .Dialog = showDialog
            .Parameters = processParameters
            .MakeUndo = createUndo
            .Tool = relevantTool
            .Recorded = recordAction
        End With
        
    End If
    
    'If a dialog is being displayed, disable Undo creation
    If showDialog Then createUndo = 0
    
    'If this action requires us to create an Undo, create it now.  (We can also use this identifier to initiate a few
    ' other, related actions.)
    If createUndo <> 0 Then
        
        'Temporarily disable drag-and-drop operations for the main form
        g_AllowDragAndDrop = False
        FormMain.OLEDropMode = 0
        
        'By default, actions are assumed to want Undo data created.  However, there are some known exceptions:
        ' 1) If a dialog is being displayed
        ' 2) If recording has been disabled for this action
        ' 3) If we are in the midst of playing back a recorded macro (Undo data takes extra time to process, so drop it)
        If MacroStatus <> MacroBATCH Then
            If (Not showDialog) And recordAction Then pdImages(g_CurrentImage).undoManager.createUndoData processID, createUndo, relevantTool
        End If
        
        'Save this information in the LastProcess variable (to be used if the user clicks on Edit -> Redo Last Action.
        FormMain.MnuEdit(2).Enabled = True
        With LastProcess
            .Id = processID
            .Dialog = showDialog
            .Parameters = processParameters
            .MakeUndo = createUndo
            .Tool = relevantTool
            .Recorded = recordAction
        End With
        
        'If the user wants us to time how long this action takes, mark the current time now
        If Not showDialog Then
            If DISPLAY_TIMINGS Then m_ProcessingTime = Timer
        End If
        
    End If
    
    'Finally, create a parameter parser to handle the parameter string.  This class will parse out individual parameters
    ' as specific data types when it comes time to use them.
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    If Len(processParameters) > 0 Then cParams.setParamString processParameters
    
    '******************************************************************************************************************
    '
    'BEGIN PROCESS SORTING
    '
    'The bulk of this routine starts here.  From this point on, the processID string is compared against a hard-coded
    ' list of every possible PhotoDemon action, filter, or other operation.  Depending on the processID, additional
    ' actions will be performed.
    '
    'Note that prior to the 6.0 release, this function used numeric identifiers instead of strings.  This has since
    ' been abandoned in favor of a string-based approach, and at present there are no plans to restore the old
    ' numeric behavior.  Strings simplify the code, they make it much easier to add new functions, and they will
    ' eventually allow for a "filter browser" that allows the user to preview any effect from a single dialog.
    ' Numeric IDs were much harder to manage in that context, and over time their numbering grew so arbitrary that
    ' it made maintenance very difficult.
    '
    'For ease of reference, the various processIDs are divided into categories of similar functions.  These categories
    ' match the organization of PhotoDemon's menus.  Please note that such organization in this function is simply to
    ' improve readability; there is no functional purpose to it.
    '
    '******************************************************************************************************************
    
    Select Case processID
    
        'FILE MENU FUNCTIONS
        ' This includes actions like opening or saving images.  These actions are never recorded.
    
        Case "Open"
            MenuOpen
            
        Case "Close"
            MenuClose
            
        Case "Close all"
            MenuCloseAll
        
        Case "Save"
            MenuSave g_CurrentImage
            
        Case "Save as"
            MenuSaveAs g_CurrentImage
            
        Case "Revert"
            If FormMain.MnuFile(7).Enabled Then
                pdImages(g_CurrentImage).undoManager.revertToLastSavedState
                
                'Also, redraw the current child form icon and the image tab-bar
                createCustomFormIcon pdImages(g_CurrentImage).containingForm
                toolbar_ImageTabs.notifyUpdatedImage g_CurrentImage
            End If
            
        Case "Batch wizard"
            If showDialog Then
                
                'Because the Batch Wizard window provides a custom drag/drop implementation, we disable regular drag/drop while it's active
                g_AllowDragAndDrop = False
                showPDDialog vbModal, FormBatchWizard
                g_AllowDragAndDrop = True
                
            End If
            
        Case "Print"
            If showDialog Then
                
                'As a temporary workaround, Vista+ users are routed through the default Windows photo printing
                ' dialog.  XP users get the old PD print dialog.
                If g_IsVistaOrLater Then
                    printViaWindowsPhotoPrinter
                Else
                    If Not FormPrint.Visible Then showPDDialog vbModal, FormPrint
                End If
                
                'In the future, the print dialog will be replaced with this new version.  However, there are
                ' bigger priorities for 6.2, so I'm putting this on hold for now...
                'If Not FormPrintNew.Visible Then showPDDialog vbModal, FormPrintNew
                
            End If
            
        Case "Exit program"
            Unload FormMain
            
            'If the user does not cancel the exit, we must forcibly exit this sub (otherwise the program will not exit)
            If g_ProgramShuttingDown Then Exit Sub
        
        Case "Select scanner or camera"
            Twain32SelectScanner
            
        Case "Scan image"
            Twain32Scan
            
        Case "Screen capture"
            If showDialog Then
                showPDDialog vbModal, FormScreenCapture
            Else
                CaptureScreen cParams.GetBool(1), cParams.GetBool(2), cParams.GetLong(3), cParams.GetBool(4), cParams.GetString(5)
            End If
            
        Case "Internet import"
            If Not FormInternetImport.Visible Then
                showPDDialog vbModal, FormInternetImport
                Exit Sub
            End If
        
        
        
        'EDIT MENU FUNCTIONS
        ' This includes things like copying or pasting an image.  These actions are never recorded.
        
        Case "Undo"
            If FormMain.MnuEdit(0).Enabled Then
                pdImages(g_CurrentImage).undoManager.restoreUndoData
                
                'Also, redraw the current child form icon and the image tab-bar
                createCustomFormIcon pdImages(g_CurrentImage).containingForm
                toolbar_ImageTabs.notifyUpdatedImage g_CurrentImage
            End If
            
        Case "Redo"
            If FormMain.MnuEdit(1).Enabled Then
                pdImages(g_CurrentImage).undoManager.RestoreRedoData
                
                'Also, redraw the current child form icon and the image tab-bar
                createCustomFormIcon pdImages(g_CurrentImage).containingForm
                toolbar_ImageTabs.notifyUpdatedImage g_CurrentImage
            End If
        
        Case "Copy to clipboard"
            ClipboardCopy
            
        Case "Paste as new image"
            ClipboardPaste
            
        Case "Empty clipboard"
            ClipboardEmpty
        
        
        
        'TOOLS MENU FUNCTIONS
        ' This includes things like macro recording.  These actions are never recorded.
        Case "Start macro recording"
            StartMacro
        
        Case "Stop macro recording"
            StopMacro
            
        Case "Play macro"
            PlayMacro
            
        
        
        'IMAGE MENU FUNCTIONS
        ' This includes all actions that can only operate on a full image (never selections).  These actions are recorded.
        
        'Duplicate image
        Case "Duplicate image"
            
            'It may seem odd, but the Duplicate function can be found in the "Loading" module; I do this because
            ' we effectively LOAD a copy of the original image, so all loading operations (create a form, catalog
            ' metadata, initialize properties) have to be repeated.
            DuplicateCurrentImage
        
        'Alpha channel addition/removal
        Case "Add alpha channel"
            If showDialog Then
                showPDDialog vbModal, FormTransparency_Basic
            Else
                FormTransparency_Basic.simpleConvert32bpp cParams.GetLong(1)
            End If
            
        Case "Color to alpha"
            If showDialog Then
                showPDDialog vbModal, FormTransparency_FromColor
            Else
                FormTransparency_FromColor.colorToAlpha cParams.GetLong(1), cParams.GetDouble(2), cParams.GetDouble(3)
            End If
            
        Case "Remove alpha channel"
            If showDialog Then
                showPDDialog vbModal, FormConvert24bpp
            Else
                ConvertImageColorDepth 24, cParams.GetLong(1)
            End If
            
        'Resize operations
        Case "Resize"
            If showDialog Then
                showPDDialog vbModal, FormResize
            Else
                FormResize.ResizeImage cParams.GetLong(1), cParams.GetLong(2), cParams.GetByte(3), cParams.GetLong(4), cParams.GetLong(5)
            End If
        
        Case "Smart resize"
            If showDialog Then
                showPDDialog vbModal, FormSmartResize
            Else
                FormSmartResize.SmartResizeImage cParams.GetLong(1), cParams.GetLong(2)
            End If
        
        Case "Canvas size"
            If showDialog Then
                showPDDialog vbModal, FormCanvasSize
            Else
                FormCanvasSize.ResizeCanvas cParams.GetLong(1), cParams.GetLong(2), cParams.GetLong(3), cParams.GetLong(4)
            End If
        
        'Crop operations
        Case "Crop"
            MenuCropToSelection
            
        Case "Autocrop"
            AutocropImage
            
        'Rotate operations
        Case "Rotate 90° clockwise"
            MenuRotate90Clockwise
            
        Case "Rotate 180°"
            MenuRotate180
            
        Case "Rotate 90° counter-clockwise"
            MenuRotate270Clockwise
            
        Case "Arbitrary rotation"
            If showDialog Then
                showPDDialog vbModal, FormRotate
            Else
                FormRotate.RotateArbitrary cParams.GetLong(1), cParams.GetDouble(2)
            End If
            
        'Other coordinate transforms
        Case "Flip vertical"
            MenuFlip
            
        Case "Flip horizontal"
            MenuMirror
            
        Case "Isometric conversion"
            FilterIsometric
            
        Case "Tile"
            If showDialog Then
                showPDDialog vbModal, FormTile
            Else
                FormTile.GenerateTile cParams.GetByte(1), cParams.GetLong(2), cParams.GetLong(3)
            End If
        
        
        'Other miscellaneous image-only items
        Case "Count image colors"
            MenuCountColors
            
        Case "Reduce colors"
            If showDialog Then
                showPDDialog vbModal, FormReduceColors
            Else
                FormReduceColors.ReduceImageColors_Auto cParams.GetLong(2)
            End If
        
        
        
        'SELECTION FUNCTIONS
        ' Any action that operates on selections - creating them, moving them, erasing them, etc
        
        
        'Create/remove selections
        Case "Create selection"
            CreateNewSelection cParams.getParamString
        
        Case "Remove selection"
            RemoveCurrentSelection cParams.getParamString
        
        
        'Backup selection settings during a recorded macro (required to avoid "lazy" tracking method used on selection changes)
        Case "Finalize selection for macro"
            backupSelectionSettingsForMacro cParams.getParamString
            
        
        'Modify the existing selection in some way
        Case "Invert selection"
            invertCurrentSelection
            
        Case "Grow selection"
            If showDialog Then
                growCurrentSelection True
            Else
                growCurrentSelection False, cParams.GetDouble(1)
            End If
            
        Case "Shrink selection"
            If showDialog Then
                shrinkCurrentSelection True
            Else
                shrinkCurrentSelection False, cParams.GetDouble(1)
            End If
        
        Case "Feather selection"
            If showDialog Then
                featherCurrentSelection True
            Else
                featherCurrentSelection False, cParams.GetDouble(1)
            End If
        
        Case "Sharpen selection"
            If showDialog Then
                sharpenCurrentSelection True
            Else
                sharpenCurrentSelection False, cParams.GetDouble(1)
            End If
            
        Case "Border selection"
            If showDialog Then
                borderCurrentSelection True
            Else
                borderCurrentSelection False, cParams.GetDouble(1)
            End If
        
        'Load/save selection from/to file
        Case "Load selection"
            If showDialog Then
                LoadSelectionFromFile True
            Else
                LoadSelectionFromFile False, cParams.getParamString
            End If
            
        Case "Save selection"
            SaveSelectionToFile
            
        'Export selected area as image (defaults to PNG, but user can select the actual format)
        Case "Export selected area as image"
            ExportSelectedAreaAsImage
        
        'Export selection mask as image (defaults to PNG, but user can select the actual format)
        Case "Export selection mask as image"
            ExportSelectionMaskAsImage
        
        ' This is a dummy entry; it only exists so that Undo/Redo data is correctly generated when a selection is moved
        Case "Move selection"
            CreateNewSelection cParams.getParamString
            
        ' This is a dummy entry; it only exists so that Undo/Redo data is correctly generated when a selection is moved
        Case "Resize selection"
            CreateNewSelection cParams.getParamString
        
        Case "Select all"
            SelectWholeImage
        
        
        
        'ADJUSTMENT FUNCTIONS
        ' Any action that is used to fix or correct problems with an image, including basic color space transformations (e.g. grayscale/sepia)
        
        'Luminance adjustment functions
        Case "Brightness and contrast"
            If showDialog Then
                showPDDialog vbModal, FormBrightnessContrast
            Else
                FormBrightnessContrast.BrightnessContrast cParams.GetLong(1), cParams.GetDouble(2), cParams.GetBool(3)
            End If
        
        Case "Curves"
            If showDialog Then
                showPDDialog vbModal, FormCurves
            Else
                FormCurves.ApplyCurveToImage cParams.getParamString
            End If
        
        Case "Exposure"
            If showDialog Then
                showPDDialog vbModal, FormExposure
            Else
                FormExposure.Exposure cParams.GetDouble(1)
            End If
            
        Case "Gamma"
            If showDialog Then
                showPDDialog vbModal, FormGamma
            Else
                FormGamma.GammaCorrect cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetDouble(3)
            End If
        
        Case "Levels"
            If showDialog Then
                showPDDialog vbModal, FormLevels
            Else
                FormLevels.MapImageLevels cParams.GetLong(1), cParams.GetDouble(2), cParams.GetLong(3), cParams.GetLong(4), cParams.GetLong(5)
            End If
            
        Case "Shadows and highlights"
            If showDialog Then
                showPDDialog vbModal, FormShadowHighlight
            Else
                FormShadowHighlight.ApplyShadowHighlight cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetLong(3)
            End If
            
        Case "White balance"
            If showDialog Then
                showPDDialog vbModal, FormWhiteBalance
            Else
                FormWhiteBalance.AutoWhiteBalance cParams.GetDouble(1)
            End If
        
        'Color adjustments
        Case "Color balance"
            If showDialog Then
                showPDDialog vbModal, FormColorBalance
            Else
                FormColorBalance.ApplyColorBalance cParams.GetLong(1), cParams.GetLong(2), cParams.GetLong(3), cParams.GetBool(4)
            End If
            
        Case "Hue and saturation"
            If showDialog Then
                showPDDialog vbModal, FormHSL
            Else
                FormHSL.AdjustImageHSL cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetDouble(3)
            End If
            
        Case "Photo filter"
            If showDialog Then
                showPDDialog vbModal, FormPhotoFilters
            Else
                FormPhotoFilters.ApplyPhotoFilter cParams.GetLong(1), cParams.GetDouble(2), cParams.GetBool(3)
            End If
            
        Case "Replace color"
            If showDialog Then
                showPDDialog vbModal, FormReplaceColor
            Else
                FormReplaceColor.ReplaceSelectedColor cParams.GetLong(1), cParams.GetLong(2), cParams.GetDouble(3), cParams.GetDouble(4)
            End If
            
        Case "Temperature"
            If showDialog Then
                showPDDialog vbModal, FormColorTemp
            Else
                FormColorTemp.ApplyTemperatureToImage cParams.GetLong(1), cParams.GetBool(2), cParams.GetDouble(3)
            End If
            
        Case "Vibrance"
            If showDialog Then
                showPDDialog vbModal, FormVibrance
            Else
                FormVibrance.Vibrance cParams.GetDouble(1)
            End If
        
        'Miscellaneous adjustments
        Case "Colorize"
            If showDialog Then
                showPDDialog vbModal, FormColorize
            Else
                FormColorize.ColorizeImage cParams.GetDouble(1), cParams.GetBool(2)
            End If
        
        'Grayscale conversions
        Case "Black and white"
            If showDialog Then
                showPDDialog vbModal, FormGrayscale
            Else
                FormGrayscale.masterGrayscaleFunction cParams.GetLong(1), cParams.GetString(2)
            End If
        
        'Invert operations
        Case "Invert RGB"
            MenuInvert
            
        Case "Compound invert"
            MenuCompoundInvert cParams.GetLong(1)
            
        Case "Film negative"
            MenuNegative
            
        Case "Invert hue"
            MenuInvertHue
        
        'Monochrome conversion
        ' (Note: all monochrome conversion operations are condensed into a single function.  (Past versions spread them across multiple functions.))
        Case "Color to monochrome"
            If showDialog Then
                showPDDialog vbModal, FormMonochrome
            Else
                FormMonochrome.masterBlackWhiteConversion cParams.GetLong(1), cParams.GetLong(2), cParams.GetLong(3), cParams.GetLong(4)
            End If
            
        Case "Monochrome to grayscale"
            If showDialog Then
                showPDDialog vbModal, FormMonoToColor
            Else
                FormMonoToColor.ConvertMonoToColor cParams.GetLong(1)
            End If
            
        Case "Sepia"
            MenuSepia
            
        'Channel operations
        Case "Channel mixer"
            If showDialog Then
                showPDDialog vbModal, FormChannelMixer
            Else
                FormChannelMixer.ApplyChannelMixer cParams.getParamString
            End If
            
        Case "Rechannel"
            If showDialog Then
                showPDDialog vbModal, FormRechannel
            Else
                FormRechannel.RechannelImage cParams.GetByte(1)
            End If
            
        Case "Shift colors (left)"
            MenuCShift 1
            
        Case "Shift colors (right)"
            MenuCShift 0
                    
        Case "Maximum channel"
            FilterMaxMinChannel True
        
        Case "Minimum channel"
            FilterMaxMinChannel False
            
        'Histogram functions
        Case "Display histogram"
            showPDDialog vbModal, FormHistogram
        
        Case "Stretch histogram"
            FormHistogram.StretchHistogram
            
        Case "Equalize"
            If showDialog Then
                showPDDialog vbModal, FormEqualize
            Else
                FormEqualize.EqualizeHistogram cParams.GetBool(1), cParams.GetBool(2), cParams.GetBool(3), cParams.GetBool(4)
            End If
        
        
        
        'EFFECT FUNCTIONS
        'Sometimes fun, sometimes practical, no real unifying factor to these.
        
        
        'Artistic
        Case "Comic book"
            MenuComicBook
            
        Case "Figured glass"
            If showDialog Then
                showPDDialog vbModal, FormFiguredGlass
            Else
                FormFiguredGlass.FiguredGlassFX cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetLong(3), cParams.GetBool(4)
            End If
        
        Case "Film noir"
            MenuFilmNoir
            
        Case "Kaleidoscope"
            If showDialog Then
                showPDDialog vbModal, FormKaleidoscope
            Else
                FormKaleidoscope.KaleidoscopeImage cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetDouble(3), cParams.GetDouble(4), cParams.GetBool(5)
            End If
            
        Case "Modern art"
            If showDialog Then
                showPDDialog vbModal, FormModernArt
            Else
                FormModernArt.ApplyModernArt cParams.GetLong(1)
            End If
            
        Case "Oil painting"
            If showDialog Then
                showPDDialog vbModal, FormOilPainting
            Else
                FormOilPainting.ApplyOilPaintingEffect cParams.GetLong(1), cParams.GetDouble(2)
            End If
            
        Case "Posterize"
            If showDialog Then
                showPDDialog vbModal, FormPosterize
            Else
                FormPosterize.ReduceImageColors_BitRGB cParams.GetByte(1), cParams.GetByte(2), cParams.GetByte(3), cParams.GetBool(4)
            End If
            
        Case "Posterize (dithered)"
            FormPosterize.ReduceImageColors_BitRGB_ErrorDif cParams.GetByte(1), cParams.GetByte(2), cParams.GetByte(3), cParams.GetBool(4)
            
        Case "Pencil drawing"
            FilterPencil
            
        Case "Relief"
            FilterRelief
            
            
        'Blur
        
        Case "Box blur"
            If showDialog Then
                showPDDialog vbModal, FormBoxBlur
            Else
                FormBoxBlur.BoxBlurFilter cParams.GetLong(1), cParams.GetLong(2)
            End If
        
        Case "Gaussian blur"
            If showDialog Then
                showPDDialog vbModal, FormGaussianBlur
            Else
                FormGaussianBlur.GaussianBlurFilter cParams.GetDouble(1), cParams.GetLong(2)
            End If
        
        Case "Grid blur"
            FilterGridBlur
            
        Case "Motion blur"
            If showDialog Then
                showPDDialog vbModal, FormMotionBlur
            Else
                FormMotionBlur.MotionBlurFilter cParams.GetDouble(1), cParams.GetLong(2), cParams.GetBool(3), cParams.GetBool(4)
            End If
            
        Case "Pixelate"
            If showDialog Then
                showPDDialog vbModal, FormPixelate
            Else
                FormPixelate.PixelateFilter cParams.GetLong(1), cParams.GetLong(2)
            End If
            
        Case "Radial blur"
            If showDialog Then
                showPDDialog vbModal, FormRadialBlur
            Else
                FormRadialBlur.RadialBlurFilter cParams.GetDouble(1), cParams.GetBool(2), cParams.GetBool(3)
            End If
            
        Case "Smart blur"
            If showDialog Then
                showPDDialog vbModal, FormSmartBlur
            Else
                FormSmartBlur.SmartBlurFilter cParams.GetDouble(1), cParams.GetByte(2), cParams.GetBool(3)
            End If
            
        Case "Zoom blur"
            If showDialog Then
                showPDDialog vbModal, FormZoomBlur
            Else
                FormZoomBlur.ZoomBlurWrapper cParams.GetBool(1), cParams.GetLong(2)
            End If
        
        'Distort filters
        
        Case "Apply lens distortion"
            If showDialog Then
                showPDDialog vbModal, FormLens
            Else
                FormLens.ApplyLensDistortion cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetBool(3)
            End If
            
        Case "Correct lens distortion"
            If showDialog Then
                showPDDialog vbModal, FormLensCorrect
            Else
                FormLensCorrect.ApplyLensCorrection cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetDouble(3), cParams.GetLong(4), cParams.GetBool(5)
            End If
        
        Case "Miscellaneous distort"
            If showDialog Then
                showPDDialog vbModal, FormMiscDistorts
            Else
                FormMiscDistorts.ApplyMiscDistort cParams.GetString(1), cParams.GetLong(2), cParams.GetLong(3), cParams.GetBool(4)
            End If
            
        Case "Pan and zoom"
            If showDialog Then
                showPDDialog vbModal, FormPanAndZoom
            Else
                FormPanAndZoom.PanAndZoomFilter cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetDouble(3), cParams.GetLong(4), cParams.GetBool(5)
            End If
        
        Case "Perspective"
            If showDialog Then
                showPDDialog vbModal, FormPerspective
            Else
                FormPerspective.PerspectiveImage cParams.getParamString
            End If
            
        Case "Pinch and whirl"
            If showDialog Then
                showPDDialog vbModal, FormPinch
            Else
                FormPinch.PinchImage cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetDouble(3), cParams.GetLong(4), cParams.GetBool(5)
            End If
            
        Case "Poke"
            If showDialog Then
                showPDDialog vbModal, FormPoke
            Else
                FormPoke.ApplyPokeDistort cParams.GetDouble(1), cParams.GetLong(2), cParams.GetBool(3)
            End If
            
        Case "Polar conversion"
            If showDialog Then
                showPDDialog vbModal, FormPolar
            Else
                FormPolar.ConvertToPolar cParams.GetLong(1), cParams.GetBool(2), cParams.GetDouble(3), cParams.GetLong(4), cParams.GetBool(5)
            End If
            
        Case "Ripple"
            If showDialog Then
                showPDDialog vbModal, FormRipple
            Else
                FormRipple.RippleImage cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetDouble(3), cParams.GetDouble(4), cParams.GetLong(5), cParams.GetBool(6)
            End If
            
        Case "Rotate"
            If showDialog Then
                showPDDialog vbModal, FormRotateDistort
            Else
                FormRotateDistort.RotateFilter cParams.GetDouble(1), cParams.GetLong(2), cParams.GetBool(3)
            End If
            
        Case "Shear"
            If showDialog Then
                showPDDialog vbModal, FormShear
            Else
                FormShear.ShearImage cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetLong(3), cParams.GetBool(4)
            End If
            
        Case "Spherize"
            If showDialog Then
                showPDDialog vbModal, FormSpherize
            Else
                FormSpherize.SpherizeImage cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetDouble(3), cParams.GetBool(4), cParams.GetLong(5), cParams.GetBool(6)
            End If
        
        Case "Squish"
            If showDialog Then
                showPDDialog vbModal, FormSquish
            Else
                FormSquish.SquishImage cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetLong(3), cParams.GetBool(4)
            End If
            
        Case "Swirl"
            If showDialog Then
                showPDDialog vbModal, FormSwirl
            Else
                FormSwirl.SwirlImage cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetLong(3), cParams.GetBool(4)
            End If
            
        Case "Waves"
            If showDialog Then
                showPDDialog vbModal, FormWaves
            Else
                FormWaves.WaveImage cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetDouble(3), cParams.GetDouble(4), cParams.GetLong(5), cParams.GetBool(6)
            End If
            
        
        'Edge filters
        Case "Emboss or engrave"
            showPDDialog vbModal, FormEmbossEngrave
            
            Case "Emboss"
                FormEmbossEngrave.FilterEmbossColor cParams.GetLong(1)
                
            Case "Engrave"
                FormEmbossEngrave.FilterEngraveColor cParams.GetLong(1)
            
        Case "Edge enhance"
            FilterEdgeEnhance
            
        Case "Find edges"
            showPDDialog vbModal, FormFindEdges
            
            Case "Artistic contour"
                FormFindEdges.FilterSmoothContour cParams.GetBool(1)
                
            Case "Find edges (Prewitt horizontal)"
                FormFindEdges.FilterPrewittHorizontal cParams.GetBool(1)
                
            Case "Find edges (Prewitt vertical)"
                FormFindEdges.FilterPrewittVertical cParams.GetBool(1)
                
            Case "Find edges (Sobel horizontal)"
                FormFindEdges.FilterSobelHorizontal cParams.GetBool(1)
                
            Case "Find edges (Sobel vertical)"
                FormFindEdges.FilterSobelVertical cParams.GetBool(1)
                
            Case "Find edges (Laplacian)"
                FormFindEdges.FilterLaplacian cParams.GetBool(1)
                
            Case "Find edges (Hilite)"
                FormFindEdges.FilterHilite cParams.GetBool(1)
                
            Case "Find edges (PhotoDemon linear)"
                FormFindEdges.PhotoDemonLinearEdgeDetection cParams.GetBool(1)
                
            Case "Find edges (PhotoDemon cubic)"
                FormFindEdges.PhotoDemonCubicEdgeDetection cParams.GetBool(1)
            
        Case "Trace contour"
            If showDialog Then
                showPDDialog vbModal, FormContour
            Else
                FormContour.TraceContour cParams.GetLong(1), cParams.GetBool(2), cParams.GetBool(3)
            End If
            
        
        'Experimental
        
        Case "Alien"
            MenuAlien
            
        Case "Black light"
            If showDialog Then
                showPDDialog vbModal, FormBlackLight
            Else
                FormBlackLight.fxBlackLight cParams.GetDouble(1)
            End If
            
        Case "Dream"
            MenuDream
            
        Case "Radioactive"
            MenuRadioactive
            
        Case "Synthesize"
            MenuSynthesize
        
        Case "Thermograph (heat map)"
            MenuHeatMap
        
        Case "Vibrate"
            MenuVibrate
        
        
        'Natural
        
        Case "Atmosphere"
            MenuAtmospheric
            
        Case "Burn"
            MenuBurn
            
        Case "Fog"
            MenuFogEffect
        
        Case "Freeze"
            MenuFrozen
            
        Case "Lava"
            MenuLava
            
        Case "Rainbow"
            MenuRainbow
        
        Case "Steel"
            MenuSteel
            
        Case "Water"
            MenuWater
            
        
        'Noise
        
        Case "Add film grain"
            If showDialog Then
                showPDDialog vbModal, FormFilmGrain
            Else
                FormFilmGrain.AddFilmGrain cParams.GetDouble(1), cParams.GetLong(2)
            End If
        
        Case "Add RGB noise"
            If showDialog Then
                showPDDialog vbModal, FormNoise
            Else
                FormNoise.AddNoise cParams.GetLong(1), cParams.GetBool(2)
            End If
            
        Case "Median"
            If showDialog Then
                FormMedian.showMedianDialog 50
            Else
                FormMedian.ApplyMedianFilter cParams.GetLong(1), cParams.GetDouble(2)
            End If
            
            
        'Sharpen
        
        Case "Sharpen"
            If showDialog Then
                showPDDialog vbModal, FormSharpen
            Else
                FormSharpen.ApplySharpenFilter cParams.GetDouble(1)
            End If
            
        Case "Unsharp mask"
            If showDialog Then
                showPDDialog vbModal, FormUnsharpMask
            Else
                FormUnsharpMask.UnsharpMask cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetLong(3)
            End If
            
            
        'Stylize
            
        Case "Antique"
            MenuAntique
        
        Case "Diffuse"
            If showDialog Then
                showPDDialog vbModal, FormDiffuse
            Else
                FormDiffuse.DiffuseCustom cParams.GetLong(1), cParams.GetLong(2), cParams.GetBool(3)
            End If
            
        Case "Dilate (maximum rank)"
            If showDialog Then
                FormMedian.showMedianDialog 100
            Else
                FormMedian.ApplyMedianFilter cParams.GetLong(1), cParams.GetDouble(2)
            End If
            
        Case "Erode (minimum rank)"
            If showDialog Then
                FormMedian.showMedianDialog 1
            Else
                FormMedian.ApplyMedianFilter cParams.GetLong(1), cParams.GetDouble(2)
            End If
        
        Case "Solarize"
            If showDialog Then
                showPDDialog vbModal, FormSolarize
            Else
                FormSolarize.SolarizeImage cParams.GetByte(1)
            End If
            
        Case "Twins"
            If showDialog Then
                showPDDialog vbModal, FormTwins
            Else
                FormTwins.GenerateTwins cParams.GetLong(1)
            End If
            
        Case "Vignetting"
            If showDialog Then
                showPDDialog vbModal, FormVignette
            Else
                FormVignette.ApplyVignette cParams.GetDouble(1), cParams.GetDouble(2), cParams.GetDouble(3), cParams.GetBool(4), cParams.GetLong(5)
            End If
            
        
        'Custom filters
        
        Case "Custom filter"
            If showDialog Then
                showPDDialog vbModal, FormCustomFilter
            Else
                DoFilter cParams.getParamString
            End If
        
        
        'SPECIAL OPERATIONS
        Case "Fade last effect"
            MenuFadeLastEffect
            
            
        'DEBUG FAILSAFE
        ' This function should never be passed a process ID it can't parse, but if that happens, ask the user to report the unparsed ID
        Case Else
            If Len(processID) > 0 Then pdMsgBox "Unknown processor request submitted: %1" & vbCrLf & vbCrLf & "Please report this bug via the Help -> Submit Bug Report menu.", vbCritical + vbOKOnly + vbApplicationModal, "Processor Error", processID
        
    End Select
    
    'If the user wants us to time this action, display the results now.  (Note - only do this for actions that will change the image
    ' in some way, as determined by the createUndo param)
    If createUndo > 0 Then
        If DISPLAY_TIMINGS Then
            Dim timingString As String
            timingString = g_Language.TranslateMessage("Time taken")
            timingString = timingString & ": " & Format$(Timer - m_ProcessingTime, "#0.####") & " "
            timingString = timingString & g_Language.TranslateMessage("seconds")
            Message timingString
        End If
    End If
    
    'Restore the mouse pointer to its default value.
    ' (NOTE: if we are in the midst of a batch conversion, leave the cursor on "busy".  The batch function will restore the cursor when done.)
    If MacroStatus <> MacroBATCH Then Screen.MousePointer = vbDefault
        
    'If the histogram form is visible and images are loaded, redraw the histogram
    'If FormHistogram.Visible Then
        'If g_OpenImageCount > 0 Then
            
            'Note that the histogram is automatically drawn when an MDI child form receives focus.  This happens below
            ' (look for pdImages(g_CurrentImage).containingForm.SetFocus), so we do not need to manually redraw the histogram here.
            'FormHistogram.TallyHistogramValues
            'FormHistogram.DrawHistogram
        'Else
            'If the histogram is visible but no images are open, unload the histogram
            'Unload FormHistogram
        'End If
    'End If
    
    'If the image has been modified and we are not performing a batch conversion (disabled to save speed!), redraw form and taskbar icons,
    ' as well as the image tab-bar.
    If (createUndo > 0) And (MacroStatus <> MacroBATCH) Then
        createCustomFormIcon pdImages(g_CurrentImage).containingForm
        toolbar_ImageTabs.notifyUpdatedImage g_CurrentImage
    End If
    
    'Unlock the main form
    FormMain.Enabled = True
    
    'If the user canceled the requested action before it completed, we need to roll back the undo data we created
    If cancelCurrentAction Then
        
        'Ask the Undo manager to roll back to a previous state
        pdImages(g_CurrentImage).undoManager.rollBackLastUndo
        
        'Reset any interface elements that may still be in "processing" mode.
        releaseProgressBar
        Message "Action canceled."
    
        'Reset the cancel trigger; if this is not done, the user will not be able to cancel subsequent actions.
        cancelCurrentAction = False
        
    End If
    
    'If a filter or tool was just used, return focus to the active form.  This will make it "flash" to catch the user's attention.
    If (createUndo > 0) Then
        If g_OpenImageCount > 0 Then pdImages(g_CurrentImage).containingForm.ActivateWorkaround "processor call complete"
    
        'Also, re-enable drag and drop operations
        g_AllowDragAndDrop = True
        FormMain.OLEDropMode = 1
    End If
    
    'The interface will automatically be synched if an image is open and some undo-related action was applied,
    ' but if either of those did not occur, sync the interface now
    syncInterfaceToCurrentImage
        
    'Finally, after all our work is done, return focus to the main PD window
    If Not g_WindowManager.getFloatState(IMAGE_WINDOW) Then g_WindowManager.requestActivation FormMain.hWnd
        
    'Mark the processor as ready
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
        Message "Out of memory.  Function canceled."
        mType = vbExclamation + vbOKOnly + vbApplicationModal
    
    'Invalid picture error
    ElseIf Err.Number = 481 Then
        AddInfo = g_Language.TranslateMessage("Unfortunately, this image file appears to be invalid.  This can happen if a file does not contain image data, or if it contains image data in an unsupported format." & vbCrLf & vbCrLf & "- If you downloaded this image from the Internet, the download may have terminated prematurely.  Please try downloading the image again." & vbCrLf & vbCrLf & "- If this image file came from a digital camera, scanner, or other image editing program, it's possible that PhotoDemon simply doesn't understand this particular file format.  Please save the image in a generic format (such as JPEG or PNG), then reload it.")
        Message "Invalid image.  Image load canceled."
        mType = vbExclamation + vbOKOnly + vbApplicationModal
    
        'Since we know about this error, there's no need to display the extended box.  Display a smaller one, then exit.
        pdMsgBox AddInfo, mType, "Invalid image file"
        
        'On an invalid picture load, there will be a blank form that needs to be dealt with.
        pdImages(g_CurrentImage).deactivateImage
        Unload pdImages(g_CurrentImage).containingForm
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
    
        'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the photodemon.org contact form
        If secondaryReturn = vbYes Then
            'Shell a browser window with the GitHub issue report form
            OpenURL "https://github.com/tannerhelland/PhotoDemon/issues/new"
            
            'Display one final message box with additional instructions
            pdMsgBox "PhotoDemon has automatically opened a GitHub bug report webpage for you.  In the Title box, please enter the following error number with a short description of the problem: " & vbCrLf & "%1" & vbCrLf & vbCrLf & "Any additional details you can provide in the large text box, including the steps that led up to this error, will help it get fixed as quickly as possible." & vbCrLf & vbCrLf & "When finished, click the Submit New Issue button.  Thank you!", vbInformation + vbApplicationModal + vbOKOnly, "GitHub bug report instructions", Err.Number
            
        Else
            'Shell a browser window with the photodemon.org contact form
            OpenURL "http://photodemon.org/about/contact/"
            
            'Display one final message box with additional instructions
            pdMsgBox "PhotoDemon has automatically opened a bug report webpage for you.  In the Comment box, please describe the steps that led to this error." & vbCrLf & vbCrLf & "Also in that box, please include the following error number: " & vbCrLf & "%1" & vbCrLf & vbCrLf & "When finished, click the Submit button.  Thank you!", vbInformation + vbApplicationModal + vbOKOnly, "Bug report instructions", Err.Number
            
        End If
    
    End If
        
End Sub
