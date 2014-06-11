Attribute VB_Name = "Loading"
'***************************************************************************
'Program/File Loading Handler
'Copyright ©2001-2014 by Tanner Helland
'Created: 4/15/01
'Last updated: 19/May/14
'Last update: final work on custom Undo/Redo file loading
'
'Module for handling any and all program loading.  This includes the program itself,
' plugins, files, and anything else the program needs to take from the hard drive.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Use these to ensure that the splash shows for a certain amount of time
Dim m_LoadTime As Double, m_StartTime As Double

'IT ALL BEGINS HERE (after Sub Main, that is).
' Note that this function is called AFTER FormMain has been loaded.  FormMain is loaded - but not visible - so it can be operated
' on by functions called from this routine.  (It is necessary to load the main form first, since a number of these operations -
' like loading all PNG menu icons from the resource file - operate on the main form.)
Public Sub LoadTheProgram()
    
    m_StartTime = Timer
    
    '*************************************************************************************************************************************
    ' Prepare the splash screen (but don't display it yet)
    '*************************************************************************************************************************************
    
    'We need GDI+ to extract a JPEG from the resource file and convert it in-memory.  (Yes, there are other ways to do this.  No, I don't
    ' care about using them.)  Check its availability.
    If isGDIPlusAvailable() Then
    
        'Load FormSplash into memory, but don't make it visible.  Then ask it to prepare itself.
        FormSplash.Visible = False
        FormSplash.prepareSplash
        
    End If
        
    'Check the environment.  If inside the the IDE, the splash needs to be modified slightly.
    CheckLoadingEnvironment
    
    If g_GDIPlusAvailable Then
        If g_IsProgramCompiled Then m_LoadTime = 1.5 Else m_LoadTime = 1#
    Else
        m_LoadTime = 0#
    End If
    
    '*************************************************************************************************************************************
    ' Determine which version of Windows the user is running (as other load functions rely on this)
    '*************************************************************************************************************************************
    
    LoadMessage "Detecting Windows® version..."
    
    'Certain features are OS-version dependent.  We must determine the OS version early in the load process to ensure that all features
    ' are loaded correctly.
    g_IsVistaOrLater = getVistaOrLaterStatus
    g_IsWin7OrLater = getWin7OrLaterStatus
    
    'If we are on Windows 7, prepare some Win7-specific features (like taskbar progress bars)
    If g_IsWin7OrLater Then prepWin7Features
    
    
    
    '*************************************************************************************************************************************
    ' If the user doesn't have font smoothing enabled, enable it now.  PD's interface looks much better with some form of antialiasing.
    '*************************************************************************************************************************************
    
    handleClearType True
    
    
    
    '*************************************************************************************************************************************
    ' Initialize the user preferences (settings) handler
    '*************************************************************************************************************************************
    
    Set g_UserPreferences = New pdPreferences
    
    'Ask the preferences handler to generate key program folders.  (If these folders don't exist, the handler will create them.)
    LoadMessage "Initializing all program directories..."
    
    g_UserPreferences.initializePaths
    
    'Now, ask the preferences handler to load all other user settings from the preferences file
    LoadMessage "Loading all user settings..."
    
    g_UserPreferences.loadUserSettings
            
    'While here, also initialize the image format handler (as plugins and other load functions interact with it)
    Set g_ImageFormats = New pdFormats
    
    
    
    '*************************************************************************************************************************************
    ' PhotoDemon works very well with multiple monitors.  Check for such a situation now.
    '*************************************************************************************************************************************
    
    LoadMessage "Analyzing current monitor setup..."
    
    Set g_cMonitors = New clsMonitors
    g_cMonitors.Refresh
    
    'While here, also cache the current color management settings in use by the system
    cacheCurrentSystemColorProfile
    
    
    
    '*************************************************************************************************************************************
    ' Now we have what we need to properly display the splash screen.  Do so now.
    '*************************************************************************************************************************************
        
    'Determine the program's previous on-screen location.  We need that to determine where to display the splash screen.
    Dim wRect As RECT
    wRect.Left = g_UserPreferences.GetPref_Long("Core", "Last Window Left", 1)
    wRect.Top = g_UserPreferences.GetPref_Long("Core", "Last Window Top", 1)
    wRect.Right = wRect.Left + g_UserPreferences.GetPref_Long("Core", "Last Window Width", 1)
    wRect.Bottom = wRect.Top + g_UserPreferences.GetPref_Long("Core", "Last Window Height", 1)
    
    'Center the splash screen on whichever monitor the user previously used.
    g_cMonitors.CenterFormOnMonitor FormSplash, , wRect.Left, wRect.Right, wRect.Top, wRect.Bottom
    
    'Make the splash screen's message display match the rest of PD
    If g_IsVistaOrLater And g_UseFancyFonts Then
        g_InterfaceFont = "Segoe UI"
    Else
        g_InterfaceFont = "Tahoma"
    End If
    FormSplash.lblMessage.FontName = g_InterfaceFont
    
    'Display the splash screen, centered on whichever monitor the user previously used the program on.
    FormSplash.Show vbModeless
    
    
    '*************************************************************************************************************************************
    ' Initialize the translation (language) engine
    '*************************************************************************************************************************************
    
    'Initialize a new language engine.
    Set g_Language = New pdTranslate
        
    LoadMessage "Scanning for language files..."
    
    'Before doing anything else, check to see what languages are available in the language folder.
    ' (Note that this function will also populate the Languages menu, though it won't place a checkmark next to an entry yet.)
    g_Language.CheckAvailableLanguages
        
    LoadMessage "Determining which language to use..."
        
    'Next, determine which language to use.  (This function will take into account the system language at first-run, so it can
    ' estimate which language to present to the user.)
    g_Language.DetermineLanguage
    
    LoadMessage "Applying selected language..."
    
    'Apply that language to the program.  This involves loading the translation file into memory, which can take a bit of time,
    ' but it only needs to be done once.  From that point forward, any text requests will operate on the in-memory copy of the file.
    g_Language.ApplyLanguage
    
    
    
    '*************************************************************************************************************************************
    ' Check for the presence of plugins (as other functions rely on these to initialize themselves)
    '*************************************************************************************************************************************
    
    LoadMessage "Loading plugins..."
    
    LoadPlugins
    
    '(Note that LoadPlugins also checks GDI+ availability, despite GDI+ not really being a "plugin")
    
    'If ExifTool was enabled successfully, ask it to double-check that its tag database has been created
    ' successfully at some point in the past.  If it hasn't, generate a new copy now.
    '
    'Now that this has been thoroughly tested, I'm postponing actual enabling of it until PD supports metadata editing.
    'If g_ExifToolEnabled Then writeTagDatabase
    
    
    
    '*************************************************************************************************************************************
    ' Based on available plugins, determine which image formats PhotoDemon can handle
    '*************************************************************************************************************************************
        
    LoadMessage "Loading import/export libraries..."
        
    g_ImageFormats.generateInputFormats
    g_ImageFormats.generateOutputFormats
    
    
    
    '*************************************************************************************************************************************
    ' Initialize the visual themes engine
    '*************************************************************************************************************************************
    
    'Because this class subclasses all forms in the project, it must be loaded very early in the start process
    LoadMessage "Initializing theme engine..."
    
    Set g_Themer = New pdVisualThemes
    
    
    
    '*************************************************************************************************************************************
    ' Get the viewport engine ready
    '*************************************************************************************************************************************
    
    'Initialize our current zoom method
    LoadMessage "Initializing viewport engine..."
    
    'Create the program's primary zoom handler
    Set g_Zoom = New pdZoom
    g_Zoom.initializeViewportEngine
    
    'Populate the main form's zoom drop-down
    g_Zoom.populateZoomComboBox FormMain.mainCanvas(0).getZoomDropDownReference()
    
    'Populate the main canvas's size unit dropdown
    FormMain.mainCanvas(0).populateSizeUnits
    
    '*************************************************************************************************************************************
    ' Initialize the window manager (the class that synchronizes all toolbox and image window positions)
    '*************************************************************************************************************************************
    
    LoadMessage "Initializing window manager..."
    Set g_WindowManager = New pdWindowManager
    
    'Register the main form
    g_WindowManager.registerParentForm FormMain
    
    'Load all tool windows.  Even though they may not be visible (as the user can elect to hide them), we still want them loaded,
    ' so we can interact with them as necessary (e.g. "enable Undo button", etc).
    Load toolbar_File
    Load toolbar_ImageTabs
    Load toolbar_Tools
        
    'Retrieve floating window status from the preferences file, mark their menus, and pass their values to the window manager
    toggleWindowFloating TOOLBAR_WINDOW, g_UserPreferences.GetPref_Boolean("Core", "Floating Toolbars", False), True
    
    'Retrieve visibility and mark those menus as well
    FormMain.MnuWindow(0).Checked = g_UserPreferences.GetPref_Boolean("Core", "Show File Toolbox", True)
    FormMain.MnuWindow(1).Checked = g_UserPreferences.GetPref_Boolean("Core", "Show Selections Toolbox", True)
    FormMain.MnuWindow(2).Checked = g_UserPreferences.GetPref_Boolean("Core", "Show Layers Toolbox", True)
    
    'Retrieve two additional settings for the image tabstrip menu: when to display the image tabstrip...
    toggleImageTabstripVisibility g_UserPreferences.GetPref_Long("Core", "Image Tabstrip Visibility", 1), True
        
    '...and the alignment of the tabstrip
    toggleImageTabstripAlignment g_UserPreferences.GetPref_Long("Core", "Image Tabstrip Alignment", vbAlignTop)
    
    
    
    '*************************************************************************************************************************************
    ' Set all default tool values
    '*************************************************************************************************************************************
        
    LoadMessage "Initializing image tools..."
        
    'Note that selection tools are initialized in the Tool toolbar's Form_Load event
    
    
    
    '*************************************************************************************************************************************
    ' PhotoDemon's complex interface requires a lot of things to be generated at run-time.
    '*************************************************************************************************************************************
    
    LoadMessage "Initializing user interface..."
    
    'Initialize the drop shadow engine
    Set g_CanvasShadow = New pdShadow
    g_CanvasShadow.initializeSquareShadow PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSTRENGTH, g_CanvasBackground
    
    'Manually create multi-line tooltips for some command buttons
    toolbar_File.cmdOpen.ToolTip = g_Language.TranslateMessage("Open one or more images for editing." & vbCrLf & vbCrLf & "(Another way to open images is dragging them from your desktop or Windows Explorer and dropping them onto PhotoDemon.)")
    If g_ConfirmClosingUnsaved Then
        toolbar_File.cmdClose.ToolTip = g_Language.TranslateMessage("Close the current image." & vbCrLf & vbCrLf & "If the current image has not been saved, you will receive a prompt to save it before it closes.")
    Else
        toolbar_File.cmdClose.ToolTip = g_Language.TranslateMessage("Close the current image." & vbCrLf & vbCrLf & "Because you have turned off save prompts (via Edit -> Preferences), you WILL NOT receive a prompt to save this image before it closes.")
    End If
    toolbar_File.cmdSave.ToolTip = g_Language.TranslateMessage("Save the current image." & vbCrLf & vbCrLf & "WARNING: this will overwrite the current image file.  To save to a different file, use the ""Save As"" button.")
    toolbar_File.cmdSaveAs.ToolTip = g_Language.TranslateMessage("Save the current image to a new file.")
                        
    'Use the API to give PhotoDemon's main form a 32-bit icon (VB is too old to support 32bpp icons)
    SetIcon FormMain.hWnd, "AAA", True
    
    'Initialize all system cursors we rely on (hand, busy, resizing, etc)
    initAllCursors
    
    'Set up the program's title bar.  Odd-numbered releases are development releases.  Even-numbered releases are formal builds.
    FormMain.Caption = getPhotoDemonNameAndVersion()
    
    'PhotoDemon renders many of its own icons dynamically.  Initialize that engine now.
    initializeIconHandler
    
    'Prepare a checkerboard pattern, which will be used behind any transparent objects.  Caching this is much more efficient.
    ' than re-creating it every time it's needed.
    Set g_CheckerboardPattern = New pdDIB
    Drawing.createAlphaCheckerboardDIB g_CheckerboardPattern
    
    'Allow drag-and-drop operations
    g_AllowDragAndDrop = True
    
    'Set the main canvas background color
    FormMain.mainCanvas(0).BackColor = g_CanvasBackground
    
    'Clear the main canvas coordinate and size displays
    FormMain.mainCanvas(0).displayCanvasCoordinates 0, 0, True
    FormMain.mainCanvas(0).displayImageSize Nothing, True
    
    'Throughout the program, g_MouseAccuracy is used to determine how close the mouse cursor must be to a point of interest to
    ' consider it "over" that point.  DPI must be accounted for when calculating this value (as it's calculated in pixels).
    g_MouseAccuracy = fixDPIFloat(6)
    
    'Apply visual styles
    FormMain.requestMakeFormPretty
    
    
    '*************************************************************************************************************************************
    ' The program's menus support many features that VB can't do natively (like icons and custom shortcuts).  Load such things now.
    '*************************************************************************************************************************************
    
    LoadMessage "Preparing program menus..."
    
    'If inside the IDE, disable the "Effects" -> "Test" menu
    If g_IsProgramCompiled Then
        FormMain.MnuTest.Visible = False
        'FormMain.MnuEffectExperimental.Visible = False
    Else
        FormMain.MnuTest.Visible = True
        'FormMain.MnuEffectExperimental.Visible = True
    End If
    
    'Create all manual shortcuts (ones VB isn't capable of generating itself)
    LoadAccelerators
            
    'Initialize the Recent Files manager and load the most-recently-used file list (MRU)
    Set g_RecentFiles = New pdRecentFiles
    g_RecentFiles.MRU_LoadFromFile
            
    'Load and draw all menu icons
    loadMenuIcons
    'resetMenuIcons
    
    'Synchronize all other interface elements to match the current program state (e.g. no images loaded).
    syncInterfaceToCurrentImage
    
    
    '*************************************************************************************************************************************
    ' Unload the splash screen and present the main form
    '*************************************************************************************************************************************
        
    'Display the splash screen for at least a second or two
    If Timer - m_StartTime < m_LoadTime Then
        Do While Timer - m_StartTime < m_LoadTime
        Loop
    End If
    
    'If this is the first time the user has run PhotoDemon, resize the window a bit to make the default position nice.
    ' (If this is *not* the first time, the window manager will automatically restore the window's last-known position and state.)
    If g_IsFirstRun Then g_WindowManager.setFirstRunMainWindowPosition
    
    'Display the main form
    FormMain.Show
    
    Unload FormSplash
        
End Sub

'If files are present in the command line, this sub will load them
Public Sub LoadImagesFromCommandLine()

    Message "Loading image(s)..."
        
    'NOTE: Windows will pass multiple filenames via the command line, but it does so in a confusing and overly complex way.
    ' Specifically, quotation marks are placed around filenames IFF they contain a space; otherwise, file names are separated from
    ' neighboring filenames by a space.  This creates a problem when passing a mixture of filenames with spaces and filenames without,
    ' because Windows will switch between using and not using quotation marks to delimit the filenames.  Thus, we must perform complex,
    ' specialized parsing of the command line.
        
    'This array will ultimately contain each filename to be loaded (one filename per index)
    Dim sFile() As String
        
    'First, check the command line for quotation marks
    If InStr(g_CommandLine, Chr(34)) = 0 Then
        
        'If there aren't any, our work is simple - simply split the array using the "space" character as the delimiter
        sFile = Split(g_CommandLine, Chr(32))
        
    'If there are quotation marks, things get a lot messier.
    Else
        
        Dim inQuotes As Boolean
        inQuotes = False
        
        Dim tChar As String
        
        'Scan the command line one character at a time
        Dim i As Long
        For i = 1 To Len(g_CommandLine)
            
            tChar = Mid(g_CommandLine, i, 1)
                
            'If the current character is a quotation mark, change inQuotes to specify that we are either inside
            ' or outside a SET of quotation marks (note: they will always occur in pairs, per the rules of
            ' how Windows handles command line parameters)
            If tChar = Chr(34) Then inQuotes = Not inQuotes
                
            'If the current character is a space...
            If tChar = Chr(32) Then
                    
                '...check to see if we are inside quotation marks.  If we are, that means this space is part of a
                ' filename and NOT a delimiter.  Replace it with an asterisk.
                If inQuotes = True Then g_CommandLine = Left(g_CommandLine, i - 1) & "*" & Right(g_CommandLine, Len(g_CommandLine) - i)
                    
            End If
            
        Next i
            
        'At this point, spaces that are parts of filenames have been replaced by asterisks.  That means we can use
        ' Split() to fill our filename array, because the only spaces remaining in the command line are delimiters
        ' between filenames.
        sFile = Split(g_CommandLine, Chr(32))
            
        'Now that our filenames are successfully inside the sFile() array, go back and replace our asterisk placeholders
        ' with spaces.  Also, remove any quotation marks (since those aren't technically part of the filename).
        For i = 0 To UBound(sFile)
            sFile(i) = Replace$(sFile(i), Chr(42), Chr(32))
            sFile(i) = Replace$(sFile(i), Chr(34), "")
        Next i
        
    End If
    
    'Finally, pass the array of filenames to the image loading routine
    LoadFileAsNewImage sFile

End Sub

'Loading an image begins here.  This routine examines a given file's extension and re-routes control based on that.
Public Sub LoadFileAsNewImage(ByRef sFile() As String, Optional ByVal ToUpdateMRU As Boolean = True, Optional ByVal imgFormTitle As String = "", Optional ByVal imgName As String = "", Optional ByVal isThisPrimaryImage As Boolean = True, Optional ByRef targetImage As pdImage, Optional ByRef targetDIB As pdDIB, Optional ByVal pageNumber As Long = 0)
        
    '*************************************************************************************************************************************
    ' Prepare all variables related to image loading
    '*************************************************************************************************************************************
    
    'Display a busy cursor
    If Screen.MousePointer <> vbHourglass Then Screen.MousePointer = vbHourglass
            
    'One of the things we'll be doing in this routine is establishing an original color depth for this image. FreeImage will return
    ' this automatically; GDI+ may not.  Use a tracking variable to determine if a manual color count needs to be performed.
    Dim mustCountColors As Boolean
    Dim colorCountCheck As Long
            
    'To improve load time, declare a variety of other variables outside the image load loop
    Dim FileExtension As String
    Dim loadSuccessful As Boolean
    
    Dim loadedByOtherMeans As Boolean
    loadedByOtherMeans = False
            
    'If multiple files are being loaded, we want to suppress all warnings and errors until the very end.
    Dim multipleFilesLoading As Boolean
    If UBound(sFile) > 0 Then multipleFilesLoading = True Else multipleFilesLoading = False
    
    Dim missingFiles As String
    missingFiles = ""
    
    Dim brokenFiles As String
    brokenFiles = ""
            
    
    '*************************************************************************************************************************************
    'Before actually loading anything, perform a one-time check to make sure the metadata engine isn't still busy
    ' processing an initial database build.
    '*************************************************************************************************************************************
    
    If g_ExifToolEnabled And isDatabaseModeActive Then
        
        'Wait for metadata parsing to finish...
        If Not isMetadataFinished Then
        
            Message "Finishing final program initialization steps..."
        
            'Forcibly disable the main form to avoid DoEvents allowing click-through
            FormMain.Enabled = False
        
            'Pause for 1/10 second
            Do
                PauseProgram 0.1
                
                'If the user shuts down the program while we are still waiting for input, exit immediately
                If g_ProgramShuttingDown Then Exit Sub
                
            Loop While (Not isMetadataFinished)
            
            'Re-enable the main form
            FormMain.Enabled = True
            
        End If
        
    End If
        
        
    '*************************************************************************************************************************************
    ' To prevent re-entry problems, forcibly disable the main form
    '*************************************************************************************************************************************
    
    FormMain.Enabled = False
    
    
            
    '*************************************************************************************************************************************
    ' Loop through each entry in the sFile() array, loading images as we go
    '*************************************************************************************************************************************
            
    'Because this routine accepts an array of images, we have to be prepared for the possibility that more than
    ' one image file is being opened.  This loop will execute until all files are loaded.
    Dim thisImage As Long
    
    For thisImage = 0 To UBound(sFile)
    
    
    
        '*************************************************************************************************************************************
        ' Reset all variables used on a per-image level
        '*************************************************************************************************************************************
    
        'Reset the multipage checker (which is now handled on a per-image basis)
        g_imageHasMultiplePages = False
        
        '...and reset the "need to check colors" variable.  If FreeImage is used, color depth of the source file is retrieved automatically.
        ' If FreeImage is not used, we manually calculate a bit-depth for incoming images.
        mustCountColors = False
    
    
    
        '*************************************************************************************************************************************
        ' Before attempting to load this image, make sure it exists
        '*************************************************************************************************************************************
    
        'If isThisPrimaryImage Then Message "Verifying that file exists..."
    
        If isThisPrimaryImage And (Not FileExist(sFile(thisImage))) Then
            
            'If multiple files are being loaded, suppress any errors until the end
            If multipleFilesLoading Then
                missingFiles = missingFiles & getFilename(sFile(thisImage)) & vbCrLf
            Else
                pdMsgBox "Unfortunately, the image '%1' could not be found." & vbCrLf & vbCrLf & "If this image was originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbApplicationModal + vbExclamation + vbOKOnly, "File not found", sFile(thisImage)
            End If
            
            'If the missing image was part of a list of images, try loading the next entry in the list
            GoTo PreloadMoreImages
            
        End If
        
        
        
        '*************************************************************************************************************************************
        ' If the image being loaded is a primary image (e.g. one opened normally), prepare a blank form to receive it
        '*************************************************************************************************************************************
        
        If isThisPrimaryImage Then
            
            CreateNewPDImage
            
            'If this is a primary image, we will automatically set the targetImage and targetDIB parameters.  If this is NOT a primary image,
            ' the calling function must have specified this for us.
            Set targetImage = pdImages(g_CurrentImage)
            
            'Create a blank layer in the receiving image, and retrieve a pointer to it
            Dim newLayerID As Long
            newLayerID = pdImages(g_CurrentImage).createBlankLayer
            
            Set targetDIB = New pdDIB
            
            g_AllowViewportRendering = False
            
            'Reset the main viewport's scroll bars
            FormMain.mainCanvas(0).setScrollValue PD_BOTH, 0
            
        End If
        
        
        
        '*************************************************************************************************************************************
        ' If the ExifTool plugin is available, initiate a separate thread for metadata extraction
        '*************************************************************************************************************************************
        
        'Note that metadata extraction is handled asynchronously (e.g. in parallel to the core image loading process), which is
        ' why we launch it so early in the load process.  If the image load fails, we simply ignore any received metadata.
        ' ExifTool is extremely robust, so any errors it experiences during processing will not affect PD.
        Set targetImage.imgMetadata = New pdMetadata
        
        If g_ExifToolEnabled And isThisPrimaryImage Then
            Message "Starting separate metadata extraction thread..."
            startMetadataProcessing sFile(thisImage), targetImage.originalFileFormat
        End If
        
        'By default, set this image to use the program's default metadata setting (settable from Tools -> Options).
        ' The user may override this setting later, but by default we always start with the user's program-wide setting.
        targetImage.imgMetadata.setMetadataExportPreference g_UserPreferences.GetPref_Long("Saving", "Metadata Export", 1)
        
        
            
        '*************************************************************************************************************************************
        ' Call the most appropriate load function for this image's format (FreeImage, GDI+, or VB's LoadPicture)
        '*************************************************************************************************************************************
            
        If isThisPrimaryImage Then Message "Determining filetype..."
        
        'Initially, set the filetype of the target image to "unknown".  If the load is successful, this value will
        ' be changed to something >= 0. (Note: if FreeImage is used to load the file, this value will be set by the
        ' LoadFreeImageV4 function.)
        If Not (targetImage Is Nothing) Then targetImage.originalFileFormat = -1
        
        'Strip the extension from the file
        FileExtension = UCase(GetExtension(sFile(thisImage)))
        
        loadSuccessful = False
        loadedByOtherMeans = False
            
        'Depending on the file's extension, load the image using the most appropriate image decoding routine
        Select Case FileExtension
        
            'PhotoDemon's custom file format must be handled specially (as obviously, FreeImage and GDI+ won't handle it!)
            Case "PDI", "PDTMP"
            
                'PDI images require zLib, and are only loaded via a custom routine (obviously, since they are PhotoDemon's native format)
                loadSuccessful = LoadPhotoDemonImage(sFile(thisImage), targetDIB, targetImage)
                
                targetImage.originalFileFormat = 100
                targetImage.currentFileFormat = 100
                mustCountColors = False
            
            'Straight TMP files are internal files (BMP format) used by PhotoDemon.  GDI+ is preferable for loading these, as
            ' it handles 32bpp images well, but if we must, we can resort to VB's internal .LoadPicture command.
            Case "TMP"
            
                If g_ImageFormats.GDIPlusEnabled Then loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetDIB)
                
                If (Not g_ImageFormats.GDIPlusEnabled) Or (Not loadSuccessful) Then loadSuccessful = LoadVBImage(sFile(thisImage), targetDIB)
                
                'Lie and say that the original file format of this image was JPEG.  We do this because tmp images are typically images
                ' captured via non-traditional means (screenshot's, scans), and when the user tries to save the file, they should not
                ' be prompted to save it as a BMP.
                targetImage.originalFileFormat = FIF_JPEG
                mustCountColors = True
            
            'All other formats follow a set pattern: try to load them via FreeImage (if available), then GDI+, then finally
            ' VB's internal LoadPicture function.
            Case Else
                                
                'If FreeImage is available, we first use it to try and load the image.
                If g_ImageFormats.FreeImageEnabled Then
                
                    'Start by seeing if the image file contains multiple pages.  If it does, we will load each page as a separate layer.
                    If isMultiImage(sFile(thisImage)) > 0 Then
                        
                        'TODO: load each page individually, to a unique layer
                        
                        'TEMP SOLUTION: just load the first page
                        pageNumber = 0
                        loadSuccessful = LoadFreeImageV4(sFile(thisImage), targetDIB, pageNumber, isThisPrimaryImage)
                     
                    'The image only has one page.  Load it!
                    Else
                        
                        pageNumber = 0
                        loadSuccessful = LoadFreeImageV4(sFile(thisImage), targetDIB, pageNumber, isThisPrimaryImage)
                    
                    End If
                    
                    'FreeImage worked!  Copy any relevant information from the DIB to the parent pdImage object (such as file format),
                    ' then continue with the load process.
                    If loadSuccessful Then
                    
                        loadedByOtherMeans = False
                        
                        'Mirror the determined file format from the DIB to the parent pdImage object
                        targetImage.originalFileFormat = targetDIB.getOriginalFormat
                        
                        'Mirror the discovered resolution, if any, from the DIB
                        targetImage.setDPI targetDIB.getDPI, targetDIB.getDPI
                        
                        'Mirror the original file's color depth
                        targetImage.originalColorDepth = targetDIB.getOriginalColorDepth
                        
                        'Finally, copy the background color (if any) from the DIB
                        targetImage.pngBackgroundColor = targetDIB.getBackgroundColor
                        
                    End If
                    
                End If
                
                'If FreeImage fails for some reason, offload the image to GDI+ - UNLESS the image is a WMF or EMF, which can cause
                ' GDI+ to experience a silent fail, thus bringing down the entire program.
                If (Not loadSuccessful) And g_ImageFormats.GDIPlusEnabled And ((FileExtension <> "EMF") And (FileExtension <> "WMF")) Then
                    
                    If isThisPrimaryImage Then Message "FreeImage refused to load image.  Dropping back to GDI+ and trying again..."
                    loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetDIB)
                    
                    'If GDI+ loaded the image successfully, note that we have to determine color depth manually.  (There may be a way
                    ' to retrieve that info from GDI+, but I haven't bothered to look!)
                    If loadSuccessful Then
                    
                        loadedByOtherMeans = True
                        mustCountColors = True
                        
                        'Also, mirror the discovered resolution, if any, from the source DIB
                        targetImage.setDPI targetDIB.getDPI, targetDIB.getDPI
                        
                    End If
                        
                End If
                
                'If both FreeImage and GDI+ failed, give the image one last try with VB's LoadPicture
                If (Not loadSuccessful) Then
                    
                    If isThisPrimaryImage Then Message "GDI+ refused to load image.  Dropping back to internal routines and trying again..."
                    loadSuccessful = LoadVBImage(sFile(thisImage), targetDIB)
                
                    'If VB managed to load the image successfully, note that we have to deteremine color depth manually
                    If loadSuccessful Then
                        loadedByOtherMeans = True
                        mustCountColors = True
                    End If
                
                End If
                    
        End Select
        
        DoEvents
        
        '*************************************************************************************************************************************
        ' Run a few checks to confirm that the image data was loaded successfully
        '*************************************************************************************************************************************
        
        'Sometimes, our image load functions will think the image loaded correctly, but they will return a blank image.  Check for
        ' non-zero width and height before continuing.
        If ((Not loadSuccessful) Or (targetDIB.getDIBWidth = 0) Or (targetDIB.getDIBHeight = 0)) And isThisPrimaryImage Then
            
            Message "Failed to load %1", sFile(thisImage)
            
            'If multiple files are being loaded, suppress any errors until the end
            If multipleFilesLoading Then
                brokenFiles = brokenFiles & getFilename(sFile(thisImage)) & vbCrLf
            Else
                If MacroStatus <> MacroBATCH Then pdMsgBox "Unfortunately, PhotoDemon was unable to load the following image:" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "Please use another program to save this image in a generic format (such as JPEG or PNG) before loading it into PhotoDemon.  Thanks!", vbExclamation + vbOKOnly + vbApplicationModal, "Image Import Failed", sFile(thisImage)
            End If
            
            'Deactivate the (now useless) pdImage object, and forcibly unload whatever resources it has claimed
            targetImage.deactivateImage
            fullPDImageUnload targetImage.imageID
            
            'Update the interface to reflect the images currently loaded
            syncInterfaceToCurrentImage
            
            GoTo PreloadMoreImages
            
        Else
            If isThisPrimaryImage Then Message "Image data loaded successfully."
        End If
        
        DoEvents
        
        
        '*************************************************************************************************************************************
        ' If the loaded image was in PDI format (PhotoDemon's internal format), skip a number of additional processing steps.
        '*************************************************************************************************************************************
        
        '100 is the magic number for PDI files.  I don't generally like using magic numbers in place of enums, but as our
        ' file format enum list is declared in the external FreeImage module - which is occasionally overwritten by official
        ' FreeImage updates - I can't easily add my own enums to theirs.  So this will do, for now...
        If targetImage.originalFileFormat <> 100 Then
            
            
            '*************************************************************************************************************************************
            ' If GDI+ or VB's LoadPicture was used to load the file, populate some data fields manually (filetype, color depth, etc)
            '*************************************************************************************************************************************
            
            If loadedByOtherMeans Then
            
                Select Case FileExtension
                    
                    Case "GIF"
                        targetImage.originalFileFormat = FIF_GIF
                        targetImage.originalColorDepth = 8
                        
                    Case "ICO"
                        targetImage.originalFileFormat = FIF_ICO
                    
                    Case "JIF", "JPG", "JPEG", "JPE"
                        targetImage.originalFileFormat = FIF_JPEG
                        targetImage.originalColorDepth = 24
                        
                    Case "PNG"
                        targetImage.originalFileFormat = FIF_PNG
                    
                    Case "TIF", "TIFF"
                        targetImage.originalFileFormat = FIF_TIFF
                        
                    'Treat anything else as a BMP file
                    Case Else
                        targetImage.originalFileFormat = FIF_BMP
                        
                End Select
            
            End If
            
            DoEvents
            
            
            '*************************************************************************************************************************************
            ' If the incoming image is 24bpp, convert it to 32bpp.  PD assumes an available alpha channel for all layers.
            '*************************************************************************************************************************************
            
            If targetDIB.getDIBColorDepth = 24 Then targetDIB.convertTo32bpp
                    
            
            '*************************************************************************************************************************************
            ' If the image contained an embedded ICC profile, apply it now (before counting colors, etc).
            '*************************************************************************************************************************************
            
            'Note that we now need to see if the ICC profile has already been applied.  For CMYK images, the ICC profile will be applied by
            ' the image load function.  If we don't do this, we'll be left with a 32bpp image that contains CMYK data instead of RGBA!
            If targetDIB.ICCProfile.hasICCData And (Not targetDIB.ICCProfile.hasProfileBeenApplied) Then
                
                '32bpp images must be un-premultiplied before the transformation
                If targetDIB.getDIBColorDepth = 32 Then targetDIB.fixPremultipliedAlpha
                
                'Apply the ICC transform
                targetDIB.ICCProfile.applyICCtoSelf targetDIB
                
                '32bpp images must be re-premultiplied after the transformation
                If targetDIB.getDIBColorDepth = 32 Then targetDIB.fixPremultipliedAlpha True
                
            End If
            
            DoEvents
            
            
            '*************************************************************************************************************************************
            ' The target DIB has been loaded successfully, so copy its contents into the main layer of the targetImage
            '*************************************************************************************************************************************
            
            If isThisPrimaryImage Then
                If Len(imgName) = 0 Then
                    targetImage.getLayerByID(newLayerID).CreateNewImageLayer targetDIB, targetImage, getFilenameWithoutExtension(sFile(thisImage))
                Else
                    targetImage.getLayerByID(newLayerID).CreateNewImageLayer targetDIB, targetImage, imgName
                End If
            End If
            
            'Update the pdImage container to be the same size as its (newly created) base layer
            targetImage.updateSize
            
            DoEvents
            
            
            '*************************************************************************************************************************************
            ' If requested by the user, manually count the number of unique colors in the image (to accurately determine color depth)
            '*************************************************************************************************************************************
                    
            'At this point, we now have loaded image data in 24 or 32bpp format.  For future reference, let's count
            ' the number of colors present in the image (if the user has allowed it).  If the user HASN'T allowed
            ' it, we have no choice but to rely on whatever color depth was returned by FreeImage or GDI+ (or was
            ' inferred by us for this format, e.g. we know that GIFs are 8bpp).
            
            If isThisPrimaryImage And (g_UserPreferences.GetPref_Boolean("Loading", "Verify Initial Color Depth", True) Or mustCountColors) Then
                
                colorCountCheck = getQuickColorCount(targetDIB, g_CurrentImage)
            
                'If 256 or less colors were found in the image, mark it as 8bpp.  Otherwise, mark it as 24 or 32bpp.
                targetImage.originalColorDepth = getColorDepthFromColorCount(colorCountCheck, targetDIB)
                
                If g_IsImageGray Then
                    Message "Color count successful (%1 BPP, grayscale)", targetImage.originalColorDepth
                Else
                    Message "Color count successful (%1 BPP, color)", targetImage.originalColorDepth
                End If
                            
            End If
            
            DoEvents
        
        End If
                
        '*************************************************************************************************************************************
        ' Determine a name for this image, and store it (along with any other relevant bits) inside the parent pdImage object
        '*************************************************************************************************************************************
        
        
        'Note: this is where PDI format processing picks up again
PDI_Load_Continuation:

        
        'Mark the original file size and file format of the image
        If FileExist(sFile(thisImage)) Then targetImage.originalFileSize = FileLen(sFile(thisImage))
        targetImage.currentFileFormat = targetImage.originalFileFormat
        
        If isThisPrimaryImage Then Message "Determining image title..."
        
        'If a different image name has been specified, we can assume the calling routine is NOT loading a file
        ' from disk (e.g. it's a scan, or Internet download, or screen capture, etc.).  Therefore, set the
        ' file name as requested but leave the .LocationOnDisk blank so that a Save command will trigger
        ' the necessary Save As... dialog.
        Dim tmpFilename As String
        
        'Autosaved images are handled differently from normal images.  In order to preserve their original data,
        ' we load certain image data from a standalone XML file.
        If FileExtension = "PDTMP" Then
            
            targetImage.locationOnDisk = sFile(thisImage)
            
            'Ask the AutoSave engine to retrieve this image's data from the matching XML autosave file
            Image_Autosave_Handler.alignLoadedImageWithAutosave targetImage
            
            'This is a bit wacky, but - the MRU engine will automatically update this entry based on its location
            ' on disk (per PD convention) AS STORED IN THE sFile ARRAY.  But as this file's location on disk is
            ' a temp file, we need to rewrite its sFile entry mid-loading!
            sFile(thisImage) = targetImage.locationOnDisk
        
        'This is a non-autosave (normal!) image.
        Else
        
            If Len(imgName) = 0 Then
                'The calling routine hasn't specified an image name, so assume this is a normal load situation.
                ' That means pulling the filename from the file itself.
                targetImage.locationOnDisk = sFile(thisImage)
                
                tmpFilename = sFile(thisImage)
                StripFilename tmpFilename
                targetImage.originalFileNameAndExtension = tmpFilename
                StripOffExtension tmpFilename
                targetImage.originalFileName = tmpFilename
                
                'Disable the save button, because this file exists on disk
                targetImage.setSaveState True
                
            Else
            
                'The calling routine has specified a file name.  Assume this is a special case, and force a Save As...
                ' dialog in the future by not specifying a location on disk
                targetImage.locationOnDisk = ""
                targetImage.originalFileNameAndExtension = imgName
                
                tmpFilename = imgName
                StripOffExtension tmpFilename
                targetImage.originalFileName = tmpFilename
                
                'Similarly, enable the save button
                targetImage.setSaveState False
                
            End If
        
        End If
        
        DoEvents
        
        
        '*************************************************************************************************************************************
        ' If this is a primary image, update all relevant interface elements (image size display, 24/32bpp options, custom form icon, etc)
        '*************************************************************************************************************************************
                
        'If this is a primary image, it needs to be rendered to the screen
        If isThisPrimaryImage Then
            
            'Create an icon-sized version of this image, which we will use as form's taskbar icon
            If MacroStatus <> MacroBATCH Then createCustomFormIcon targetImage
            
            'Synchronize all other interface elements to match the newly loaded image
            syncInterfaceToCurrentImage
            
            DoEvents
        
        
        
        '*************************************************************************************************************************************
        ' If this is a primary image, do a few additional preparations, then render the image to the screen
        '*************************************************************************************************************************************
            
            
            'Register this image with the image tab bar
            toolbar_ImageTabs.registerNewImage g_CurrentImage
            
            'Just to be safe, update the color management profile of the current monitor
            checkParentMonitor True
            
            Message "Resizing image to fit screen..."
    
            'If the user wants us to resize the image to fit on-screen, do that now
            If g_AutozoomLargeImages = 0 Then FitImageToViewport True
            
            'g_AllowViewportRendering may have been reset by this point (by the FitImageToViewport sub, among others), so set it back to False, then
            ' update the zoom combo box to match the zoom assigned by the window-fit function.
            g_AllowViewportRendering = False
            FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = targetImage.currentZoomValue
        
            'Now that the image's window has been fully sized and moved around, use PrepareViewport to set up any scrollbars and a back-buffer
            g_AllowViewportRendering = True
            PrepareViewport targetImage, FormMain.mainCanvas(0), "LoadFileAsNewImage"
                                    
            'Add this file to the MRU list (unless specifically told not to)
            If ToUpdateMRU And (pageNumber = 0) And (MacroStatus <> MacroBATCH) Then g_RecentFiles.MRU_AddNewFile sFile(thisImage), targetImage
            
            'Reflow any image-window-specific display elements on the actual image form (status bar, rulers, etc)
            FormMain.mainCanvas(0).fixChromeLayout
            
            DoEvents
            
        End If
        
        '*************************************************************************************************************************************
        ' Hopefully metadata processing has finished, but if it hasn't, start a timer on the main form, which will wait for it to complete.
        '*************************************************************************************************************************************
        
        'Ask the metadata handler if it has finished parsing the image
        If g_ExifToolEnabled And isThisPrimaryImage And (targetImage.originalFileFormat <> 100) Then
        
            'Wait for metadata parsing to finish...
            If isMetadataFinished Then
                
                Message "Metadata retrieved successfully."
                targetImage.imgMetadata.loadAllMetadata retrieveMetadataString
                
            Else
            
                Message "Finishing image metadata parsing..."
            
                'Forcibly disable the main form to avoid DoEvents allowing click-through
                FormMain.Enabled = False
                
                'We don't want to pause for more than 4 additional seconds, so make a note of the time.
                Dim timeWaitMetadata As Double
                timeWaitMetadata = Timer
                
                'Pause for 1/2 second
                Do
                    PauseProgram 0.5
                    
                    'If the user shuts down the program while we are still waiting for input, exit immediately
                    If g_ProgramShuttingDown Then Exit Sub
                
                'Give ExifTool 3.5 seconds to complete its work.  If it hasn't succeeded within 3.5 seconds, assume
                ' some kind of internal failure and abandon metadata importing.
                Loop While (Not isMetadataFinished) And ((Timer - timeWaitMetadata) < 3.5)
                
                'Re-enable the main form
                FormMain.Enabled = True
                
                If isMetadataFinished Then
                    Message "Metadata retrieved successfully."
                    targetImage.imgMetadata.loadAllMetadata retrieveMetadataString
                End If
                
            End If
            
            'Next, retrieve any specific metadata-related entries that may be useful to further processing
            
            'First is resolution
            Dim xResolution As Double, yResolution As Double
            If targetImage.imgMetadata.getResolution(xResolution, yResolution) Then
                targetImage.setDPI xResolution, yResolution
            End If
            
            'I hate doing this, but we need to resync the interface to match any metadata discoveries
            syncInterfaceToCurrentImage
            
        End If
        
        
        
        '*************************************************************************************************************************************
        ' As of 2014, the new Active Undo/Redo engine requires a base pdImage copy as the starting point for Undo/Redo diffs.
        '*************************************************************************************************************************************
        
        'If this is a primary image, force an immediate Undo/Redo write to file.  This serves multiple purposes: it is our
        ' baseline for calculating future Undo/Redo diffs, and it can be used to recover the original file if something
        ' goes wrong before the user performs a manual save (e.g. AutoSave).
        '
        '(Note that all Undo behavior is disabled during batch processing, to improve performance, so we can skip this step.)
        If isThisPrimaryImage And (MacroStatus <> MacroBATCH) Then
            Message "Creating initial auto-save entry (this may take a moment)..."
            targetImage.undoManager.createUndoData "Initial image load", "", UNDO_EVERYTHING
        End If
        
        'Also, set an initial image checkpoint, in case the user decides to immediately start applying non-destructive
        ' edits to the image.
        Processor.setImageCheckpoint
        
        
        '*************************************************************************************************************************************
        ' Image loaded successfully.  Carry on.
        '*************************************************************************************************************************************
        
        targetImage.loadedSuccessfully = True
        
        If isThisPrimaryImage Then Message "Image loaded successfully."
        
        
        
        '*************************************************************************************************************************************
        ' If the just-loaded image was in a multipage format (icon, animated GIF, multipage TIFF), perform a few extra checks.
        '*************************************************************************************************************************************
        
        'Before continuing on to the next image (if any), see if the just-loaded image was in multipage format.  If it was, the user
        ' may have requested that we load all frames from this image.
        If g_imageHasMultiplePages Then
            
            Dim pageTracker As Long
            
            Dim tmpStringArray(0) As String
            tmpStringArray(0) = sFile(thisImage)
            
            'Call LoadFileAsNewImage again for each individual frame in the multipage file
            For pageTracker = 1 To g_imagePageCount
                If UCase(GetExtension(sFile(thisImage))) = "GIF" Then
                    LoadFileAsNewImage tmpStringArray, False, targetImage.originalFileName & " (" & g_Language.TranslateMessage("frame") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.originalFileName & " (" & g_Language.TranslateMessage("frame") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                ElseIf UCase(GetExtension(sFile(thisImage))) = "ICO" Then
                    LoadFileAsNewImage tmpStringArray, False, targetImage.originalFileName & " (" & g_Language.TranslateMessage("icon") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.originalFileName & " (" & g_Language.TranslateMessage("icon") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                Else
                    LoadFileAsNewImage tmpStringArray, False, targetImage.originalFileName & " (" & g_Language.TranslateMessage("page") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.originalFileName & " (" & g_Language.TranslateMessage("page") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                End If
            Next pageTracker
        
        End If
        
    '*************************************************************************************************************************************
    ' Move on to the next image.
    '*************************************************************************************************************************************
        
PreloadMoreImages:

    'If we have more images to process, now's the time to do it!
    Next thisImage
        
    
    '*************************************************************************************************************************************
    ' As all images have now loaded, re-enable the main form
    '*************************************************************************************************************************************
    
    FormMain.Enabled = True
    
    
    
    
    '*************************************************************************************************************************************
    ' Before finishing, display any relevant load problems (missing files, invalid formats, etc)
    '*************************************************************************************************************************************
    
    'Restore the screen cursor if necessary
    If pageNumber <= 0 Then Screen.MousePointer = vbNormal
    
    'If multiple images were loaded and everything went well, display a success message
    If multipleFilesLoading And (Len(missingFiles) = 0) And (Len(brokenFiles) = 0) And isThisPrimaryImage Then Message "All images loaded successfully."
        
    'Finally, if we were loading multiple images and something went wrong (missing files, broken files), let the user know about them.
    If multipleFilesLoading And (Len(missingFiles) > 0) Then
        Message "All images loaded, except for those that could not be found."
        pdMsgBox "Unfortunately, PhotoDemon was unable to find the following image(s):" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "If these images were originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbApplicationModal + vbExclamation + vbOKOnly, "Image files missing", missingFiles
    End If
        
    If multipleFilesLoading And (Len(brokenFiles) > 0) Then
        Message "All images loaded, except for those in invalid formats."
        pdMsgBox "Unfortunately, PhotoDemon was unable to load the following image(s):" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "Please use another program to save these images in a generic format (such as JPEG or PNG) before loading them into PhotoDemon. Thanks!", vbExclamation + vbOKOnly + vbApplicationModal, "Image Formats Not Supported", brokenFiles
    End If
        
End Sub

'Quick and dirty function for loading an image file to a containing DIB.  This function provides none of the extra scans or features
' that the more advanced LoadFileAsNewImage does; instead, it is assumed that the calling function will handle any extra work.
' (Note that things like metadata will not be processed *at all* for the image file.)
'
'That said, FreeImage/GDI+ are still used intelligently.  The function will return TRUE if successful.
Public Function QuickLoadImageToDIB(ByVal imagePath As String, ByRef targetDIB As pdDIB) As Boolean
    
    'Even though this function is designed to operate as quickly as possible, some images may take a long time to load.
    ' Display a busy cursor
    If Screen.MousePointer <> vbHourglass Then Screen.MousePointer = vbHourglass
            
    'To improve load time, declare a variety of other variables outside the image load loop
    Dim FileExtension As String
    Dim loadSuccessful As Boolean
    
    'To prevent re-entry problems, forcibly disable the main form until loading is complete
    FormMain.Enabled = False
    
    'Before attempting to load an image, make sure it exists
    If Not FileExist(imagePath) Then
        pdMsgBox "Unfortunately, the image '%1' could not be found." & vbCrLf & vbCrLf & "If this image was originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbApplicationModal + vbExclamation + vbOKOnly, "File not found", imagePath
        QuickLoadImageToDIB = False
        FormMain.Enabled = True
        Screen.MousePointer = vbNormal
        Exit Function
    End If
        
    'Prepare a dummy pdImage object, which some external functions may require
    Dim tmpPDImage As pdImage
    Set tmpPDImage = New pdImage
    
    'Determine the most appropriate load function for this image's format (FreeImage, GDI+, or VB's LoadPicture)
    
    'Start by stripping the extension from the file path
    FileExtension = UCase(GetExtension(imagePath))
    loadSuccessful = False
    
    'Depending on the file's extension, load the image using the most appropriate image decoding routine
    Select Case FileExtension
    
        'PhotoDemon's custom file format must be handled specially (as obviously, FreeImage and GDI+ won't handle it!)
        Case "PDI"
        
            'PDI images require zLib, and are only loaded via a custom routine (obviously, since they are PhotoDemon's native format)
            loadSuccessful = LoadPhotoDemonImage(imagePath, targetDIB, tmpPDImage)
            
        'TMP files are internal files (BMP format) used by PhotoDemon.  GDI+ is preferable for loading these, as it handles
        ' 32bpp images as well, but if we must, we can use VB's internal .LoadPicture command.
        Case "TMP"
        
            If g_ImageFormats.GDIPlusEnabled Then loadSuccessful = LoadGDIPlusImage(imagePath, targetDIB)
            If (Not g_ImageFormats.GDIPlusEnabled) Or (Not loadSuccessful) Then loadSuccessful = LoadVBImage(imagePath, targetDIB)
            
        'PDTMP files are raw image buffers saved as part of Undo/Redo or Autosaving.
        Case "PDTMP"
            
            loadSuccessful = LoadRawImageBuffer(imagePath, targetDIB, tmpPDImage)
            
        'All other formats follow a set pattern: try to load them via FreeImage (if it's available), then GDI+, then finally
        ' VB's internal LoadPicture function.
        Case Else
            
            'If FreeImage is available, use it to try and load the image.
            If g_ImageFormats.FreeImageEnabled Then loadSuccessful = LoadFreeImageV4(imagePath, targetDIB, 0, False)
                
            'If FreeImage fails for some reason, offload the image to GDI+ - UNLESS the image is a WMF or EMF, which can cause
            ' GDI+ to experience a silent fail, thus bringing down the entire program.
            If (Not loadSuccessful) And g_ImageFormats.GDIPlusEnabled And ((FileExtension <> "EMF") And (FileExtension <> "WMF")) Then loadSuccessful = LoadGDIPlusImage(imagePath, targetDIB)
            
            'If both FreeImage and GDI+ failed, give the image one last try with VB's LoadPicture
            If (Not loadSuccessful) Then loadSuccessful = LoadVBImage(imagePath, targetDIB)
                    
    End Select
    
    
    'Sometimes, our image load functions will think the image loaded correctly, but they will return a blank image.  Check for
    ' non-zero width and height before continuing.
    If (Not loadSuccessful) Or (targetDIB.getDIBWidth = 0) Or (targetDIB.getDIBHeight = 0) Then
        
        Message "Failed to load %1", imagePath
        pdMsgBox "Unfortunately, PhotoDemon was unable to load the following image:" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "Please use another program to save this image in a generic format (such as JPEG or PNG) before loading it into PhotoDemon.  Thanks!", vbExclamation + vbOKOnly + vbApplicationModal, "Image Import Failed", imagePath
        
        'Deactivate the (now useless) DIB
        targetDIB.eraseDIB
        
        'Re-enable the main interface
        FormMain.Enabled = True
        Screen.MousePointer = vbNormal
        
        'Exit with failure status
        QuickLoadImageToDIB = False
        
        Exit Function
        
    End If
    
    'If the loaded image contains alpha data, verify it.  If the alpha channel is blank (e.g. all 0 or all 255), convert it to 24bpp
    If targetDIB.getDIBColorDepth = 32 Then
        
        'Make sure the user hasn't disabled alpha channel validation
        If g_UserPreferences.GetPref_Boolean("Transparency", "Validate Alpha Channels", True) Then
            
            'Verify the alpha channel.  If this function returns FALSE, the current alpha channel is unnecessary.
            If Not targetDIB.verifyAlphaChannel Then targetDIB.convertTo24bpp
            
        
        End If
        
    End If
    
    'If the image contained an embedded ICC profile, apply it now.
    '
    'Note that we need to check if the ICC profile has already been applied.  For CMYK images, the ICC profile will be applied by
    ' the image load function.  (If we don't do this, we'll be left with a 32bpp image that contains CMYK data instead of RGBA!)
    If targetDIB.ICCProfile.hasICCData And (Not targetDIB.ICCProfile.hasProfileBeenApplied) Then
        
        '32bpp images must be un-premultiplied before the transformation
        If targetDIB.getDIBColorDepth = 32 Then targetDIB.fixPremultipliedAlpha
        
        'Apply the ICC transform
        targetDIB.ICCProfile.applyICCtoSelf targetDIB
        
        '32bpp images must be re-premultiplied after the transformation
        If targetDIB.getDIBColorDepth = 32 Then targetDIB.fixPremultipliedAlpha True
    
    End If

    'Restore the main interface
    FormMain.Enabled = True
    Screen.MousePointer = vbNormal

    'If we made it all the way here, the image file was loaded successfully!
    QuickLoadImageToDIB = True

End Function

'PDI loading.  "PhotoDemon Image" files are the only format PD supports for saving layered images.  PDI to PhotoDemon is like
' PSD to PhotoShop, or XCF to Gimp.
'
'Note the unique "sourceIsUndoFile" parameter for this load function.  PDI files are used to store undo/redo data, and when one of their
' kind is loaded as part of an Undo/Redo action, we must ignore certain elements stored in the file (e.g. settings like "LastSaveFormat"
' which we do not want to Undo/Redo).  This parameter is passed to the pdImage initializer, and it tells it to ignore certain settings.
Public Function LoadPhotoDemonImage(ByVal PDIPath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByVal sourceIsUndoFile As Boolean = False) As Boolean
    
    On Error GoTo LoadPDIFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    If g_ZLibEnabled Then pdiReader.init_ZLib g_PluginPath & "zlibwapi.dll"
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.readPackageFromFile(PDIPath, PD_IMAGE_IDENTIFIER) Then
    
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.getNodeDataByIndex(0, True, retBytes, sourceIsUndoFile) Then
        
            'Convert the received bytes into a string
            retString = StrConv(retBytes, vbUnicode)
            
            'Pass the string to the target pdImage, which will read the XML data and initialize itself accordingly
            dstImage.readExternalData retString, True, sourceIsUndoFile
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'With the main pdImage now assembled, the next task is to populate all layers with two pieces of information a piece:
        ' 1) The layer header, which contains stuff like layer name, opacity, blend mode, etc
        ' 2) The layer DIB, which is a raw stream of bytes containing the DIB's data
        
        Dim i As Long
        For i = 0 To dstImage.getNumOfLayers - 1
        
            'First, retrieve the layer's header
            If pdiReader.getNodeDataByIndex(i + 1, True, retBytes, sourceIsUndoFile) Then
            
                'Convert the received bytes into a string
                retString = StrConv(retBytes, vbUnicode)
                
                'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                If Not dstImage.getLayerByIndex(i).CreateNewImageLayerFromXML(retString) Then
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
                End If
                
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
            
            'Next, retrieve the layer's raw DIB data
            If pdiReader.getNodeDataByIndex(i + 1, False, retBytes, sourceIsUndoFile) Then
                
                'Pass the raw bytes to the target layer's DIB, which will copy the bytes over its existing DIB data.
                dstImage.getLayerByIndex(i).layerDIB.copyStreamOverImageArray retBytes
                
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
        
        Next i
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
        ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
        dstDIB.createBlank 4, 4
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonImage = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadPhotoDemonImage = False
    
    End If
    
    Exit Function
    
LoadPDIFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadPhotoDemonImage = False
    Exit Function

End Function

'Load just the layer stack from a standard PDI file, and non-destructively align our current layer stack to match.
' At present, this function is only used internally by the Undo/Redo engine.
Public Function LoadPhotoDemonImageHeaderOnly(ByVal PDIPath As String, ByRef dstImage As pdImage) As Boolean
    
    On Error GoTo LoadPDIHeaderFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    If g_ZLibEnabled Then pdiReader.init_ZLib g_PluginPath & "zlibwapi.dll"
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.readPackageFromFile(PDIPath, PD_IMAGE_IDENTIFIER) Then
    
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.getNodeDataByIndex(0, True, retBytes, True) Then
        
            'Convert the received bytes into a string
            retString = StrConv(retBytes, vbUnicode)
            
            'Pass the string to the target pdImage, which will read the XML data and initialize itself accordingly
            dstImage.readExternalData retString, True, True, True
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'With the main pdImage now assembled, the next task is to populate all layer headers.  This is a bit more
        ' confusing than a regular PDI load, because we have to maintain existing layer DIB data (ugh!).
        ' So basically, we must:
        ' 1) Extract each layer header from file, in turn
        ' 2) See if the current pdImage layer version is in the proper position in the layer stack; if it isn't,
        '    move it into place.
        ' 3) Non-destructively ask the layer to overwrite its header with the header from the file.
        
        Dim layerNodeName As String, layerNodeID As Long, layerNodeType As Long
        
        Dim i As Long
        For i = 0 To dstImage.getNumOfLayers - 1
        
            'Before doing anything else, retrieve the ID of the node at this position.  (Retrieve the rest of the node
            ' header too, although we don't actually have a use for those values at present.)
            pdiReader.getNodeInfo i + 1, layerNodeName, layerNodeID, layerNodeType
            
            'We now know what layer ID is supposed to appear at this position in the layer stack.  If that layer ID
            ' is *not* in its proper position, move it now.
            If dstImage.getLayerIndexFromID(layerNodeID) <> i Then dstImage.swapTwoLayers dstImage.getLayerIndexFromID(layerNodeID), i
            
            'Now that the node is in place, we can retrieve its header.
            If pdiReader.getNodeDataByIndex(i + 1, True, retBytes, True) Then
            
                'Convert the received bytes into a string
                retString = StrConv(retBytes, vbUnicode)
                
                'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                If Not dstImage.getLayerByIndex(i).CreateNewImageLayerFromXML(retString, , True) Then
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
                End If
                
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
            
            'Normally we would load the layer's DIB data here, but we don't care about that when loading just the headers!
            ' Continue to the next layer.
        
        Next i
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonImageHeaderOnly = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadPhotoDemonImageHeaderOnly = False
    
    End If
    
    Exit Function
    
LoadPDIHeaderFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadPhotoDemonImageHeaderOnly = False
    Exit Function

End Function

'Load a single layer from a standard PDI file.
' At present, this function is only used internally by the Undo/Redo engine.  If the nearest diff to a layer-specific change is a
' full pdImage stack, this function is used to extract only the relevant layer (or layer header) from the PDI file.
Public Function LoadSingleLayerFromPDI(ByVal PDIPath As String, ByRef dstLayer As pdLayer, ByVal targetLayerID As Long, Optional ByVal loadHeaderOnly As Boolean = False) As Boolean
    
    On Error GoTo LoadLayerFromPDIFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    If g_ZLibEnabled Then pdiReader.init_ZLib g_PluginPath & "zlibwapi.dll"
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.readPackageFromFile(PDIPath, PD_IMAGE_IDENTIFIER) Then
    
        'PDI files all follow a standard format: a pdImage node at the top, which contains the full pdImage header,
        ' followed by individual nodes for each layer.  Layers are stored in stack order, which makes it very fast and easy
        ' to reconstruct the layer stack.
        
        'Unfortunately, stack order is not helpful in this function, because the target layer's position may have changed
        ' since the time this pdImage file was created.  To work around that, we must located the layer using its cardinal
        ' ID value, which is helpfully stored as the node ID parameter for a given layer node.
        
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.getNodeDataByID(targetLayerID, True, retBytes, True) Then
        
            'Convert the received bytes into a string
            retString = StrConv(retBytes, vbUnicode)
            
            'Pass the string to the target layer, which will read the XML data and initialize itself accordingly.
            ' Note that we also pass along the loadHeaderOnly flag, which will instruct the layer to erase its current
            ' DIB as necessary.
            If Not dstLayer.CreateNewImageLayerFromXML(retString, , loadHeaderOnly) Then
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'If this is not a header-only operation, repeat the above steps, but for the layer DIB this time
        If Not loadHeaderOnly Then
        
            If pdiReader.getNodeDataByID(targetLayerID, False, retBytes, True) Then
            
                'Pass the raw bytes to the target layer's DIB, which will copy the bytes over its existing DIB data.
                dstLayer.layerDIB.copyStreamOverImageArray retBytes
            
            'Bytes could not be read, or alternately, checksums didn't match for the first node.
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
                
        End If
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadSingleLayerFromPDI = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadSingleLayerFromPDI = False
    
    End If
    
    Exit Function
    
LoadLayerFromPDIFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadSingleLayerFromPDI = False
    Exit Function

End Function

'Load a single PhotoDemon layer from a standalone pdLayer file (which is really just a modified PDI file).
' At present, this function is only used internally by the Undo/Redo engine.  Its counterpart is SavePhotoDemonLayer in
' the Saving module; any changes there should be mirrored here.
Public Function LoadPhotoDemonLayer(ByVal PDIPath As String, ByRef dstLayer As pdLayer, Optional ByVal loadHeaderOnly As Boolean = False) As Boolean
    
    On Error GoTo LoadPDLayerFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    If g_ZLibEnabled Then pdiReader.init_ZLib g_PluginPath & "zlibwapi.dll"
    
    'Load the file into the pdPackager instance.  pdPackager It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.readPackageFromFile(PDIPath, PD_LAYER_IDENTIFIER) Then
    
        'Layer variants of PDI files contain a single node.  The layer's header is stored to the node's header chunk
        ' (in XML format, as expected).  The layer's DIB data is stored to the node's data chunk (in binary format, as expected).
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.getNodeDataByIndex(0, True, retBytes, True) Then
        
            'Convert the received bytes into a string
            retString = StrConv(retBytes, vbUnicode)
            
            'Pass the string to the target layer, which will read the XML data and initialize itself accordingly.
            ' Note that we pass the loadHeaderOnly request to this function; if this is a header-only load, the target
            ' layer must retain its current DIB.  This functionality is used by PD's Undo/Redo engine.
            dstLayer.CreateNewImageLayerFromXML retString, , loadHeaderOnly
            
        'Bytes could not be read, or alternately, checksums didn't match.  (Note that checksums are currently disabled
        ' for this function, for performance reasons, but I'm leaving this check in case we someday decide to re-enable them.)
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'Unless a header-only load was requested, we will now repeat the steps above, but for the layer's DIB data
        If Not loadHeaderOnly Then
        
            If pdiReader.getNodeDataByIndex(0, False, retBytes, True) Then
            
                'Pass the raw bytes to the target layer's DIB, which will copy the bytes over its existing DIB data.
                dstLayer.layerDIB.copyStreamOverImageArray retBytes
                
            'Bytes could not be read, or alternately, checksums didn't match.  (Note that checksums are currently disabled
            ' for this function, for performance reasons, but I'm leaving this check in case we someday decide to re-enable them.)
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
            
        End If
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonLayer = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadPhotoDemonLayer = False
    
    End If
    
    Exit Function
    
LoadPDLayerFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadPhotoDemonLayer = False
    Exit Function

End Function

'Use GDI+ to load an image.  This does very minimal error checking (which is a no-no with GDI+) but because it's only a
' fallback when FreeImage can't be found, I'm postponing further debugging for now.
'Used for PNG and TIFF files if FreeImage cannot be located.
Public Function LoadGDIPlusImage(ByVal imagePath As String, ByRef dstDIB As pdDIB) As Boolean
            
    Dim verifyGDISuccess As Boolean
    
    verifyGDISuccess = GDIPlusLoadPicture(imagePath, dstDIB)
    
    If verifyGDISuccess And (dstDIB.getDIBWidth <> 0) And (dstDIB.getDIBHeight <> 0) Then
        LoadGDIPlusImage = True
    Else
        LoadGDIPlusImage = False
    End If
    
End Function

'BITMAP loading
Public Function LoadVBImage(ByVal imagePath As String, ByRef dstDIB As pdDIB) As Boolean
    
    On Error GoTo LoadVBImageFail
    
    'Create a temporary StdPicture object that will be used to load the image
    Dim tmpPicture As StdPicture
    Set tmpPicture = New StdPicture
    Set tmpPicture = LoadPicture(imagePath)
    
    If tmpPicture.Width = 0 Or tmpPicture.Height = 0 Then
        LoadVBImage = False
        Exit Function
    End If
    
    'Copy the image into the current pdImage object
    dstDIB.CreateFromPicture tmpPicture
    
    LoadVBImage = True
    Exit Function
    
LoadVBImageFail:

    LoadVBImage = False
    Exit Function
    
End Function

'Load data from a PD-generated Undo file.  This function is fairly complex, on account of PD's new diff-based Undo engine.
' Note that two types of Undo data must be specified: the Undo type of the file requested (because this function has no
' knowledge of that, by design), and what type of Undo data the caller wants extracted from the file.
Public Sub LoadUndo(ByVal undoFile As String, ByVal undoTypeOfFile As Long, ByVal undoTypeOfAction As Long, Optional ByVal targetLayerID As Long = -1, Optional ByVal suspendRedraw As Boolean = False)
    
    'Certain load functions require access to a DIB, so declare a generic one in advance
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'If selection data was loaded as part of this diff, this value will be set to TRUE.  We check it at the end of
    ' the load function, and activate various selection-related items as necessary.
    Dim selectionDataLoaded As Boolean
    selectionDataLoaded = False
    
    'Depending on the Undo data requested, we may end up loading one or more diff files at this location
    Select Case undoTypeOfAction
    
        'UNDO_EVERYTHING: a full copy of both the pdImage stack and all selection data is wanted
        Case UNDO_EVERYTHING
            Loading.LoadPhotoDemonImage undoFile, tmpDIB, pdImages(g_CurrentImage), True
            pdImages(g_CurrentImage).mainSelection.readSelectionFromFile undoFile & ".selection"
            selectionDataLoaded = True
            
        'UNDO_IMAGE: a full copy of the pdImage stack is wanted
        '             Because the underlying file data must be of type UNDO_EVERYTHING or UNDO_IMAGE, we don't have to do
        '             any special processing to the file - just load the whole damn thing.
        Case UNDO_IMAGE
            Loading.LoadPhotoDemonImage undoFile, tmpDIB, pdImages(g_CurrentImage), True
            
            'Once the full image has been loaded, we now know that at least the *existence* of all layers is correct.
            ' Unfortunately, subsequent changes to the pdImage header (or individual layers/layer headers) still need
            ' to be manually reconstructed, because they may have changed between the last full pdImage write and the
            ' current image state.  This step is handled by the Undo/Redo engine, which will call this LoadUndo function
            ' as many times as necessary to reconstruct each individual layer against its most recent diff.
        
        'UNDO_IMAGEHEADER: a full copy of the pdImage stack is wanted, but with all DIB data ignored (if present)
        '             For UNDO_IMAGEHEADER requests, we know the underlying file data is a PDI file.  We don't actually
        '             care if it has DIB data or not, because we'll just ignore it - but a special load function is
        '             required, due to the messy business of non-destructively aligning the current layer stack with
        '             the layer stack described by the file.
        Case UNDO_IMAGEHEADER
            Loading.LoadPhotoDemonImageHeaderOnly undoFile, pdImages(g_CurrentImage)
            
            'Once the full image has been loaded, we now know that at least the *existence* of all layers is correct.
            ' Unfortunately, subsequent changes to the pdImage header (or individual layers/layer headers) still need
            ' to be manually reconstructed, because they may have changed between the last full pdImage write and the
            ' current image state.  This step is handled by the Undo/Redo engine, which will call this LoadUndo function
            ' as many times as necessary to reconstruct each individual layer against its most recent diff.
        
        'UNDO_LAYER: a full copy of the saved layer data at this position.
        '             Because the underlying file data can be different types (layer data can be loaded from standalone layer saves,
        '             or from a full pdImage stack save), we must check the undo type of the saved file, and modify our load
        '             behavior accordingly.
        Case UNDO_LAYER
            
            'Layer data can appear in multiple types of Undo files
            Select Case undoTypeOfFile
            
                'The underlying save file is a standalone layer entry.  Simply overwrite the target layer with the data from the file.
                Case UNDO_LAYER
                    Loading.LoadPhotoDemonLayer undoFile & ".layer", pdImages(g_CurrentImage).getLayerByID(targetLayerID), False
            
                'The underlying save file is a full pdImage stack.  Extract only the relevant layer data from the stack.
                Case UNDO_EVERYTHING, UNDO_IMAGE
                    Loading.LoadSingleLayerFromPDI undoFile, pdImages(g_CurrentImage).getLayerByID(targetLayerID), targetLayerID, False
                
            End Select
        
        'UNDO_LAYERHEADER: a full copy of the saved layer header data at this position.  Layer DIB data is ignored.
        '             Because the underlying file data can be many different types (layer data header can be loaded from
        '             standalone layer header saves, or full layer saves, or even a full pdImage stack), we must check the
        '             undo type of the saved file, and modify our load behavior accordingly.
        Case UNDO_LAYERHEADER
            
            'Layer header data can appear in multiple types of Undo files
            Select Case undoTypeOfFile
            
                'The underlying save file is a standalone layer entry.  Simply overwrite the target layer header with the
                ' header data from this file.
                Case UNDO_LAYER, UNDO_LAYERHEADER
                    Loading.LoadPhotoDemonLayer undoFile & ".layer", pdImages(g_CurrentImage).getLayerByID(targetLayerID), True
            
                'The underlying save file is a full pdImage stack.  Extract only the relevant layer data from the stack.
                Case UNDO_EVERYTHING, UNDO_IMAGE, UNDO_IMAGEHEADER
                    Loading.LoadSingleLayerFromPDI undoFile, pdImages(g_CurrentImage).getLayerByID(targetLayerID), targetLayerID, True
                
            End Select
        
        'UNDO_SELECTION: a full copy of the saved selection data is wanted
        '                 Because the underlying file data must be of type UNDO_EVERYTHING or UNDO_SELECTION, we don't have to do
        '                 any special processing.
        Case UNDO_SELECTION
            pdImages(g_CurrentImage).mainSelection.readSelectionFromFile undoFile & ".selection"
            selectionDataLoaded = True
            
            
        'For now, any unhandled Undo types result in a request for the full pdImage stack.  This line can be removed when
        ' all Undo types finally have their own custom handling implemented.
        Case Else
            Loading.LoadPhotoDemonImage undoFile, tmpDIB, pdImages(g_CurrentImage), True
            
        
    End Select
    
    'If a selection was loaded, activate all selection-related stuff now
    If selectionDataLoaded Then
    
        'Activate the selection as necessary
        pdImages(g_CurrentImage).selectionActive = pdImages(g_CurrentImage).mainSelection.isLockedIn
        
        'Synchronize the text boxes as necessary
        syncTextToCurrentSelection g_CurrentImage
    
    End If
    
    'If a selection is active, request a redraw of the selection mask before rendering the image to the screen.  (If we are
    ' "undoing" an action that changed the image's size, the selection mask will be out of date.  Thus we need to re-render
    ' it before rendering the image or OOB errors may occur.)
    If pdImages(g_CurrentImage).selectionActive Then pdImages(g_CurrentImage).mainSelection.requestNewMask
        
    'Render the image to the screen, if requested
    If Not suspendRedraw Then PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "LoadUndo"
    
End Sub

'Load a raw image buffer (.pdtmp) into the destination image and DIB
Public Function LoadRawImageBuffer(ByVal imagePath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage) As Boolean

    On Error GoTo LoadRawImageBufferFail
    
    'Ask the destination DIB to create itself using the raw image buffer data
    dstDIB.createFromFile imagePath
    
    LoadRawImageBuffer = True
    Exit Function
    
LoadRawImageBufferFail:

    LoadRawImageBuffer = False
    Exit Function

End Function

'This routine sets the message on the splash screen (used only when the program is first started)
Public Sub LoadMessage(ByVal sMsg As String)

    Dim warnIDE As String
    warnIDE = "(IDE NOT RECOMMENDED - PLEASE COMPILE)"
    
    'Load messages are translatable, but we don't want to translate them if the translation object isn't ready yet
    If (Not (g_Language Is Nothing)) Then
        If g_Language.readyToTranslate Then
            If g_Language.translationActive Then
                sMsg = g_Language.TranslateMessage(sMsg)
                warnIDE = g_Language.TranslateMessage("(IDE NOT RECOMMENDED - PLEASE COMPILE)")
            End If
        End If
    End If
    
    If Not g_IsProgramCompiled Then sMsg = sMsg & "  " & warnIDE
    
    If FormSplash.Visible Then
        FormSplash.lblMessage = sMsg
        FormSplash.lblMessage.Refresh
    End If
    'DoEvents
    
End Sub

'Generates all shortcuts that VB can't; many thanks to Steve McMahon for his accelerator class, which helps a great deal
Public Sub LoadAccelerators()

    'Don't allow custom shortcuts in the IDE, as they require subclassing and might crash
    'If Not g_IsProgramCompiled Then Exit Sub

    With FormMain.ctlAccelerator
    
        'File menu
        .AddAccelerator vbKeyO, vbCtrlMask, "Open", FormMain.MnuFile(0), True, False, True, UNDO_NOTHING
        .AddAccelerator vbKeyF4, vbCtrlMask, "Close", FormMain.MnuFile(4), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyF4, vbCtrlMask Or vbShiftMask, "Close all", FormMain.MnuFile(5), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyS, vbCtrlMask, "Save", FormMain.MnuFile(7), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyS, vbCtrlMask Or vbShiftMask, "Save as", FormMain.MnuFile(8), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyF12, 0, "Revert", FormMain.MnuFile(9), True, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyB, vbCtrlMask, "Batch wizard", FormMain.MnuFile(11), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyP, vbCtrlMask, "Print", FormMain.MnuFile(13), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyQ, vbCtrlMask, "Exit program", FormMain.MnuFile(15), True, False, True, UNDO_NOTHING
        
            'File -> Import submenu
            .AddAccelerator vbKeyI, vbCtrlMask Or vbShiftMask Or vbAltMask, "Scan image", FormMain.MnuScanImage, True, False, True, UNDO_NOTHING
            .AddAccelerator vbKeyD, vbCtrlMask Or vbShiftMask, "Internet import", FormMain.MnuImportFromInternet, True, True, True, UNDO_NOTHING
            .AddAccelerator vbKeyI, vbCtrlMask Or vbAltMask, "Screen capture", FormMain.MnuScreenCapture, True, True, True, UNDO_NOTHING
        
            'Most-recently used files.  Note that we cannot automatically associate these with a menu, as these menus may not
            ' exist at run-time.  (They are dynamically created as necessary.)
            .AddAccelerator vbKey0, vbCtrlMask, "MRU_0"
            .AddAccelerator vbKey1, vbCtrlMask, "MRU_1"
            .AddAccelerator vbKey2, vbCtrlMask, "MRU_2"
            .AddAccelerator vbKey3, vbCtrlMask, "MRU_3"
            .AddAccelerator vbKey4, vbCtrlMask, "MRU_4"
            .AddAccelerator vbKey5, vbCtrlMask, "MRU_5"
            .AddAccelerator vbKey6, vbCtrlMask, "MRU_6"
            .AddAccelerator vbKey7, vbCtrlMask, "MRU_7"
            .AddAccelerator vbKey8, vbCtrlMask, "MRU_8"
            .AddAccelerator vbKey9, vbCtrlMask, "MRU_9"
            
        'Edit menu
        .AddAccelerator vbKeyZ, vbCtrlMask, "Undo", FormMain.MnuEdit(0), True, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyY, vbCtrlMask, "Redo", FormMain.MnuEdit(1), True, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyF, vbCtrlMask, "Repeat last action", FormMain.MnuEdit(2), True, True, False, UNDO_IMAGE
        
        .AddAccelerator vbKeyC, vbCtrlMask, "Copy", FormMain.MnuEdit(4), True, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyC, vbCtrlMask Or vbShiftMask, "Copy merged", FormMain.MnuEdit(5), True, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyV, vbCtrlMask, "Paste as new layer", FormMain.MnuEdit(6), True, False, False, UNDO_IMAGE
        .AddAccelerator vbKeyV, vbCtrlMask Or vbShiftMask, "Paste as new image", FormMain.MnuEdit(7), True, False, False, UNDO_NOTHING
        
        'View menu
        .AddAccelerator vbKey0, 0, "FitOnScreen", FormMain.MnuFitOnScreen, False, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyAdd, 0, "Zoom_In", FormMain.MnuZoomIn, False, True, False, UNDO_NOTHING
        .AddAccelerator vbKeySubtract, 0, "Zoom_Out", FormMain.MnuZoomOut, False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey5, 0, "Zoom_161", FormMain.MnuSpecificZoom(0), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey4, 0, "Zoom_81", FormMain.MnuSpecificZoom(1), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey3, 0, "Zoom_41", FormMain.MnuSpecificZoom(2), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey2, 0, "Zoom_21", FormMain.MnuSpecificZoom(3), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey1, 0, "Actual_Size", FormMain.MnuSpecificZoom(4), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey2, vbShiftMask, "Zoom_12", FormMain.MnuSpecificZoom(5), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey3, vbShiftMask, "Zoom_14", FormMain.MnuSpecificZoom(6), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey4, vbShiftMask, "Zoom_18", FormMain.MnuSpecificZoom(7), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey5, vbShiftMask, "Zoom_116", FormMain.MnuSpecificZoom(8), False, True, False, UNDO_NOTHING
        
        'Image menu
        .AddAccelerator vbKeyA, vbCtrlMask Or vbShiftMask, "Duplicate image", FormMain.MnuImage(0), True, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyR, vbCtrlMask, "Resize image", FormMain.MnuImage(2), True, True, True, UNDO_IMAGE
        .AddAccelerator vbKeyR, vbCtrlMask Or vbAltMask, "Canvas size", FormMain.MnuImage(4), True, True, True, UNDO_IMAGEHEADER
        .AddAccelerator vbKeyX, vbCtrlMask Or vbShiftMask, "Crop", FormMain.MnuImage(8), True, True, False, UNDO_IMAGE
        .AddAccelerator vbKeyX, vbCtrlMask Or vbAltMask, "Trim", FormMain.MnuImage(9), True, True, False, UNDO_IMAGEHEADER
        'KeyCode 188 = <,  (next to the letter M)
        .AddAccelerator 188, vbCtrlMask Or vbAltMask, "Reduce colors", FormMain.MnuImage(15), True, True, False, UNDO_IMAGE
        
            'Image -> Rotate submenu
            .AddAccelerator vbKeyR, 0, "Rotate image 90° clockwise", FormMain.MnuRotate(2), True, True, False, UNDO_IMAGE
            .AddAccelerator vbKeyL, 0, "Rotate image 90° counter-clockwise", FormMain.MnuRotate(3), True, True, False, UNDO_IMAGE
            .AddAccelerator vbKeyR, vbCtrlMask Or vbShiftMask Or vbAltMask, "Arbitrary image rotation", FormMain.MnuRotate(5), True, True, True, UNDO_NOTHING
        
        'Layer Menu
        '(none yet)
        
        
        'Select Menu
        .AddAccelerator vbKeyA, vbCtrlMask, "Select all", FormMain.MnuSelect(0), True, True, False, UNDO_SELECTION
        .AddAccelerator vbKeyD, vbCtrlMask, "Remove selection", FormMain.MnuSelect(1), False, True, False, UNDO_SELECTION
        .AddAccelerator vbKeyI, vbCtrlMask Or vbShiftMask, "Invert selection", FormMain.MnuSelect(2), True, True, False, UNDO_SELECTION
        'KeyCode 219 = {[  (next to the letter P), 221 = }]
        .AddAccelerator 221, vbCtrlMask Or vbAltMask, "Grow selection", FormMain.MnuSelect(4), True, True, True, UNDO_NOTHING
        .AddAccelerator 219, vbCtrlMask Or vbAltMask, "Shrink selection", FormMain.MnuSelect(5), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyD, vbCtrlMask Or vbAltMask, "Feather selection", FormMain.MnuSelect(7), True, True, True, UNDO_NOTHING
        
        'Adjustments Menu
        
        'Adjustments top shortcut menu
        .AddAccelerator vbKeyU, vbCtrlMask Or vbShiftMask, "Black and white", FormMain.MnuAdjustments(0), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyB, vbCtrlMask Or vbShiftMask, "Brightness and contrast", FormMain.MnuAdjustments(1), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyC, vbCtrlMask Or vbAltMask, "Color balance", FormMain.MnuAdjustments(2), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyM, vbCtrlMask, "Curves", FormMain.MnuAdjustments(3), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyL, vbCtrlMask, "Levels", FormMain.MnuAdjustments(4), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyAdd, vbCtrlMask Or vbAltMask, "Vibrance", FormMain.MnuAdjustments(5), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyW, vbCtrlMask, "White balance", FormMain.MnuAdjustments(6), True, True, True, UNDO_NOTHING
        
            'Color adjustments
            .AddAccelerator vbKeyH, vbCtrlMask, "Hue and saturation", FormMain.MnuColor(3), True, True, True, UNDO_NOTHING
            
            'Lighting adjustments
            .AddAccelerator vbKeyG, vbCtrlMask, "Gamma", FormMain.MnuLighting(2), True, True, True, UNDO_NOTHING
            .AddAccelerator vbKeyH, vbCtrlMask Or vbShiftMask, "Shadows and highlights", FormMain.MnuLighting(4), True, True, True, UNDO_NOTHING
            .AddAccelerator vbKeyT, vbCtrlMask, "Temperature", FormMain.MnuLighting(5), True, True, True, UNDO_NOTHING
            
            'Adjustments -> Invert submenu
            .AddAccelerator vbKeyI, vbCtrlMask, "Invert RGB", FormMain.mnuInvert, True, True, False, UNDO_LAYER
            
            'Adjustments -> Monochrome submenu
            .AddAccelerator vbKeyB, vbCtrlMask Or vbAltMask Or vbShiftMask, "Color to monochrome", FormMain.MnuMonochrome(0), True, True, True, UNDO_NOTHING
            
            'Adjustments -> Photography submenu
            .AddAccelerator vbKeyE, vbCtrlMask Or vbAltMask, "Exposure", FormMain.MnuAdjustmentsPhoto(0), True, True, True, UNDO_NOTHING
            .AddAccelerator vbKeyP, vbCtrlMask Or vbAltMask, "Photo filter", FormMain.MnuAdjustmentsPhoto(1), True, True, True, UNDO_NOTHING
            
        
        'Effects Menu
        '.AddAccelerator vbKeyZ, vbCtrlMask Or vbAltMask Or vbShiftMask, "Add RGB noise", FormMain.MnuNoise(1), True, True, True, False
        '.AddAccelerator vbKeyG, vbCtrlMask Or vbAltMask Or vbShiftMask, "Gaussian blur", FormMain.MnuBlurFilter(1), True, True, True, False
        '.AddAccelerator vbKeyY, vbCtrlMask Or vbAltMask Or vbShiftMask, "Correct lens distortion", FormMain.MnuDistortEffects(1), True, True, True, False
        '.AddAccelerator vbKeyU, vbCtrlMask Or vbAltMask Or vbShiftMask, "Unsharp mask", FormMain.MnuSharpen(1), True, True, True, False
        
        'Tools menu
        .AddAccelerator vbKeyReturn, vbAltMask, "Preferences", FormMain.mnuTool(5), False, False, True, UNDO_NOTHING
        .AddAccelerator vbKeyM, vbCtrlMask Or vbAltMask, "Plugin manager", FormMain.mnuTool(6), False, False, True, UNDO_NOTHING
        'KeyCode 190 = >.  (two over from the letter M)
        .AddAccelerator 190, vbCtrlMask Or vbAltMask, "Play macro", FormMain.MnuPlayMacroRecording, True, True, True, UNDO_NOTHING
        
        'Window menu
        .AddAccelerator vbKeyPageDown, 0, "Next_Image", FormMain.MnuWindow(7), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyPageUp, 0, "Prev_Image", FormMain.MnuWindow(8), False, True, False, UNDO_NOTHING
                
        'No equivalent menu
        .AddAccelerator vbKeyEscape, 0, "Escape"
        
        .Enabled = True
    End With

    DrawAccelerators
    
End Sub

'After all menu shortcuts (accelerators) are loaded above, the custom shortcuts need to be added to the menu entries themselves.
' If we don't do this, the user won't know how to trigger the shortcuts!
Public Sub DrawAccelerators()

    Dim i As Long
    
    For i = 1 To FormMain.ctlAccelerator.Count
        With FormMain.ctlAccelerator
            If .hasMenu(i) Then
                .associatedMenu(i).Caption = .associatedMenu(i).Caption & vbTab & .stringRep(i)
            End If
        End With
    Next i

    'A few menu shortcuts must be drawn manually.
    
    'Because the Import -> From Clipboard menu shares the same shortcut as Edit -> Paste as new image, we must
    ' manually add its shortcut (as only the Edit -> Paste will be handled automatically).
    FormMain.MnuImportClipboard.Caption = FormMain.MnuImportClipboard.Caption & vbTab & "Ctrl+Shift+V"
    
    'Similarly for the Layer -> New -> From Clipboard menu
    FormMain.MnuLayerNew(3).Caption = FormMain.MnuLayerNew(3).Caption & vbTab & "Ctrl+V"
    
    'NOTE: Drawing of MRU shortcuts is handled in the MRU module
    
End Sub

'This subroutine handles the detection of the three core plugins strongly recommended for an optimal PhotoDemon
' experience: zLib, EZTwain32, and FreeImage.  For convenience' sake, it also checks for GDI+ availability.
Public Sub LoadPlugins()
    
    'Plugin files are located in the \Data\Plugins subdirectory
    g_PluginPath = g_UserPreferences.getAppPath & "Plugins\"
    
    'Make sure the plugin path exists
    If Not DirectoryExist(g_PluginPath) Then MkDir g_PluginPath
    
    'Old versions of PhotoDemon kept plugins in a different directory. Check the old location,
    ' and if plugin-related files are found, copy them to the new directory
    On Error Resume Next
    Dim tmpg_PluginPath As String
    tmpg_PluginPath = g_UserPreferences.getDataPath & "Plugins\"
    
    If DirectoryExist(tmpg_PluginPath) Then
        LoadMessage "Moving plugins to updated folder location..."
        
        Dim pluginName As String
        pluginName = "EZTW32.dll"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "EZTWAIN_README.TXT"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "FreeImage.dll"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "license-fi.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "license-freeimage.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "license-gplv2.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "license-gplv3.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "zlibwapi.dll"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "pngnq-s9.exe"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "PNGNQ-S9-LICENSE"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "PNGNQ-S9-LICENSE.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "PNGNQ-S9-LICENCE.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        'After all files have been removed, kill the old Plugin directory
        RmDir tmpg_PluginPath
        
    End If
        
    'Check for image scanning
    'First, make sure we have our dll file
    If isEZTwainAvailable Then
                
        'If we do find the DLL, check to see if EZTwain has been forcibly disabled by the user.
        If g_UserPreferences.GetPref_Boolean("Plugins", "Force EZTwain Disable", False) Then
            g_ScanEnabled = False
        Else
            g_ScanEnabled = True
        End If
        
    Else
        
        'If we can't find the DLL, hide the menu options and internally disable scanning
        '(perhaps overkill, but it acts as a safeguard to prevent bad DLL-based crashes)
        g_ScanEnabled = False
        
    End If
    
        'Additionally related to EZTwain - enable/disable the various scanner options contigent on EZTwain's enabling
        FormMain.MnuScanImage.Visible = g_ScanEnabled
        FormMain.MnuSelectScanner.Visible = g_ScanEnabled
        FormMain.MnuImportSepBar1.Visible = g_ScanEnabled
    
    'Check for zLib compression capabilities
    If isZLibAvailable Then
    
        'Check to see if zLib has been forcibly disabled.
        If g_UserPreferences.GetPref_Boolean("Plugins", "Force ZLib Disable", False) Then
            g_ZLibEnabled = False
        Else
            g_ZLibEnabled = True
        End If
        
    Else
        g_ZLibEnabled = False
    End If
    
    'Check for FreeImage file interface
    If isFreeImageAvailable Then
        
        'Check to see if FreeImage has been forcibly disabled
        If g_UserPreferences.GetPref_Boolean("Plugins", "Force FreeImage Disable", False) Then
            g_ImageFormats.FreeImageEnabled = False
        Else
            g_ImageFormats.FreeImageEnabled = True
            
            'Because FreeImage is used so frequently throughout PhotoDemon, we only load it once - now - rather than having each
            ' individual function load it.
            g_FreeImageHandle = LoadLibrary(g_PluginPath & "FreeImage.dll")
            
        End If
        
    Else
        g_ImageFormats.FreeImageEnabled = False
    End If
    
        'Additionally related to FreeImage - enable/disable the arbitrary rotation option contingent on FreeImage's enabling
        FormMain.MnuRotate(3).Visible = g_ImageFormats.FreeImageEnabled
    
    'Check for pngnq interface
    If isPngnqAvailable Then
        
        'Check to see if pngnq-s9 has been forcibly disabled
        If g_UserPreferences.GetPref_Boolean("Plugins", "Force Pngnq Disable", False) Then
            g_ImageFormats.pngnqEnabled = False
        Else
            g_ImageFormats.pngnqEnabled = True
        End If
        
    Else
        g_ImageFormats.pngnqEnabled = False
    End If
    
    'Check for ExifTool metadata interface
    If isExifToolAvailable Then
        
        'Check to see if ExifTool has been forcibly disabled
        If g_UserPreferences.GetPref_Boolean("Plugins", "Force ExifTool Disable", False) Then
            g_ExifToolEnabled = False
        Else
            
            'Attempt to start ExifTool.  Because we interact with it asynchronously, we do not need to wait for an image to be loaded
            ' before executing it.
            If startExifTool() Then
                g_ExifToolEnabled = True
            Else
                g_ExifToolEnabled = False
            End If
            
        End If
        
    Else
        g_ExifToolEnabled = False
    End If
    
    'Finally, check GDI+ availability
    If g_GDIPlusAvailable Then
        g_ImageFormats.GDIPlusEnabled = True
    Else
        g_ImageFormats.GDIPlusEnabled = False
    End If
    
    
End Sub

'Make a copy of the current image.  Thanks to PSC user "Achmad Junus" for this suggestion.
Public Sub DuplicateCurrentImage()
    
    Message "Duplicating current image..."
    
    'Ask the currently active image to write itself out to file
    Dim tmpDuplicationFile As String
    tmpDuplicationFile = g_UserPreferences.GetTempPath & "PDDuplicate.pdi"
    SavePhotoDemonImage pdImages(g_CurrentImage), tmpDuplicationFile, True, True, True, False
    
    'We can now use the standard image load routine to import the temporary file
    Dim sFile() As String, sTitle As String, sFilename As String
    ReDim sFile(0) As String
    sFile(0) = tmpDuplicationFile
    sTitle = pdImages(g_CurrentImage).originalFileName & " - " & g_Language.TranslateMessage("Copy")
    sFilename = sTitle
    
    LoadFileAsNewImage sFile, False, sTitle, sFilename
                    
    'Be polite and remove the temporary file
    If FileExist(tmpDuplicationFile) Then Kill tmpDuplicationFile
    
    Message "Image duplication complete."
        
End Sub

'Check for IDE or compiled EXE, and set program parameters accordingly
Private Sub CheckLoadingEnvironment()
    If App.logMode = 1 Then
        g_IsProgramCompiled = True
    Else
        g_IsProgramCompiled = False
    End If
End Sub
