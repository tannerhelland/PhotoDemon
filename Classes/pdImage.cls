VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image class
'Copyright ©2006-2014 by Tanner Helland
'Created: sometime 2006
'Last updated: 14/February/14
'Last update: added save tracking variables for WebP and JPEG-XR formats
'
'The pdImage class is used to store information on each image loaded by the user.  One copy of this
' class exists for each loaded image.  These copies are stored in the pdImages() array, which is declared
' in the MDIWindow module.
'
'The purpose of this class is to store a bunch of variables related to a given image - things like size, zoom,
' name, file path, format containing form, Undo/Redo tracking - as well as a reference to a pdDIB object
' called "mainDIB", which contains an in-memory copy of the relevant image's data.  In the future,
' additional DIBs may be utilized within each image.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit


'Each active image in the program exists within one of these pdImage objects.  This class stores all relevant information
' about the image, including references to its layers, metadata, Undo/Redo, and much more.

    'This ID number matches the .Tag value of the containing form.
    Public imageID As Long

    'Is the form associated with this image still loaded?
    Public IsActive As Boolean

    'Was the image loaded successfully?
    Public loadedSuccessfully As Boolean
    
    'Image dimensions
    Public Width As Long
    Public Height As Long
    
    'Image resolution (in DPI).  In general, we don't deal with differing x/y resolutions - in that case,
    ' use the imageDPI value, which will return the average resolution of the two.
    Private xResolution As Double
    Private yResolution As Double
    Private imageDPI As Double
    
    'Current index of the zoom combo box
    Public currentZoomValue As Long
    
    'Size of the file when loaded
    Public originalFileSize As Long
    
    'Original name of the image
    Public originalFileName As String
    
    'Original name of the image AND its extension
    Public originalFileNameAndExtension As String
    
    'Original path of the image (full path, including filename); blank if image does not exist on disk
    Public locationOnDisk As String
    
    'Original image format.  This is derived from the equivalent FreeImage constant - see the enum for FREE_IMAGE_FORMAT for details
    Public originalFileFormat As Long
    
    'Current file format.  For example, if the user loads a .BMP file and uses "Save As" to save it as a .JPG, this variable will
    ' represent "JPG" - while OriginalFileFormat will still represent "BMP".  (OriginalFileFormat is only set once, at load-time.)
    Public currentFileFormat As Long
    
    'Original color depth (a BPP value, most commonly 24 or 32, may be 8)
    Public originalColorDepth As Long
    
    'saveParameters let us save the image using settings the user has already specified (so we don't have to pester
    ' them every time they save)
    Public saveParameters As String
    
    'For JPEGs specifically, has a quality value been set?  If it hasn't, the user needs to be prompted at least once.
    Public hasSeenJPEGPrompt As Boolean
    
    'For JPEG-2000 files, has a compression ratio been set?  If it hasn't, the user needs to be prompted at least once.
    Public hasSeenJP2Prompt As Boolean
    
    'For WEBP files, has a compression ratio been set?  If it hasn't, the user needs to be prompted at least once.
    Public hasSeenWebPPrompt As Boolean
    
    'For JPEG XR files, has a compression ratio been set?  If it hasn't, the user needs to be prompted at least once.
    Public hasSeenJXRPrompt As Boolean
    
    'For PNG files, if the original PNG file has a background color, it will be stored here.
    Public pngBackgroundColor As Long
    
    'Has this image been saved?  Access this variable via the getSaveState and setSaveState functions.
    Private hasBeenSaved As Boolean
    
    'These two variables are only used at form load; they can be safely ignored once the image is visible.
    Public WindowLeft As Long
    Public WindowTop As Long
        
    'A copy of the form's current icon
    Public curFormIcon32 As Long, curFormIcon16 As Long
    
    'If this image is simply being used for internal PhotoDemon processing, hide it from the user
    Public forInternalUseOnly As Boolean
    
    'Image's actual image data (stored as a DIB inside the mainDIB object)
    Public mainDIB As pdDIB
    
    'Image's back buffer (also a DIB; this is used to hold the final composite before everything is rendered on-screen
    Public backBuffer As pdDIB
    
    'If this image is 32bpp then it requires a specialized compositing layer for rendering the transparency
    Public alphaFixDIB As pdDIB
    
    'Image's selection data (stored and handled by a mainSelection object)
    Public mainSelection As pdSelection
    Public selectionActive As Boolean
    
    'Metadata handler/storage class.  This class is automatically filled with an image file's metadata at load-time.
    ' NOTE: EXIFTOOL IS REQUIRED FOR METADATA HANDLING.
    Public imgMetadata As pdMetadata
    
    'All Undo/Redo actions are handled through this publicly available Undo/Redo handler class
    Public undoManager As pdUndo
    
    'All ICC Profile interactions are handled through a publicly available ICCProfile class
    Public ICCProfile As pdICCProfile
    
    'Various viewport actions are now handled through a separate pdViewport class
    Public imgViewport As pdViewport
    
    'Image layers!  All layers are stored in this array.
    Private imgLayers() As pdLayer
    
    'Current layer.  Layers in PD are zero-based, with the 0 layer being created by default when a pdImage is created.
    ' To ensure proper behavior for things like caching viewports, this value is not publicly accessible - you must access
    ' it via the relevant get/set functions.
    Private curLayer As Long
    
    'Total number of layers.
    Private numOfLayers As Long
    
'Write all relevant image information to an external file (in XML format).  Note that this will not be a perfect copy of the current image's state,
' but it should be good enough to reconstruct an approximation for data recovery purposes.
Public Function writeInternalDataToFile(ByVal dstFilename As String) As Boolean

    On Error GoTo failedInternalDataWrite

    'Prepare an XML engine, which will handle the actual writing of the file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Add a basic header and explanatory comment
    xmlEngine.prepareNewXML "pdImage Backup"
    
    xmlEngine.writeBlankLine
    xmlEngine.writeComment "This file contains a summary of crucial data for a given pdImage entry.  It is not comprehensive, but should be sufficient for autosave purposes."
    xmlEngine.writeBlankLine
    
    'Start by writing out the imageID.  This value will not be re-used in subsequent session, but it's helpful for matching up
    ' this image's data with any saved Undo/Redo entries (which are created by imageID for privacy reasons)
    xmlEngine.writeTag "ID", imageID
    
    'Start writing any tag that may be useful for reconstructing an autosave version of this image.  We limit the information to
    ' properties independent of a given session (e.g. imageID); session-dependent properties will be recreated anew by PD.
    xmlEngine.writeTag "Width", Width
    xmlEngine.writeTag "Height", Height
    xmlEngine.writeTag "xResolution", xResolution
    xmlEngine.writeTag "yResolution", yResolution
    xmlEngine.writeTag "DPI", imageDPI
    xmlEngine.writeTag "OriginalFileSize", originalFileSize
    xmlEngine.writeTag "OriginalFileName", originalFileName
    xmlEngine.writeTag "OriginalFileNameAndExtension", originalFileNameAndExtension
    xmlEngine.writeTag "LocationOnDisk", locationOnDisk
    xmlEngine.writeTag "OriginalFileFormat", originalFileFormat
    xmlEngine.writeTag "CurrentFileFormat", currentFileFormat
    xmlEngine.writeTag "OriginalColorDepth", originalColorDepth
    xmlEngine.writeTag "pngBackgroundColor", pngBackgroundColor
    xmlEngine.writeBlankLine
    
    'That should be enough information to reconstruct the original image state.  Write it out to the file we were supplied.
    xmlEngine.writeXMLToFile dstFilename
    
    writeInternalDataToFile = True
    Exit Function
    
failedInternalDataWrite:

    Message "Could not write Autosave data - did the containing drive run out of free space?"
    writeInternalDataToFile = False
    
End Function

'Sister function to writeInternalDataToFile, above.  If you add a variable to one function, make sure to include it in the other.
' If the function fails, it will return FALSE.
Public Function readInternalDataFromFile(ByVal srcFilename As String) As Boolean

    On Error GoTo failedInternalDataRead

    'Prepare an XML engine, which will handle the actual reading and parsing of the file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Make sure the supplied file is valid
    If xmlEngine.loadXMLFile(srcFilename) And xmlEngine.isPDDataType("pdImage Backup") Then
    
        'The XML file checked out.  Start retrieving relevant values.
        Me.Width = xmlEngine.getUniqueTag_Long("Width")
        Me.Height = xmlEngine.getUniqueTag_Long("Height")
        xResolution = xmlEngine.getUniqueTag_Long("xResolution")
        yResolution = xmlEngine.getUniqueTag_Long("yResolution")
        imageDPI = xmlEngine.getUniqueTag_Long("DPI")
        Me.originalFileSize = xmlEngine.getUniqueTag_Long("originalFileSize")
        Me.originalFileName = xmlEngine.getUniqueTag_String("originalFileName")
        Me.originalFileNameAndExtension = xmlEngine.getUniqueTag_String("originalFileNameAndExtension")
        Me.locationOnDisk = xmlEngine.getUniqueTag_String("locationOnDisk")
        Me.originalFileFormat = xmlEngine.getUniqueTag_Long("originalFileFormat")
        Me.currentFileFormat = xmlEngine.getUniqueTag_Long("currentFileFormat")
        Me.originalColorDepth = xmlEngine.getUniqueTag_Long("originalColorDepth")
        Me.pngBackgroundColor = xmlEngine.getUniqueTag_Long("pngBackgroundColor")
    
    Else
        GoTo failedInternalDataRead
    End If
    
    xmlEngine.writeBlankLine
    xmlEngine.writeComment "This file contains a summary of crucial pdImage data.  It is not comprehensive, but should be sufficient for autosave purposes."
    xmlEngine.writeBlankLine
    
    readInternalDataFromFile = True
    
    Exit Function
    
failedInternalDataRead:
    
    Debug.Print "WARNING: could not read Autosave data.  pdImage internals were not updated."
    readInternalDataFromFile = False
    
End Function

'Get/Set image resolution (in DPI).  Note that the vertical resolution is optional; if the two values
' differ, PD will average them when image DPI is requested.
Public Function getDPI() As Double
    
    If imageDPI = 0 Then
        getDPI = 96
    Else
        getDPI = imageDPI
    End If
    
End Function

Public Sub setDPI(ByVal xRes As Double, ByVal yRes As Double, Optional ByVal sourceIsFreeImage As Boolean = False)
    
    'Many image types do not store resolution information; default to 96 in this case
    If xRes = 0 Then xRes = 96
    If yRes = 0 Then yRes = 96
    
    'FreeImage is pretty damn stupid when it comes to DPI.  If no DPI information is found, it will return "72" by default.
    ' If the source of this call is FreeImage, we want to replace 72 with a more modern default of 96.  Obviously this creates
    ' a problem when images have an actual resolution of 72 DPI, as there is no way to tell if FreeImage set that by default,
    ' or if it is the actual DPI value of the file.  If this occurs, we can use ExifTool to retrieve an actual value as
    ' stored in the file.
    If sourceIsFreeImage Then
        If xRes = 72 Then xRes = 96
        If yRes = 72 Then yRes = 96
    End If
    
    xResolution = xRes
    yResolution = yRes
    
    'It is extremely rare for x/y resolution to differ, but just in case, calculate an average resolution as well
    imageDPI = (xRes + yRes) \ 2

End Sub

'If the image has been saved to file in its current state, this will return TRUE.  Use this value to determine
' whether to enable a Save button, for example.
Public Function getSaveState() As Boolean
    getSaveState = hasBeenSaved
End Function

'Returns all layers of the image as a single, composited image (in pdDIB format, of course).
Public Function getCompositedImage() As pdDIB
    Set getCompositedImage = mainDIB
End Function

'Get the currently active layer index
Public Function getActiveLayerIndex() As Long
    getActiveLayerIndex = curLayer
End Function

'Get the number of layers in this image.  (Note: this function might return zero, so handle that condition correctly
' in calling functions!)
Public Function getNumOfLayers() As Long
    getNumOfLayers = numOfLayers
End Function

'Retrieve a reference to the currently active layer's DIB.  This is effectively a shortcut function when someone
' needs access to the current layer's DIB, and they don't want to retrieve it via a bunch of helper functions.
Public Function getActiveDIB() As pdDIB
    If numOfLayers > 0 Then
        Set getActiveDIB = imgLayers(curLayer).layerDIB
    Else
        Set getActiveDIB = Nothing
    End If
End Function

'Retrieve the currently active layer
Public Function getActiveLayer() As pdLayer
    If numOfLayers > 0 Then
        Set getActiveLayer = imgLayers(curLayer)
    Else
        Set getActiveLayer = Nothing
    End If
End Function

'Retrieve a layer at an arbitrary position.  Remember that layers in PD are zero-based, so the base layer is layer 0, not 1.
' Also, this function does not check bounds, so make sure the passed value is valid!
Public Function getLayer(ByVal layerIndex As Long) As pdLayer
    Set getLayer = imgLayers(layerIndex)
End Function

'External functions can use this function to request a thumbnail version of the contained image.  Previously, thumbnail
' generation was handled externally, but to prepare for the addition of layers, I will be rewriting all thumbnail-related
' functions to use this function.  (This is necessary so that this pdImage object can composite a multi-layer image
' before passing along a thumbnail version.)
'
'FreeImage is required for thumbnail resizing.
Public Function requestThumbnail(ByRef dstThumbnailDIB As pdDIB, Optional ByVal thumbnailSize As Long = 64) As Boolean
    
    'Thumbnails have some interesting requirements.  We always want them to be square, with the image set in the middle
    ' of the thumbnail (with aspect ratio preserved) and any empty edges made transparent.
    
    'Start by determining an aspect ratio for the current image.
    Dim aspectRatio As Double
    aspectRatio = CSng(Width) / CSng(Height)
    
    'We also need to determine the thumbnail's actual width and height, and any x and y offset necessary to preserve the
    ' aspect ratio and center the image on the thumbnail.
    Dim tIcoWidth As Double, tIcoHeight As Double, tX As Double, tY As Double
    
    'If the form is wider than it is tall...
    If aspectRatio > 1 Then
        
        'Determine proper sizes and (x, y) positioning so the icon will be centered
        tIcoWidth = thumbnailSize
        tIcoHeight = thumbnailSize * (1 / aspectRatio)
        tX = 0
        tY = (thumbnailSize - tIcoHeight) / 2
        
    Else
    
        'Same thing, but with the math adjusted for images taller than they are wide
        tIcoHeight = thumbnailSize
        tIcoWidth = thumbnailSize * aspectRatio
        tY = 0
        tX = (thumbnailSize - tIcoWidth) / 2
        
    End If
    
    'There are two possible ways to create a thumbnail image.  If FreeImage is available, we prefer to use it, as it
    ' provides superior results, but if it is not available, GDI+ will suffice.
    If g_ImageFormats.FreeImageEnabled Then
        
        'Convert our current DIB to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(getCompositedImage().getDIBDC)
        
        'Use that handle to request an image resize
        If fi_DIB <> 0 Then
            
            'Rescale the image
            Dim returnDIB As Long
            returnDIB = FreeImage_RescaleByPixel(fi_DIB, CLng(tIcoWidth), CLng(tIcoHeight), True, FILTER_BILINEAR)
            
            'Make sure the image is 32bpp (returns a clone of the image if it's already 32bpp, so no harm done)
            Dim newDIB32 As Long
            newDIB32 = FreeImage_ConvertTo32Bits(returnDIB)
            
            'Unload the original DIB
            If newDIB32 <> returnDIB Then FreeImage_UnloadEx returnDIB
            'If the image isn't square-shaped, we need to enlarge the DIB accordingly. FreeImage provides a function for that.
            
            'Start by preparing a transparent quad, which we'll assign to the background of the enlarged area
            Dim newColor As RGBQUAD
            With newColor
                .rgbBlue = 255
                .rgbGreen = 255
                .rgbRed = 255
                .rgbReserved = 0
            End With
                
            'Enlarge the canvas as necessary
            Dim finalDIB As Long
            finalDIB = FreeImage_EnlargeCanvas(newDIB32, tX, tY, tX, tY, newColor, FI_COLOR_IS_RGBA_COLOR)
            
            'Unload the original DIB
            If finalDIB <> newDIB32 Then FreeImage_UnloadEx newDIB32
                
            'At this point, finalDIB contains the 32bpp alpha icon exactly how we want it.  Copy it into the destination DIB.
            dstThumbnailDIB.createBlank thumbnailSize, thumbnailSize, 32
            SetDIBitsToDevice dstThumbnailDIB.getDIBDC, 0, 0, thumbnailSize, thumbnailSize, 0, 0, 0, thumbnailSize, ByVal FreeImage_GetBits(finalDIB), ByVal FreeImage_GetInfo(finalDIB), 0&
                
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            requestThumbnail = True
        
        Else
            requestThumbnail = False
        End If
        
    Else
    
        dstThumbnailDIB.createBlank thumbnailSize, thumbnailSize, 32, 0
        
        Dim compositedImage As New pdDIB
        Set compositedImage = getCompositedImage()
        
        requestThumbnail = GDIPlusResizeDIB(dstThumbnailDIB, tX, tY, tIcoWidth, tIcoHeight, compositedImage, 0, 0, compositedImage.getDIBWidth, compositedImage.getDIBHeight, InterpolationModeHighQualityBilinear)
        
    End If
    
End Function

'When this DIB is no longer being used, we can deactivate it to save on resources.
Public Sub deactivateImage()
    
    'Erase the main DIB
    If Not (mainDIB Is Nothing) Then
        mainDIB.eraseDIB
        Set mainDIB = Nothing
    End If
    
    'Erase the back buffer
    If Not (backBuffer Is Nothing) Then
        backBuffer.eraseDIB
        Set backBuffer = Nothing
    End If
    
    'If the alpha adjustment DIB exists, erase it too
    If Not (alphaFixDIB Is Nothing) Then
        alphaFixDIB.eraseDIB
        Set alphaFixDIB = Nothing
    End If
    
    'If a selection exists, wipe it
    If Not (mainSelection Is Nothing) Then Set mainSelection = Nothing
        
    'Deactivate the Undo/Redo handler
    If Not (undoManager Is Nothing) Then
        undoManager.clearUndos
        Set undoManager.parentPDImage = Nothing
        Set undoManager = Nothing
    End If
    
    'Release the viewport renderer
    If Not (imgViewport Is Nothing) Then Set imgViewport = Nothing
    
    'It doesn't make a big difference, but we can also empty out this image's String-type variables to save a bit of space.
    originalFileName = ""
    originalFileNameAndExtension = ""
    locationOnDisk = ""
    
    'Delete the file's Autosave entry, if any
    If FileExist(g_UserPreferences.getAutosavePath & imageID & ".xml") Then Kill g_UserPreferences.getAutosavePath & imageID & ".xml"
    
    'Mark this image as inactive
    IsActive = False
    
End Sub
    
'Outside actions (such as saving) can affect the HasBeenSaved variable.  However, because we need to do additional
' processing based on the state of this variable, we provide this interface.
Public Sub setSaveState(ByVal newSaveState As Boolean)

    If newSaveState Then
        
        hasBeenSaved = True
        
        'Remember the undo value at this juncture; if the user performs additional actions, but "Undos" to this point,
        ' we want to disable the save button for them
        undoManager.notifyImageSaved
        
    Else
        hasBeenSaved = False
    End If
    
    'Any time the save state is modified, write a new AutoSave entry out to file.
    If IsActive And (Not forInternalUseOnly) Then writeInternalDataToFile g_UserPreferences.getAutosavePath & imageID & ".xml"
    
End Sub

'If the images's size has somehow changed, this can be called to update it.  (At present, it simply copies over the size of
' the main DIB, but in the future it may need to take into account multiple layers.)
Public Function updateSize()
    Me.Width = mainDIB.getDIBWidth
    Me.Height = mainDIB.getDIBHeight
End Function

'INITIALIZE class
Private Sub Class_Initialize()

    'Initially, mark the image as *not* having been saved
    IsActive = False
    loadedSuccessfully = False
    hasBeenSaved = False
    forInternalUseOnly = False
    hasSeenJPEGPrompt = False
    pngBackgroundColor = -1
    
    'Initialize the main DIB (currently, just the image data - layers aren't actually implemented)
    Set mainDIB = New pdDIB
    
    'Initialize the back buffer
    Set backBuffer = New pdDIB
    
    'Initialize the alpha composite layer
    Set alphaFixDIB = New pdDIB
    
    'Initialize the main selection
    Set mainSelection = New pdSelection
    selectionActive = False
    mainSelection.setSelectionShape sRectangle
    Set mainSelection.containingPDImage = Me
    
    'Initialize the metadata object (which may not get used, but this prevents errors if other functions try to access metadata)
    Set imgMetadata = New pdMetadata
    
    'Initialize the Undo/Redo handler
    Set undoManager = New pdUndo
    Set undoManager.parentPDImage = Me
    
    'Initialize the ICC profile manager
    Set ICCProfile = New pdICCProfile
    
    'Initialize the viewport manager
    Set imgViewport = New pdViewport
    
    'Create at least one blank layer
    ReDim imgLayers(0) As pdLayer
    curLayer = 0
        
End Sub

'Return a DIB that contains the currently selected area, fully processed according to the selection mask
Public Function retrieveProcessedSelection(ByRef dstDIB As pdDIB, Optional ByVal preMultipliedAlphaState As Boolean = False) As Boolean

    'If this image does not contain an active selection, exit now.
    If (Not IsActive) Or (Not selectionActive) Then
        retrieveProcessedSelection = False
        Exit Function
    End If
    
    'Selections can be one of several types.  Right now, we don't give special handling to simple rectangular selections - all selections
    ' are fully processed according to the contents of the mask.  Also, all selections are currently created as 32bpp DIBs.
    
    'Start by initializing the destination DIB to the size of the active selection
    Set dstDIB = New pdDIB
    dstDIB.createBlank mainSelection.boundWidth, mainSelection.boundHeight, 32
    
    'We now need pointers to three separate sets of image data: destination DIB, source DIB, and selection mask.
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, getActiveDIB()
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    Dim selData() As Byte
    Dim selSA As SAFEARRAY2D
    prepSafeArray selSA, mainSelection.selMask
    CopyMemory ByVal VarPtrArray(selData()), VarPtr(selSA), 4
            
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    Dim leftOffset As Long, topOffset As Long
    leftOffset = mainSelection.boundLeft
    topOffset = mainSelection.boundTop
    
    Dim x As Long, y As Long
    Dim r As Long, g As Long, b As Long
    Dim thisAlpha As Long, origAlpha As Long
    Dim blendAlpha As Double
    
    Dim srcDIBCD As Long, srcQuickX As Long, srcQuickY As Long
    srcDIBCD = getActiveDIB().getDIBColorDepth \ 8
    
    Dim dstDIBCD As Long, dstQuickX As Long
    dstDIBCD = dstDIB.getDIBColorDepth \ 8
            
    For x = 0 To dstDIB.getDIBWidth - 1
        dstQuickX = x * dstDIBCD
        srcQuickX = (leftOffset + x) * srcDIBCD
    For y = 0 To dstDIB.getDIBHeight - 1
    
        srcQuickY = topOffset + y
        thisAlpha = selData((leftOffset + x) * 3, srcQuickY)
        
        'Check for the presence of alpha data in the original image.  If found, use that AND the selection mask to set the
        ' alpha value of the destination pixel.  If the source data is 32bpp, use only the selection mask data.
        If srcDIBCD = 4 Then
        
            'Source pixel data will be premultiplied, so we must un-premultiply it now
            origAlpha = srcImageData(srcQuickX + 3, srcQuickY)
            If origAlpha > 0 Then
                blendAlpha = 255 / origAlpha
            Else
                blendAlpha = 0
            End If
            
            r = srcImageData(srcQuickX + 2, srcQuickY)
            g = srcImageData(srcQuickX + 1, srcQuickY)
            b = srcImageData(srcQuickX, srcQuickY)
            
            r = r * blendAlpha
            g = g * blendAlpha
            b = b * blendAlpha
            
            If r > 255 Then r = 255
            If g > 255 Then g = 255
            If b > 255 Then b = 255
            
            dstImageData(dstQuickX + 2, y) = r
            dstImageData(dstQuickX + 1, y) = g
            dstImageData(dstQuickX, y) = b
            
            'We have now set un-premultiplied alpha data to the destination.  Finish our work by calculating
            ' a new alpha channel value for this pixel, which is a blend of the original alpha value, and the
            ' selection mask value at this location.
            blendAlpha = thisAlpha / 255
            dstImageData(dstQuickX + 3, y) = origAlpha * blendAlpha
            
        Else
            dstImageData(dstQuickX, y) = srcImageData(srcQuickX, srcQuickY)
            dstImageData(dstQuickX + 1, y) = srcImageData(srcQuickX + 1, srcQuickY)
            dstImageData(dstQuickX + 2, y) = srcImageData(srcQuickX + 2, srcQuickY)
            dstImageData(dstQuickX + 3, y) = thisAlpha
        End If
        
    Next y
    Next x
    
    'Clear all array references
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(selData), 0&, 4
    
    Erase srcImageData
    Erase dstImageData
    Erase selData
    
    'If the calling function requested premultiplied alpha, apply it now
    If preMultipliedAlphaState Then dstDIB.fixPremultipliedAlpha True
        
    retrieveProcessedSelection = True

End Function
