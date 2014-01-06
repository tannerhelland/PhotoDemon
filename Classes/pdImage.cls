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
'Last updated: 13/September/13
'Last update: added a new requestThumbnail function, which external functions can use to request a square thumbnail
'              of an arbitrary size.
'
'The pdImage class is used to store information on each image loaded by the user.  One copy of this
' class exists for each loaded image.  These copies are stored in the pdImages() array, which is declared
' in the MDIWindow module.
'
'The purpose of this class is to store a bunch of variables related to a given image - things like size, zoom,
' name, file path, format containing form, Undo/Redo tracking - as well as a reference to a pdLayer object
' called "mainLayer", which contains an in-memory copy of the relevant image's data.  In the future,
' additional layers may be utilized within each image.
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

    'The ID of the MDI child form containing this particular image
    Public containingForm As Form
    
    'Image dimensions
    Public Width As Long
    Public Height As Long
    
    'Image resolution (in DPI).  In general, we don't deal with differing x/y resolutions - in that case,
    ' use the imageDPI value, which will return the average resolution of the two.
    Private xResolution As Long
    Private yResolution As Long
    Private imageDPI As Long
    
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
    
    'For PNG files, if the original PNG file has a background color, it will be stored here.
    Public pngBackgroundColor As Long
    
    'Has this image been saved?  Access this variable via the getSaveState and setSaveState functions.
    Private hasBeenSaved As Boolean
    
    'These two variables are only used at form load; they can be safely ignored once the image is visible.
    Public WindowLeft As Long
    Public WindowTop As Long
    
    'Track the current window state; if the image's containing form goes from maximized to normal, all images need their
    ' buffer refreshed (as scroll bars may then be needed)
    Public WindowState As Long
        
    'A copy of the form's current icon
    Public curFormIcon32 As Long, curFormIcon16 As Long
    
    'If this image is simply being used for internal PhotoDemon processing, hide it from the user
    Public forInternalUseOnly As Boolean
    
    'Image's actual image data (stored as a DIB inside the mainLayer object)
    Public mainLayer As pdLayer
    
    'Image's back buffer (also a DIB; this is used to hold the final composite before everything is rendered on-screen
    Public backBuffer As pdLayer
    
    'If this image is 32bpp then it requires a specialized compositing layer for rendering the transparency
    Public alphaFixLayer As pdLayer
    
    'Image's selection data (stored and handled by a mainSelection object)
    Public mainSelection As pdSelection
    Public selectionActive As Boolean
    
    'Metadata handler/storage class.  This class is automatically filled with an image file's metadata at load-time.
    ' NOTE: EXIFTOOL IS REQUIRED FOR METADATA HANDLING.
    Public imgMetadata As pdMetadata
    
    'This image's location in the window manager.  We cache this value so we can quickly access information about this image's window.
    Public indexInWindowManager As Long
    
    'All Undo/Redo actions are handled through this publicly available Undo/Redo handler class
    Public undoManager As pdUndo
    
    'All ICC Profile interactions are handled through a publicly available ICCProfile class
    Public ICCProfile As pdICCProfile
    
    'Various viewport actions are now handled through a separate pdViewport class
    Public imgViewport As pdViewport

'Get/Set image resolution (in DPI).  Note that the vertical resolution is optional; if the two values
' differ, PD will average them when image DPI is requested.
Public Function getDPI() As Long
    
    If imageDPI = 0 Then
        getDPI = 96
    Else
        getDPI = imageDPI
    End If
    
End Function

Public Sub setDPI(ByVal xRes As Long, ByVal yRes As Long)
    
    'Many image types do not store resolution information; default to 96 in this case
    If xRes = 0 Then xRes = 96
    If yRes = 0 Then yRes = 96
    
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

'Returns all layers of the image as a single, composited image (in pdLayer format, of course).
Public Function getCompositedImage() As pdLayer
    Set getCompositedImage = mainLayer
End Function

'In preparation for layers support, I will be switching various external modules to use this function instead of a
' direct reference to the mainLayer object.  DO NOT DIRECTLY ACCESS THIS pdImage's MAINLAYER OBJECT IN NEW CODE!
Public Function getActiveLayer() As pdLayer
    Set getActiveLayer = mainLayer
End Function

'External functions can use this function to request a thumbnail version of the contained image.  Previously, thumbnail
' generation was handled externally, but to prepare for the addition of layers, I will be rewriting all thumbnail-related
' functions to use this function.  (This is necessary so that this pdImage object can composite a multi-layer image
' before passing along a thumbnail version.)
'
'FreeImage is required for thumbnail resizing.
Public Function requestThumbnail(ByRef dstThumbnailLayer As pdLayer, Optional ByVal thumbnailSize As Long = 64) As Boolean
    
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
        
        'Convert our current layer to a FreeImage-type DIB
        Dim fi_DIB As Long
        fi_DIB = FreeImage_CreateFromDC(getCompositedImage().getLayerDC)
        
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
                
            'At this point, finalDIB contains the 32bpp alpha icon exactly how we want it.  Copy it into the destination layer.
            dstThumbnailLayer.createBlank thumbnailSize, thumbnailSize, 32
            SetDIBitsToDevice dstThumbnailLayer.getLayerDC, 0, 0, thumbnailSize, thumbnailSize, 0, 0, 0, thumbnailSize, ByVal FreeImage_GetBits(finalDIB), ByVal FreeImage_GetInfo(finalDIB), 0&
                
            'With the transfer complete, release the FreeImage DIB and unload the library
            If returnDIB <> 0 Then FreeImage_UnloadEx returnDIB
            requestThumbnail = True
        
        Else
            requestThumbnail = False
        End If
        
    Else
    
        dstThumbnailLayer.createBlank thumbnailSize, thumbnailSize, 32, 0
        
        Dim compositedImage As New pdLayer
        Set compositedImage = getCompositedImage()
        
        requestThumbnail = GDIPlusResizeLayer(dstThumbnailLayer, tX, tY, tIcoWidth, tIcoHeight, compositedImage, 0, 0, compositedImage.getLayerWidth, compositedImage.getLayerHeight, InterpolationModeHighQualityBilinear)
        
    End If
    
End Function

'When this layer is no longer being used, we can deactivate it to save on resources.
Public Sub deactivateImage()
    
    'Erase the main layer
    If Not (mainLayer Is Nothing) Then
        mainLayer.eraseLayer
        Set mainLayer = Nothing
    End If
    
    'Erase the back buffer
    If Not (backBuffer Is Nothing) Then
        backBuffer.eraseLayer
        Set backBuffer = Nothing
    End If
    
    'If the alpha adjustment layer exists, erase it too
    If Not (alphaFixLayer Is Nothing) Then
        alphaFixLayer.eraseLayer
        Set alphaFixLayer = Nothing
    End If
    
    'If a selection exists, wipe it
    If Not (mainSelection Is Nothing) Then Set mainSelection = Nothing
    
    'Remove the reference to the parent form so VB can wipe it
    If Not (containingForm Is Nothing) Then Set containingForm = Nothing
    
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
    
End Sub

'If the images's size has somehow changed, this can be called to update it.  (At present, it simply copies over the size of
' the main layer, but in the future it may need to take into account multiple layers.)
Public Function updateSize()
    Me.Width = mainLayer.getLayerWidth
    Me.Height = mainLayer.getLayerHeight
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
    
    'Initialize the main layer (currently, just the image data - layers aren't actually implemented)
    Set mainLayer = New pdLayer
    
    'Initialize the back buffer
    Set backBuffer = New pdLayer
    
    'Initialize the alpha composite layer
    Set alphaFixLayer = New pdLayer
    
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
        
End Sub

'Return a layer that contains the currently selected area, fully processed according to the selection mask
Public Function retrieveProcessedSelection(ByRef dstLayer As pdLayer, Optional ByVal preMultipliedAlphaState As Boolean = False) As Boolean

    'If this image does not contain an active selection, exit now.
    If (Not IsActive) Or (Not selectionActive) Then
        retrieveProcessedSelection = False
        Exit Function
    End If
    
    'Selections can be one of several types.  Right now, we don't give special handling to simple rectangular selections - all selections
    ' are fully processed according to the contents of the mask.  Also, all selections are currently created as 32bpp layers.
    
    'Start by initializing the destination layer to the size of the active selection
    Set dstLayer = New pdLayer
    dstLayer.createBlank mainSelection.boundWidth, mainSelection.boundHeight, 32
    
    'We now need pointers to three separate sets of image data: destination layer, source layer, and selection mask.
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, getActiveLayer()
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    Dim selData() As Byte
    Dim selSA As SAFEARRAY2D
    prepSafeArray selSA, mainSelection.selMask
    CopyMemory ByVal VarPtrArray(selData()), VarPtr(selSA), 4
            
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstLayer
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    Dim leftOffset As Long, topOffset As Long
    leftOffset = mainSelection.boundLeft
    topOffset = mainSelection.boundTop
    
    Dim x As Long, y As Long
    Dim r As Long, g As Long, b As Long
    Dim thisAlpha As Long, origAlpha As Long
    Dim blendAlpha As Double
    
    Dim srcLayerCD As Long, srcQuickX As Long, srcQuickY As Long
    srcLayerCD = getActiveLayer().getLayerColorDepth \ 8
    
    Dim dstLayerCD As Long, dstQuickX As Long
    dstLayerCD = dstLayer.getLayerColorDepth \ 8
            
    For x = 0 To dstLayer.getLayerWidth - 1
        dstQuickX = x * dstLayerCD
        srcQuickX = (leftOffset + x) * srcLayerCD
    For y = 0 To dstLayer.getLayerHeight - 1
    
        srcQuickY = topOffset + y
        thisAlpha = selData((leftOffset + x) * 3, srcQuickY)
        
        'Check for the presence of alpha data in the original image.  If found, use that AND the selection mask to set the
        ' alpha value of the destination pixel.  If the source data is 32bpp, use only the selection mask data.
        If srcLayerCD = 4 Then
        
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
    If preMultipliedAlphaState Then dstLayer.fixPremultipliedAlpha True
        
    retrieveProcessedSelection = True

End Function
