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
'Copyright 2006-2015 by Tanner Helland
'Created: sometime 2006
'Last updated: 13/September/15
'Last update: convert imgStorage to pdDictionary, so we can drop the run-time scrrun.dll requirement
'
'The pdImage class is used to store information on each image loaded by the user.  As such, one instance of this class
' exists for each loaded image.  These instances are stored in the public pdImages() array.
'
'This class doesn't do a lot of its own processing.  At present, it's mainly just a container for key image properties,
' e.g. size, name, file path, format, Undo/Redo tracking - as well as the all-important array of pdLayer objects that
' comprise the image's actual bitmap data.
'
'Generally speaking, this class should not be used separate from its parent pdImages() array.  If you need access to
' a temporary image, use pdDIB, or if layer interactions are required, a single pdLayer instance (and its corresponding
' layerDIB).  This class is quite heavy as it must track things like metadata, undo/redo handling, and other items,
' so avoid using it for any temporary image operations.
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
    
    'Any random data relevant to the image can be stored in this Dictionary object.  It replaces the wide array of specific
    ' public variables we previously used.
    Public imgStorage As pdDictionary
        
    'Has this image been saved?  Access this variable via the getSaveState and setSaveState functions.
    Private hasBeenSavedPDI As Boolean, hasBeenSavedFlat As Boolean
        
    'A copy of the image's current icon (at 16 and 32px sizes)
    Public curFormIcon32 As Long, curFormIcon16 As Long
    
    'If this image is simply being used for internal PhotoDemon processing, hide it from the user
    Public forInternalUseOnly As Boolean
    
    'Persistent buffer, for storing the fully composited image.  This composite is only updated when absolutely necessary;
    ' otherwise, look at grabbing the dedicated canvasBuffer, below.
    Public compositeBuffer As pdDIB
    
    'Image's canvas buffer.  This holds a composited version of the image, extracted and cropped to reflect the current viewport.
    ' It is cached here to make viewport rendering faster and more reliable.
    Public canvasBuffer As pdDIB
    
    'If this image is 32bpp then it requires a specialized compositing layer for rendering the transparency
    Public alphaFixDIB As pdDIB
    
    'Image's selection data (stored and handled by a mainSelection object)
    Public mainSelection As pdSelection
    Public selectionActive As Boolean
    
    'Metadata handler/storage class.  This class is automatically filled with an image file's metadata at load-time.
    ' NOTE: THE EXIFTOOL PLUGIN IS REQUIRED FOR METADATA HANDLING.
    Public imgMetadata As pdMetadata
    
    'All Undo/Redo actions are handled through this publicly available Undo/Redo handler class
    Public undoManager As pdUndo
        
    'Various viewport actions are now handled through a separate pdViewport class
    Public imgViewport As pdViewport
    
    'Image layers!  All layers are stored in this array.
    Private imgLayers() As pdLayer
    
    'Current layer.  Layers in PD are zero-based, with the 0 layer being created by default when a pdImage is created.
    ' To ensure proper behavior for things like caching viewports, this value is not publicly accessible - you must access
    ' it via the relevant get/set functions.
    Private m_curLayer As Long
    
    'Total number of layers EVER created for this image.  This is used to assign canonical layer IDs to individual layers.
    ' It has no relation to layer position or order, and will be constant for the life of both the layer itself, and this
    ' pdImage object (including persistence when saving to/from file).
    Private m_numOfLayersEverCreated As Long
    
    'Number of ACTIVE layers in the current image.  This is equal to or less than m_numOfLayers, above.  Note that it is not
    ' zero-based; e.g. right after an image file is loaded, this will be set to 1.
    Private m_numOfLayers As Long
    
    'Image compositor.  Previously, all compositing was done internally, but as compositing has grown more complex, I have decided
    ' to move all compositing work to an external class (to keep pdImage from getting too bloated).
    Private imgCompositor As pdCompositor
    
    'As of April 2015, each pdImage object now maintains a persistent copy of its fully composited layer stack (inside the
    ' compositeBuffer pdDIB).  To avoid recompositing the image unless absolutely necessary, new trackers were required.
    Private m_CompositeCacheClean As Boolean
    
    
'Create a new, blank layer in this image.  Note that layer type is not provided - that is handled separately, via the
' layer object itself.  All this function does is create a new pdLayer object and rearrange the layer stack to fit.

'Optionally, a layer position can be passed.  The position is assumed to be the position of the
' new layer; e.g. if zero is passed, the layer will be placed at the bottom of the stack.  If no position is requested,
' the layer will be created above the currently active layer.
'
'The returned Long-type value is the CANONICAL ID of the new layer, NOT ITS INDEX.  (Returning its index would be pointless,
' as the calling function likely knows that in advance.)  Note that layers can be accessed by either ID or index; after
' calling this function, make sure to use the CANONICAL ID function with the returned value.
Public Function createBlankLayer(Optional ByVal layerPosition As Long = -1) As Long

    'Start by seeing if the image has zero layers.  If it does, we can ignore layerPosition entirely.
    If m_numOfLayersEverCreated = 0 Then
    
        'Make sure the current layer pointer is correct
        m_curLayer = 0
        
    'The layer has images.  Extra work is involved
    Else

        'If no layer has been specified, insert the new layer above the requested layer
        If layerPosition = -1 Then m_curLayer = m_curLayer + 1 Else m_curLayer = layerPosition + 1
        
        'Resize the layers array by one, and shift all existing layers upward.
        ReDim Preserve imgLayers(0 To m_numOfLayers + 1) As pdLayer
        
        Dim i As Long
        For i = UBound(imgLayers) To m_curLayer Step -1
            Set imgLayers(i) = imgLayers(i - 1)
        Next i
        
    End If
    
    'Initialize the new layer.  (This should already have been done at creation time, but it doesn't hurt to make sure.)
    Set imgLayers(m_curLayer) = New pdLayer
    
    'Assign the layer a canonical ID value.
    imgLayers(m_curLayer).assignLayerID m_numOfLayersEverCreated
    
    'Increment our internal layer counts
    m_numOfLayersEverCreated = m_numOfLayersEverCreated + 1
    m_numOfLayers = m_numOfLayers + 1
    
    'Return the new layer's ID value (which is the number of layers ever created for this image)
    createBlankLayer = m_numOfLayersEverCreated - 1

End Function

Public Sub rebuildCompositeBuffer()

    'If the buffer is already clean, ignore this request
    If m_CompositeCacheClean Then
        Debug.Print "WARNING!  pdImage.rebuildCompositeBuffer was called, but the compositeBuffer is marked as clean!"
    Else
        
        'Rebuild the composite cache
        imgCompositor.getCompositedImage Me, compositeBuffer
        
        'Mark the cache as clean
        'm_CompositeCacheClean = True
        
    End If

End Sub

'If an outside function directly modifies something inside this image (e.g. a layer), it must notify us via this function.
' The modification type mirrors PD's Undo types, to simply the work external functions have to do (as typically, they are
' notifying the central processor of these changes, so Undo data is generated correctly).
'
'If the change involves a layer, an index *must* be specified, so the object can notify the layer of the change (allowing the
' layer to rebuild any internal caches or buffers, like layer styles).
Public Function notifyImageChanged(ByVal typeOfChange As PD_UNDO_TYPE, Optional ByVal layerIndex As Long = -1)

    'Separate handling by type.
    
    'TODO: in the future, it would be awesome to store a partial composite up-to-but-NOT-including the current layer.
    ' This would allow for faster recomposites when the current layer is modified.
    ' However, until such time as that feature is implemented, I have no choice but to recomposite the entire image
    ' after changes are made.
    
    Select Case typeOfChange
    
        Case UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE
            m_CompositeCacheClean = False
        
        Case UNDO_IMAGEHEADER
            m_CompositeCacheClean = False
        
        Case UNDO_LAYER, UNDO_LAYER_VECTORSAFE
            m_CompositeCacheClean = False
            
            If layerIndex <> -1 Then
                Me.getLayerByIndex(layerIndex).notifyOfDestructiveChanges
            Else
                Debug.Print "WARNING!  Invalid params passed to pdImage.notifyImageChanged (-1 layer index for a layer action)."
            End If
        
        Case UNDO_LAYERHEADER
            m_CompositeCacheClean = False
        
        Case Else
            Debug.Print "WARNING!  pdImage.notifyImageChanged received an unknown change request: " & typeOfChange
    
    End Select

End Function

'Some functions (like saving the image) need to know whether to present the user with 24 or 32bpp options.  Because the composited
' image's color depth depends on a variety of factors, the calling function can simply rely on this function to determine the
' appropriate color depth for us.
Public Function getCompositeImageColorDepth() As Long

    'NOTE: as of April '14 (v6.4) newly added layer support dictates that all layers are 32bpp by default.
    getCompositeImageColorDepth = 32
    
    'Single-layer images are much easier to deal with!
    'If m_numOfLayers = 0 Then
    '    getCompositeImageColorDepth = imgLayers(m_curLayer).layerDIB.getDIBColorDepth
    'Else
        'Use magic to determine the ideal color depth of the composited image??
    'End If

End Function

'Merge two layers together.  Note this can be used to merge any two arbitrary layers, with the bottom layer holding the result
' of the merge.  It is up to the caller to deal with any subsequent layer deletions, etc - this sub just performs the merge.
Public Sub mergeTwoLayers(ByRef topLayer As pdLayer, ByRef bottomLayer As pdLayer, Optional ByVal bottomLayerIsFullSize As Boolean = True)
    imgCompositor.mergeLayers topLayer, bottomLayer, bottomLayerIsFullSize
End Sub

'Returns all layers of the image as a single, composited image (in pdDIB format, of course).  Because of the way VB handles
' object references, we ask the calling function to supply the DIB they want filled.  Optionally, they can also request a
' particular premultiplication status of the composited DIB's alpha values.  (This is helpful for save functions, which
' require non-premultiplied alpha, vs viewport functions, which require premultiplied alpha).
Public Sub getCompositedImage(ByRef dstDIB As pdDIB, Optional ByVal premultiplicationStatus As Boolean = True)
    
    'If the current composite cache is dirty, rebuild it
    If Not m_CompositeCacheClean Then rebuildCompositeBuffer
    
    'Give the caller a copy of the composited image
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.createFromExistingDIB compositeBuffer
    
    'Change premultiplication, if requested by the caller.  (By default, the composite image is stored in premultiplied format.)
    If dstDIB.getAlphaPremultiplication <> premultiplicationStatus Then dstDIB.SetAlphaPremultiplication premultiplicationStatus
    
End Sub

'Returns a subsection of the fully composited image (in pdDIB format, of course).  This is helpful for rendering the main viewport,
' as we only composite the relevant portions of the image.
Public Sub getCompositedRect(ByRef dstDIB As pdDIB, ByVal dstX As Single, ByVal dstY As Single, ByVal dstWidth As Single, ByVal dstHeight As Single, ByVal srcX As Single, ByVal srcY As Single, ByVal srcWidth As Single, ByVal srcHeight As Single, ByVal interpolationType As InterpolationMode, Optional ByVal ignoreInternalCaches As Boolean = False, Optional ByVal levelOfDetail As COMPOSITOR_LOD_CLUE = CLC_Generic)
    imgCompositor.getCompositedRect Me, dstDIB, dstX, dstY, dstWidth, dstHeight, srcX, srcY, srcWidth, srcHeight, interpolationType, ignoreInternalCaches, levelOfDetail
End Sub

'Get the currently active layer index
Public Function getActiveLayerIndex() As Long
    getActiveLayerIndex = m_curLayer
End Function

'Get the currently active layer canonical ID
Public Function getActiveLayerID() As Long
    If m_curLayer > UBound(imgLayers) Then
        getActiveLayerID = imgLayers(0).getLayerID
    Else
        getActiveLayerID = imgLayers(m_curLayer).getLayerID
    End If
End Function

'Get the number of layers in this image.  (Note: this function might return zero, so handle that condition correctly
' in calling functions!)
Public Function getNumOfLayers() As Long
    getNumOfLayers = m_numOfLayers
End Function

'Retrieve a reference to the currently active layer's DIB.  This is effectively a shortcut function when we
' need quick access to the current layer's DIB.
Public Function getActiveDIB() As pdDIB
    Set getActiveDIB = imgLayers(m_curLayer).layerDIB
End Function

'Retrieve the currently active layer
Public Function getActiveLayer() As pdLayer
    If m_numOfLayers > 0 Then
        
        'Failsafe check for a valid current layer
        If m_curLayer > UBound(imgLayers) Then m_curLayer = 0
        Set getActiveLayer = imgLayers(m_curLayer)
        
    Else
        Set getActiveLayer = Nothing
    End If
End Function

'Set the currently active layer by its cardinal ID value
Public Sub setActiveLayerByID(ByVal newActiveLayerID As Long)
    
    Dim i As Long
    For i = 0 To UBound(imgLayers)
        If imgLayers(i).getLayerID = newActiveLayerID Then
            m_curLayer = i
            Exit For
        End If
    Next i
    
    'If we made it all the way here, the requested layer was not found.  Do nothing (e.g. do not change the active layer).
    
End Sub

Public Sub setActiveLayerByIndex(ByVal newActiveLayerIndex As Long)

    'Validate the incoming layer index
    If newActiveLayerIndex < 0 Then newActiveLayerIndex = 0
    If newActiveLayerIndex > m_numOfLayers - 1 Then newActiveLayerIndex = m_numOfLayers - 1
    
    m_curLayer = newActiveLayerIndex
    
End Sub

'Retrieve a layer index using its canonical ID.
Public Function getLayerIndexFromID(ByVal requestedID As Long) As Long
    
    Dim i As Long
    For i = 0 To UBound(imgLayers)
        If imgLayers(i).getLayerID = requestedID Then
            getLayerIndexFromID = i
            Exit Function
        End If
    Next i
    
    'If we made it all the way here, the requested layer was not found.  Return 0.
    getLayerIndexFromID = 0
    
End Function

'Retrieve a layer at an arbitrary position.  Remember that layers in PD are zero-based, so the base layer is layer 0, not 1.
' Also, this function does not check bounds, so make sure the passed index is valid!
Public Function getLayerByIndex(ByVal layerIndex As Long) As pdLayer
    
    'As a failsafe, validate the incoming layer index
    If layerIndex > UBound(imgLayers) Then
        layerIndex = UBound(imgLayers)
        Debug.Print "WARNING! Invalid layerIndex requested (" & layerIndex & ":" & UBound(imgLayers) & ")"
    End If
    
    If layerIndex < 0 Then layerIndex = 0
    
    'Return the requested layer
    'If Not imgLayers(layerIndex) Is Nothing Then
        Set getLayerByIndex = imgLayers(layerIndex)
    'Else
    '    Set getLayerByIndex = Nothing
    'End If
    
End Function

'Retrieve a layer using its canonical ID.  Remember that layers in PD are zero-based, so the base layer is layer 0, not 1.
' Also, this function does not check bounds, so make sure the passed index is valid!
Public Function getLayerByID(ByVal requestedID As Long) As pdLayer
    
    Dim i As Long
    For i = 0 To m_numOfLayers - 1
        If imgLayers(i).getLayerID = requestedID Then
            Set getLayerByID = imgLayers(i)
            Exit Function
        End If
    Next i
    
    'If we made it all the way here, the requested layer was not found.  Return Nothing.
    Debug.Print "WARNING!  getLayerByID received an invalid ID: " & requestedID
    Set getLayerByID = Nothing
    
End Function

'Simple function for moving a given layer up/down in the layer stack.
Public Sub moveLayerByIndex(ByVal srcLayerIndex As Long, ByVal moveLayerUp As Boolean)

    'Before doing anything else, make sure the requested move is a valid one
    If moveLayerUp Then
        If srcLayerIndex = m_numOfLayers - 1 Then Exit Sub
    Else
        If srcLayerIndex = 0 Then Exit Sub
    End If
    
    'Process the actual move
    If moveLayerUp Then
    
        'Move the selected layer UP.
        swapTwoLayers srcLayerIndex, srcLayerIndex + 1
    
    Else
    
        'Move the selected layer DOWN.
        swapTwoLayers srcLayerIndex, srcLayerIndex - 1
    
    End If

End Sub

'Complicated function for moving a layer around in the layer stack.  A current layer index and destination layer index are required.
' The destination layer index will be the *new* index of the source layer.  If the two indices match, the function will terminate.
'
'The function will return TRUE if the layers were moved successfully.
Public Function moveLayerToArbitraryIndex(ByVal srcLayerIndex As Long, ByVal dstLayerIndex As Long) As Boolean

    'Before doing anything else, make sure the requested move is a valid one
    If srcLayerIndex = dstLayerIndex Then
        moveLayerToArbitraryIndex = False
        Exit Function
    End If
        
    If (srcLayerIndex < 0) Or (srcLayerIndex > m_numOfLayers - 1) Then
        moveLayerToArbitraryIndex = False
        Exit Function
    End If
    
    'Validate the destination layer differently; if it lies out-of-bounds, automatically correct it to the nearest
    ' relevant position.
    If dstLayerIndex < 0 Then dstLayerIndex = 0
    If dstLayerIndex > m_numOfLayers - 1 Then dstLayerIndex = m_numOfLayers - 1
    
    'We now know several things:
    ' 1) srcLayerIndex and dstLayerIndex are not the same
    ' 2) srcLayerIndex and dstLayerIndex are valid entries in the layer stack
    
    Dim i As Long
    
    'We can now process the actual layer rearranging, and we will do so in two separate chunks of code: up vs down.
    If dstLayerIndex > srcLayerIndex Then
    
        'The source layer is moving HIGHER in the stack
    
        'Iterate through all layers between the current position and the new one, swapping layers as we go
        For i = srcLayerIndex To dstLayerIndex - 1
            swapTwoLayers i, i + 1
        Next i
    
    Else
    
        'The source layer is moving LOWER in the stack
    
        'Iterate through all layers between the current position and the new one, swapping layers as we go
        For i = srcLayerIndex To dstLayerIndex + 1 Step -1
            swapTwoLayers i, i - 1
        Next i
    
    End If
    
    moveLayerToArbitraryIndex = True

End Function

'Used by various layer movement functions.  Given two layer indices, swap them.
Public Sub swapTwoLayers(ByVal srcLayerIndex_1 As Long, ByVal srcLayerIndex_2 As Long)

    'Create a temporary reference to the first layer
    Dim tmpLayerRef As pdLayer
    Set tmpLayerRef = imgLayers(srcLayerIndex_1)
    
    'Overwrite the first layer's reference with the second one
    Set imgLayers(srcLayerIndex_1) = imgLayers(srcLayerIndex_2)
    
    'Overwrite the second layer's reference with our temporary copy of the first layer
    Set imgLayers(srcLayerIndex_2) = tmpLayerRef
    
    'Release our temporary reference
    Set tmpLayerRef = Nothing

End Sub

'Delete a given layer.  The pdLayers stack will automatically be resized to match.
Public Sub deleteLayerByIndex(ByVal srcLayerIndex As Long)

    'Validate the layer index
    If srcLayerIndex < 0 Or srcLayerIndex >= m_numOfLayers Then Exit Sub
    
    'If this is the last remaining layer, exit
    If m_numOfLayers = 1 Then Exit Sub
    
    'Shift all layer references above this one downward
    Dim i As Long
    For i = srcLayerIndex To m_numOfLayers - 2
        Set imgLayers(i) = imgLayers(i + 1)
    Next i
    
    'Shrink the layers array
    m_numOfLayers = m_numOfLayers - 1
    ReDim Preserve imgLayers(0 To m_numOfLayers - 1) As pdLayer

End Sub

'Write all relevant image information to an external file (in XML format).  Note that this will not be a perfect copy of the
' current image's state, because some pdImage data is only relevant at run-time.  Also, some data structures (like metadata) are
' very difficult to handle properly; I am working on a solution to this, but have not finished it yet.
'
'In the event that you want to do something else with the XML data, I have now added an optional parameter.
' Mark "filePathIsActuallyString" as TRUE to treat the filename string as a destination buffer, and this function will
' write the XML data there (instead of treating it as a filename).
Public Function writeExternalData(ByRef dstFilename As String, Optional ByVal filePathIsActuallyString As Boolean = False) As Boolean

    On Error GoTo failedDataWrite

    'Prepare an XML engine, which will handle the actual writing of the file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Add a basic header and explanatory comment
    xmlEngine.prepareNewXML "pdImage"
    
    xmlEngine.writeBlankLine
    xmlEngine.writeComment "This file contains a summary of crucial data for a given pdImage entry."
    xmlEngine.writeBlankLine
    
    'Start by writing out the imageID.  This value will not be re-used in subsequent session, but it's helpful for matching up
    ' this image's data with any saved Undo/Redo entries (which are created by imageID for privacy reasons)
    xmlEngine.writeTag "ID", imageID
    
    'Start writing any tags that may be useful for reconstructing an autosave version of this image.  We limit the information to
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
    xmlEngine.writeTag "numOfLayersEverCreated", m_numOfLayersEverCreated
    xmlEngine.writeTag "numOfLayers", m_numOfLayers
    xmlEngine.writeTag "curLayer", m_curLayer
    
    'Next comes conditional data, which may not exist in the scripting dictionary.  If it's available, we'll write it out.
    If imgStorage.doesKeyExist("pngBackgroundColor") Then xmlEngine.writeTag "pngBackgroundColor", imgStorage.getEntry_Long("pngBackgroundColor")
    
    xmlEngine.writeBlankLine
    
    'That should be enough information to reconstruct the original image state.  We now have two options, depending on the value
    ' of the "filePathIsActuallyString" parameter.
    
    'The string we were supplied is a blank buffer, not a filename.  Write the XML data to it.  Note that by default, XML indentation
    ' is suppressed; this is for performance reasons.
    If filePathIsActuallyString Then
        dstFilename = xmlEngine.returnCurrentXMLString(True)
    
    'The string we were supplied is a filename.  Write the XML data to it.
    Else
        xmlEngine.writeXMLToFile dstFilename
    End If
    
    writeExternalData = True
    Exit Function
    
failedDataWrite:

    Debug.Print "WARNING: could not write XML data to file.  Are you sure this drive has enough free space?"
    writeExternalData = False
    
End Function

'Sister function to writeExternalDataToFile, above.  If you add a variable to one function, make sure to include it in the other.
' If the function fails, it will return FALSE.
'
'In the event that you have already loaded the XML data yourself (or have compiled it manually), I have added an optional parameter.
' Mark "filePathIsActuallyString" as TRUE to treat the filename string as actual XML data instead.
'
'Finally, if this function is being called as part of an Undo/Redo action, we can safely ignore certain entries in the file, in favor
' of the values we already have stored for this image.  The Undo/Redo engine may also specify non-destructive loading, in which case
' we assume that our existing layer stack is valid, but in the wrong order - so simply reorder the stack according to the data in the
' XML, instead of full-on recreating it from scratch.
Public Function readExternalData(ByRef srcFilename As String, Optional ByVal filePathIsActuallyString As Boolean = False, Optional ByVal sourceIsUndoFile As Boolean = False, Optional ByVal loadNonDestructively As Boolean = False) As Boolean

    On Error GoTo failedDataRead
    
    Dim xmlLoadSuccessful As Boolean
    xmlLoadSuccessful = False

    'Prepare an XML engine, which will handle the actual reading and parsing of the file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'We now have two choices, depending on the value of filePathIsActuallyString.  Load the supposed XML contents
    ' (either from file or directly) into the XML engine; note that both steps validate the incoming data automatically.
    If filePathIsActuallyString Then
        xmlLoadSuccessful = xmlEngine.loadXMLFromString(srcFilename)
    Else
        xmlLoadSuccessful = xmlEngine.loadXMLFile(srcFilename)
    End If
    
    'Make sure the supplied XML actually contains pdImage data.
    If xmlLoadSuccessful And xmlEngine.isPDDataType("pdImage") Then
    
        'The XML file checked out.  Start retrieving relevant values.
        Me.Width = xmlEngine.getUniqueTag_Long("Width")
        Me.Height = xmlEngine.getUniqueTag_Long("Height")
        xResolution = xmlEngine.getUniqueTag_Long("xResolution")
        yResolution = xmlEngine.getUniqueTag_Long("yResolution")
        imageDPI = xmlEngine.getUniqueTag_Long("DPI")
        
        'All settings past this point are only relevant when loading a file anew, not when it's being used as part of
        ' the Undo/Redo stack.
        If Not sourceIsUndoFile Then
        
            'These values are all stored in the PDI file, but they're not really relevant.  I will look at not storing
            ' them in the first place, but for now, just ignore them.
            'TODO: make a firm decision on what data does not belong in PDI files!
            'Me.originalFileSize = xmlEngine.getUniqueTag_Long("originalFileSize")
            'Me.originalFileName = xmlEngine.getUniqueTag_String("originalFileName")
            'Me.originalFileNameAndExtension = xmlEngine.getUniqueTag_String("originalFileNameAndExtension")
            'Me.locationOnDisk = xmlEngine.getUniqueTag_String("locationOnDisk")
            'Me.originalFileFormat = xmlEngine.getUniqueTag_Long("originalFileFormat")
            'Me.currentFileFormat = xmlEngine.getUniqueTag_Long("currentFileFormat")
            Me.originalColorDepth = xmlEngine.getUniqueTag_Long("originalColorDepth")
            
            'Dictionary tags are conditional, and may not exist in the PDI file.  Check for their existence, and load them
            ' only as necessary.
            If xmlEngine.doesTagExist("pngBackgroundColor") Then imgStorage.AddEntry "pngBackgroundColor", xmlEngine.getUniqueTag_Long("pngBackgroundColor")
            
        End If
        
        'Finally, retrieve layer values from the file
        m_numOfLayersEverCreated = xmlEngine.getUniqueTag_Long("numOfLayersEverCreated")
        m_numOfLayers = xmlEngine.getUniqueTag_Long("numOfLayers")
        m_curLayer = xmlEngine.getUniqueTag_Long("curLayer")
        
        'If this is a standard, destructive load, prepare our image layer array to receive the upcoming layer data.
        ' (Otherwise, leave it as it is, because the calling function will simply reorder it for us.)
        If Not loadNonDestructively Then
        
            'Now that all data has been read, we can initialize our layers array.  (The calling function will presumably proceed to
            ' fill each layer with usable data, because we can't do that here!)
            ReDim imgLayers(0 To m_numOfLayers - 1) As pdLayer
            
            Dim i As Long
            For i = 0 To m_numOfLayers - 1
                Set imgLayers(i) = New pdLayer
            Next i
            
        End If
    
    Else
        GoTo failedDataRead
    End If
    
    'The pdImage struct has been created successfully!
    readExternalData = True
    
    Exit Function
    
failedDataRead:
    
    Debug.Print "WARNING: could not validate XML data.  pdImage internals were not updated."
    readExternalData = False
    
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
Public Function getSaveState(ByVal desiredSaveType As PD_SAVE_EVENT) As Boolean
    
    Select Case desiredSaveType
    
        Case pdSE_AnySave
            getSaveState = hasBeenSavedPDI Or hasBeenSavedFlat
        
        Case pdSE_SavePDI
            getSaveState = hasBeenSavedPDI
        
        Case pdSE_SaveFlat
            getSaveState = hasBeenSavedFlat
    
    End Select
    
End Function

'External functions can use this function to request a thumbnail version of the contained image.  Note that the full
' image must be re-composited on every request, so please do not call this function unless absolutely necessary!
Public Function requestThumbnail(ByRef dstThumbnailDIB As pdDIB, Optional ByVal thumbnailSize As Long = 64) As Boolean
    
    'Uncomment the two lines below (and another at the end of the function) to profile different thumbnail resize strategies
    'Dim startTime As Double
    'startTime = Timer
    
    'Thumbnails have some interesting requirements.  We always want them to be square, with the image set in the middle
    ' of the thumbnail (with aspect ratio preserved) and any empty edges made transparent.
    
    'We also need to determine the thumbnail's actual width and height, and any x and y offset necessary to preserve the
    ' aspect ratio and center the image on the thumbnail.
    Dim tIcoWidth As Long, tIcoHeight As Long, tX As Double, tY As Double
    
    'Start by determining proper dimensions for the resized thumbnail image.
    Math_Functions.convertAspectRatio Me.Width, Me.Height, thumbnailSize, thumbnailSize, tIcoWidth, tIcoHeight
    
    'If the form is wider than it is tall, center the thumbnail vertically
    If tIcoWidth > tIcoHeight Then
        tX = 0
        tY = (thumbnailSize - tIcoHeight) / 2
    
    '...otherwise, center it horizontally
    Else
        tY = 0
        tX = (thumbnailSize - tIcoWidth) / 2
    End If
    
    'Retrieve a composite version of the image
    Dim compositeImage As pdDIB
    Set compositeImage = New pdDIB
    compositeImage.createBlank tIcoWidth, tIcoHeight, 32, 0
    
    'Note that the user's thumbnail performance setting affects the quality used here.
    If g_ThumbnailPerformance = PD_PERF_FASTEST Then
        Me.getCompositedRect compositeImage, 0, 0, tIcoWidth, tIcoHeight, 0, 0, Me.Width, Me.Height, InterpolationModeNearestNeighbor, , CLC_Thumbnail
    Else
        Me.getCompositedRect compositeImage, 0, 0, tIcoWidth, tIcoHeight, 0, 0, Me.Width, Me.Height, InterpolationModeHighQualityBicubic, , CLC_Thumbnail
    End If
    
    'Prepare the destination DIB
    If dstThumbnailDIB Is Nothing Then Set dstThumbnailDIB = New pdDIB
    If (dstThumbnailDIB.getDIBWidth <> thumbnailSize) Or (dstThumbnailDIB.getDIBHeight <> thumbnailSize) Then
        dstThumbnailDIB.createBlank thumbnailSize, thumbnailSize, 32, 0
    Else
        dstThumbnailDIB.resetDIB 0
    End If
    
    'There are two possible ways to create a thumbnail image.  If FreeImage is available, we prefer to use it, as it
    ' provides superior results with better performance, but if it is not available, GDI+ will suffice.
    If g_ImageFormats.FreeImageEnabled Then
        requestThumbnail = FreeImageResizeDIBFast(dstThumbnailDIB, tX, tY, tIcoWidth, tIcoHeight, compositeImage, FILTER_BILINEAR, True)
    Else
        requestThumbnail = GDIPlusResizeDIB(dstThumbnailDIB, tX, tY, tIcoWidth, tIcoHeight, compositeImage, 0, 0, compositeImage.getDIBWidth, compositeImage.getDIBHeight, InterpolationModeHighQualityBicubic)
    End If
    
    'Uncomment the line below to profile different thumbnail resize strategies
    'Debug.Print "Time taken for image thumbnail generation: " & CStr((Timer - startTime) * 1000) & " ms"
    
End Function

'When this DIB is no longer being used, we can deactivate it to save on resources.
Public Sub deactivateImage()
    
    'Erase the composite buffer
    If Not (compositeBuffer Is Nothing) Then
        compositeBuffer.eraseDIB
        Set compositeBuffer = Nothing
    End If
    
    'Erase the canvas buffer
    If Not (canvasBuffer Is Nothing) Then
        canvasBuffer.eraseDIB
        Set canvasBuffer = Nothing
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
    
    'Release the compositor
    If Not (imgCompositor Is Nothing) Then Set imgCompositor = Nothing
    
    'It doesn't make a big difference, but we can also empty out this image's String-type variables to save a bit of space.
    originalFileName = ""
    originalFileNameAndExtension = ""
    locationOnDisk = ""
    
    'Mark this image as inactive
    IsActive = False
    
End Sub
    
'Outside actions (such as saving) can affect the HasBeenSaved variable.  However, because we need to do additional
' processing based on the state of this variable, we provide this interface.
Public Sub setSaveState(ByVal newSaveState As Boolean, ByVal typeOfSaveEvent As PD_SAVE_EVENT)
    
    'The image has just been saved
    If newSaveState Then
        
        Select Case typeOfSaveEvent
            
            'This value should never be passed in!
            Case pdSE_AnySave
                Debug.Print "The pdSE_AnySave type is only for GETTING image save state, not SETTING it!  Fix this!"
                
            'Image has been saved to PDI format, meaning layers are intact
            Case pdSE_SavePDI
                hasBeenSavedPDI = True
                
            'Image has been saved to some flat format (JPEG, PNG, etc)
            Case pdSE_SaveFlat
                hasBeenSavedFlat = True
            
        End Select
        
        'Remember the undo value at this juncture; if the user performs additional actions, but "Undos" to this point,
        ' we want to disable the save button for them
        undoManager.notifyImageSaved typeOfSaveEvent
    
    'Some change has occurred, meaning the image has not been saved (to any format) in its current state
    Else
        hasBeenSavedPDI = False
        hasBeenSavedFlat = False
    End If
        
End Sub

'If the images's size has somehow changed, this can be called to update it.
'
'The first optional parameter, assumed to be TRUE, makes the image the same size as its base layer.  This is helpful when
' an image file is first loaded, as the pdImage container should be set to an identical size.
'
'If this is another type of resize action, such as a canvas resize, the caller must supply the new width/height values desired.
'
'Note that this function performs NO VALIDATION on the passed width/height values.  Make sure they are correct in advance!
Public Function updateSize(Optional ByVal useBaseLayer As Boolean = True, Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0)
    
    If useBaseLayer Then
        Me.Width = imgLayers(0).getLayerWidth(False)
        Me.Height = imgLayers(0).getLayerHeight(False)
    Else
        Me.Width = newWidth
        Me.Height = newHeight
    End If
    
    'Changes to image size require a rebuild of our composited image cache
    Me.notifyImageChanged UNDO_IMAGE
    
End Function

'INITIALIZE class
Private Sub Class_Initialize()

    'Initially, mark the image as *not* having been saved
    IsActive = False
    loadedSuccessfully = False
    hasBeenSavedPDI = False
    hasBeenSavedFlat = False
    forInternalUseOnly = False
    
    'Initialize the public storage dictionary
    Set imgStorage = New pdDictionary
    
    'Add initial key/value pairs to the dictionary.  Doing this here saves us from having to check key existence in subsequent functions.
    
        'Lossless filetypes always present an export dialog at least once.  Beyond that point, any use of the Save command will
        ' re-save the file using previous settings.  Save As will still trigger a new dialog, however.
        imgStorage.AddEntry "hasSeenJPEGPrompt", False
        imgStorage.AddEntry "hasSeenJP2Prompt", False
        imgStorage.AddEntry "hasSeenWebPPrompt", False
        imgStorage.AddEntry "hasSeenJXRPrompt", False
        
    'Initialize the composite and canvas buffers
    Set compositeBuffer = New pdDIB
    Set canvasBuffer = New pdDIB
    
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
    
    'Initialize the viewport manager
    Set imgViewport = New pdViewport
    
    'Create at least one blank layer, so that functions referencing a layer object don't break
    ReDim imgLayers(0) As pdLayer
    Set imgLayers(0) = New pdLayer
    
    'The current layer marker always starts at zero (which is considered the background layer).  To check for the case of
    ' no active layers, do not use m_curLayer, use m_numOfLayers.
    m_curLayer = 0
    
    'By default, list this image as having zero layers
    m_numOfLayers = 0
    m_numOfLayersEverCreated = 0
    
    'Initialize this image's compositor
    Set imgCompositor = New pdCompositor
        
End Sub

'Pass a given layer's DIB into a new DIB, but pad it against null pixels so that it is the same size as the image.
' This may cause cropping if the layer currently lies off the image - this is a known "issue", though I haven't quite made up my
' mind as to whether it's an issue that needs resolving or not.  I will revisit in the future, once I have a better idea of
' the workflow implications.
' TODO: determine if off-image layer cropping is an acceptable solution in all scenarios.
Public Sub retrieveNullPaddedLayer(ByRef dstDIB As pdDIB, ByVal srcLayerIndex As Long)

    'Create a blank destination DIB at the size of the image
    dstDIB.createBlank Width, Height, 32, 0
    
    'Copy the source layer into the blank DIB
    BitBlt dstDIB.getDIBDC, imgLayers(srcLayerIndex).getLayerOffsetX, imgLayers(srcLayerIndex).getLayerOffsetY, imgLayers(srcLayerIndex).getLayerWidth(False), imgLayers(srcLayerIndex).getLayerHeight(False), imgLayers(srcLayerIndex).layerDIB.getDIBDC, 0, 0, vbSrcCopy
    
End Sub

'Return a DIB that contains the currently selected area, fully processed according to the selection mask
Public Function retrieveProcessedSelection(ByRef dstDIB As pdDIB, Optional ByVal preMultipliedAlphaState As Boolean = False, Optional ByVal useMergedImage As Boolean = False) As Boolean

    'If this image does not contain an active selection, exit now.
    If (Not IsActive) Or (Not selectionActive) Then
        retrieveProcessedSelection = False
        Exit Function
    End If
    
    'Before doing anything else, make a temporary copy of the source data.  The source data varies depending on the useMergedImage
    ' parameter - if it is TRUE, we use a composite version of the image.  If it is FALSE, we use a copy of the current layer,
    ' but null-padded to be the same size as the image.
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    
    If useMergedImage Then
        getCompositedImage srcDIB, True
    Else
        retrieveNullPaddedLayer srcDIB, m_curLayer
    End If
    
    'Selections can be one of several types.  Right now, we don't give special handling to simple rectangular selections - all selections
    ' are fully processed according to the contents of the mask.  Also, all selections are currently created as 32bpp DIBs.
    
    'Start by initializing the destination DIB to the size of the active selection
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.createBlank mainSelection.boundWidth, mainSelection.boundHeight, 32, 0
    dstDIB.setInitialAlphaPremultiplicationState True
    
    'We now need pointers to three separate sets of image data: destination DIB, source DIB, and selection mask.
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
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
    
    Dim srcDIBCD As Long, srcQuickX As Long, srcQuickY As Long, selQuickX As Long
    srcDIBCD = srcDIB.getDIBColorDepth \ 8
    
    Dim dstDIBCD As Long, dstQuickX As Long
    dstDIBCD = dstDIB.getDIBColorDepth \ 8
            
    For x = 0 To dstDIB.getDIBWidth - 1
        dstQuickX = x * dstDIBCD
        srcQuickX = (leftOffset + x) * srcDIBCD
        selQuickX = (leftOffset + x) * 3
    For y = 0 To dstDIB.getDIBHeight - 1
    
        srcQuickY = topOffset + y
        thisAlpha = selData(selQuickX, srcQuickY)
        
        'If the selection does not exist at this pixel, we have no reason to process this pixel further!
        If thisAlpha > 0 Then
            
            'Check for the presence of alpha data in the original image.  If found, use that AND the selection mask to set the
            ' alpha value of the destination pixel.  If the source data is 24bpp, use only the selection mask data.
            If srcDIBCD = 4 Then
            
                'Check the image's alpha value.  If it's zero, we have no reason to process it further
                origAlpha = srcImageData(srcQuickX + 3, srcQuickY)
                
                If origAlpha > 0 Then
                    
                    'Source pixel data will be premultiplied, which saves us a bunch of processing time.  (That is why
                    ' we premultiply alpha, after all!)
                    r = srcImageData(srcQuickX + 2, srcQuickY)
                    g = srcImageData(srcQuickX + 1, srcQuickY)
                    b = srcImageData(srcQuickX, srcQuickY)
                    
                    'Calculate a new multiplier, based on the strength of the selection at this location
                    blendAlpha = thisAlpha / 255
                    
                    'Apply the multiplier to the existing pixel data (which is already premultiplied, saving us a bunch of time now)
                    dstImageData(dstQuickX + 2, y) = r * blendAlpha
                    dstImageData(dstQuickX + 1, y) = g * blendAlpha
                    dstImageData(dstQuickX, y) = b * blendAlpha
                    
                    'Finish our work by calculating a new alpha channel value for this pixel, which is a blend of
                    ' the original alpha value, and the selection mask value at this location.
                    dstImageData(dstQuickX + 3, y) = origAlpha * blendAlpha
                    
                End If
                
            Else
                dstImageData(dstQuickX, y) = srcImageData(srcQuickX, srcQuickY)
                dstImageData(dstQuickX + 1, y) = srcImageData(srcQuickX + 1, srcQuickY)
                dstImageData(dstQuickX + 2, y) = srcImageData(srcQuickX + 2, srcQuickY)
                dstImageData(dstQuickX + 3, y) = thisAlpha
            End If
            
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
    
    'If the calling function requested un-premultiplied alpha, apply it now
    If (Not preMultipliedAlphaState) Then dstDIB.SetAlphaPremultiplication False
        
    retrieveProcessedSelection = True

End Function

'Erase the currently selected area of a layer, with new alpha applied according to the selection mask
Public Function eraseProcessedSelection(ByVal targetLayerIndex As Long) As Boolean

    'If this image does not contain an active selection, exit now.
    If (Not IsActive) Or (Not selectionActive) Or (targetLayerIndex < 0) Or (targetLayerIndex >= m_numOfLayers) Then
        eraseProcessedSelection = False
        Exit Function
    End If
    
    'Before doing anything else, null-pad the layer we will be processing.  (At present, this function only operates on a single layer.)
    imgLayers(targetLayerIndex).convertToNullPaddedLayer Me.Width, Me.Height
    
    'Selections can be one of several types.  Right now, we don't give special handling to simple rectangular selections - all selections
    ' are fully processed according to the contents of the mask.  Also, all selections are currently created as 32bpp DIBs.
    
    'We now need pointers to two separate sets of image data: destination DIB, and selection mask.
    Dim selData() As Byte
    Dim selSA As SAFEARRAY2D
    prepSafeArray selSA, mainSelection.selMask
    CopyMemory ByVal VarPtrArray(selData()), VarPtr(selSA), 4
            
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, imgLayers(targetLayerIndex).layerDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    Dim leftOffset As Long, topOffset As Long
    leftOffset = mainSelection.boundLeft
    topOffset = mainSelection.boundTop
    
    Dim x As Long, y As Long
    Dim r As Long, g As Long, b As Long
    Dim thisAlpha As Long, origAlpha As Long
    Dim blendAlpha As Double
    
    Dim dstDIBCD As Long, dstQuickX As Long, selQuickX As Long
    dstDIBCD = imgLayers(targetLayerIndex).layerDIB.getDIBColorDepth \ 8
            
    For x = mainSelection.boundLeft To mainSelection.boundLeft + mainSelection.boundWidth - 1
        dstQuickX = x * dstDIBCD
        selQuickX = x * 3
    For y = mainSelection.boundTop To mainSelection.boundTop + mainSelection.boundHeight - 1
        
        thisAlpha = selData(selQuickX, y)
        
        'If the selection does not exist at this pixel, we have no reason to process this pixel further!
        If thisAlpha > 0 Then
        
            thisAlpha = 255 - thisAlpha
            
            'Check for the presence of alpha data in the original image.  If found, use that AND the selection mask to set the
            ' alpha value of the destination pixel.  If the source data is 24bpp, use only the selection mask data.
            If dstDIBCD = 4 Then
            
                'Check the image's alpha value.  If it's zero, we have no reason to process it further
                origAlpha = dstImageData(dstQuickX + 3, y)
                
                If origAlpha > 0 Then
                    
                    'Source pixel data will be premultiplied, which saves us a bunch of processing time.  (That is why
                    ' we premultiply alpha, after all!)
                    r = dstImageData(dstQuickX + 2, y)
                    g = dstImageData(dstQuickX + 1, y)
                    b = dstImageData(dstQuickX, y)
                    
                    'Calculate a new multiplier, based on the strength of the selection at this location
                    blendAlpha = thisAlpha / 255
                    
                    'Apply the multiplier to the existing pixel data (which is already premultiplied, saving us a bunch of time now)
                    dstImageData(dstQuickX + 2, y) = r * blendAlpha
                    dstImageData(dstQuickX + 1, y) = g * blendAlpha
                    dstImageData(dstQuickX, y) = b * blendAlpha
                    
                    'Finish our work by calculating a new alpha channel value for this pixel, which is a blend of
                    ' the original alpha value, and the selection mask value at this location.
                    dstImageData(dstQuickX + 3, y) = origAlpha * blendAlpha
                    
                End If
                
            Else
            
                'Premultiply alpha
                blendAlpha = thisAlpha / 255
                
                dstImageData(dstQuickX, y) = dstImageData(dstQuickX, y) * blendAlpha
                dstImageData(dstQuickX + 1, y) = dstImageData(dstQuickX + 1, y) * blendAlpha
                dstImageData(dstQuickX + 2, y) = dstImageData(dstQuickX + 2, y) * blendAlpha
                dstImageData(dstQuickX + 3, y) = thisAlpha
                
            End If
            
        End If
        
    Next y
    Next x
    
    'Clear all array references
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(selData), 0&, 4
    
    Erase dstImageData
    Erase selData
    
    'Shrink the layer to its maximum relevant size
    imgLayers(targetLayerIndex).cropNullPaddedLayer
    
    'Notify ourselves of the change
    pdImages(g_CurrentImage).notifyImageChanged UNDO_LAYER, targetLayerIndex
    
    eraseProcessedSelection = True

End Function

Private Sub Class_Terminate()

    'Release our image compositor
    Set imgCompositor = Nothing

    'Forcibly release all image layers
    Dim i As Long
    For i = 0 To UBound(imgLayers)
        Set imgLayers(i) = Nothing
    Next i

End Sub

'Get number of visible/hidden layers
Public Function getNumOfVisibleLayers() As Long

    Dim i As Long, visLayerCount As Long
    visLayerCount = 0
    
    For i = 0 To m_numOfLayers - 1
        If imgLayers(i).getLayerVisibility Then visLayerCount = visLayerCount + 1
    Next i
    
    getNumOfVisibleLayers = visLayerCount

End Function

Public Function getNumOfHiddenLayers() As Long

    Dim i As Long, hidLayerCount As Long
    hidLayerCount = 0
    
    For i = 0 To m_numOfLayers - 1
        If Not imgLayers(i).getLayerVisibility Then hidLayerCount = hidLayerCount + 1
    Next i
    
    getNumOfHiddenLayers = hidLayerCount

End Function

'Estimate the RAM usage for this image, based on the number of layers and the size of each layer
Public Function estimateRAMUsage() As Double

    estimateRAMUsage = 0

    Dim i As Long
    
    For i = 0 To m_numOfLayers - 1
        
        'Layers provide their own RAM estimation, based on their current cache state, layer DIB, active styles, and more
        estimateRAMUsage = estimateRAMUsage + imgLayers(i).estimateRAMUsage
        
    Next i
    
    'Assume a 10% overhead for the pdImage as a whole, to cover things like metadata, loose ICC profiles, class instances, and more
    estimateRAMUsage = CDbl(estimateRAMUsage) * 1.1

End Function

'Helper function to identify the presence of vector layers.  PD's UI synchonization service makes use of this.
Public Function getNumOfVectorLayers() As Long

    Dim i As Long, vectorLayerCount As Long
    vectorLayerCount = 0
    
    For i = 0 To m_numOfLayers - 1
        If imgLayers(i).isLayerVector Then vectorLayerCount = vectorLayerCount + 1
    Next i
    
    getNumOfVectorLayers = vectorLayerCount

End Function

'Helper function to identify the presence of raster layers.  PD's UI synchonization service makes use of this.
Public Function getNumOfRasterLayers() As Long

    Dim i As Long, rasterLayerCount As Long
    rasterLayerCount = 0
    
    For i = 0 To m_numOfLayers - 1
        If imgLayers(i).isLayerRaster Then rasterLayerCount = rasterLayerCount + 1
    Next i
    
    getNumOfRasterLayers = rasterLayerCount

End Function
