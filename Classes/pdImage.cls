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
'Copyright 2006-2017 by Tanner Helland
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

'Canonical image ID value.  This matches the image's position in the main pdImages array.
Public imageID As Long

'The main pdImages array isn't resized when images are deleted (unless *all* images are deleted).  Images that have
' been unloaded are marked as "not active" via this variable.
Private m_IsActive As Boolean

'Image dimensions
Public Width As Long
Public Height As Long

'Any random data relevant to the image can be stored in this Dictionary object.  It replaces the wide array of specific
' public variables we previously used.  Currently stored values include, but are not limited to:
' - "OriginalFileName"
' - "OriginalFileExtension"
' - "CurrentLocationOnDisk"
Public ImgStorage As pdDictionary

'Persistent buffer, for storing the fully composited image.  This composite is only updated when absolutely necessary;
' otherwise, look at grabbing the dedicated canvasBuffer, below.
Public CompositeBuffer As pdDIB

'Image's canvas buffer.  This holds a composited version of the image, extracted and cropped to reflect the current viewport.
' It is cached here to make viewport rendering faster and more reliable.
Public CanvasBuffer As pdDIB

'Scratch layer.  This layer can be used for any temporary work relevant to this image (e.g. storing paintbrush strokes).
' Note that no guarantees are made regarding the scratch layer's initialization, size, or other attributes.  You should
' always verify scratch layer attributes (or call ResetScratchLayer) prior to utilizing it.
Public ScratchLayer As pdLayer

'Image's selection data (stored and handled by a pdSelection instance)
Public MainSelection As pdSelection
Private m_SelectionIsActive As Boolean

'Metadata handler/storage class.  This class is automatically filled with an image file's metadata at load-time.
' NOTE: THE EXIFTOOL PLUGIN IS REQUIRED FOR METADATA HANDLING.
Public ImgMetadata As pdMetadata

'All Undo/Redo actions are handled through this publicly available Undo/Redo handler class
Public UndoManager As pdUndo
    
'Various viewport actions are now handled through a separate pdViewport class
Public ImgViewport As pdViewport

'Image resolution (in DPI).  In general, we don't deal with differing x/y resolutions - in that case,
' use the m_DPI value, which will return the average resolution of the two.
Private m_XResolution As Double
Private m_YResolution As Double
Private m_DPI As Double

'Current index of the associated zoom combo box; this is stored locally so zoom is preserved when switching images
Private m_ZoomIndex As Long

'Each pdImage stores two file format constants:
' 1) Original file format.  This is the format the image was in when it was loaded from file (if applicable).
'    Original file format is only set once, at load-time, and not changed during a given session.  We must retain
'    it because this value tells us how to parse things like cached metadata (which is often format-specific).
' 2) Current file format.  For example, if the user loads a .BMP file and uses "Save As" to save it as a .JPG,
'    this will state "JPG" (while original file format is still "BMP").
Private m_originalFileFormat As PD_IMAGE_FORMAT, m_currentFileFormat As PD_IMAGE_FORMAT

'Original color depth (expressed as a standard BPP value, most commonly 32)
Private m_OriginalColorDepth As Long
    
'Has this image been saved?  Access this variable via the getSaveState and setSaveState functions.
Private m_hasBeenSavedPDI As Boolean, m_hasBeenSavedFlat As Boolean

'A copy of the current image, in small and large icon form (16 and 32px at 100% system DPI)
Private m_curFormIcon16 As Long, m_curFormIcon32 As Long

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

'As of v7.0, PD's color management engine now supports a "preserve embedded profile" type of policy.  This means that
' individual images may operate in different working spaces, even during the same session.  The current working space
' of this image is stored here; you must access it via matching Get/Set functions.
'
'If the working space index is -1, the image has no working space.  Assume sRGB.
Private m_WorkingSpaceIndex As Long

'As a convenience to outside functions, the image class grabs a high-resolution timer value every time it receives an
' "image data has changed" notification.  Outside functions can track this value via the wrapper functions, and use it
' to determine if the underlying image may have changed since they last operated on it.  (PD uses this to accelerate
' certain internal rendering tasks, like regenerating "magic wand selection" image copies.)
Private m_TimeAtLastModify As Currency

Friend Function GetCurrentFileFormat() As PD_IMAGE_FORMAT
    GetCurrentFileFormat = m_currentFileFormat
End Function

Friend Sub SetCurrentFileFormat(ByVal newFormat As PD_IMAGE_FORMAT)
    m_currentFileFormat = newFormat
End Sub

Friend Function GetImageIcon(ByVal getLargeIcon As Boolean)
    If getLargeIcon Then
        GetImageIcon = m_curFormIcon32
    Else
        GetImageIcon = m_curFormIcon16
    End If
End Function

Friend Sub SetImageIcon(ByVal setLargeIcon As Boolean, ByVal newHandle As Long)
    If setLargeIcon Then
        m_curFormIcon32 = newHandle
    Else
        m_curFormIcon16 = newHandle
    End If
End Sub

'Original color depth of this image, expressed as BPP (bits per pixel).  If this image was created internally,
' the value will almost certainly be 32.  If imported from file, it will mirror the original value of the initial
' pdDIB instance it was created from.
Friend Function GetOriginalColorDepth() As Long
    GetOriginalColorDepth = m_OriginalColorDepth
End Function

Friend Sub SetOriginalColorDepth(ByVal origColorDepth As Long)
    m_OriginalColorDepth = origColorDepth
End Sub

Friend Function GetOriginalFileFormat() As PD_IMAGE_FORMAT
    GetOriginalFileFormat = m_originalFileFormat
End Function

Friend Sub SetOriginalFileFormat(ByVal newFormat As PD_IMAGE_FORMAT)
    m_originalFileFormat = newFormat
End Sub

'Zoom is stored as an index into PD's default zoom table.  To resolve these to actual floating-point modifiers,
' use a pdZoom class instance.
Friend Function GetZoom() As Long
    GetZoom = m_ZoomIndex
End Function

Friend Sub SetZoom(ByVal newZoomIndex As Long)
    m_ZoomIndex = newZoomIndex
End Sub

'Time at last modification.  This cannot be directly set by external sources; instead, use the safe "NotifyImageChanged" function.
Friend Function GetTimeOfLastChange() As Currency
    GetTimeOfLastChange = m_TimeAtLastModify
End Function

'To improve performance, PD doesn't always free image resources when an image is finished.  To see if an image is
' "active" (the PD word for "open and editable by a user"), use this function.
Friend Function IsActive() As Boolean
    IsActive = m_IsActive
End Function

Friend Sub ChangeActiveState(ByVal newActiveState As Boolean)
    m_IsActive = newActiveState
End Sub

'Create a new, blank layer in this image.  Note that layer type is not provided - that is handled separately, via the
' layer object itself.  All this function does is create a new pdLayer object and rearrange the layer stack to fit.

'Optionally, a layer position can be passed.  The position is assumed to be the position of the
' new layer; e.g. if zero is passed, the layer will be placed at the bottom of the stack.  If no position is requested,
' the layer will be created above the currently active layer.
'
'The returned Long-type value is the CANONICAL ID of the new layer, NOT ITS INDEX.  (Returning its index would be pointless,
' as the calling function likely knows that in advance.)  Note that layers can be accessed by either ID or index; after
' calling this function, make sure to use the CANONICAL ID function with the returned value.
Friend Function CreateBlankLayer(Optional ByVal layerPosition As Long = -1) As Long

    'Start by seeing if the image has zero layers.  If it does, we can ignore layerPosition entirely.
    If (m_numOfLayersEverCreated = 0) Then
        m_curLayer = 0
        
    'The image has layers.  Extra work is involved
    Else

        'If no layer has been specified, insert the new layer above the requested layer
        If (layerPosition = -1) Then m_curLayer = m_curLayer + 1 Else m_curLayer = layerPosition + 1
        
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
    imgLayers(m_curLayer).AssignLayerID m_numOfLayersEverCreated
    
    'Increment our internal layer counts
    m_numOfLayersEverCreated = m_numOfLayersEverCreated + 1
    m_numOfLayers = m_numOfLayers + 1
    
    'Return the new layer's ID value (which is the number of layers ever created for this image)
    CreateBlankLayer = m_numOfLayersEverCreated - 1

End Function

Friend Sub RebuildCompositeBuffer(Optional ByVal renderScratchLayerIndex As Long = -1)

    'If the buffer is already clean, ignore this request
    If m_CompositeCacheClean Then
        Debug.Print "WARNING!  pdImage.rebuildCompositeBuffer was called, but the compositeBuffer is marked as clean!"
    Else
        
        'Rebuild the composite cache
        imgCompositor.GetCompositedImage Me, CompositeBuffer, , renderScratchLayerIndex
        
        'Mark the cache as clean.  (NOTE: this is disabled pending further testing; I'm not sure that all changes
        ' are being tracked correctly, so composites are currently generated on every request as a workaround)
        'm_CompositeCacheClean = True
        
    End If

End Sub

'If an outside function directly modifies something inside this image (e.g. a layer), it must notify us via this function.
' The modification type mirrors PD's Undo types, to simply the work external functions have to do (as typically, they are
' notifying the central processor of these changes, so Undo data is generated correctly).
'
'If the change involves a layer, an index *must* be specified, so the object can notify the layer of the change (allowing the
' layer to rebuild any internal caches or buffers, like layer styles).
Friend Sub NotifyImageChanged(ByVal typeOfChange As PD_UNDO_TYPE, Optional ByVal layerIndex As Long = -1)
    
    'Note a new "image changed" timestamp
    VBHacks.GetHighResTime m_TimeAtLastModify
    
    'Separate handling by type.
    
    'TODO: in the future, it would be awesome to store a partial composite up-to-but-NOT-including the current layer.
    ' This would allow for faster recomposites when the current layer is modified.
    ' However, until such time as that feature is implemented, I have no choice but to recomposite the entire image
    ' after changes are made.
    
    Select Case typeOfChange
    
        Case UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE, UNDO_EVERYTHING
            m_CompositeCacheClean = False
            
            'Notify all layers of potentially destructive changes
            Dim i As Long
            For i = 0 To Me.GetNumOfLayers() - 1
                Me.GetLayerByIndex(i).NotifyOfDestructiveChanges
            Next i
        
        Case UNDO_IMAGEHEADER
            m_CompositeCacheClean = False
        
        Case UNDO_LAYER, UNDO_LAYER_VECTORSAFE
            m_CompositeCacheClean = False
            
            If (layerIndex <> -1) Then
                Me.GetLayerByIndex(layerIndex).NotifyOfDestructiveChanges
            Else
                Debug.Print "WARNING!  Invalid params passed to pdImage.notifyImageChanged (-1 layer index for a layer action)"
            End If
        
        Case UNDO_LAYERHEADER
            m_CompositeCacheClean = False
        
        Case Else
            Debug.Print "WARNING!  pdImage.notifyImageChanged received an unknown change request: " & typeOfChange
    
    End Select

End Sub

'Merge two layers together.  Note this can be used to merge any two arbitrary layers, with the bottom layer holding the result
' of the merge.  It is up to the caller to deal with any subsequent layer deletions, etc - this sub just performs the merge.
Friend Sub MergeTwoLayers(ByRef topLayer As pdLayer, ByRef bottomLayer As pdLayer, Optional ByVal bottomLayerIsFullSize As Boolean = True, Optional ByVal usePaintStyleMerge As Boolean = False, Optional ByVal ptrToPaintbrushRectF As Long = 0)
    If usePaintStyleMerge Then
        imgCompositor.MergeLayers_PaintStyle topLayer, bottomLayer, bottomLayerIsFullSize, ptrToPaintbrushRectF, Me
    Else
        imgCompositor.MergeLayers topLayer, bottomLayer, bottomLayerIsFullSize
    End If
End Sub

'Returns all layers of the image as a single, composited image (in pdDIB format, of course).  Because of the way VB handles
' object references, we ask the calling function to supply the DIB they want filled.  Optionally, they can also request a
' particular premultiplication status of the composited DIB's alpha values.  (This is helpful for save functions, which
' require non-premultiplied alpha, vs viewport functions, which require premultiplied alpha).
Friend Sub GetCompositedImage(ByRef dstDIB As pdDIB, Optional ByVal premultiplicationStatus As Boolean = True)
    
    'If the current composite cache is dirty, rebuild it
    If (Not m_CompositeCacheClean) Then Me.RebuildCompositeBuffer
    
    'Give the caller a copy of the composited image
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.CreateFromExistingDIB CompositeBuffer
    
    'Change premultiplication, if requested by the caller.  (By default, the composite image is stored in premultiplied format.)
    If (dstDIB.GetAlphaPremultiplication <> premultiplicationStatus) Then dstDIB.SetAlphaPremultiplication premultiplicationStatus
    
End Sub

'Returns a subsection of the fully composited image (in pdDIB format, of course).  This is helpful for rendering the main viewport,
' as we only composite the relevant portions of the image.
Friend Sub GetCompositedRect(ByRef dstDIB As pdDIB, ByRef dstViewportRect As RECTF, ByRef srcImageRect As RECTF, ByVal interpolationType As GP_InterpolationMode, Optional ByVal ignoreInternalCaches As Boolean = False, Optional ByVal levelOfDetail As COMPOSITOR_LOD_CLUE = CLC_Generic, Optional ByVal renderScratchLayerIndex As Long = -1)
    imgCompositor.GetCompositedRect Me, dstDIB, dstViewportRect, srcImageRect, interpolationType, ignoreInternalCaches, levelOfDetail, renderScratchLayerIndex
End Sub

'Get the currently active layer index
Friend Function GetActiveLayerIndex() As Long
    GetActiveLayerIndex = m_curLayer
End Function

'Get the currently active layer canonical ID
Friend Function GetActiveLayerID() As Long
    If (m_curLayer > UBound(imgLayers)) Then
        GetActiveLayerID = imgLayers(0).GetLayerID
    Else
        GetActiveLayerID = imgLayers(m_curLayer).GetLayerID
    End If
End Function

'Get the number of layers in this image.  (Note: this function might return zero, so handle that condition correctly
' in calling functions!)
Friend Function GetNumOfLayers() As Long
    GetNumOfLayers = m_numOfLayers
End Function

'Retrieve a reference to the currently active layer's DIB.  This is effectively a shortcut function when we
' need quick access to the current layer's DIB.
Friend Function GetActiveDIB() As pdDIB
    If (m_numOfLayers <> 0) Then
        Set GetActiveDIB = imgLayers(m_curLayer).layerDIB
    Else
        Set GetActiveDIB = Nothing
    End If
End Function

'Retrieve the currently active layer
Friend Function GetActiveLayer() As pdLayer
    If (m_numOfLayers > 0) Then
        If (m_curLayer > UBound(imgLayers)) Then m_curLayer = 0
        Set GetActiveLayer = imgLayers(m_curLayer)
    Else
        Set GetActiveLayer = Nothing
    End If
End Function

'Set the currently active layer by its cardinal ID value
Friend Sub SetActiveLayerByID(ByVal newActiveLayerID As Long)
    
    Dim i As Long
    For i = 0 To UBound(imgLayers)
        If (imgLayers(i).GetLayerID = newActiveLayerID) Then
            m_curLayer = i
            Exit For
        End If
    Next i
    
    'If we made it all the way here, the requested layer was not found.  Do nothing (e.g. do not change the active layer).
    
End Sub

Friend Sub SetActiveLayerByIndex(ByVal newActiveLayerIndex As Long)

    'Validate the incoming layer index
    If (newActiveLayerIndex < 0) Then newActiveLayerIndex = 0
    If (newActiveLayerIndex > m_numOfLayers - 1) Then newActiveLayerIndex = m_numOfLayers - 1
    
    m_curLayer = newActiveLayerIndex
    
End Sub

'Retrieve a layer index using its canonical ID.
Friend Function GetLayerIndexFromID(ByVal requestedID As Long) As Long
    
    Dim i As Long
    For i = 0 To UBound(imgLayers)
        If (imgLayers(i).GetLayerID = requestedID) Then
            GetLayerIndexFromID = i
            Exit Function
        End If
    Next i
    
    'If we made it all the way here, the requested layer was not found.  Return 0.
    GetLayerIndexFromID = 0
    
End Function

'Retrieve a layer at an arbitrary position.  Remember that layers in PD are zero-based, so the base layer is layer 0, not 1.
' Also, this function does not check bounds, so make sure the passed index is valid!
Friend Function GetLayerByIndex(ByVal layerIndex As Long) As pdLayer
    
    'As a failsafe, validate the incoming layer index
    If (layerIndex > UBound(imgLayers)) Then
        layerIndex = UBound(imgLayers)
        Debug.Print "WARNING! Invalid layerIndex requested (" & layerIndex & ":" & UBound(imgLayers) & ")"
    End If
    
    If (layerIndex < 0) Then layerIndex = 0
    
    'Return the requested layer.  Note that by design, we allow the target layer to be returned even if it
    ' does not exist.
    Set GetLayerByIndex = imgLayers(layerIndex)
    
End Function

'Retrieve a layer using its canonical ID.  Remember that layers in PD are zero-based, so the base layer is layer 0, not 1.
' Also, this function does not check bounds, so make sure the passed index is valid!
Friend Function GetLayerByID(ByVal requestedID As Long) As pdLayer
    
    Dim i As Long
    For i = 0 To m_numOfLayers - 1
        If (imgLayers(i).GetLayerID = requestedID) Then
            Set GetLayerByID = imgLayers(i)
            Exit Function
        End If
    Next i
    
    'If we made it all the way here, the requested layer was not found.  Return Nothing.
    Debug.Print "WARNING!  GetLayerByID received an invalid ID: " & requestedID
    Set GetLayerByID = Nothing
    
End Function

Friend Function PointLayerAtNewObject(ByVal layerID As Long, ByRef newLayer As pdLayer) As Boolean

    Dim targetLayerIndex As Long
    targetLayerIndex = -1
    
    Dim i As Long
    For i = 0 To m_numOfLayers - 1
        If (imgLayers(i).GetLayerID = layerID) Then
            targetLayerIndex = i
            Exit For
        End If
    Next i
    
    If (targetLayerIndex >= 0) Then
        
        PointLayerAtNewObject = True
        Set imgLayers(targetLayerIndex) = newLayer
        
        'Assign the new layer object the *old* layer's ID.
        imgLayers(targetLayerIndex).AssignLayerID layerID
        
    Else
        PointLayerAtNewObject = False
    End If
    
End Function

'Simple function for moving a given layer up/down in the layer stack.
Friend Sub MoveLayerByIndex(ByVal srcLayerIndex As Long, ByVal moveLayerUp As Boolean)

    'Before doing anything else, make sure the requested move is a valid one
    If moveLayerUp Then
        If srcLayerIndex = m_numOfLayers - 1 Then Exit Sub
    Else
        If srcLayerIndex = 0 Then Exit Sub
    End If
    
    'Process the actual move
    If moveLayerUp Then
    
        'Move the selected layer UP.
        SwapTwoLayers srcLayerIndex, srcLayerIndex + 1
    
    Else
    
        'Move the selected layer DOWN.
        SwapTwoLayers srcLayerIndex, srcLayerIndex - 1
    
    End If

End Sub

'Complicated function for moving a layer around in the layer stack.  A current layer index and destination layer index are required.
' The destination layer index will be the *new* index of the source layer.  If the two indices match, the function will terminate.
'
'The function will return TRUE if the layers were moved successfully.
Friend Function MoveLayerToArbitraryIndex(ByVal srcLayerIndex As Long, ByVal dstLayerIndex As Long) As Boolean

    'Before doing anything else, make sure the requested move is a valid one
    If srcLayerIndex = dstLayerIndex Then
        MoveLayerToArbitraryIndex = False
        Exit Function
    End If
        
    If (srcLayerIndex < 0) Or (srcLayerIndex > m_numOfLayers - 1) Then
        MoveLayerToArbitraryIndex = False
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
            SwapTwoLayers i, i + 1
        Next i
    
    Else
    
        'The source layer is moving LOWER in the stack
    
        'Iterate through all layers between the current position and the new one, swapping layers as we go
        For i = srcLayerIndex To dstLayerIndex + 1 Step -1
            SwapTwoLayers i, i - 1
        Next i
    
    End If
    
    MoveLayerToArbitraryIndex = True

End Function

'Used by various layer movement functions.  Given two layer indices, swap them.
Friend Sub SwapTwoLayers(ByVal srcLayerIndex_1 As Long, ByVal srcLayerIndex_2 As Long)

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
Friend Sub DeleteLayerByIndex(ByVal srcLayerIndex As Long)

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

'Get/set active selection state
Friend Function IsSelectionActive() As Boolean
    IsSelectionActive = m_SelectionIsActive
End Function

Friend Sub SetSelectionActive(ByVal newState As Boolean)
    m_SelectionIsActive = newState
End Sub

'Write all relevant image information to an external file (in XML format).  Note that this will not be a perfect copy of the
' current image's state, because some pdImage data is only relevant at run-time.  Also, some data structures (like metadata) are
' very difficult to handle properly; I am working on a solution to this, but have not finished it yet.
'
'In the event that you want to do something else with the XML data, I have now added an optional parameter.
' Mark "filePathIsActuallyString" as TRUE to treat the filename string as a destination buffer, and this function will
' write the XML data there (instead of treating it as a filename).
Friend Function WriteExternalData(ByRef dstFilename As String, Optional ByVal filePathIsActuallyString As Boolean = False) As Boolean

    On Error GoTo failedDataWrite

    'Prepare an XML engine, which will handle the actual writing of the file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Add a basic header and explanatory comment
    xmlEngine.PrepareNewXML "pdImage"
    
    xmlEngine.WriteBlankLine
    xmlEngine.WriteComment "This file contains a summary of crucial data for a given pdImage entry."
    xmlEngine.WriteBlankLine
    
    'Start by writing out the imageID.  This value will not be re-used in subsequent session, but it's helpful for matching up
    ' this image's data with any saved Undo/Redo entries (which are created by imageID for privacy reasons)
    xmlEngine.WriteTag "ID", imageID
    
    'Start writing any tags that may be useful for reconstructing an autosave version of this image.  We limit the information to
    ' properties independent of a given session (e.g. imageID); session-dependent properties will be recreated anew by PD.
    xmlEngine.WriteTag "Width", Width
    xmlEngine.WriteTag "Height", Height
    xmlEngine.WriteTag "XResolution", m_XResolution
    xmlEngine.WriteTag "YResolution", m_YResolution
    xmlEngine.WriteTag "DPI", m_DPI
    xmlEngine.WriteTag "OriginalFileFormat", m_originalFileFormat
    xmlEngine.WriteTag "CurrentFileFormat", m_currentFileFormat
    xmlEngine.WriteTag "OriginalColorDepth", m_OriginalColorDepth
    xmlEngine.WriteTag "numOfLayersEverCreated", m_numOfLayersEverCreated
    xmlEngine.WriteTag "numOfLayers", m_numOfLayers
    xmlEngine.WriteTag "curLayer", m_curLayer
    
    'Next comes conditional data, which may not exist in the scripting dictionary.  If it's available, we'll write it out.
    If ImgStorage.DoesKeyExist("pngBackgroundColor") Then xmlEngine.WriteTag "pngBackgroundColor", ImgStorage.GetEntry_Long("pngBackgroundColor")
    If ImgStorage.DoesKeyExist("OriginalFileSize") Then xmlEngine.WriteTag "OriginalFileSize", ImgStorage.GetEntry_Long("OriginalFileSize")
    
    xmlEngine.WriteBlankLine
    
    'That should be enough information to reconstruct the original image state.  We now have two options, depending on the value
    ' of the "filePathIsActuallyString" parameter.
    
    'The string we were supplied is a blank buffer, not a filename.  Write the XML data to it.  Note that by default, XML indentation
    ' is suppressed; this is for performance reasons.
    If filePathIsActuallyString Then
        dstFilename = xmlEngine.ReturnCurrentXMLString(True)
    
    'The string we were supplied is a filename.  Write the XML data to it.
    Else
        xmlEngine.WriteXMLToFile dstFilename
    End If
    
    WriteExternalData = True
    Exit Function
    
failedDataWrite:

    Debug.Print "WARNING: could not write XML data to file.  Are you sure this drive has enough free space?"
    WriteExternalData = False
    
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
Friend Function ReadExternalData(ByRef srcFilename As String, Optional ByVal filePathIsActuallyString As Boolean = False, Optional ByVal sourceIsUndoFile As Boolean = False, Optional ByVal loadNonDestructively As Boolean = False) As Boolean

    On Error GoTo failedDataRead
    
    Dim xmlLoadSuccessful As Boolean
    xmlLoadSuccessful = False

    'Prepare an XML engine, which will handle the actual reading and parsing of the file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'We now have two choices, depending on the value of filePathIsActuallyString.  Load the supposed XML contents
    ' (either from file or directly) into the XML engine; note that both steps validate the incoming data automatically.
    If filePathIsActuallyString Then
        xmlLoadSuccessful = xmlEngine.LoadXMLFromString(srcFilename)
    Else
        xmlLoadSuccessful = xmlEngine.LoadXMLFile(srcFilename)
    End If
    
    'Make sure the supplied XML actually contains pdImage data.
    If xmlLoadSuccessful And xmlEngine.IsPDDataType("pdImage") Then
    
        'The XML file checked out.  Start retrieving relevant values.
        Me.Width = xmlEngine.GetUniqueTag_Long("Width")
        Me.Height = xmlEngine.GetUniqueTag_Long("Height")
        m_XResolution = xmlEngine.GetUniqueTag_Long("XResolution")
        m_YResolution = xmlEngine.GetUniqueTag_Long("YResolution")
        m_DPI = xmlEngine.GetUniqueTag_Long("DPI")
        
        'All settings past this point are only relevant when loading a file anew, not when it's being used as part of
        ' the Undo/Redo stack.
        If (Not sourceIsUndoFile) Then
        
            'These values are all stored in the PDI file, but they're not really relevant.  I will look at not storing
            ' them in the first place, but for now, just ignore them.
            'TODO: make a firm decision on what data does not belong in PDI files!
            Me.SetOriginalColorDepth xmlEngine.GetUniqueTag_Long("originalColorDepth")
            
            'Dictionary tags are conditional, and may not exist in the PDI file.  Check for their existence, and load them
            ' only as necessary.
            If xmlEngine.DoesTagExist("pngBackgroundColor") Then ImgStorage.AddEntry "pngBackgroundColor", xmlEngine.GetUniqueTag_Long("pngBackgroundColor")
            
        End If
        
        'Finally, retrieve layer values from the file
        m_numOfLayersEverCreated = xmlEngine.GetUniqueTag_Long("numOfLayersEverCreated")
        m_numOfLayers = xmlEngine.GetUniqueTag_Long("numOfLayers")
        m_curLayer = xmlEngine.GetUniqueTag_Long("curLayer")
        
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
    ReadExternalData = True
    
    Exit Function
    
failedDataRead:
    
    Debug.Print "WARNING: could not validate XML data.  pdImage internals were not updated."
    ReadExternalData = False
    
End Function

'Get/Set image resolution (in DPI).  Note that the vertical resolution is optional; if the two values
' differ, PD will average them when image DPI is requested.
Friend Function GetDPI() As Double
    
    If m_DPI = 0 Then
        GetDPI = 96
    Else
        GetDPI = m_DPI
    End If
    
End Function

Friend Sub SetDPI(ByVal xRes As Double, ByVal yRes As Double, Optional ByVal sourceIsFreeImage As Boolean = False)
    
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
    
    m_XResolution = xRes
    m_YResolution = yRes
    
    'It is extremely rare for x/y resolution to differ, but just in case, calculate an average resolution as well
    m_DPI = (xRes + yRes) \ 2

End Sub

'If the image has been saved to file in its current state, this will return TRUE.  Use this value to determine
' whether to enable a Save button, for example.
Friend Function GetSaveState(ByVal desiredSaveType As PD_SAVE_EVENT) As Boolean
    
    Select Case desiredSaveType
    
        Case pdSE_AnySave
            GetSaveState = m_hasBeenSavedPDI Or m_hasBeenSavedFlat
        
        Case pdSE_SavePDI
            GetSaveState = m_hasBeenSavedPDI
        
        Case pdSE_SaveFlat
            GetSaveState = m_hasBeenSavedFlat
    
    End Select
    
End Function

'Outside actions (such as saving) can affect the m_hasBeenSaved variable.  However, because we need to do additional
' processing based on the state of this variable, we provide this interface.
Friend Sub SetSaveState(ByVal newSaveState As Boolean, ByVal typeOfSaveEvent As PD_SAVE_EVENT)
    
    'The image has just been saved
    If newSaveState Then
        
        Select Case typeOfSaveEvent
            
            'This value should never be passed in!
            Case pdSE_AnySave
                Debug.Print "The pdSE_AnySave type is only for GETTING image save state, not SETTING it!  Fix this!"
                
            'Image has been saved to PDI format, meaning layers are intact
            Case pdSE_SavePDI
                m_hasBeenSavedPDI = True
                
            'Image has been saved to some flat format (JPEG, PNG, etc)
            Case pdSE_SaveFlat
                m_hasBeenSavedFlat = True
            
        End Select
        
        'Remember the undo value at this juncture; if the user performs additional actions, but "Undos" to this point,
        ' we want to disable the save button for them
        UndoManager.NotifyImageSaved typeOfSaveEvent
    
    'Some change has occurred, meaning the image has not been saved (to any format) in its current state
    Else
        m_hasBeenSavedPDI = False
        m_hasBeenSavedFlat = False
    End If
        
End Sub

'Working space index related to color management.  It is the index of this image's current working spcae profile, as it is
' found in the ColorManagement module's ICC profile cache.  (Accessing by index is way faster than doing large byte array
' comparisons, which is why we do it this way.)
Friend Function GetWorkingSpaceIndex() As Long
    GetWorkingSpaceIndex = m_WorkingSpaceIndex
End Function

Friend Sub SetWorkingSpaceIndex(ByVal newIndex As Long)
    m_WorkingSpaceIndex = newIndex
End Sub

'External functions can use this function to request a thumbnail version of the contained image.  Note that the full
' image must be re-composited on every request, so please do not call this function unless absolutely necessary!
Friend Function RequestThumbnail(ByRef dstThumbnailDIB As pdDIB, Optional ByVal thumbnailSize As Long = 64) As Boolean
    
    'Uncomment the two lines below (and another at the end of the function) to profile different thumbnail resize strategies
    'Dim startTime As Double
    'startTime = Timer
    
    'Thumbnails have some interesting requirements.  We always want them to be square, with the image set in the middle
    ' of the thumbnail (with aspect ratio preserved) and any empty edges made transparent.
    
    'We also need to determine the thumbnail's actual width and height, and any x and y offset necessary to preserve the
    ' aspect ratio and center the image on the thumbnail.
    Dim tIcoWidth As Long, tIcoHeight As Long, tX As Single, tY As Single
    
    'Start by determining proper dimensions for the resized thumbnail image.
    PDMath.ConvertAspectRatio Me.Width, Me.Height, thumbnailSize, thumbnailSize, tIcoWidth, tIcoHeight
    
    'If the form is wider than it is tall, center the thumbnail vertically
    If (tIcoWidth > tIcoHeight) Then
        tX = 0
        tY = (thumbnailSize - tIcoHeight) * 0.5
    
    '...otherwise, center it horizontally
    Else
        tY = 0
        tX = (thumbnailSize - tIcoWidth) * 0.5
    End If
    
    'Prepare the destination DIB
    If (dstThumbnailDIB Is Nothing) Then Set dstThumbnailDIB = New pdDIB
    If (dstThumbnailDIB.GetDIBWidth <> thumbnailSize) Or (dstThumbnailDIB.GetDIBHeight <> thumbnailSize) Then
        dstThumbnailDIB.CreateBlank thumbnailSize, thumbnailSize, 32, 0
    Else
        dstThumbnailDIB.ResetDIB 0
    End If
    
    'Retrieve a composited thumbnail.  (Note that the user's thumbnail performance setting affects the interpolation
    ' method used.)
    Dim dstRectF As RECTF, srcRectF As RECTF
    With dstRectF
        .Left = tX
        .Top = tY
        .Width = tIcoWidth
        .Height = tIcoHeight
    End With
    
    With srcRectF
        .Left = 0#
        .Top = 0#
        .Width = Me.Width
        .Height = Me.Height
    End With
    
    Me.GetCompositedRect dstThumbnailDIB, dstRectF, srcRectF, g_UserPreferences.GetThumbnailInterpolationPref(), , CLC_Thumbnail
    
    'Uncomment the line below to profile different thumbnail resize strategies
    'Debug.Print "Time taken for image thumbnail generation: " & CStr((Timer - startTime) * 1000) & " ms"
    
    RequestThumbnail = True
    
End Function

'When the attached image object is deactivated (e.g. the user leaves it loaded, but switches to a
' different image), you can call this function to free up some internal memory caches and other objects.
' This can free a non-trivial amount of resources when system memory is tight.
Friend Sub DeactivateImage()

    'Scratch layers aren't needed unless the image is actively being edited
    Set ScratchLayer = Nothing
    
    'Some internal selection caches are free-able
    If (Not MainSelection Is Nothing) Then MainSelection.FreeNonEssentialResources

End Sub

'When the attached image object is closed, we can deactivate a ton of items to save on resources.
' (At present, note that this class itself is *not* freed; we do this to save time when reloading
'  subsequent images - a very old design decision that should probably be revisited.)
Friend Sub FreeAllImageResources()
    
    FreeInternalIcons
    
    'Erase any internal image buffers
    If (Not CompositeBuffer Is Nothing) Then Set CompositeBuffer = Nothing
    If (Not CanvasBuffer Is Nothing) Then Set CanvasBuffer = Nothing
    
    'Erase the selection manager (*very* important because of circular references)
    If (Not MainSelection Is Nothing) Then
        MainSelection.SetParentReference Nothing
        Set MainSelection = Nothing
    End If
        
    'Deactivate the Undo/Redo handler (*very* important because of circular references)
    If (Not UndoManager Is Nothing) Then
        UndoManager.ClearUndos
        Set UndoManager.parentPDImage = Nothing
        Set UndoManager = Nothing
    End If
    
    'Release the viewport renderer
    If (Not ImgViewport Is Nothing) Then Set ImgViewport = Nothing
    
    'Release the compositor
    If (Not imgCompositor Is Nothing) Then Set imgCompositor = Nothing
    
    'Release all layers
    Dim i As Long
    If (m_numOfLayers > 0) Then
        For i = 0 To UBound(imgLayers)
            Set imgLayers(i) = Nothing
        Next i
        m_numOfLayers = 0
    End If
    
    'It doesn't make a big difference, but we can also empty out this image's String-type variables to save a bit of space.
    Set ImgStorage = Nothing
    
    'Mark this image as inactive
    Me.ChangeActiveState False
    
End Sub

Private Sub FreeInternalIcons()
    If (m_curFormIcon16 <> 0) Then
        IconsAndCursors.ReleaseIcon m_curFormIcon16
        m_curFormIcon16 = 0
    End If
    If (m_curFormIcon32 <> 0) Then
        IconsAndCursors.ReleaseIcon m_curFormIcon32
        m_curFormIcon32 = 0
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
Friend Function UpdateSize(Optional ByVal useBaseLayer As Boolean = True, Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0)
    
    If useBaseLayer Then
        Me.Width = imgLayers(0).GetLayerWidth(False)
        Me.Height = imgLayers(0).GetLayerHeight(False)
    Else
        Me.Width = newWidth
        Me.Height = newHeight
    End If
    
    'Changes to image size require a rebuild of our composited image cache
    Me.NotifyImageChanged UNDO_IMAGE
    
End Function

'INITIALIZE class
Private Sub Class_Initialize()

    'Initially, mark the image as *not* having been saved
    m_IsActive = False
    m_hasBeenSavedPDI = False
    m_hasBeenSavedFlat = False

    'Initialize the public storage dictionary
    Set ImgStorage = New pdDictionary

    'Add initial key/value pairs to the dictionary.  Doing this here saves us from having to check key existence in subsequent functions.

        'Lossless filetypes always present an export dialog at least once.  Beyond that point, any use of the Save command will
        ' re-save the file using previous settings.  Save As will still trigger a new dialog, however.
        ImgStorage.AddEntry "hasSeenJPEGPrompt", False
        ImgStorage.AddEntry "hasSeenJP2Prompt", False
        ImgStorage.AddEntry "hasSeenWebPPrompt", False
        ImgStorage.AddEntry "hasSeenJXRPrompt", False

    'Initialize the composite and canvas buffers
    Set CompositeBuffer = New pdDIB
    Set CanvasBuffer = New pdDIB

    'Initialize the main selection
    Set MainSelection = New pdSelection
    m_SelectionIsActive = False
    MainSelection.SetSelectionShape ss_Rectangle
    MainSelection.SetParentReference Me

    'Initialize the metadata object (which may not get used, but this prevents errors if other functions try to access metadata)
    Set ImgMetadata = New pdMetadata

    'Initialize the Undo/Redo handler
    Set UndoManager = New pdUndo
    Set UndoManager.parentPDImage = Me

    'Initialize the viewport manager
    Set ImgViewport = New pdViewport

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
    
    'Initialize any color management trackers
    m_WorkingSpaceIndex = -1
    
End Sub

Private Sub Class_Terminate()
    
    'Manually release any remaining classes and/or resources
    FreeInternalIcons
    
    Set imgCompositor = Nothing
    Set ImgViewport = Nothing
    If (Not UndoManager Is Nothing) Then Set UndoManager.parentPDImage = Nothing
    Set UndoManager = Nothing
    Set ImgMetadata = Nothing
    
    'Note: if the main selection object has an active reference to us, it's circular, so this _Terminate() event won't be
    ' called at all!
    If (Not MainSelection Is Nothing) Then MainSelection.SetParentReference Nothing
    Set MainSelection = Nothing
    
    Set CanvasBuffer = Nothing
    Set CompositeBuffer = Nothing
    Set ImgStorage = Nothing
    
    'Forcibly release all image layers
    Dim i As Long
    If VBHacks.IsArrayInitialized(imgLayers) Then
        For i = 0 To UBound(imgLayers)
            Set imgLayers(i) = Nothing
        Next i
        Erase imgLayers
    End If
    
End Sub

'Pass a given layer's DIB into a new DIB, but pad it against null pixels so that it is the same size as the image.
' This may cause cropping if the layer currently lies off the image - this is a known "issue", though I haven't quite made up my
' mind as to whether it's an issue that needs resolving or not.  I will revisit in the future, once I have a better idea of
' the workflow implications.
' TODO: determine if off-image layer cropping is an acceptable solution in all scenarios.
Friend Sub RetrieveNullPaddedLayer(ByRef dstDIB As pdDIB, ByVal srcLayerIndex As Long)

    'Create a blank destination DIB at the size of the image
    dstDIB.CreateBlank Width, Height, 32, 0
    
    'Copy the source layer into the blank DIB
    BitBlt dstDIB.GetDIBDC, imgLayers(srcLayerIndex).GetLayerOffsetX, imgLayers(srcLayerIndex).GetLayerOffsetY, imgLayers(srcLayerIndex).GetLayerWidth(False), imgLayers(srcLayerIndex).GetLayerHeight(False), imgLayers(srcLayerIndex).layerDIB.GetDIBDC, 0, 0, vbSrcCopy
    
End Sub

'Return a DIB that contains the currently selected area, fully processed according to the selection mask
Friend Function RetrieveProcessedSelection(ByRef dstDIB As pdDIB, Optional ByVal preMultipliedAlphaState As Boolean = False, Optional ByVal useMergedImage As Boolean = False) As Boolean

    'If this image does not contain an active selection, exit now.
    If (Not Me.IsActive) Or (Not m_SelectionIsActive) Then
        RetrieveProcessedSelection = False
        Exit Function
    End If
    
    'Before doing anything else, make a temporary copy of the source data.  The source data varies depending on the useMergedImage
    ' parameter - if it is TRUE, we use a composite version of the image.  If it is FALSE, we use a copy of the current layer,
    ' but null-padded to be the same size as the image.
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    
    If useMergedImage Then
        GetCompositedImage srcDIB, True
    Else
        RetrieveNullPaddedLayer srcDIB, m_curLayer
    End If
    
    'Selections can be one of several types.  Right now, we don't give special handling to simple rectangular selections - all selections
    ' are fully processed according to the contents of the mask.  Also, all selections are currently created as 32bpp DIBs.
    
    'Start by initializing the destination DIB to the size of the active selection
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    Dim maskBounds As RECTF
    maskBounds = MainSelection.GetBoundaryRect
    dstDIB.CreateBlank maskBounds.Width, maskBounds.Height, 32, 0
    dstDIB.SetInitialAlphaPremultiplicationState True
    
    'We now need pointers to three separate sets of image data: destination DIB, source DIB, and selection mask.
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    Dim selData() As Byte
    Dim selSA As SAFEARRAY2D
    PrepSafeArray selSA, MainSelection.GetMaskDIB
    CopyMemory ByVal VarPtrArray(selData()), VarPtr(selSA), 4
            
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    Dim leftOffset As Long, topOffset As Long
    leftOffset = maskBounds.Left
    topOffset = maskBounds.Top
    
    Dim x As Long, y As Long
    Dim r As Long, g As Long, b As Long
    Dim thisAlpha As Long, origAlpha As Long
    Dim blendAlpha As Double
    
    Dim srcDIBCD As Long, srcQuickX As Long, srcQuickY As Long, selQuickX As Long
    srcDIBCD = srcDIB.GetDIBColorDepth \ 8
    
    Dim dstDIBCD As Long, dstQuickX As Long
    dstDIBCD = dstDIB.GetDIBColorDepth \ 8
    
    Dim selDIBCD As Long
    selDIBCD = MainSelection.GetMaskDIB.GetDIBColorDepth \ 8
            
    For x = 0 To dstDIB.GetDIBWidth - 1
        dstQuickX = x * dstDIBCD
        srcQuickX = (leftOffset + x) * srcDIBCD
        selQuickX = (leftOffset + x) * selDIBCD
    For y = 0 To dstDIB.GetDIBHeight - 1
    
        srcQuickY = topOffset + y
        thisAlpha = selData(selQuickX, srcQuickY)
        
        'If the selection does not exist at this pixel, we have no reason to process this pixel further!
        If (thisAlpha > 0) Then
            
            'Check for the presence of alpha data in the original image.  If found, use that AND the selection mask to set the
            ' alpha value of the destination pixel.  If the source data is 24bpp, use only the selection mask data.
            If (srcDIBCD = 4) Then
            
                'Check the image's alpha value.  If it's zero, we have no reason to process it further
                origAlpha = srcImageData(srcQuickX + 3, srcQuickY)
                
                If (origAlpha > 0) Then
                    
                    'Source pixel data will be premultiplied, which saves us a bunch of processing time.  (That is why
                    ' we premultiply alpha, after all!)
                    b = srcImageData(srcQuickX, srcQuickY)
                    g = srcImageData(srcQuickX + 1, srcQuickY)
                    r = srcImageData(srcQuickX + 2, srcQuickY)
                    
                    'Calculate a new multiplier, based on the strength of the selection at this location
                    blendAlpha = thisAlpha / 255
                    
                    'Apply the multiplier to the existing pixel data (which is already premultiplied, saving us a bunch of time now)
                    dstImageData(dstQuickX, y) = b * blendAlpha
                    dstImageData(dstQuickX + 1, y) = g * blendAlpha
                    dstImageData(dstQuickX + 2, y) = r * blendAlpha
                    
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
    
    'If the calling function requested un-premultiplied alpha, apply it now
    If (Not preMultipliedAlphaState) Then dstDIB.SetAlphaPremultiplication False
        
    RetrieveProcessedSelection = True

End Function

'Erase the currently selected area of a layer, with new alpha applied according to the selection mask
Friend Function EraseProcessedSelection(ByVal targetLayerIndex As Long) As Boolean

    'If this image does not contain an active selection, exit now.
    If (Not Me.IsActive) Or (Not m_SelectionIsActive) Or (targetLayerIndex < 0) Or (targetLayerIndex >= m_numOfLayers) Then
        EraseProcessedSelection = False
        Exit Function
    End If
    
    'Before doing anything else, null-pad the layer we will be processing.  (At present, this function only operates on a single layer.)
    imgLayers(targetLayerIndex).ConvertToNullPaddedLayer Me.Width, Me.Height
    
    'Selections can be one of several types.  Right now, we don't give special handling to simple rectangular selections - all selections
    ' are fully processed according to the contents of the mask.  Also, all selections are currently created as 32bpp DIBs.
    
    'We now need pointers to two separate sets of image data: destination DIB, and selection mask.
    Dim selData() As Byte
    Dim selSA As SAFEARRAY2D
    PrepSafeArray selSA, MainSelection.GetMaskDIB
    CopyMemory ByVal VarPtrArray(selData()), VarPtr(selSA), 4
            
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    PrepSafeArray dstSA, imgLayers(targetLayerIndex).layerDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    Dim maskBounds As RECTF
    maskBounds = MainSelection.GetBoundaryRect
    
    Dim leftBound As Long, topBound As Long, rightBound As Long, bottomBound As Long
    leftBound = maskBounds.Left
    topBound = maskBounds.Top
    rightBound = leftBound + maskBounds.Width - 1
    bottomBound = topBound + maskBounds.Height - 1
    
    Dim x As Long, y As Long
    Dim r As Long, g As Long, b As Long
    Dim thisAlpha As Long, origAlpha As Long
    Dim blendAlpha As Double
    
    Dim dstDIBCD As Long, selDIBCD As Long, dstQuickX As Long, selQuickX As Long
    dstDIBCD = imgLayers(targetLayerIndex).layerDIB.GetDIBColorDepth \ 8
    selDIBCD = MainSelection.GetMaskDIB.GetDIBColorDepth \ 8
    
    Const BYTE_TO_FLOAT As Single = 1# / 255#
    
    For x = leftBound To rightBound
        dstQuickX = x * dstDIBCD
        selQuickX = x * selDIBCD
    For y = topBound To bottomBound
        
        thisAlpha = selData(selQuickX, y)
        
        'If the selection does not exist at this pixel, we have no reason to process this pixel further!
        If (thisAlpha > 0) Then
        
            thisAlpha = 255 - thisAlpha
            
            'Check for the presence of alpha data in the original image.  If found, use that AND the selection mask to set the
            ' alpha value of the destination pixel.  If the source data is 24bpp, use only the selection mask data.
            If (dstDIBCD = 4) Then
            
                'Check the image's alpha value.  If it's zero, we have no reason to process it further
                origAlpha = dstImageData(dstQuickX + 3, y)
                
                If (origAlpha > 0) Then
                    
                    'Source pixel data will be premultiplied, which saves us a bunch of processing time.  (That is why
                    ' we premultiply alpha, after all!)
                    b = dstImageData(dstQuickX, y)
                    g = dstImageData(dstQuickX + 1, y)
                    r = dstImageData(dstQuickX + 2, y)
                    
                    'Calculate a new multiplier, based on the strength of the selection at this location
                    blendAlpha = thisAlpha * BYTE_TO_FLOAT
                    
                    'Apply the multiplier to the existing pixel data (which is already premultiplied, saving us a bunch of time now)
                    dstImageData(dstQuickX, y) = b * blendAlpha
                    dstImageData(dstQuickX + 1, y) = g * blendAlpha
                    dstImageData(dstQuickX + 2, y) = r * blendAlpha
                    
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
    
    'Shrink the layer to its maximum relevant size
    imgLayers(targetLayerIndex).CropNullPaddedLayer
    
    'Notify ourselves of the change
    pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, targetLayerIndex
    
    EraseProcessedSelection = True

End Function

'Get number of visible/hidden layers
Friend Function GetNumOfVisibleLayers() As Long

    Dim i As Long, visLayerCount As Long
    visLayerCount = 0
    
    For i = 0 To m_numOfLayers - 1
        If imgLayers(i).GetLayerVisibility Then visLayerCount = visLayerCount + 1
    Next i
    
    GetNumOfVisibleLayers = visLayerCount

End Function

Friend Function GetNumOfHiddenLayers() As Long

    Dim i As Long, hidLayerCount As Long
    hidLayerCount = 0
    
    For i = 0 To m_numOfLayers - 1
        If Not imgLayers(i).GetLayerVisibility Then hidLayerCount = hidLayerCount + 1
    Next i
    
    GetNumOfHiddenLayers = hidLayerCount

End Function

'Estimate the RAM usage for this image, based on the number of layers and the size of each layer
Friend Function EstimateRAMUsage() As Double

    EstimateRAMUsage = 0

    Dim i As Long
    
    For i = 0 To m_numOfLayers - 1
        
        'Layers provide their own RAM estimation, based on their current cache state, layer DIB, active styles, and more
        EstimateRAMUsage = EstimateRAMUsage + imgLayers(i).EstimateRAMUsage
        
    Next i
    
    'Assume a 10% overhead for the pdImage as a whole, to cover things like metadata, loose ICC profiles, class instances, and more
    EstimateRAMUsage = CDbl(EstimateRAMUsage) * 1.1

End Function

'Helper function to identify the presence of vector layers.  PD's UI synchonization service makes use of this.
Friend Function GetNumOfVectorLayers() As Long

    Dim i As Long, vectorLayerCount As Long
    vectorLayerCount = 0
    
    For i = 0 To m_numOfLayers - 1
        If imgLayers(i).IsLayerVector Then vectorLayerCount = vectorLayerCount + 1
    Next i
    
    GetNumOfVectorLayers = vectorLayerCount

End Function

'Helper function to identify the presence of raster layers.  PD's UI synchonization service makes use of this.
Friend Function GetNumOfRasterLayers() As Long

    Dim i As Long, rasterLayerCount As Long
    rasterLayerCount = 0
    
    For i = 0 To m_numOfLayers - 1
        If imgLayers(i).IsLayerRaster Then rasterLayerCount = rasterLayerCount + 1
    Next i
    
    GetNumOfRasterLayers = rasterLayerCount

End Function

'Want to use this image's scratch layer for something?  Call this function to validate the layer's current attributes.
Friend Sub ResetScratchLayer(Optional ByVal resetLayerSettingsToo As Boolean = False)

    'In the future, it may be helpful to know when the scratch layer is first being created
    Dim firstInitialization As Boolean
    If (ScratchLayer Is Nothing) Then
        Set ScratchLayer = New pdLayer
        firstInitialization = True
    End If
    
    'By design, the scratch layer should always match the current image size.
    If (ScratchLayer.layerDIB.GetDIBWidth <> Me.Width) Or (ScratchLayer.layerDIB.GetDIBHeight <> Me.Height) Then
        ScratchLayer.layerDIB.CreateBlank Me.Width, Me.Height, 32, 0, 0
        ScratchLayer.layerDIB.SetInitialAlphaPremultiplicationState True
    Else
        ScratchLayer.layerDIB.ResetDIB 0
    End If
    
    'No other scratch layer attributes are set by this function, unless specifically requested
    If resetLayerSettingsToo Then ScratchLayer.ResetLayerParameters

End Sub

'Want to suspend (or subsequently re-enable) animations?  Use this sub.  (At present, this only affects the
' "marching ant" selection animation.)
Friend Sub NotifyAnimationsAllowed(ByVal allowedState As Boolean)
    If (Not MainSelection Is Nothing) Then MainSelection.NotifyAnimationsAllowed allowedState
End Sub

