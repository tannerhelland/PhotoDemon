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
'Copyright 2006-2026 by Tanner Helland
'Created: sometime 2006
'Last updated: 10/March/25
'Last update: expose "suspend to disk" behavior to callers
'
'The pdImage class is used to store information on each image loaded by the user.
' One instance of this class exists for each loaded image.
'
'This class doesn't do much its own processing.  At present, it's mainly just a container for
' key image properties like name, file path, format, Undo/Redo tracking - as well as the
' all-important array of pdLayer objects that comprise an image's actual bitmap data.
'
'Generally speaking, this class should rarely be instantiated directly.  If you need a
' temporary image, use pdDIB, or if layer interactions are required, a single pdLayer instance
' (and its corresponding layerDIB).  This class is quite heavy as it must track things like
' metadata, undo/redo handling, and other items, so avoid using it for things like temporary
' image operations.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************


Option Explicit


'Each active image in the program exists within one of these pdImage objects.  This class stores all relevant information
' about the image, including references to its layers, metadata, Undo/Redo, and much more.

'Canonical image ID value.  (This currently matches the image's index in the main pdImages array, but this
' behavior is *not* guaranteed forever.)
Public imageID As Long

'The main pdImages array isn't resized when images are deleted (unless *all* images are deleted).  Images that have
' been unloaded are marked as "not active" via this variable.
Private m_IsActive As Boolean

'Image dimensions
Public Width As Long, Height As Long

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

'As of v8.0, pdImage objects now associate with two different color profiles.
' 1) an "original" profile; this points to image's ICC profile, if any, when it was first loaded/created.
' 2) a "working space" profile; this points to the working space that all layers *must* be converted to
'    prior to merging / blending / compositing / etc.
'Why two profiles?  Because the original profile can always be exported via File > Export, even if the
' current sessions's color management settings result in things like hard conversions to sRGB.
Private m_colorProfileOriginal As String, m_colorProfileWorkingSpace As String

'Image resolution (in DPI).  In general, we don't deal with differing x/y resolutions - in that case,
' use the m_DPI value, which will return the average resolution of the two.
Private m_XResolution As Double, m_YResolution As Double
Private m_DPI As Double

'If the current image is intended to be animated (e.g. an animated GIF or PNG), set this flag to TRUE.
' PD uses this to activate contextual "play" and "pause" buttons in the navigator.
Private m_Animated As Boolean

'Each pdImage stores two file format constants:
' 1) Original file format.  This is the format the image was in when it was loaded from file (if applicable).
'    Original file format is only set once, at load-time, and not changed during a given session.  We must retain
'    it because this value tells us how to parse things like cached metadata (which is often format-specific).
' 2) Current file format.  For example, if the user loads a .BMP file and uses "Save As" to save it as a .JPG,
'    this will state "JPG" (while original file format is still "BMP").
Private m_originalFileFormat As PD_IMAGE_FORMAT, m_currentFileFormat As PD_IMAGE_FORMAT

'Original color depth (expressed as a standard BPP value, most commonly 32), whether the original image was
' saved in a grayscale format, and whether the original image contained alpha data.
Private m_originalColorDepth As Long, m_originalGrayscale As Boolean, m_originalAlpha As Boolean

'Original palette, if any.  As you can imagine, this is only relevant if m_OriginalColorDepth is <= 8,
' so you must *always* check for a null object, as many pdImage instances won't initialize this object.
Private m_origPalette As pdPalette

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

'As a convenience to outside functions, the image class grabs a high-resolution timer value every time it receives an
' "image data has changed" notification.  Outside functions can track this value via the wrapper functions, and use it
' to determine if the underlying image may have changed since they last operated on it.  (PD uses this to accelerate
' certain internal rendering tasks, like regenerating "magic wand selection" image copies.)
Private m_timeAtLastModify As Currency

'If the current image thumbnail is clean or dirty.  Clean means that to the best of this class's knowledge, the current
' thumbnail accurately reflects the composited image's contents.  Dirty means that we know the current thumbnail is
' invalid, so it must be regenerated before continuing.
Private m_IsThumbClean As Boolean

'Composited image thumbnail.  A DIB with maximum width/height matching this size (whatever preserves aspect ratio)
' is cached locally, then passed to external functions only when necessary.  Note that requests for sizes larger
' than the hard-coded thumb size (in either dimension) will force creation of a custom thumbnail from scratch,
' so try to avoid that if at all possible!
Private Const IMAGE_THUMB_SIZE As Long = 256
Private m_ImageThumbnail As pdDIB, m_currentImageThumbSize As Long

'Unique hash string identifying this image.  This value is appended to Undo/Redo files, so it must be
' unique to this object for the lifetime of the current session.
Private m_uniqueHash As String

Friend Function GetColorProfile_Original() As String
    GetColorProfile_Original = m_colorProfileOriginal
End Function

Friend Sub SetColorProfile_Original(ByRef srcHash As String)
    m_colorProfileOriginal = srcHash
End Sub

Friend Function GetColorProfile_WorkingSpace() As String
    GetColorProfile_WorkingSpace = m_colorProfileWorkingSpace
End Function

Friend Sub SetColorProfile_WorkingSpace(ByRef srcHash As String)
    m_colorProfileWorkingSpace = srcHash
End Sub

Friend Function GetCurrentFileFormat() As PD_IMAGE_FORMAT
    GetCurrentFileFormat = m_currentFileFormat
End Function

Friend Sub SetCurrentFileFormat(ByVal newFormat As PD_IMAGE_FORMAT)
    m_currentFileFormat = newFormat
End Sub

Friend Function GetImageIcon(ByVal getLargeIcon As Boolean) As Long
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
    GetOriginalColorDepth = m_originalColorDepth
End Function

Friend Sub SetOriginalColorDepth(ByVal origColorDepth As Long)
    m_originalColorDepth = origColorDepth
End Sub

Friend Function GetOriginalGrayscale() As Boolean
    GetOriginalGrayscale = m_originalGrayscale
End Function

Friend Sub SetOriginalGrayscale(ByVal origGrayscale As Boolean)
    m_originalGrayscale = origGrayscale
End Sub

Friend Function GetOriginalAlpha() As Boolean
    GetOriginalAlpha = m_originalAlpha
End Function

Friend Sub SetOriginalAlpha(ByVal origAlpha As Boolean)
    m_originalAlpha = origAlpha
End Sub

'Functions related to storage/retrieval of this image's original palette, if any.
Friend Function HasOriginalPalette() As Boolean
    HasOriginalPalette = (Not m_origPalette Is Nothing)
End Function

Friend Function GetOriginalPalette(ByRef dstPalette As pdPalette) As Boolean
    GetOriginalPalette = Me.HasOriginalPalette()
    If GetOriginalPalette Then Set dstPalette = m_origPalette
End Function

Friend Sub SetOriginalPalette(ByRef srcQuads() As RGBQuad, ByVal palSize As Long)
    If (palSize > 0) Then
        Set m_origPalette = New pdPalette
        m_origPalette.CreateFromPaletteArray srcQuads, palSize
    End If
End Sub

Friend Sub NotifyOriginalPaletteCount(ByVal newPalSize As Long)
    If (Not m_origPalette Is Nothing) Then m_origPalette.SetNewPaletteCount newPalSize
End Sub

Friend Function GetOriginalFileFormat() As PD_IMAGE_FORMAT
    GetOriginalFileFormat = m_originalFileFormat
End Function

Friend Sub SetOriginalFileFormat(ByVal newFormat As PD_IMAGE_FORMAT)
    m_originalFileFormat = newFormat
End Sub

'Zoom is stored as an index into PD's default zoom table.  To resolve these to actual floating-point modifiers,
' use a pdZoom class instance.
Friend Function GetZoomIndex() As Long
    GetZoomIndex = ImgViewport.GetZoomIndex()
End Function

Friend Sub SetZoomIndex(ByVal newIndex As Long)
    ImgViewport.SetZoomIndex newIndex
End Sub

'Time at last modification.  This cannot be directly set by external sources; instead, use the safe "NotifyImageChanged" function.
Friend Function GetTimeOfLastChange() As Currency
    GetTimeOfLastChange = m_timeAtLastModify
End Function

'Return a pseudo-unique ID string.  This *must* be unique for the life of this image object,
' and it should never be shared by two image objects.
Friend Function GetUniqueID() As String
    
    'Generate a unique ID string.
    If (LenB(m_uniqueHash) = 0) Then
        Dim cCrypto As pdCrypto: Set cCrypto = New pdCrypto
        Dim tmpString As String
        tmpString = OS.GetArbitraryGUID()
        m_uniqueHash = cCrypto.QuickHashString(tmpString, Len(tmpString))
        If (Len(m_uniqueHash) > 16) Then m_uniqueHash = Left$(m_uniqueHash, 16)
    End If
    
    GetUniqueID = m_uniqueHash
    
End Function

'This function should *ONLY* be called when restoring data from Autosave data.
' (Otherwise, an ID will be auto-generated the first time GetUniqueID(), above, is called.)
Friend Sub SetUniqueID(ByRef srcID As String)
    
    'If we already have an ID, erase any files associated with that ID
    If (LenB(m_uniqueHash) <> 0) Then
        Set UndoManager = New pdUndo
        Set UndoManager.parentPDImage = Me
    End If
    
    m_uniqueHash = srcID
    
End Sub

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
'The returned Long-type value is the CANONICAL ID of the new layer, NOT ITS INDEX.  (Returning its index would be
' pointless, as the calling function likely knows that in advance.)  Note that layers can be accessed by either ID
' or index; after calling this function, make sure to use the CANONICAL ID function with the returned value.
Friend Function CreateBlankLayer(Optional ByVal layerPosition As Long = -1) As Long

    'Start by seeing if the image has zero layers.  If it does, we can ignore layerPosition entirely.
    If (m_numOfLayersEverCreated = 0) Then
        m_curLayer = 0
        
    'The image has layers.  Extra work is involved
    Else
        
        'At load-time, memory is under especially tight crunch (as the source file or intermediary data
        ' may occupy a lot of memory).  To try and improve the odds of successfully loading large images,
        ' suspend the previous layer, if any, to memory before creating this one.
        
        'TEMPORARILY DISABLED PENDING FURTHER TESTING: this idea is interesting, but at load-time the layer
        ' will simply be un-suspended when the initial composite is created; I'm not sure how to suspend *after*
        ' that (because that's where we could save a *lot* of memory).
        '
        'Until that's solved, this optimization is basically undone by that composite render; solving is TBD.
        'Dim idxPrevLayer As Long
        'idxPrevLayer = m_curLayer
        'If (Not imgLayers(idxPrevLayer).GetLayerDIB() Is Nothing) Then imgLayers(idxPrevLayer).SuspendLayer True
        
        'If no layer has been specified, insert the new layer above the requested layer
        If (layerPosition = -1) Then m_curLayer = m_curLayer + 1 Else m_curLayer = layerPosition + 1
        
        'Resize the layers array by one, and shift all existing layers upward.
        ReDim Preserve imgLayers(0 To m_numOfLayers + 1) As pdLayer
        
        Dim i As Long
        For i = UBound(imgLayers) To m_curLayer Step -1
            Set imgLayers(i) = imgLayers(i - 1)
        Next i
        
    End If
    
    'Initialize the new layer.
    ' (This should already have been done at creation time, but it doesn't hurt to make sure.)
    Set imgLayers(m_curLayer) = New pdLayer
    
    'Assign the layer a canonical ID value.
    imgLayers(m_curLayer).AssignLayerID m_numOfLayersEverCreated
    
    'Return the new layer's ID value (which is just the number of layers ever created for this image)
    CreateBlankLayer = imgLayers(m_curLayer).GetLayerID
    
    'Increment our internal layer counts
    m_numOfLayersEverCreated = m_numOfLayersEverCreated + 1
    m_numOfLayers = m_numOfLayers + 1
    
End Function

Friend Sub RebuildCompositeBuffer(Optional ByVal renderScratchLayerIndex As Long = -1)

    'If the buffer is already clean, ignore this request
    If m_CompositeCacheClean Then
        Debug.Print "WARNING!  pdImage.rebuildCompositeBuffer was called, but the compositeBuffer is marked as clean!"
    Else
        
        'Rebuild the composite cache
        imgCompositor.GetCompositedImage Me, CompositeBuffer, True, renderScratchLayerIndex
        
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
Friend Sub NotifyImageChanged(ByVal typeOfChange As PD_UndoType, Optional ByVal layerIndex As Long = -1)
    
    'Cache a new "image changed" timestamp
    VBHacks.GetHighResTime m_timeAtLastModify
    
    'Separate handling by type.
    
    'TODO: in the future, it would be awesome to store a partial composite up-to-but-NOT-including the current layer.
    ' This would allow for faster recomposites when the current layer is modified.
    ' However, until such time as that feature is implemented, I have no choice but to recomposite the entire image
    ' after changes are made.
    
    Select Case typeOfChange
    
        Case UNDO_Image, UNDO_Image_VectorSafe, UNDO_Everything
            m_CompositeCacheClean = False
            m_IsThumbClean = False
            
            'Notify all layers of potentially destructive changes
            Dim i As Long
            For i = 0 To Me.GetNumOfLayers() - 1
                Me.GetLayerByIndex(i).NotifyOfDestructiveChanges
            Next i
            
            'Notify the tool engine that image size may not be the same
            Tools.NotifyImageSizeChanged
            
        Case UNDO_ImageHeader
            m_CompositeCacheClean = False
            m_IsThumbClean = False
        
        Case UNDO_Layer, UNDO_Layer_VectorSafe
            m_CompositeCacheClean = False
            m_IsThumbClean = False
            
            If (layerIndex <> -1) Then
                Me.GetLayerByIndex(layerIndex).NotifyOfDestructiveChanges
            Else
                Debug.Print "WARNING!  Invalid params passed to pdImage.NotifyImageChanged (-1 layer index for a layer action)"
            End If
        
        Case UNDO_LayerHeader
            m_CompositeCacheClean = False
            m_IsThumbClean = False
        
        'Selection change notifications can also occur, but they don't require anything from us; the selectino object
        ' handles it specially
        Case UNDO_Selection
        
        Case Else
            Debug.Print "WARNING!  pdImage.NotifyImageChanged received an unknown change request: " & typeOfChange
    
    End Select

End Sub

'Merge two layers together.  Note this can be used to merge any two arbitrary layers, with the bottom layer holding the result
' of the merge.  It is up to the caller to deal with any subsequent layer deletions, etc - this sub just performs the merge.
Friend Sub MergeTwoLayers(ByRef topLayer As pdLayer, ByRef bottomLayer As pdLayer, Optional ByVal usePaintStyleMerge As Boolean = False, Optional ByVal ptrToPaintbrushRectF As Long = 0)
    If usePaintStyleMerge Then
        imgCompositor.MergeLayers_PaintStyle topLayer, bottomLayer, ptrToPaintbrushRectF, Me
    Else
        imgCompositor.MergeLayers topLayer, bottomLayer
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
Friend Sub GetCompositedRect(ByRef dstDIB As pdDIB, ByRef dstViewportRect As RectF, ByRef srcImageRect As RectF, ByVal interpolationType As GP_InterpolationMode, Optional ByVal ignoreInternalCaches As Boolean = False, Optional ByVal levelOfDetail As PD_CompositorLOD = CLC_Generic, Optional ByVal renderScratchLayerIndex As Long = -1, Optional ByVal ptrToAlternateScratch As Long = 0&)
    imgCompositor.GetCompositedRect Me, dstDIB, dstViewportRect, srcImageRect, interpolationType, ignoreInternalCaches, levelOfDetail, renderScratchLayerIndex, ptrToAlternateScratch
End Sub

'See if a specific layer ID exists
Friend Function DoesLayerIDExist(ByVal requestedID As Long) As Boolean
    
    DoesLayerIDExist = False
    
    Dim i As Long
    For i = 0 To m_numOfLayers - 1
        If (Not imgLayers(i) Is Nothing) Then
            If (imgLayers(i).GetLayerID = requestedID) Then
                DoesLayerIDExist = True
                Exit Function
            End If
        End If
    Next i
    
End Function

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
        Set GetActiveDIB = imgLayers(m_curLayer).GetLayerDIB
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
    
    Dim origLayerID As Long
    origLayerID = Me.GetActiveLayerID
    
    If (origLayerID <> newActiveLayerID) Then
        
        Dim i As Long
        For i = 0 To UBound(imgLayers)
            If (imgLayers(i).GetLayerID = newActiveLayerID) Then
                m_curLayer = i
                Exit For
            End If
        Next i
        
        Me.NotifyImageChanged UNDO_ImageHeader
        
    End If
        
    'If we made it all the way here, the requested layer was not found.  Do nothing (e.g. do not change the active layer).
    
End Sub

Friend Sub SetActiveLayerByIndex(ByVal newActiveLayerIndex As Long)

    'Validate the incoming layer index
    If (newActiveLayerIndex < 0) Then newActiveLayerIndex = 0
    If (newActiveLayerIndex > m_numOfLayers - 1) Then newActiveLayerIndex = m_numOfLayers - 1
    
    If (m_curLayer <> newActiveLayerIndex) Then
        m_curLayer = newActiveLayerIndex
        Me.NotifyImageChanged UNDO_ImageHeader
    End If
    
End Sub

'Retrieve a layer index using its canonical ID.
Friend Function GetLayerIndexFromID(ByVal requestedID As Long) As Long
    
    Dim i As Long
    For i = 0 To UBound(imgLayers)
        If (Not imgLayers(i) Is Nothing) Then
            If (imgLayers(i).GetLayerID = requestedID) Then
                GetLayerIndexFromID = i
                Exit Function
            End If
        End If
    Next i
    
    'If we made it all the way here, the requested layer was not found.  Return 0.
    GetLayerIndexFromID = 0
    
End Function

'Retrieve a layer at an arbitrary position.  Remember that layers in PD are zero-based, so the base layer is layer 0, not 1.
' Also, this function does not check bounds, so make sure the passed index is valid!
Friend Function GetLayerByIndex(ByVal layerIndex As Long, Optional ByVal forceToValidIndex As Boolean = True) As pdLayer
    
    'As a failsafe, validate the incoming layer index
    If (layerIndex > UBound(imgLayers)) And forceToValidIndex Then
        Debug.Print "WARNING! Invalid layerIndex requested (" & layerIndex & ":" & UBound(imgLayers) & ")"
        layerIndex = UBound(imgLayers)
    End If
    
    If (layerIndex < 0) And forceToValidIndex Then layerIndex = 0
    
    'Return the requested layer.  Note that by design, we allow the target layer to be returned even if it
    ' does not exist.
    If (layerIndex >= 0) And (layerIndex <= UBound(imgLayers)) Then Set GetLayerByIndex = imgLayers(layerIndex)
    
End Function

'Retrieve a layer using its canonical ID.  (ID is unique and guaranteed valid for the life of a layer,
' even across sessions.)
Friend Function GetLayerByID(ByVal requestedID As Long) As pdLayer
    
    Dim i As Long
    For i = 0 To m_numOfLayers - 1
        If (Not imgLayers(i) Is Nothing) Then
            If (imgLayers(i).GetLayerID = requestedID) Then
                Set GetLayerByID = imgLayers(i)
                Exit Function
            End If
        End If
    Next i
    
    'If we made it all the way here, the requested layer was not found.  Return Nothing.
    PDDebug.LogAction "WARNING!  GetLayerByID received an invalid ID: " & requestedID
    For i = 0 To m_numOfLayers - 1
        If (Not imgLayers(i) Is Nothing) Then
            PDDebug.LogAction i & ": " & imgLayers(i).GetLayerID
        Else
            PDDebug.LogAction "(nothing)"
        End If
    Next i
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

'Reverse z-order of all layers in the layer stack
Friend Sub ReverseLayerOrder()
    
    'As a convenience to the user, preserve the active layer
    Dim curLayerID As Long
    curLayerID = Me.GetActiveLayerID
    
    'Reverse layer order
    Dim i As Long
    For i = 0 To (m_numOfLayers \ 2) - 1
        Me.SwapTwoLayers i, (m_numOfLayers - 1) - i
    Next i
    
    'Restore the originally selected layer
    Me.SetActiveLayerByID curLayerID
    
    'Notify ourselves that the entire image has change
    Me.NotifyImageChanged UNDO_Image
    
End Sub

'Returns TRUE if this image is currently suspended (to either memory or disk; see params for details)
Friend Function IsSuspended(Optional ByVal checkSuspendedToDisk As Boolean = False) As Boolean
    
    IsSuspended = True
    
    'Because PD automatically un-suspends layers when they are accessed, it's nontrivial to assume an entire
    ' image is either "suspended" or "not suspended".
    '
    'As such, just scan layers and if *any* are unsuspended, treat the whole image as unsuspended.
    
    'Start by suspending all layers
    Dim i As Long
    For i = 0 To UBound(imgLayers)
        If (Not imgLayers(i) Is Nothing) Then
            If (Not imgLayers(i).IsSuspended(checkSuspendedToDisk)) Then
                IsSuspended = False
                Exit For
            End If
        End If
    Next i
    
    
End Function

'Non-destructively suspend as many image resources as we can to either disk or memory.
Friend Sub SuspendImage(Optional ByVal suspendToDisk As Boolean = False, Optional ByVal alsoSuspendSharedBuffer As Boolean = True)
    
    'Normally we just suspend to a compressed memory stream, but if the user wants the image suspended to disk,
    ' that's a little more aggressive.
    
    'Start by suspending all layers
    Dim i As Long
    For i = 0 To UBound(imgLayers)
        If (Not imgLayers(i) Is Nothing) Then
            If (Not imgLayers(i).IsSuspended(suspendToDisk)) Then imgLayers(i).SuspendLayer True, suspendToDisk
        End If
    Next i
    
    'Suspend any other internal DIBs
    If (Not m_ImageThumbnail Is Nothing) Then m_ImageThumbnail.SuspendDIB cf_Lz4, False
    If (Not CompositeBuffer Is Nothing) Then CompositeBuffer.SuspendDIB cf_Lz4, False
    If (Not CanvasBuffer Is Nothing) Then CanvasBuffer.SuspendDIB cf_Lz4, False
    
    'After suspending everything, we're also going to release the shared compression buffer.
    ' The main benefit of this is freeing memory - the shared compression buffer has to maintain
    ' the worst-case post-compression size of any object that attempts to use it.  This is
    ' usually the size of the object plus some trivial amount of bytes.  For an image-sized buffer,
    ' this is a large ask, and while there may be a slight stutter the next time the buffer needs
    ' to be compressed, that's a lesser usability evil than leaving behind a massive, momentarily
    ' unnecessary memory chunk.
    If alsoSuspendSharedBuffer Then UIImages.FreeSharedCompressBuffer
    
End Sub

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

'Delete a given layer by ID.  The pdLayers stack will automatically be resized to match.
Friend Sub DeleteLayerByID(ByVal srcLayerID As Long)
    
    Dim i As Long
    For i = 0 To m_numOfLayers - 1
        If (imgLayers(i).GetLayerID = srcLayerID) Then
            Me.DeleteLayerByIndex i
            Exit Sub
        End If
    Next i
    
    'If we reach this function, the ID in question wasn't found
    PDDebug.LogAction "WARNING! pdImage.DeleteLayerByID couldn't find the requested ID: " & srcLayerID
    
End Sub

'Delete a given layer by index.  The pdLayers stack will automatically be resized to match.
Friend Sub DeleteLayerByIndex(ByVal srcLayerIndex As Long)

    'Validate the layer index
    If (srcLayerIndex < 0) Or (srcLayerIndex >= m_numOfLayers) Then
        PDDebug.LogAction "WARNING: pdImage.DeleteLayerByIndex received a bad index: " & srcLayerIndex
        Exit Sub
    End If
    
    'If this is the last remaining layer, exit
    If (m_numOfLayers = 1) Then Exit Sub
    
    'Shift all layer references above this one downward
    Dim i As Long
    If (srcLayerIndex < m_numOfLayers - 1) Then
        For i = srcLayerIndex To m_numOfLayers - 2
            Set imgLayers(i) = imgLayers(i + 1)
        Next i
    End If
    
    'Shrink the layers array
    m_numOfLayers = m_numOfLayers - 1
    ReDim Preserve imgLayers(0 To m_numOfLayers - 1) As pdLayer
    
    'If this causes the current layer index to point at an invalid index, adjust it automatically
    If (m_curLayer > m_numOfLayers - 1) Then m_curLayer = m_numOfLayers - 1

End Sub

'Get/set active selection state
Friend Function IsSelectionActive() As Boolean
    IsSelectionActive = m_SelectionIsActive
End Function

Friend Sub SetSelectionActive(ByVal newState As Boolean)
    m_SelectionIsActive = newState
End Sub

'Serialize all relevant image header data to an XML string.
' (This uses the new, much-improved 2020 header format.  ALL new PD code must use this function from now on.)
Friend Function GetHeaderAsXML() As String

    'Prepare an XML engine, which will handle the actual writing of the file
    Dim xmlEngine As pdSerialize
    Set xmlEngine = New pdSerialize
    
    'Add a basic header and version info
    xmlEngine.AddXMLString "<photodemon-image>"
    xmlEngine.AddParam "image-version", 1&, True, True
    
    'Start by writing out the hard-coded image ID.  This value can effectively be ignored at load-time;
    ' we only use it when recovering autosave data after a crash (as this ID helps us match up autosaved
    ' file data).
    xmlEngine.AddParam "image-id", imageID, True, True
    
    'Write relevant parent image data
    xmlEngine.AddParam "width", Me.Width, True, True
    xmlEngine.AddParam "height", Me.Height, True, True
    xmlEngine.AddParam "x-resolution", m_XResolution, True, True
    xmlEngine.AddParam "y-resolution", m_YResolution, True, True
    xmlEngine.AddParam "dpi", m_DPI, True, True
    xmlEngine.AddParam "file-format-original", ImageFormats.GetExtensionFromPDIF(m_originalFileFormat), True, True
    xmlEngine.AddParam "file-format-current", ImageFormats.GetExtensionFromPDIF(m_currentFileFormat), True, True
    xmlEngine.AddParam "color-depth-original", m_originalColorDepth, True, True
    xmlEngine.AddParam "layer-count-max", m_numOfLayersEverCreated, True, True
    xmlEngine.AddParam "layer-count-active", m_numOfLayers, True, True
    xmlEngine.AddParam "active-layer", m_curLayer, True, True
    xmlEngine.AddParam "animated", m_Animated, True, True
    
    'Next comes conditional data, which may not exist in the catch-all property dictionary.
    ' (Note that we manually pull out tags that we care about.  At present, we don't bother writing
    ' the full dictionary, as most dictionary data is only useful per-session anyway.)
    
    'Return the finished XML string
    GetHeaderAsXML = xmlEngine.GetParamString()
    
End Function

'Sister function to GetHeaderAsXML(), above.  If you report a value in that function, make sure to
' include it here!
'
'If this function is being called as part of an Undo/Redo action, we can safely ignore certain entries
' in the file (in favor of the values already have stored for this image).  The Undo/Redo engine may
' also specify non-destructive loading, in which case we assume that our existing layer stack is valid,
' but in the wrong order - so we'll simply reorder the stack according to the data in the XML,
' instead of full-on recreating it from scratch.  (This is a high-performance way to handle Undo/Redo
' data for operations like reordering the layer stack.)
Friend Function SetHeaderFromXML(ByRef srcText As String, Optional ByVal sourceIsUndoFile As Boolean = False, Optional ByVal loadNonDestructively As Boolean = False) As Boolean
    
    'Prep a fast XML param engine, which greatly simplifies the process of retrieving XML data
    Dim xmlEngine As pdSerialize
    Set xmlEngine = New pdSerialize
    xmlEngine.SetParamString srcText
    
    'Validate the XML header...
    If (xmlEngine.DoesParamExist("photodemon-image") And (xmlEngine.GetLong("image-version", 0) >= 1)) Then
    
        'The XML package appears to be valid.  Start retrieving values.
        Me.Width = xmlEngine.GetLong("width", 0, True)
        Me.Height = xmlEngine.GetLong("height", 0, True)
        m_XResolution = xmlEngine.GetDouble("x-resolution", 0#, True)
        m_YResolution = xmlEngine.GetDouble("y-resolution", 0#, True)
        m_DPI = xmlEngine.GetDouble("dpi", 0#, True)
        m_Animated = xmlEngine.GetBool("animated", False, True)
        
        'All settings past this point are only relevant when loading a file anew,
        ' *not* when it's being used as part of the Undo/Redo stack.
        If (Not sourceIsUndoFile) Then
            
            'These values are all stored in the PDI file, but they're not really relevant.  I will look at not storing
            ' them in the first place, but for now, just ignore them.
            Me.SetOriginalColorDepth xmlEngine.GetLong("color-depth-original", 32, True)
            Me.SetOriginalFileFormat ImageFormats.GetPDIFFromExtension(xmlEngine.GetString("file-format-original", vbNullString, True))
            Me.SetCurrentFileFormat ImageFormats.GetPDIFFromExtension(xmlEngine.GetString("file-format-current", vbNullString, True))
            
            'Dictionary tags are conditional, and may not exist in the PDI file.
            ' Check for their existence, and load only as necessary.
            
        End If
        
        'Finally, retrieve layer values from the file
        m_numOfLayersEverCreated = xmlEngine.GetLong("layer-count-max", 0, True)
        m_numOfLayers = xmlEngine.GetLong("layer-count-active", 0, True)
        m_curLayer = xmlEngine.GetLong("active-layer", 0, True)
        
        'If this is a standard "destructive" load (e.g. we are overwriting all existing properties
        ' with incoming ones), we can now prepare the layer collection to receive layer data.
        
        '(Otherwise, we must leave the current layer collection as it is - this operation will
        ' only reorder existing data!)
        If (Not loadNonDestructively) Then
        
            'Initialize the layer array
            ReDim imgLayers(0 To m_numOfLayers - 1) As pdLayer
            Dim i As Long
            For i = 0 To m_numOfLayers - 1
                Set imgLayers(i) = New pdLayer
            Next i
            
        End If
        
        SetHeaderFromXML = True
    
    'Basic validation failed.  Attempt the legacy loader.
    Else
        Set xmlEngine = Nothing
        PDDebug.LogAction "pdImage.SetHeaderFromXML() is trying again using the legacy serialization engine..."
        SetHeaderFromXML = SetHeaderFromXML_Legacy(srcText, sourceIsUndoFile, loadNonDestructively)
    End If
    
End Function

Friend Function IsAnimated() As Boolean
    IsAnimated = m_Animated
End Function

Friend Sub SetAnimated(ByVal newState As Boolean)
    m_Animated = newState
End Sub

'Get/Set image resolution (in DPI).  Note that the vertical resolution is optional; if the two values
' differ, PD will average them when image DPI is requested.
Friend Function GetDPI() As Double
    If (m_DPI = 0#) Then GetDPI = 96# Else GetDPI = m_DPI
End Function

Friend Sub SetDPI(ByVal xRes As Double, ByVal yRes As Double)
    
    'Many image types do not store resolution information; default to 96 in this case
    If (xRes = 0#) Then xRes = 96#
    If (yRes = 0#) Then yRes = 96#
    
    m_XResolution = xRes
    m_YResolution = yRes
    
    'It is extremely rare for x/y resolution to differ, but just in case, calculate an average resolution as well
    m_DPI = (xRes + yRes) * 0.5

End Sub

'Helper function to fill a RectF with current image boundaries (as used in hit-testing).
' Left/top will always be (0, 0).
Friend Function GetBoundaryRectF() As RectF
    With GetBoundaryRectF
        .Left = 0!
        .Top = 0!
        .Width = Me.Width
        .Height = Me.Height
    End With
End Function

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

'Update the internal thumbnail.  Only call this function if the current thumbnail has been marked as DIRTY;
' otherwise, just use the existing thumbnail as-is.
Private Sub UpdateInternalThumbnail(Optional ByVal requiredSize As Long = 0)
    
    'If the image has not been instantiated properly, reject the thumbnail request
    If (Not Me.IsActive) Or (m_numOfLayers = 0) Then Exit Sub
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'To help us coalesce multiple thumbnail requests together, we simply cache a "large enough" copy
    ' and then produce smaller copies on-demand.
    If (requiredSize < IMAGE_THUMB_SIZE) Then requiredSize = IMAGE_THUMB_SIZE
    m_currentImageThumbSize = requiredSize
    
    'Determine proper dimensions for the thumbnail image.
    Dim newThumbWidth As Long, newThumbHeight As Long
    PDMath.ConvertAspectRatio Me.Width, Me.Height, requiredSize, requiredSize, newThumbWidth, newThumbHeight
    
    'Prepare the thumbnail DIB
    If (m_ImageThumbnail.GetDIBWidth <> newThumbWidth) Or (m_ImageThumbnail.GetDIBHeight <> newThumbHeight) Then
        m_ImageThumbnail.CreateBlank newThumbWidth, newThumbHeight, 32, 0
    Else
        m_ImageThumbnail.ResetDIB 0
    End If
    
    'Retrieve a composited thumbnail.  (Note that the user's thumbnail performance setting affects the interpolation
    ' method used.)
    Dim dstRectF As RectF, srcRectF As RectF
    With dstRectF
        .Left = 0!
        .Top = 0!
        .Width = newThumbWidth
        .Height = newThumbHeight
    End With
    
    With srcRectF
        .Left = 0!
        .Top = 0!
        .Width = Me.Width
        .Height = Me.Height
    End With
    
    Dim gdipTimeStart As Currency, gdipTimeFinal As String
    VBHacks.GetHighResTime gdipTimeStart
    Me.GetCompositedRect m_ImageThumbnail, dstRectF, srcRectF, UserPrefs.GetThumbnailInterpolationPref(), , CLC_Thumbnail
    m_ImageThumbnail.SetInitialAlphaPremultiplicationState True
    gdipTimeFinal = VBHacks.GetTimeDiffNowAsString(gdipTimeStart)
    
    'Before exiting, apply color-management.  This spares callers from needing to do it (and when wouldn't we want
    ' a color-managed thumbnail anyway?)
    Dim cmTime As Currency
    VBHacks.GetHighResTime cmTime
    ColorManagement.ApplyDisplayColorManagement m_ImageThumbnail, , True
    
    'Mark the thumbnail state as clean
    m_IsThumbClean = True
    PDDebug.LogAction "pdImage had to generate a thumbnail; it took " & VBHacks.GetTimeDiffNowAsString(startTime) & " total (color-management: " & VBHacks.GetTimeDiffNowAsString(cmTime) & ", GDI+: " & gdipTimeFinal & ")"
    
End Sub

'External functions can use this function to request a thumbnail version of the contained image.  By default,
' the thumbnail is assumed to be square, and this function will automatically center the image inside the
' square thumbnail.  (Optional parameters allow you to bypass this; they should be self-explanatory.)
'
'IMPORTANT NOTE: the returned thumbnail *has already been color-managed for the active display*.  Do not reapply
' color-management settings.
'
'IMPORTANT NOTE: the full image stack may need to be re-composited on weird size requests, so please
' 1) call this function sparingly, and...
' 2) stick to the default IMAGE_THUMB_SIZE value or smaller
Friend Function RequestThumbnail(ByRef dstThumbnailDIB As pdDIB, Optional ByVal thumbnailSize As Long = IMAGE_THUMB_SIZE, Optional ByVal thumbIsSquare As Boolean = True, Optional ByVal ptrToRectF As Long = 0) As Boolean
    
    'Is the current thumbnail dirty?  If so, regenerate it.
    If (Not m_IsThumbClean) Or (thumbnailSize > m_currentImageThumbSize) Then UpdateInternalThumbnail thumbnailSize
    
    'We also need to determine the thumbnail's actual width and height, and any x and y offset necessary to preserve the
    ' aspect ratio and center the image on the thumbnail.
    Dim thumbWidth As Long, thumbHeight As Long, thumbLeft As Single, thumbTop As Single
    PDMath.ConvertAspectRatio m_ImageThumbnail.GetDIBWidth, m_ImageThumbnail.GetDIBHeight, thumbnailSize, thumbnailSize, thumbWidth, thumbHeight
        
    'Thumbnails have some interesting requirements.  We always want them to be square, with the image set in the middle
    ' of the thumbnail (with aspect ratio preserved) and any empty edges made transparent.
    If thumbIsSquare Then
        
        'If the image is wider than it is tall, center the thumbnail vertically
        If (thumbWidth > thumbHeight) Then
            thumbLeft = 0!
            thumbTop = (thumbnailSize - thumbHeight) * 0.5
        
        '...otherwise, center it horizontally
        Else
            thumbTop = 0!
            thumbLeft = (thumbnailSize - thumbWidth) * 0.5
        End If
        
        'Prep the destination thumbnail
        If (dstThumbnailDIB Is Nothing) Then Set dstThumbnailDIB = New pdDIB
        If (dstThumbnailDIB.GetDIBWidth <> thumbnailSize) Or (dstThumbnailDIB.GetDIBHeight <> thumbnailSize) Or (dstThumbnailDIB.GetDIBWidth = 0) Then
            dstThumbnailDIB.CreateBlank thumbnailSize, thumbnailSize, 32, 0, 0
        Else
            dstThumbnailDIB.ResetDIB 0
        End If
        
        'Paint the thumbnail into place
        GDI_Plus.GDIPlus_StretchBlt dstThumbnailDIB, thumbLeft, thumbTop, thumbWidth, thumbHeight, m_ImageThumbnail, 0!, 0!, m_ImageThumbnail.GetDIBWidth, m_ImageThumbnail.GetDIBHeight, , UserPrefs.GetThumbnailInterpolationPref(), , , , True
    
    'If this is a non-standard request (e.g. a square thumbnail isn't needed), paint out the thumbnail at whatever
    ' size the caller wants.  Note that they are typically responsible for preparing the target DIB in this scenario,
    ' but PD will provide a "failsafe" check if the passed tmpRectF doesn't exist.
    Else
        
        Dim tmpRectF As RectF
        If (ptrToRectF <> 0) Then
            CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
        Else
            
            With tmpRectF
                .Left = 0!
                .Top = 0!
                .Width = thumbWidth
                .Height = thumbHeight
            End With
            
            If (dstThumbnailDIB Is Nothing) Then
                Set dstThumbnailDIB = New pdDIB
            Else
                If (dstThumbnailDIB.GetDIBWidth > 0) Then dstThumbnailDIB.ResetDIB 0
            End If
            
            dstThumbnailDIB.CreateBlank thumbWidth, thumbHeight, 32, 0, 0
            
        End If
        
        GDI_Plus.GDIPlus_StretchBlt dstThumbnailDIB, tmpRectF.Left, tmpRectF.Top, tmpRectF.Width, tmpRectF.Height, m_ImageThumbnail, 0!, 0!, m_ImageThumbnail.GetDIBWidth, m_ImageThumbnail.GetDIBHeight, , UserPrefs.GetThumbnailInterpolationPref(), , , , True
        
    End If
    
    dstThumbnailDIB.SetInitialAlphaPremultiplicationState True
    
    'NOTE: the thumbnail *has already been color-managed*, so the caller doesn't need to manage it further.
    RequestThumbnail = True
    
End Function

'When the attached image object is deactivated (e.g. the user leaves it loaded, but switches to a
' different image), you can call this function to free up some internal memory caches and other objects.
' This can free a non-trivial amount of resources when system memory is tight.
'
'For even more resource reduction, ask for a full layer suspension too - this will ask all child layers
' to suspend their DIBs to compressed buffers.
Friend Sub DeactivateImage(Optional ByVal suspendLayersToo As Boolean = True, Optional ByVal suspendToDisk As Boolean = False)

    'Scratch layers aren't needed unless the image is actively being edited
    Set ScratchLayer = Nothing
    
    'Some internal selection caches are free-able
    If (Not MainSelection Is Nothing) Then MainSelection.FreeNonEssentialResources
    
    'Layer suspension is the biggest gain, but it also has a perf penalty so the caller may
    ' choose to avoid this.
    If suspendLayersToo Then Me.SuspendImage suspendToDisk
    
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

'When the images's size changes, call this function.
'
'The first optional parameter (TRUE by default), makes the image the same size as its base layer.
' This is used when an image file is first loaded, as the pdImage container should be set to an
' identical size.
'
'If this is another type of resize action, such as a canvas resize, the caller must supply new
' width/height values.
'
'Note that this function performs NO VALIDATION on the passed width/height values.
' Make sure they are correct in advance!
Friend Sub UpdateSize(Optional ByVal useBaseLayer As Boolean = True, Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0)
    
    If useBaseLayer Then
        Me.Width = imgLayers(0).GetLayerWidth(False)
        Me.Height = imgLayers(0).GetLayerHeight(False)
    Else
        Me.Width = newWidth
        Me.Height = newHeight
    End If
    
    'Changes to image size require a rebuild of our composited image cache
    Me.NotifyImageChanged UNDO_Image
    
End Sub

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
    
    'Initialize the image thumbnail and note that it is not clean, which will force a regeneration when the thumbnail
    ' is next requested.
    Set m_ImageThumbnail = New pdDIB
    m_IsThumbClean = False
    
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
    
    Set m_ImageThumbnail = Nothing
    
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
' Note that this is a very fast operation for normal layers, but if affine transforms are active, PD has to do a
' lot more work (and significantly more memory will be required).
Friend Sub RetrieveNullPaddedLayer(ByRef dstDIB As pdDIB, ByVal srcLayerIndex As Long)

    'Create a blank destination DIB at the size of the image
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.CreateBlank Me.Width, Me.Height, 32, 0
    
    'Two branches: one for normal layers, another for affine-transformed ones.
    If imgLayers(srcLayerIndex).AffineTransformsActive(True) Then
    
        'A temporary DIB is required to handle the intermediate copy of the source layer.
        ' (This could be optimized away, but it's a non-trivial optimization depending on the complexity
        ' of the source layer.)
        Dim tmpDIB As pdDIB, xOffset As Long, yOffset As Long
        If imgLayers(srcLayerIndex).GetAffineTransformedDIB(tmpDIB, xOffset, yOffset) Then
        
            'Paint the transformed DIB into position
            GDI.BitBltWrapper dstDIB.GetDIBDC, xOffset, yOffset, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'The entire layer lies off-image, meaning there's nothing to paint.
        ' (Fortunately, the target DIB is already blank, so we don't need to do anything!)
        Else
        
        End If
    
    'Normal layers are very simple: create the destination DIB and paint us directly into position.
    Else
        GDI.BitBltWrapper dstDIB.GetDIBDC, imgLayers(srcLayerIndex).GetLayerOffsetX, imgLayers(srcLayerIndex).GetLayerOffsetY, imgLayers(srcLayerIndex).GetLayerWidth(False), imgLayers(srcLayerIndex).GetLayerHeight(False), imgLayers(srcLayerIndex).GetLayerDIB.GetDIBDC, 0, 0, vbSrcCopy
    End If
    
    'The target DIB will *always* be premultiplied
    dstDIB.SetInitialAlphaPremultiplicationState True
    
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
    Dim maskBounds As RectF
    maskBounds = MainSelection.GetCompositeBoundaryRect
    dstDIB.CreateBlank maskBounds.Width, maskBounds.Height, 32, 0
    dstDIB.SetInitialAlphaPremultiplicationState True
    
    'We now need pointers to three separate sets of image data: destination DIB, source DIB, and selection mask.
    Dim srcImageData() As Byte, srcSA As SafeArray1D
    Dim selData() As Byte, selSA As SafeArray1D
    Dim dstImageData() As Byte, dstSA As SafeArray1D
    
    Dim leftOffset As Long, topOffset As Long
    leftOffset = maskBounds.Left
    topOffset = maskBounds.Top
    
    Dim x As Long, y As Long
    Dim r As Long, g As Long, b As Long
    Dim thisAlpha As Long, origAlpha As Long
    Dim blendAlpha As Double
    
    Dim dstQuickX As Long, srcQuickX As Long, srcQuickY As Long
    Const BYTE_TO_FLOAT As Single = 1! / 255!
    
    For y = 0 To dstDIB.GetDIBHeight - 1
        
        srcQuickY = topOffset + y
        srcDIB.WrapArrayAroundScanline srcImageData, srcSA, srcQuickY
        MainSelection.GetCompositeMaskDIB.WrapArrayAroundScanline selData, selSA, srcQuickY
        dstDIB.WrapArrayAroundScanline dstImageData, dstSA, y
        
    For x = 0 To dstDIB.GetDIBWidth - 1
        
        dstQuickX = x * 4
        srcQuickX = (leftOffset + x) * 4
        thisAlpha = selData(srcQuickX)
        
        'If the selection does not exist at this pixel, we have no reason to process this pixel further!
        If (thisAlpha > 0) Then
        
            'Check the image's alpha value.  If it's zero, we have no reason to process it further
            origAlpha = srcImageData(srcQuickX + 3)
            
            If (origAlpha > 0) Then
                
                'Source pixel data will be premultiplied, which saves us a bunch of processing time.
                ' (That is why we premultiply alpha, after all!)
                b = srcImageData(srcQuickX)
                g = srcImageData(srcQuickX + 1)
                r = srcImageData(srcQuickX + 2)
                
                'Calculate a new multiplier, based on the strength of the selection at this location
                blendAlpha = thisAlpha * BYTE_TO_FLOAT
                
                'Apply the multiplier to the existing pixel data (which is already premultiplied, saving us a bunch of time now)
                dstImageData(dstQuickX) = b * blendAlpha
                dstImageData(dstQuickX + 1) = g * blendAlpha
                dstImageData(dstQuickX + 2) = r * blendAlpha
                
                'Finish our work by calculating a new alpha channel value for this pixel, which is a blend of
                ' the original alpha value, and the selection mask value at this location.
                dstImageData(dstQuickX + 3) = origAlpha * blendAlpha
                
            End If
            
        End If
        
    Next x
    Next y
    
    'Clear all array references
    srcDIB.UnwrapArrayFromDIB srcImageData
    MainSelection.GetCompositeMaskDIB.UnwrapArrayFromDIB selData
    dstDIB.UnwrapArrayFromDIB dstImageData
    
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
    
    'If this is a vector layer, rasterize it
    If imgLayers(targetLayerIndex).IsLayerVector Then Layers.RasterizeLayer targetLayerIndex
    
    'Before doing anything else, backup the current layer's position and dimensions, then null-pad
    ' the layer.  (This simplifies the process of deleting selection areas that may lie outside
    ' the current layer's dimensions - but note that for layers with affine transforms, we fall back
    ' to a simple expand-process-autocrop function.)
    Dim isLayerAffineTransformed As Boolean, origLayerRectF As RectF
    isLayerAffineTransformed = imgLayers(targetLayerIndex).AffineTransformsActive(False)
    If (Not isLayerAffineTransformed) Then imgLayers(targetLayerIndex).GetLayerBoundaryRect origLayerRectF
    imgLayers(targetLayerIndex).ConvertToNullPaddedLayer Me.Width, Me.Height
    
    'Selections can be one of several types.  Right now, we don't give special handling to simple
    ' rectangular selections - all selections are fully processed according to the contents of the
    ' selection mask.  (Also, all selections are currently created as 32bpp DIBs.)
    
    'We now need pointers to two separate sets of image data: destination DIB, and selection mask.
    Dim selData() As Byte, selSA As SafeArray1D
    Dim dstImageData() As Byte, dstSA As SafeArray1D
    
    Dim maskBounds As RectF
    maskBounds = MainSelection.GetCompositeBoundaryRect
    
    Dim leftBound As Long, topBound As Long, rightBound As Long, bottomBound As Long
    leftBound = Int(maskBounds.Left)
    topBound = Int(maskBounds.Top)
    rightBound = leftBound + Int(maskBounds.Width) - 1
    bottomBound = topBound + Int(maskBounds.Height) - 1
    
    Dim x As Long, y As Long
    Dim r As Long, g As Long, b As Long
    Dim thisAlpha As Long, origAlpha As Long
    Dim blendAlpha As Double
    
    Dim dstQuickX As Long, selQuickX As Long
    
    Const BYTE_TO_FLOAT As Single = 1! / 255!
    
    For y = topBound To bottomBound
        MainSelection.GetCompositeMaskDIB.WrapArrayAroundScanline selData, selSA, y
        imgLayers(targetLayerIndex).GetLayerDIB.WrapArrayAroundScanline dstImageData, dstSA, y
    For x = leftBound To rightBound
    
        selQuickX = x * 4
        thisAlpha = selData(selQuickX)
        
        'If the selection does not exist at this pixel, we have no reason to process this pixel further!
        If (thisAlpha > 0) Then
        
            thisAlpha = 255 - thisAlpha
        
            'Check the image's alpha value.  If it's zero, we have no reason to process it further
            dstQuickX = x * 4
            origAlpha = dstImageData(dstQuickX + 3)
            If (origAlpha > 0) Then
                
                'Source pixel data will be premultiplied, which saves us a bunch of processing time.  (That is why
                ' we premultiply alpha, after all!)
                b = dstImageData(dstQuickX)
                g = dstImageData(dstQuickX + 1)
                r = dstImageData(dstQuickX + 2)
                
                'Calculate a new multiplier, based on the strength of the selection at this location
                blendAlpha = thisAlpha * BYTE_TO_FLOAT
                
                'Apply the multiplier to the existing pixel data (which is already premultiplied, saving us a bunch of time now)
                dstImageData(dstQuickX) = b * blendAlpha
                dstImageData(dstQuickX + 1) = g * blendAlpha
                dstImageData(dstQuickX + 2) = r * blendAlpha
                
                'Finish our work by calculating a new alpha channel value for this pixel, which is a blend of
                ' the original alpha value, and the selection mask value at this location.
                dstImageData(dstQuickX + 3) = origAlpha * blendAlpha
                
            End If
            
        End If
        
    Next x
    Next y
    
    'Clear all array references
    MainSelection.GetCompositeMaskDIB.UnwrapArrayFromDIB selData
    imgLayers(targetLayerIndex).GetLayerDIB.UnwrapArrayFromDIB dstImageData
    
    'Shrink the layer to its maximum relevant size, either using the original layer rect (if the layer
    ' *wasn't affine-transformed) or via auto-crop.
    If (Not isLayerAffineTransformed) Then
        imgLayers(targetLayerIndex).CropNullPaddedLayerToRect origLayerRectF
    Else
        imgLayers(targetLayerIndex).CropNullPaddedLayer
    End If
    
    'Notify ourselves of the change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, targetLayerIndex
    
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

    EstimateRAMUsage = 0#

    Dim i As Long
    
    For i = 0 To m_numOfLayers - 1
        
        'Layers provide their own RAM estimation, based on their current cache state, layer DIB, active styles, and more
        EstimateRAMUsage = EstimateRAMUsage + imgLayers(i).EstimateRAMUsage()
        
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
    If (ScratchLayer.GetLayerDIB.GetDIBWidth <> Me.Width) Or (ScratchLayer.GetLayerDIB.GetDIBHeight <> Me.Height) Then
        ScratchLayer.GetLayerDIB.CreateBlank Me.Width, Me.Height, 32, 0, 0
        ScratchLayer.GetLayerDIB.SetInitialAlphaPremultiplicationState True
    Else
        ScratchLayer.GetLayerDIB.ResetDIB 0
    End If
    
    'No other scratch layer attributes are set by this function, unless specifically requested
    If resetLayerSettingsToo Then ScratchLayer.ResetLayerParameters

End Sub

'Want to suspend (or subsequently re-enable) animations?  Use this sub.  (At present, this only affects the
' "marching ant" selection animation.)
Friend Sub NotifyAnimationsAllowed(ByVal allowedState As Boolean)
    If (Not MainSelection Is Nothing) Then MainSelection.NotifyAnimationsAllowed allowedState
End Sub




'**********************************************************************************************
'Legacy support functions follow.  Do NOT modify these functions except to fix security issues.

'Legacy header creator
Private Function SetHeaderFromXML_Legacy(ByRef srcText As String, Optional ByVal sourceIsUndoFile As Boolean = False, Optional ByVal loadNonDestructively As Boolean = False) As Boolean
    
    'Prepare an XML engine, which will handle the actual reading and parsing of the file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'We now have two choices, depending on the value of filePathIsActuallyString.  Load the supposed XML contents
    ' (either from file or directly) into the XML engine; note that both steps validate the incoming data automatically.
    Dim xmlLoadSuccessful As Boolean
    xmlLoadSuccessful = xmlEngine.LoadXMLFromString(srcText)
    
    'Make sure the supplied XML actually contains pdImage data.
    If xmlLoadSuccessful And xmlEngine.IsPDDataType("pdImage") Then
    
        'The XML file checked out.  Start retrieving relevant values.
        Me.Width = xmlEngine.GetUniqueTag_Long("Width")
        Me.Height = xmlEngine.GetUniqueTag_Long("Height")
        m_XResolution = xmlEngine.GetUniqueTag_Double("XResolution")
        m_YResolution = xmlEngine.GetUniqueTag_Double("YResolution")
        m_DPI = xmlEngine.GetUniqueTag_Double("DPI")
        m_Animated = xmlEngine.GetUniqueTag_Boolean("animated")
        
        'All settings past this point are only relevant when loading a file anew, not when it's being used as part of
        ' the Undo/Redo stack.
        If (Not sourceIsUndoFile) Then
            
            'These values are all stored in the PDI file, but they're not really relevant.  I will look at not storing
            ' them in the first place, but for now, just ignore them.
            Me.SetOriginalColorDepth xmlEngine.GetUniqueTag_Long("originalColorDepth")
            
            'Dictionary tags are conditional, and may not exist in the PDI file.  Check for their existence, and load them
            ' only as necessary.
            If xmlEngine.DoesTagExist("png-background-color") Then ImgStorage.AddEntry "png-background-color", xmlEngine.GetUniqueTag_Long("png-background-color")
            
        End If
        
        'Finally, retrieve layer values from the file
        m_numOfLayersEverCreated = xmlEngine.GetUniqueTag_Long("numOfLayersEverCreated")
        m_numOfLayers = xmlEngine.GetUniqueTag_Long("numOfLayers")
        m_curLayer = xmlEngine.GetUniqueTag_Long("curLayer")
        
        'If this is a standard, destructive load, prepare our image layer array to receive the upcoming layer data.
        ' (Otherwise, leave it as it is, because the calling function will simply reorder it for us.)
        If (Not loadNonDestructively) Then
        
            'Now that all data has been read, we can initialize our layers array.  (The calling function will presumably proceed to
            ' fill each layer with usable data, because we can't do that here!)
            ReDim imgLayers(0 To m_numOfLayers - 1) As pdLayer
            
            Dim i As Long
            For i = 0 To m_numOfLayers - 1
                Set imgLayers(i) = New pdLayer
            Next i
            
        End If
        
        SetHeaderFromXML_Legacy = True
    
    Else
        PDDebug.LogAction "WARNING!  pdImage.SetHeaderFromXML_Legacy() couldn't validate XML header."
    End If
    
End Function

