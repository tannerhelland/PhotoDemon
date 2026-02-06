VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Layers class
'Copyright 2012-2026 by Tanner Helland
'Created: 29/August/12
'Last updated: 10/March/25
'Last update: expose "suspend to disk" behavior to callers
'
'Once upon a time, the pdLayer class was a simple pixel array wrapper.  Once PhotoDemon grew into
' a full layer-based photo editor, however, it became necessary to separate Layer-specific functions
' from generic pixel array functions.
'
'The current pdLayer class is a comprehensive wrapper for all PhotoDemon layers, regardless of type.
' All layers share key components, like opacity and blend mode, and all raster layers (and vector
' layers, for performance reasons) also provide a copy of the layer's contents in pdDIB format.
' Note that PD's main viewport renderer doesn't care how layer raster data is created - e.g. whether
' it comes from text or pixel arrays or vector data - it just cares that the associated pdDIB object
' is present and appropriately sized for its contents.
'
'Because layers interact heavily with PDI files (PhotoDemon's internal file format), each layer must
' be capable of serializing all meaningful properties to XML, and recreating itself from such an XML
' packet when requested.  Layer pixel data is always handled separately, as some layer types may need
' to recreate their raster contents on-the-fly (e.g. rasterizing text layers).
'
'pdLayer objects have no knowledge of their parent pdImage object, by design.  If interaction with
' the parent object is required, functions should explicitly request a pdImage reference as a
' function parameter, and you should always guarantee that the reference will only be used locally.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Layer type is set at creation time.  Once a layer has been created, its type should not be changed.
' (The one exception to this rule is converting vector layers (text, etc) to image layers. This can
'  be done, but it is a one-way conversion. Once a layer is converted to PDL_IMAGE type, it cannot
'  be reverted.)
Private m_LayerType As PD_LayerType

'Canonical ID value for this layer.  This value is set once, at layer creation time,
' and can never be changed again.  It is persistent for the life of the layer,
' the life of the parent pdImage object (meaning no other layer in that image will
' ever be assigned this ID value), and the life of any copies of the parent pdImage
' saved to file.
'
'Because this value cannot be changed once created, I keep it separate from the
' LayerData struct (whose data *can* be modified).
Private m_layerID As Long

'Layers store a lot of layer-type-agnostic metadata.
' **All of these entries can be changed by the user at run-time.**
' Default values are set at creation time, but never assume that these values are constant,
' and never attempt to reference a layer by one of these properties (INCLUDING NAME, as it can be
' modified without warning)
Private Type PD_LayerData
    l_Name As String                    'Layer name, as entered by the user
    l_GroupID As Long                   'Layer group (0 if not assigned to a group)
    l_Opacity As Single                 'Layer opacity ([0, 100] - defaults to 100)
    l_BlendMode As PD_BlendMode         'Layer blend mode; defaults to BM_Normal
    l_AlphaMode As PD_AlphaMode         'Layer alpha mode; defaults to AM_Normal
    l_OffsetX As Long                   'X/Y offset of the layer's top-left corner
    l_OffsetY As Long
    l_CanvasXModifier As Double         'Width/height modifier for the image; by default, these are set to 1.0.
    l_CanvasYModifier As Double
    l_Angle As Double                   'Layer angle (implemented in PD 7.0)
    l_RotateCenterX As Double           'Layer rotation center point (implemented in PD 7.0)
    l_RotateCenterY As Double
    l_ResizeQuality As PD_LayerResizeQuality   'Resampling algorithm used when non-destructive resizes are active
    l_ShearX As Double                  'Layer shear (added in PD 7.0)
    l_ShearY As Double
    l_FrameTimeInMS As Long             'Frame time, in MS; relevant only for images flagged as animated
    l_Visibility As Boolean             'Layer visibility
    l_MaskExists As Boolean             'Layer mask exists at all.
    l_MaskActive As Boolean             'Layer mask is currently active (a mask may exist, but not be active)
                                        ' (For other mask attributes, query m_LayerMask directly)
End Type

'Local instance of layer data for this layer
Private myLayerData As PD_LayerData

'Vector-type layers store many things that image layers do not.  Generic data shared between all vector layer types are
' stored in this struct.
Private Type GenericVectorData
    vd_Width As Single          'Unlike raster layers, vector layers must explicitly store their width and height
    vd_Height As Single
End Type

Private myVectorData As GenericVectorData

'Vector layers must be converted into raster format for rendering on the screen.
' If the DIB is synched against the the current vector contents, this value will be TRUE.
' If it is FALSE, the DIB must be (re)generated prior to display.
Private m_VectorDIBSynched As Boolean

'Next comes various structs for vector data in a specific format.
' These structs must be comprehensive, because their data is used to reconstruct a raster copy
' of this layer between sessions.

'TEXT LAYER STRUCTS

'A dedicated text renderer is used to render the backer layer for text layers.
' This object also manages all relevant text/typography layer data.
Private myTextRenderer As pdTextRenderer

'All layers are ultimately backed by a pdDIB surface.
' - For image layers, the layer's pixel contents are stored here.
' - For non-image layers, a pixel-level rendering of the layer contents are stored here.
'   This DIB is used to render the layer onto the viewport, rather than constantly recreating whatever
'   the layer's contents may be (text, etc)
' - For adjustment and other non-data layers, this object may not be used at all.
'
'For performance reasons, outside callers can modify this object freely via the GetLayerDIB() function.
' That said, any function that directly modifies the DIB must also take care of any management functions
' (e.g. redrawing the viewport), because layers have no way to propagate change notifications up to a
' parent pdImage object. (This propagation is typically handled by a call to the parent pdImage's
' .NotifyImageChanged function(), while passing the relevant layer index.)
Private m_LayerDIB As pdDIB

'Temporary level-of-detail (LOD) compositing DIB(s).
' At present, use of these DIBs is restricted to the pdCompositor class.  Because it's often necessary to
' generate temporary DIBs on a per-layer basis, these level-of-detail DIBs can save a ton of time during
' performance-crucial operations (e.g. painting, compositing).
Private m_lodDIB() As pdDIB

'Layer thumbnail.  A DIB of this size is cached locally, then passed to external functions only when necessary.
' Note that requests for sizes larger than the hard-coded thumb size (in either dimension) will force creation of a
' custom thumbnail from scratch, so try to avoid that if at all possible!
Private Const LAYER_THUMB_SIZE As Long = 256
Private m_LayerThumbnail As pdDIB, m_CurrentLayerThumbSize As Long

'If the current thumbnail is clean or dirty.  Clean means that to the best of this class's knowledge, the current thumbnail accurately
' reflects the layer's contents.  Dirty means that we know the current thumbnail is invalid, so regenerate it before continuing.
Private m_IsThumbClean As Boolean

'For some parts of the rendering pipeline, we can cache intermediate copies of the layer, cropped and stretched and prepped for
' viewport rendering.  To simplify the work performed by the compositor, layers can generate a quick-and-dirty "viewport state hash"
' that represents current layer state.  This hash can be used to detect changes to the layer, which in turn tells us if we need to
' regenerate the relevant LOD DIB, or if we can just use the existing copy.
Private m_ViewportStateHashes() As Long

'If a layer has changed since the last hash cache, we must still recomposite the DIB, even though the automatic hash function won't
' generate a new value (because it primarily relies on layer metadata, not actual DIB contents)
Private m_DIBChangedSinceLastHash() As Boolean

'Whenever we are notified of destructive changes, we make a current high-res time stamp.
' Outside callers can use this to synchronize any state trackers, and they can safely skip updates
' if their stored timestamp matches this one.
Private m_TimeAtLastDestructiveChange As Currency

'If a layer has non-destructive, non-standard transformations active (e.g. rotation, skew),
' we use a special, parallelogram-based rendering function.  To cut down on variable declarations,
' we declare this array once, and initialize it when a compositor instance is initialized.
Private m_PlgPoints() As PointFloat

'This layer's mask, if any, will be stored here.
' (To see if this layer has meaningful mask data, query myLayerData.l_MaskActive.)
Private m_LayerMask As pdLayerMask

'If an external function modifies this layer's DIB contents (or the vector data from which the DIB is generated),
' it must call this function afterward.
'
'IMPORTANT NOTE! As of v7.0, PD functions should not call this sub directly.
' Instead, they must call NotifyImageChanged on the *parent* pdImage object, so the parent image
' knows to recomposite itself (which is necessary for some viewport models).
'
'ANOTHER IMPORTANT NOTE!  This function only needs to be called if a change was *destructive*.
' Non-destructive changes - per the name - do not require us to rebuild internal caches,
' because the effects and/or resizing will be handled by a compositor object.
'
'Finally, this function will mark all layer caches as dirty, causing a lot of subsequent work,
' so use it only when absolutely necessary.
Friend Sub NotifyOfDestructiveChanges()

    'Mark the layer thumbnail as dirty and update our tracking timestamp
    m_IsThumbClean = False
    m_TimeAtLastDestructiveChange = VBHacks.GetHighResTimeInMSEx()
    
    'Note that the layer has changed since a past hash request
    Dim i As Long
    For i = 0 To NUM_OF_LOD_CACHES - 1
        m_DIBChangedSinceLastHash(i) = True
    Next i
    
    'If this is a vector layer, note that the backing DIB is now out of sync
    m_VectorDIBSynched = False
    
End Sub

'Direct access to the underlying DIB object.  Do not access for vector layers, as any changes will
' be overwritten by the vector renderer.  Also, remember to notify the parent image of changes via
' its .NotifyImageChanged() function.
Friend Function GetLayerDIB() As pdDIB
    Set GetLayerDIB = m_LayerDIB
End Function

Friend Sub SetLayerDIB(ByRef newDIB As pdDIB)
    Set m_LayerDIB = newDIB
End Sub

'Get/set this layer's canonical ID value.  Note that this value is valid for both the life of the layer, and the life of its
' parent image (including persistence when writing image data to/from file).
'
'Note that the set function is preferenced by "assign" - I do this on purpose, to help me remember that the function should only
' ever be called once, right after the layer is first created.  After that, it should never, ever be changed!
Friend Function GetLayerID() As Long
    GetLayerID = m_layerID
End Function

Friend Sub AssignLayerID(ByVal thisLayerID As Long)
    
    'As a failsafe, warn me if this layer has already been assigned an ID.
    ' (This should *never* happen, but I like to verify.)
    If (Not OS.IsProgramCompiled) Then
        If (m_layerID <> -1) Then InternalError "AssignLayerID", "this layer already has an ID"
    End If
    
    m_layerID = thisLayerID
    
End Sub

'Directly access our child layer mask object.  Guarantees that a pdLayerMask object exists,
' but *not* that the class contains an actual mask.  (Query this layer's HasMask member for that.)
Friend Function GetLayerMask() As pdLayerMask
    If (m_LayerMask Is Nothing) Then Set m_LayerMask = New pdLayerMask
    Set GetLayerMask = m_LayerMask
End Function

Friend Function GetLayerMaskActive() As Boolean
    GetLayerMaskActive = myLayerData.l_MaskActive
End Function

Friend Sub SetLayerMaskActive(ByVal newState As Boolean)
    myLayerData.l_MaskActive = newState
End Sub

Friend Function GetLayerMaskExists() As Boolean
    GetLayerMaskExists = myLayerData.l_MaskExists
End Function

Friend Sub SetLayerMaskExists(ByVal newState As Boolean)
    myLayerData.l_MaskExists = newState
End Sub

'Get layer type.
Friend Function GetLayerType() As PD_LayerType
    GetLayerType = m_LayerType
End Function

'Set layer type.  Note that this only sets the layer type flag; changing the actual layer contents to match (e.g. rasterization)
' must be handled separately.
Friend Sub SetLayerType(ByVal newLayerType As PD_LayerType)
    m_LayerType = newLayerType
    m_VectorDIBSynched = False
End Sub

'Counterpart to IsLayerVector(), IsLayerText(), below
Friend Function IsLayerRaster() As Boolean
    IsLayerRaster = Not IsLayerVector()
End Function

'If a caller simply wants to determine text vs other layers, this helper function takes care of it.
' Text layers currently include basic and advanced text layers, and may include additional types
' (e.g. "WordArt") in the future.
Friend Function IsLayerText() As Boolean

    Select Case m_LayerType
    
        Case PDL_TextBasic, PDL_TextAdvanced
            IsLayerText = True
        
        Case Else
            IsLayerText = False
    
    End Select

End Function

'If a caller simply wants to determine vector vs raster layers, this helper function takes care of it.  Note that text layers are
' considered vector layers, because all changes are non-destructive.
Friend Function IsLayerVector() As Boolean

    Select Case m_LayerType
    
        Case PDL_Image
            IsLayerVector = False
        
        Case PDL_TextBasic, PDL_TextAdvanced
            IsLayerVector = True
        
        Case Else
            InternalError "IsLayerVector", "Layer type unknown!  Returning VECTOR = FALSE by default."
    
    End Select

End Function

'Rasterize a layer.  In the future, it may be possible to "Undo" this (e.g. restore the original vector state, since that data is
' easy to retain), but I haven't implemented it at present.  Rasterization is currently immutable, short of the Undo/Redo menu.
Friend Sub RasterizeVectorData()

    'If the layer is already rasterized, ignore this request
    If Me.IsLayerRaster Then
        InternalError "RasterizeVectorData", "can't rasterize a raster layer: " & Me.GetLayerName()
        Exit Sub
    End If
    
    'Rasterizing is a lot simpler than it looks.  Basically, we just change the layer type to PDL_IMAGE and discard any superfluous vector data.
    m_LayerType = PDL_Image
        
    'While possibly unnecessary for some vector layer types, rasterization is a serious enough change that it's wise to regenerate
    ' any/all internal caches.
    Me.NotifyOfDestructiveChanges

End Sub

'Get/set name
Friend Function GetLayerName() As String
    GetLayerName = myLayerData.l_Name
End Function

Friend Sub SetLayerName(ByRef newLayerName As String)
    myLayerData.l_Name = newLayerName
End Sub

'Get/set layer group
Friend Function GetLayerGroup() As Long
    GetLayerGroup = myLayerData.l_GroupID
End Function

Friend Sub SetLayerGroup(ByVal newLayerGroup As Long)
    myLayerData.l_GroupID = newLayerGroup
End Sub

'Get/set opacity
Friend Function GetLayerOpacity() As Single
    GetLayerOpacity = myLayerData.l_Opacity
End Function

Friend Sub SetLayerOpacity(ByVal newLayerOpacity As Single)
    If (newLayerOpacity > 100!) Then newLayerOpacity = 100!
    If (newLayerOpacity < 0!) Then newLayerOpacity = 0!
    myLayerData.l_Opacity = newLayerOpacity
End Sub

'Get/set alpha and blend modes
Friend Function GetLayerAlphaMode() As PD_AlphaMode
    GetLayerAlphaMode = myLayerData.l_AlphaMode
End Function

Friend Sub SetLayerAlphaMode(ByVal newLayerAlphaMode As PD_AlphaMode)
    myLayerData.l_AlphaMode = newLayerAlphaMode
End Sub

Friend Function GetLayerBlendMode() As PD_BlendMode
    GetLayerBlendMode = myLayerData.l_BlendMode
End Function

Friend Sub SetLayerBlendMode(ByVal newLayerBlendMode As PD_BlendMode)
    myLayerData.l_BlendMode = newLayerBlendMode
End Sub

'Get/set visibility
Friend Function GetLayerVisibility() As Boolean
    GetLayerVisibility = myLayerData.l_Visibility
End Function

Friend Sub SetLayerVisibility(ByVal newVisibility As Boolean)
    myLayerData.l_Visibility = newVisibility
End Sub

'Get/set layer offsets
Friend Function GetLayerOffsetX() As Long
    GetLayerOffsetX = myLayerData.l_OffsetX
End Function

Friend Function GetLayerOffsetY() As Long
    GetLayerOffsetY = myLayerData.l_OffsetY
End Function

Friend Sub SetLayerOffsetX(ByVal newOffsetX As Long)
    myLayerData.l_OffsetX = newOffsetX
End Sub

Friend Sub SetLayerOffsetY(ByVal newOffsetY As Long)
    myLayerData.l_OffsetY = newOffsetY
End Sub

'Helper function for setting layer offset and non-destructive width/height all at once.
Friend Sub SetOffsetsAndModifiersTogether(ByVal topLeftX As Long, ByVal topLeftY As Long, ByVal bottomRightX As Long, ByVal bottomRightY As Long)
    
    With myLayerData
        
        .l_OffsetX = topLeftX
        .l_OffsetY = topLeftY
        
        'Image layers have their own base size, and a modifier is used to "stretch" that size accordingly
        If (m_LayerType = PDL_Image) Then
            
            If (Me.GetLayerWidth(False) <> 0#) Then
                Me.SetLayerCanvasXModifier (bottomRightX - topLeftX) / Me.GetLayerWidth(False)
            Else
                Me.SetLayerCanvasXModifier 1#
            End If
            
            If (Me.GetLayerHeight(False) <> 0#) Then
                Me.SetLayerCanvasYModifier (bottomRightY - topLeftY) / Me.GetLayerHeight(False)
            Else
                Me.SetLayerCanvasYModifier 1#
            End If
        
        'Text layers are simply resized, while their modifier is kept at 1.0
        ElseIf (m_LayerType = PDL_TextBasic) Or (m_LayerType = PDL_TextAdvanced) Then
            Me.SetLayerWidth bottomRightX - topLeftX
            Me.SetLayerHeight bottomRightY - topLeftY
        End If
        
    End With
    
End Sub

'Get/set layer size.  Note that the corresponding "set size" functions are ONLY VALID FOR VECTOR LAYERS; because size changes are destructive
' for raster layers, they must be handled manually.  (The exception, of course, being the non-destructive X/Y canvas modifier values, which exist
' for this explicit purpose!)
'
'Also, at present, these values are not stored persistently.  They are simply generated on-the-fly by taking the m_layerDIB object's dimensions,
' and multiplying them by the current canvas x/y modifiers.  The caller can override the canvas modifier calculation by passing FALSE to the function.
Friend Function GetLayerWidth(Optional ByVal applyCanvasModifier As Boolean = True) As Double
    
    'The returned width varies by layer type
    Select Case m_LayerType
    
        'Image layers do not store persistent width/height values.  Return the layer DIB's dimensions in its place
        Case PDL_Image
            
            If (Not m_LayerDIB Is Nothing) Then
        
                If applyCanvasModifier Then
                    GetLayerWidth = m_LayerDIB.GetDIBWidth * GetLayerCanvasXModifier
                    If (GetLayerWidth < 1#) Then GetLayerWidth = 1#
                Else
                    GetLayerWidth = m_LayerDIB.GetDIBWidth
                End If
                
            Else
                GetLayerWidth = 0#
            End If
        
        'Vector layers do actually store persistent width/height values.  This is important because these values may fall
        ' out of sync with the backer DIB, which is only updated on an as-needed basis
        Case PDL_TextBasic, PDL_TextAdvanced
            GetLayerWidth = myVectorData.vd_Width
            If (GetLayerWidth < 1#) Then GetLayerWidth = 1#
            
    End Select
    
End Function

Friend Function GetLayerHeight(Optional ByVal applyCanvasModifier As Boolean = True) As Double
    
    'The returned height varies by layer type
    Select Case m_LayerType
    
        'Image layers do not store persistent width/height values.  Return the layer DIB's dimensions in its place
        Case PDL_Image
            
            If (Not m_LayerDIB Is Nothing) Then
        
                If applyCanvasModifier Then
                    GetLayerHeight = m_LayerDIB.GetDIBHeight * GetLayerCanvasYModifier
                    If (GetLayerHeight < 1#) Then GetLayerHeight = 1#
                Else
                    GetLayerHeight = m_LayerDIB.GetDIBHeight
                End If
                
            Else
                GetLayerHeight = 0#
            End If
        
        'Vector layers do actually store persistent width/height values.  This is important because these values may fall
        ' out of sync with the backer DIB, which is only updated on an as-needed basis
        Case PDL_TextBasic, PDL_TextAdvanced
            GetLayerHeight = myVectorData.vd_Height
            If (GetLayerHeight < 1#) Then GetLayerHeight = 1#
    
    End Select
    
End Function

Friend Sub SetLayerWidth(ByVal newSize As Single)
    
    'Force size in-bounds
    If (newSize < 1!) Then newSize = 1!
    
    'Make sure the layer is a vector layer
    If (m_LayerType <> PDL_Image) And (newSize <> myVectorData.vd_Width) Then
        
        'Store the new size and mark the DIB cache as dirty
        myVectorData.vd_Width = newSize
        m_VectorDIBSynched = False
        
    Else
        If (m_LayerType = PDL_Image) Then InternalError "SetLayerWidth", "setLayerWidth and setLayerHeight functions are only valid for vector layers"
    End If
    
End Sub

Friend Sub SetLayerHeight(ByVal newSize As Single)
    
    'Force size in-bounds
    If (newSize < 1!) Then newSize = 1!
    
    'Make sure the layer is a vector layer
    If (m_LayerType <> PDL_Image) And (newSize <> myVectorData.vd_Height) Then
        
        'Store the new size and mark the DIB cache as dirty
        myVectorData.vd_Height = newSize
        m_VectorDIBSynched = False
        
    Else
        If (m_LayerType = PDL_Image) Then InternalError "SetLayerHeight", "setLayerWidth and setLayerHeight functions are only valid for vector layers"
    End If
    
End Sub

'Get/set layer canvas modifiers (necessary for non-destructive on-canvas resizing).  For vector layers, these have no effect.
' Vector layers always maintain canvas x/y modifiers of 1, and they simply translate these requests into setLayerWidth/Height calls.
Friend Function GetLayerCanvasXModifier() As Double
    GetLayerCanvasXModifier = myLayerData.l_CanvasXModifier
End Function

Friend Function GetLayerCanvasYModifier() As Double
    GetLayerCanvasYModifier = myLayerData.l_CanvasYModifier
End Function

Friend Sub SetLayerCanvasXModifier(ByVal newXModifier As Double)
    
    'Check bounds
    If (newXModifier <= 0#) Then newXModifier = 0.00000001
    
    'Separate handling by layer type
    Select Case m_LayerType
    
        Case PDL_Image
            myLayerData.l_CanvasXModifier = newXModifier
            
        Case PDL_TextBasic, PDL_TextAdvanced
            myLayerData.l_CanvasXModifier = 1#
            Me.SetLayerWidth newXModifier * Me.GetLayerWidth(False)
    
    End Select
    
End Sub

Friend Sub SetLayerCanvasYModifier(ByVal newYModifier As Double)
    
    'Check bounds
    If (newYModifier <= 0#) Then newYModifier = 0.00000001
    
    'Separate handling by layer type
    Select Case m_LayerType
    
        Case PDL_Image
            myLayerData.l_CanvasYModifier = newYModifier
            
        Case PDL_TextBasic, PDL_TextAdvanced
            myLayerData.l_CanvasYModifier = 1#
            Me.SetLayerHeight newYModifier * Me.GetLayerHeight(False)
    
    End Select
    
End Sub

'Get/set layer frame time, always in MS.  This data is only useful to PD under certain circumstances
' (e.g. if the parent image is flagged as "animated").
Friend Function GetLayerFrameTimeInMS() As Long
    GetLayerFrameTimeInMS = myLayerData.l_FrameTimeInMS
End Function

Friend Sub SetLayerFrameTimeInMS(ByVal newTime As Long)
    myLayerData.l_FrameTimeInMS = newTime
End Sub

'Get/set non-destructive resize algorithm.  Note that there are two GET functions:
' - One that uses PD's internal PD_LayerResizeQuality enum
' - One that converts the internal PD_LayerResizeQuality value into a GDI+ resize constant
Friend Function GetLayerResizeQuality() As PD_LayerResizeQuality
    GetLayerResizeQuality = myLayerData.l_ResizeQuality
End Function

Friend Function GetLayerResizeQuality_GDIPlus() As GP_InterpolationMode
    
    Select Case myLayerData.l_ResizeQuality
        
        Case LRQ_NearestNeighbor
            GetLayerResizeQuality_GDIPlus = GP_IM_NearestNeighbor
        
        Case LRQ_Bilinear
            GetLayerResizeQuality_GDIPlus = GP_IM_Bilinear
        
        Case LRQ_Bicubic
            GetLayerResizeQuality_GDIPlus = GP_IM_HighQualityBicubic
        
    End Select
    
End Function

Friend Sub SetLayerResizeQuality(ByVal newQuality As PD_LayerResizeQuality)
    myLayerData.l_ResizeQuality = newQuality
End Sub

'Generic layer properties can be get/set via these singular functions.  Note that these functions get and return Variants,
' as layer properties include a lot of different types.  As such, these aren't suitable for use inside a performance-sensitive
' loop (like pdCompositor, for example).  Instead, these are primarily used by UI elements and the Undo/Redo processor, where
' it is much simpler to use a single generic function.
Friend Function GetGenericLayerProperty(ByVal desiredProperty As PD_LayerGenericProperty) As Variant

    Select Case desiredProperty
    
        Case pgp_Name
            GetGenericLayerProperty = myLayerData.l_Name
            
        Case pgp_GroupID
            GetGenericLayerProperty = myLayerData.l_GroupID
            
        Case pgp_Opacity
            GetGenericLayerProperty = myLayerData.l_Opacity
            
        Case pgp_BlendMode
            GetGenericLayerProperty = myLayerData.l_BlendMode
            
        Case pgp_OffsetX
            GetGenericLayerProperty = myLayerData.l_OffsetX
            
        Case pgp_OffsetY
            GetGenericLayerProperty = myLayerData.l_OffsetY
            
        Case pgp_CanvasXModifier
            GetGenericLayerProperty = myLayerData.l_CanvasXModifier
            
        Case pgp_CanvasYModifier
            GetGenericLayerProperty = myLayerData.l_CanvasYModifier
            
        Case pgp_Angle
            GetGenericLayerProperty = myLayerData.l_Angle
            
        Case pgp_Visibility
            GetGenericLayerProperty = myLayerData.l_Visibility
            
        Case pgp_ResizeQuality
            GetGenericLayerProperty = myLayerData.l_ResizeQuality
            
        Case pgp_ShearX
            GetGenericLayerProperty = myLayerData.l_ShearX
            
        Case pgp_ShearY
            GetGenericLayerProperty = myLayerData.l_ShearY
            
        Case pgp_AlphaMode
            GetGenericLayerProperty = myLayerData.l_AlphaMode
            
        Case pgp_RotateCenterX
            GetGenericLayerProperty = myLayerData.l_RotateCenterX
        
        Case pgp_RotateCenterY
            GetGenericLayerProperty = myLayerData.l_RotateCenterY
        
        Case pgp_FrameTime
            GetGenericLayerProperty = myLayerData.l_FrameTimeInMS
        
        Case pgp_MaskExists
            GetGenericLayerProperty = myLayerData.l_MaskExists
        
        Case pgp_MaskActive
            GetGenericLayerProperty = myLayerData.l_MaskActive
        
        Case Else
            GetGenericLayerProperty = 0
    
    End Select

End Function

Friend Sub SetGenericLayerProperty(ByVal desiredProperty As PD_LayerGenericProperty, ByVal newValue As Variant)

    Select Case desiredProperty
        
        Case pgp_Name
            If (myLayerData.l_Name <> CStr(newValue)) Then myLayerData.l_Name = CStr(newValue)
            
        Case pgp_GroupID
            If (myLayerData.l_GroupID <> CLng(newValue)) Then myLayerData.l_GroupID = CLng(newValue)
            
        Case pgp_Opacity
            If (myLayerData.l_Opacity <> CSng(newValue)) Then Me.SetLayerOpacity CSng(newValue)
            
        Case pgp_BlendMode
            If (myLayerData.l_BlendMode <> CLng(newValue)) Then myLayerData.l_BlendMode = CLng(newValue)
            
        Case pgp_OffsetX
            If (myLayerData.l_OffsetX <> Int(newValue + 0.5)) Then myLayerData.l_OffsetX = Int(newValue + 0.5)
            
        Case pgp_OffsetY
            If (myLayerData.l_OffsetY <> Int(newValue + 0.5)) Then myLayerData.l_OffsetY = Int(newValue + 0.5)
            
        Case pgp_CanvasXModifier
            If (myLayerData.l_CanvasXModifier <> CDbl(newValue)) Then Me.SetLayerCanvasXModifier CDbl(newValue)
            
        Case pgp_CanvasYModifier
            If (myLayerData.l_CanvasYModifier <> CDbl(newValue)) Then Me.SetLayerCanvasYModifier CDbl(newValue)
            
        Case pgp_Angle
            If (myLayerData.l_Angle <> CDbl(newValue)) Then myLayerData.l_Angle = CDbl(newValue)
            
        Case pgp_Visibility
            If (myLayerData.l_Visibility <> CBool(newValue)) Then myLayerData.l_Visibility = CBool(newValue)
            
        Case pgp_ResizeQuality
            If (myLayerData.l_ResizeQuality <> CLng(newValue)) Then myLayerData.l_ResizeQuality = CLng(newValue)
            
        Case pgp_ShearX
            If (myLayerData.l_ShearX <> CDbl(newValue)) Then myLayerData.l_ShearX = CDbl(newValue)
            
        Case pgp_ShearY
            If (myLayerData.l_ShearY <> CDbl(newValue)) Then myLayerData.l_ShearY = CDbl(newValue)
            
        Case pgp_AlphaMode
            If (myLayerData.l_AlphaMode <> CLng(newValue)) Then myLayerData.l_AlphaMode = CLng(newValue)
            
        Case pgp_RotateCenterX
            If (myLayerData.l_RotateCenterX <> CDbl(newValue)) Then myLayerData.l_RotateCenterX = CDbl(newValue)
        
        Case pgp_RotateCenterY
            If (myLayerData.l_RotateCenterY <> CDbl(newValue)) Then myLayerData.l_RotateCenterY = CDbl(newValue)
        
        Case pgp_FrameTime
            myLayerData.l_FrameTimeInMS = CLng(newValue)
            
        Case pgp_MaskExists
            myLayerData.l_MaskExists = CBool(newValue)
        
        Case pgp_MaskActive
            myLayerData.l_MaskActive = CBool(newValue)
            
    End Select
    
End Sub

'Text layer properties can be get/set via these singular functions.
Friend Function GetTextLayerProperty(ByVal desiredProperty As PD_TextProperty) As Variant

    'All text properties are managed by the pdTextRenderer class.  This function is simply a passthrough.
    GetTextLayerProperty = myTextRenderer.GetGenericTextProperty(desiredProperty)
    
End Function

Friend Sub SetTextLayerProperty(ByVal desiredProperty As PD_TextProperty, ByVal newValue As Variant)
    
    'All text properties are managed by the pdTextRenderer class.  This function is simply a passthrough.
    
    'Note that the setGenericTextProperty function will return TRUE if the requested change matters (e.g. if the text needs to be re-rendered
    ' due to this property change).  We modify our internal VectorDIBSynched tracker accordingly.
    m_VectorDIBSynched = m_VectorDIBSynched And (Not myTextRenderer.SetGenericTextProperty(desiredProperty, newValue))
    If (Not m_VectorDIBSynched) Then m_IsThumbClean = False
    
End Sub

Friend Function GetTimeOfLastChange() As Currency
    GetTimeOfLastChange = m_TimeAtLastDestructiveChange
End Function

'Returns TRUE if the underlying layer pixel bits have been suspended to either a compressed memory stream
' or disk.
Friend Function IsSuspended(Optional ByVal checkSuspendedToDisk As Boolean = False) As Boolean
    If (Not m_LayerDIB Is Nothing) Then IsSuspended = m_LayerDIB.IsSuspended(checkSuspendedToDisk)
End Function

'If this layer needs to have its on-canvas transformations made permanent (for example, when applying an action
' that involves a selection), this function can be used.
Friend Sub MakeCanvasTransformsPermanent(Optional ByVal maxWidth As Long = 0, Optional ByVal maxHeight As Long = 0)

    'If this layer does not have any active transforms, ignore this request
    If Me.AffineTransformsActive(True) Then
        
        'Unfortunately, a temporary DIB will be required for the rescale operation
        Dim tmpTransformDIB As pdDIB
        Set tmpTransformDIB = New pdDIB
        
        'If complex affine transforms are active (e.g. rotation), this step becomes more complicated.
        If Me.AffineTransformsActive(False) Then
        
            'Affine transforms can cause the layer's x/y offset to change
            Dim newOffsetX As Long, newOffsetY As Long
            
            'Retrieve a copy of the current DIB using the layer's automated affine-transformation function
            Me.GetAffineTransformedDIB tmpTransformDIB, newOffsetX, newOffsetY, maxWidth, maxHeight
            
            'Update our internal DIB and offsets to match
            Set m_LayerDIB = tmpTransformDIB
            With myLayerData
                .l_OffsetX = newOffsetX
                .l_OffsetY = newOffsetY
                .l_CanvasXModifier = 1
                .l_CanvasYModifier = 1
                .l_Angle = 0
                .l_ShearX = 0
                .l_ShearY = 0
                .l_RotateCenterX = 0.5
                .l_RotateCenterY = 0.5
            End With
        
        'If rescaling is the only active transform, we can shortcut this function
        Else
            
            With tmpTransformDIB
                .CreateBlank GetLayerWidth(True), GetLayerHeight(True), 32, 0
                GDIPlusResizeDIB tmpTransformDIB, 0, 0, .GetDIBWidth, .GetDIBHeight, m_LayerDIB, 0, 0, GetLayerWidth(False), GetLayerHeight(False), Me.GetLayerResizeQuality_GDIPlus
            End With
            
            Set m_LayerDIB = tmpTransformDIB
            
            'Reset the width/height modifiers for this layer
            With myLayerData
                .l_CanvasXModifier = 1#
                .l_CanvasYModifier = 1#
            End With
        
        End If
                
        'Mark this layer as dirty
        Me.NotifyOfDestructiveChanges
        
    End If

End Sub

'Get/set layer angle and rotation center point.  (Rotation center point is a fraction on the scale [0, 1])
Friend Function GetLayerAngle() As Double
    GetLayerAngle = myLayerData.l_Angle
End Function

Friend Sub SetLayerAngle(ByVal newAngle As Double)
    myLayerData.l_Angle = newAngle
End Sub

Friend Function GetLayerRotateCenterX() As Double
    GetLayerRotateCenterX = myLayerData.l_RotateCenterX
End Function

Friend Sub SetLayerRotateCenterX(ByVal newCenterX As Double)
    myLayerData.l_RotateCenterX = newCenterX
End Sub

Friend Function GetLayerRotateCenterY() As Double
    GetLayerRotateCenterY = myLayerData.l_RotateCenterY
End Function

Friend Sub SetLayerRotateCenterY(ByVal newCenterY As Double)
    myLayerData.l_RotateCenterY = newCenterY
End Sub

'Get/set layer shear
Friend Function GetLayerShearX() As Double
    GetLayerShearX = myLayerData.l_ShearX
End Function

Friend Sub SetLayerShearX(ByVal newShear As Double)
    myLayerData.l_ShearX = newShear
End Sub

Friend Function GetLayerShearY() As Double
    GetLayerShearY = myLayerData.l_ShearY
End Function

Friend Sub SetLayerShearY(ByVal newShear As Double)
    myLayerData.l_ShearY = newShear
End Sub

'If one or more affine transformations are active (rotation, shear, etc), this function will return TRUE.
' This is used, for example, by pdCompositor to branch across several rendering paths.
Friend Function AffineTransformsActive(Optional ByVal includeScaling As Boolean = False) As Boolean
    
    'Rotate and shear are the primary relevant affine transforms.  Note that we don't care about rotation center
    ' point here; when the layer angle is 0, rotation center point is irrelevant.
    AffineTransformsActive = (myLayerData.l_Angle <> 0#) Or (myLayerData.l_ShearX <> 0#) Or (myLayerData.l_ShearY <> 0#)
    If includeScaling Then AffineTransformsActive = AffineTransformsActive Or (myLayerData.l_CanvasXModifier <> 1#) Or (myLayerData.l_CanvasYModifier <> 1#)
    
End Function

'If one or more affine transformations are active (rotation, skew, etc), you can call this function to retrieve
' three corner points of the layer *IN IMAGE COORDINATES*.
'
'dstPoints needs to have the bounds (0, 2) or (0, 3).  It will be filled with the following points in this order:
' 0) Top-Left
' 1) Top-Right
' 2) Bottom-Left
' 3) (optionally) Bottom-Right.  This point is optional because it can be inferred from the previous three
'
'The optional "usePathOrder" parameter will swap the order of (2) and (3).
Friend Sub GetLayerCornerCoordinates(ByRef dstPoints() As PointFloat, Optional ByVal usePathOrder As Boolean = False)
    
    'To simplify the work of translating layer boundaries by one or more affine transformations, we use a GraphicsMatrix object.
    Dim tmpMatrix As pd2DTransform
    
    'Retrieve a copy of the current layer transform matrix
    Me.GetCopyOfLayerTransformationMatrix_Full tmpMatrix
    
    'Fill the dstPoints array with the current image corners *without* translation applied (but *with* scaling)
    dstPoints(0).x = 0!
    dstPoints(0).y = 0!
    
    dstPoints(1).x = Me.GetLayerWidth(False)
    dstPoints(1).y = 0!
    
    If usePathOrder Then
        dstPoints(2).x = Me.GetLayerWidth(False)
        dstPoints(2).y = Me.GetLayerHeight(False)
    Else
        dstPoints(2).x = 0!
        dstPoints(2).y = Me.GetLayerHeight(False)
    End If
    
    'If the third point exists, fill it too
    If (UBound(dstPoints) >= 3) Then
        If usePathOrder Then
            dstPoints(3).x = 0
            dstPoints(3).y = Me.GetLayerHeight(False)
        Else
            dstPoints(3).x = Me.GetLayerWidth(False)
            dstPoints(3).y = Me.GetLayerHeight(False)
        End If
    End If
    
    'Translate each point in turn
    tmpMatrix.ApplyTransformToPointFs VarPtr(dstPoints(0)), 3
    If (UBound(dstPoints) >= 3) Then tmpMatrix.ApplyTransformToPointF dstPoints(3)
    
    'The temporary translation matrix will automatically be cleaned up when the function exits
    
    'Debug information can be displayed here:
    'Debug.Print "Layer corners: 0 - (" & dstPoints(0).x & ", " & dstPoints(0).y & ") 1 - (" & dstPoints(1).x & ", " & dstPoints(1).y & ") 1 - (" & dstPoints(2).x & ", " & dstPoints(2).y & ")"
    
End Sub

'As of PhotoDemon 7.0, non-destructive layer rotation is supported.  Call this function to retrieve the position of the layer's on-canvas
' rotation node *IN IMAGE COORDINATES*.  Note that an array is required, and it must be at least (0, 4) in size.  The first point contains
' the center point of the rotation node (e.g. the center of the current image), while subsequent points contain valid rotation nodes, which
' sit on the halfway point at each image boundary.
Friend Sub GetLayerRotationNodeCoordinates(ByRef dstPoints() As PointFloat)
    
    'Retrieve a copy of the current layer transform matrix
    Dim tmpMatrix As pd2DTransform
    Me.GetCopyOfLayerTransformationMatrix_Full tmpMatrix
    
    'Fill the destination with two coordinates: the center if the image, and the position of the current rotation node.
    '  Note that both if these points are generated *without* translation applied (but *with* scaling).
    dstPoints(0).x = Me.GetLayerWidth(False) * 0.5
    dstPoints(0).y = Me.GetLayerHeight(False) * 0.5
    
    'Also, note that changes to the outer node position *must be mirrored* to the checkPointOfInterest() function below!
    dstPoints(1).x = Me.GetLayerWidth(False)
    dstPoints(1).y = Me.GetLayerHeight(False) * 0.5
    
    dstPoints(2).x = Me.GetLayerWidth(False) * 0.5
    dstPoints(2).y = Me.GetLayerHeight(False)
    
    dstPoints(3).x = 0!
    dstPoints(3).y = Me.GetLayerHeight(False) * 0.5
    
    dstPoints(4).x = Me.GetLayerWidth(False) * 0.5
    dstPoints(4).y = 0!
    
    'Apply the transformation matrix to each point
    Dim i As Long
    For i = 0 To 4
        tmpMatrix.ApplyTransformToPointF dstPoints(i)
    Next i
    
    'The temporary translation matrix will automatically be cleaned up when the function exits
        
End Sub

'As of PhotoDemon 7.0, layers support one or more affine transformations (rotation, skew, etc).
' If a caller needs a copy of the layer's current transformation matrix, use this function to retrieve it.
'
'Note that scaling is not applied, by design.  The caller should handle scaling manually.
Friend Sub GetCopyOfLayerTransformationMatrix(ByRef dstMatrix As pd2DTransform, Optional ByVal ApplyTranslation As Boolean = True)

    'Create and/or reset the destination matrix, as necessary
    If (dstMatrix Is Nothing) Then Set dstMatrix = New pd2DTransform Else dstMatrix.Reset
    
    With myLayerData
    
        'If an angle is active, rotate the matrix accordingly
        If (.l_Angle <> 0#) Then dstMatrix.ApplyRotation .l_Angle, Me.GetLayerWidth(True) * .l_RotateCenterX, Me.GetLayerHeight(True) * .l_RotateCenterY
        
        'If shearing is active, shear the matrix accordingly
        If (.l_ShearX <> 0#) Or (.l_ShearY <> 0#) Then dstMatrix.ApplyShear .l_ShearX, .l_ShearY, Me.GetLayerWidth(True) * .l_RotateCenterX, Me.GetLayerHeight(True) * .l_RotateCenterY
        
        'Finally, if translation is desired, scale all points by the current layer offset
        If ApplyTranslation Then
            
            Dim tmpOffsetX As Single, tmpOffsetY As Single
            tmpOffsetX = .l_OffsetX
            tmpOffsetY = .l_OffsetY
            dstMatrix.ApplyTranslation tmpOffsetX, tmpOffsetY
            .l_OffsetX = Int(tmpOffsetX + 0.5)
            .l_OffsetY = Int(tmpOffsetY + 0.5)
            
        End If
        
    End With
    
End Sub

'As of PhotoDemon 7.0, layers support one or more affine transformations (resize, rotation, skew, etc).
' If a caller needs a copy of the layer's current transformation matrix, it can use this function to retrieve it.
Friend Sub GetCopyOfLayerTransformationMatrix_Full(ByRef dstMatrix As pd2DTransform)

    'Create and/or reset the destination matrix, as necessary
    If (dstMatrix Is Nothing) Then Set dstMatrix = New pd2DTransform Else dstMatrix.Reset
    
    With myLayerData
        
        'If the image has been non-destructively resized, scale accordingly
        If (.l_CanvasXModifier <> 1#) Or (.l_CanvasYModifier <> 1#) Then dstMatrix.ApplyScaling .l_CanvasXModifier, .l_CanvasYModifier
        
        'If an angle is active, rotate the matrix accordingly
        If (.l_Angle <> 0#) Then dstMatrix.ApplyRotation .l_Angle, Me.GetLayerWidth(True) * .l_RotateCenterX, Me.GetLayerHeight(True) * .l_RotateCenterY
        
        'If shearing is active, shear the matrix accordingly
        If (.l_ShearX <> 0#) Or (.l_ShearY <> 0#) Then dstMatrix.ApplyShear .l_ShearX, .l_ShearY, Me.GetLayerWidth(True) * .l_RotateCenterX, Me.GetLayerHeight(True) * .l_RotateCenterY
        
        'Finally, apply translation matching the current layer offset
        dstMatrix.ApplyTranslation .l_OffsetX, .l_OffsetY
        
    End With
    
End Sub

'As of PhotoDemon 7.0, layers support one or more affine transformations (rotation, skew, etc).  If the caller needs to determine
' layer boundaries with all affine transform operations considered, use this function.
Friend Sub GetLayerBoundaryRect(ByRef dstRect As RectF)
    
    'Get a copy of the current layer coordinates
    Dim layerCorners() As PointFloat
    ReDim layerCorners(0 To 3) As PointFloat
    
    Me.GetLayerCornerCoordinates layerCorners
    
    'Find the largest and smallest (x, y) coordinates
    Dim minX As Single, maxX As Single, minY As Single, maxY As Single
    FindMaxMinOfPointFs layerCorners, 4, minX, minY, maxX, maxY
    
    'The difference between max and min values define the boundary rect for this layer
    With dstRect
        .Left = minX
        .Top = minY
        .Width = maxX - minX
        .Height = maxY - minY
    End With
    
End Sub

'Generally speaking, extracting the portion of this layer's DIB relevant to the image is
' really easy - just BitBlt it directly from the layer DIB itself!  Unfortunately, PD 7.0
' greatly complicates this by introducing non-destructive affine transformations, which
' can make a layer occupy a non-rectangular portion of the image.
'
'This helper function can be used to retrieve a DIB - at 100% zoom - that contains an
' affine-transformed copy of the layer DIB.  By necessity, it will also return new (x, y)
' offsets (which are relevant for things like rotation, as the affine-transformed image
' will have negative offsets relative to the "original" layer offsets).  Note that the
' destination (x, y) values will already have been transformed to integers for you; this
' function automatically calculates offsets in a sub-pixel-friendly way.
'
'I have optimized this function as much as I physically can.  Separate paths are used
' depending on the complexity of the transform (if the only active non-destructive
' transform is scaling, we can bypass certain steps, for example).
'
'RETURNS: TRUE if the destination DIB was filled, FALSE if the DIB lies outside the crop area (e.g. entirely off the image)
Friend Function GetAffineTransformedDIB(ByRef dstDIB As pdDIB, ByRef dstX As Long, ByRef dstY As Long, Optional ByVal cropWidth As Long = 0, Optional ByVal cropHeight As Long = 0) As Boolean
    
    'Retrieve a copy of the layer corner points, in image coordinates (of course!)
    'Get a copy of the current layer coordinates
    Me.GetLayerCornerCoordinates m_PlgPoints
    
    'From our boundary points, find the largest and smallest (x, y) values.
    ' These give us a bounding box for the layer as a whole.
    Dim minX As Single, maxX As Single, minY As Single, maxY As Single
    FindMaxMinOfPointFs m_PlgPoints, 4, minX, minY, maxX, maxY
    
    'Copy those values into a RECTF struct, for convenience
    Dim layerRect As RectF
    With layerRect
        .Left = minX
        .Top = minY
        .Width = maxX - minX
        .Height = maxY - minY
    End With
    
    'We now know two crucial pieces of data:
    ' 1) the final locations of the affine-transformed layer corner points
    '     (in image coordinate space)
    ' 2) a bounding box for the affine-transformed layer
    '     (in image coordinate space)
    
    'Next, we need to calculate the minimum boundary size required by the destination DIB.
    ' This is the lesser of two things:
    ' 1) A crop width/height supplied by the caller, if any (typically the size of the
    '     compositing surface itself; pixels outside this are irrelevant, so we can
    '     completely ignore them for a speed boost).
    ' 2) If the layer left/top values are negative, the max x/y values calculated above
    ' 3) If the layer left/top values are positive, the width of the layer boundary rect
    
    'During this step, we will also calculate fractional values I call "normalizeX/Y".
    ' PD's internal compositor can only render to integer offsets, for performance reasons.
    ' Affine transformations almost always result in an image with fractional offsets;
    ' to make sure these offsets are preserved, we "normalize" the coordinates by offsetting
    ' them by their fractional amount.  This ensures that the final composite looks identical
    ' between GDI+ and PD's internal compositing code.
    
    'Start by calculating (2) or (3), and padding it to account for fractional offsets
    Dim intLayerWidth As Long, intLayerHeight As Long
    Dim normalizeX As Single, normalizeY As Single
    
    If (minX < 0) Then
        dstX = 0
        normalizeX = 0!
        intLayerWidth = Int(maxX + 0.9999)
    Else
        dstX = Int(minX)
        normalizeX = minX - dstX
        intLayerWidth = Int(layerRect.Width + 0.9999 + normalizeX)
    End If
    
    If (minY < 0) Then
        dstY = 0
        normalizeY = 0!
        intLayerHeight = Int(maxY + 0.9999)
    Else
        dstY = Int(minY)
        normalizeY = minY - dstY
        intLayerHeight = Int(layerRect.Height + 0.9999 + normalizeY)
    End If
    
    'Next, replace the results of (2) or (3) with (1), as relevant
    If (cropWidth > 0) Then
        If (dstX + intLayerWidth > cropWidth) Then intLayerWidth = cropWidth - dstX
    End If
                    
    If (cropHeight > 0) Then
        If (dstY + intLayerHeight > cropHeight) Then intLayerHeight = cropHeight - dstY
    End If
    
    'If the final calculated width and height are positive and non-zero, we can continue; otherwise, exit now.
    If (intLayerWidth > 0) And (intLayerHeight > 0) Then
    
        'Next, prepare the destination DIB.  If at all possible, we try to optimize this by reusing DIBs that are already
        ' an acceptable size.
        If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
        If (dstDIB.GetDIBWidth <> intLayerWidth) Or (dstDIB.GetDIBHeight <> intLayerHeight) Then
            dstDIB.CreateBlank intLayerWidth, intLayerHeight, 32, 0, 0
            dstDIB.SetInitialAlphaPremultiplicationState True
        Else
            dstDIB.ResetDIB
        End If
        
        'It is now time to render the affine-transformed DIB.  If normalization factors or custom destination offsets are active,
        ' we must modify our corner points accordingly, so that they are 0-based.
        Dim i As Long
        For i = 0 To 3
            m_PlgPoints(i).x = m_PlgPoints(i).x + (normalizeX - dstX)
            m_PlgPoints(i).y = m_PlgPoints(i).y + (normalizeY - dstY)
        Next i
        
        'Copy a resized chunk of the source image onto the temporary DIB.  (Note that this branches according to the
        ' presence of more complicated affine functions, which require a parallelogram-style blit process).
        If Me.AffineTransformsActive(False) Then
            GDI_Plus.GDIPlus_PlgBlt dstDIB, m_PlgPoints, m_LayerDIB, 0, 0, Me.GetLayerWidth(False), Me.GetLayerHeight(False), 1, Me.GetLayerResizeQuality_GDIPlus
        Else
            GDI_Plus.GDIPlus_StretchBlt dstDIB, layerRect.Left + (normalizeX - dstX), layerRect.Top + (normalizeY - dstY), Me.GetLayerWidth(True), Me.GetLayerHeight(True), m_LayerDIB, 0, 0, Me.GetLayerWidth(False), Me.GetLayerHeight(False), 1, Me.GetLayerResizeQuality_GDIPlus
        End If
        
        dstDIB.SetInitialAlphaPremultiplicationState True
        
        'Return success
        GetAffineTransformedDIB = True
        
    Else
    
        'Return failure (e.g. the transformed DIB lies completely off the image, so there is no need for a temporary DIB)
        GetAffineTransformedDIB = False
        
    End If
    
End Function

'If one or more affine transformations are active (rotation, skew, etc), you can call this function to retrieve three corner points
' of the layer *IN IMAGE COORDINATES*.

'Given an array with at least upper bound 2, find the max and min x/y values.
Private Sub FindMaxMinOfPointFs(ByRef srcPoints() As PointFloat, ByVal numOfPoints As Long, ByRef minX As Single, ByRef minY As Single, ByRef maxX As Single, ByRef maxY As Single)
    
    If (srcPoints(0).x < srcPoints(1).x) Then
        minX = srcPoints(0).x
        maxX = srcPoints(1).x
    Else
        minX = srcPoints(1).x
        maxX = srcPoints(0).x
    End If
    
    If (srcPoints(0).y < srcPoints(1).y) Then
        minY = srcPoints(0).y
        maxY = srcPoints(1).y
    Else
        minY = srcPoints(1).y
        maxY = srcPoints(0).y
    End If
    
    Dim i As Long
    For i = 2 To numOfPoints - 1
    
        If (srcPoints(i).x < minX) Then
            minX = srcPoints(i).x
        ElseIf (srcPoints(i).x > maxX) Then
            maxX = srcPoints(i).x
        End If
        
        If (srcPoints(i).y < minY) Then
            minY = srcPoints(i).y
        ElseIf (srcPoints(i).y > maxY) Then
            maxY = srcPoints(i).y
        End If
        
    Next i
    
End Sub

Private Sub Class_Initialize()

    'Assign default values to this instance
    With myLayerData
        .l_Name = g_Language.TranslateMessage("New layer")
        .l_GroupID = 0
        .l_Opacity = 100!
        .l_BlendMode = BM_Normal
        .l_OffsetX = 0
        .l_OffsetY = 0
        .l_CanvasXModifier = 1
        .l_CanvasYModifier = 1
        .l_Angle = 0
        .l_RotateCenterX = 0.5
        .l_RotateCenterY = 0.5
        .l_Visibility = True
        .l_ResizeQuality = LRQ_Bilinear
        .l_ShearX = 0
        .l_ShearY = 0
        .l_AlphaMode = AM_Normal
        .l_FrameTimeInMS = 0
        .l_MaskExists = False
        .l_MaskActive = False
    End With
    
    With myVectorData
        .vd_Width = 0
        .vd_Height = 0
    End With
    
    'Initialize a default text renderer, regardless of layer type
    Set myTextRenderer = New pdTextRenderer
    
    'Initialize the layer's DIB
    Set m_LayerDIB = New pdDIB
    
    'Initialize all temporary LOD compositing DIBs and corresponding hashes
    ReDim m_ViewportStateHashes(0 To NUM_OF_LOD_CACHES - 1) As Long
    ReDim m_DIBChangedSinceLastHash(0 To NUM_OF_LOD_CACHES - 1) As Boolean
    
    ReDim m_lodDIB(0 To NUM_OF_LOD_CACHES - 1) As pdDIB
    Dim i As Long
    For i = 0 To NUM_OF_LOD_CACHES - 1
        Set m_lodDIB(i) = New pdDIB
    Next i
    
    'Initialize the layer thumbnail and note that it is not clean, which will force a regeneration when the thumbnail is next requested.
    Set m_LayerThumbnail = New pdDIB
    m_IsThumbClean = False
    
    'Set the canonical ID to -1.  This can be used to determine if the layer has been activated.
    m_layerID = -1
    
    'By default, vector layers do not have a backer DIB generated
    m_VectorDIBSynched = False
    
    'Prep a default set of parallelogram corner points, in case the user decides to activate non-destructive transforms
    ReDim m_PlgPoints(0 To 3) As PointFloat
    
End Sub

Private Sub Class_Terminate()
    
    Set m_LayerDIB = Nothing
    Set m_LayerThumbnail = Nothing
    Set myTextRenderer = Nothing
    
    Dim i As Long
    For i = 0 To NUM_OF_LOD_CACHES - 1
        Set m_lodDIB(i) = Nothing
    Next i
    
End Sub

'This function initializes a layer header from an XML-format string created by GetLayerHeaderAsXML().
'
'Note that if the optional createNonDestructively parameter is FALSE (as it is by default), this function
' will create a blank DIB for the layer DIB, using the dimensions and color depth stored in the XML string.
' However, it will obviously NOT apply any stored image data.  That must be handled separately, as the
' format of the underlying layer's data varies by layer type.  (Raster layers store their data as a raw
' byte stream, for example, while vector and text layers store XML data and we simply recreate the raster
' copy at run-time.)
'
'Note also that this function will try to match color-managed layers against matching ICC profiles in the
' central ColorManagement cache.  Make *sure* you have added all embedded color profiles to the central
' cache before invoking this function.
Friend Function CreateNewLayerFromXML(ByRef xmlString As String, Optional ByVal useCustomLayerID As Long = -1, Optional ByVal createNonDestructively As Boolean = False) As Boolean

    'Prep a fast XML param engine, which greatly simplifies the process of assembling XML data
    Dim xmlEngine As pdSerialize
    Set xmlEngine = New pdSerialize
    xmlEngine.SetParamString xmlString
    
    'Validate the XML header...
    If (xmlEngine.DoesParamExist("photodemon-layer") And (xmlEngine.GetLong("layer-version", 0) >= 1)) Then
        
        'If the caller has not specified their own layer ID, use the one from file.
        If (useCustomLayerID = -1) Then m_layerID = xmlEngine.GetLong("layer-id") Else m_layerID = useCustomLayerID
        
        'Read in the layer type.  Remember that this value is immutable, and cannot be changed
        ' for the life of the layer.
        m_LayerType = Layers.GetLayerTypeIDFromString(xmlEngine.GetString("type", Layers.GetLayerTypeStringFromID(PDL_Image), True))
        
        'Read in the contents of the LayerData header.  (NOTE: this must be manually modified if new entries are
        ' added to the LayerData type.
        With myLayerData
            .l_Name = xmlEngine.GetString("name", vbNullString, True)
            .l_GroupID = xmlEngine.GetLong("group-id", 0, True)
            .l_Visibility = xmlEngine.GetBool("visible", True, True)
            .l_Opacity = xmlEngine.GetSingle("opacity", 100!, True)
            .l_BlendMode = Colors.GetBlendModeIDFromString(xmlEngine.GetString("blend-mode", Colors.GetBlendModeStringFromID(BM_Normal), True))
            .l_AlphaMode = Colors.GetAlphaModeIDFromString(xmlEngine.GetString("alpha-mode", Colors.GetAlphaModeStringFromID(AM_Normal), True))
            .l_OffsetX = xmlEngine.GetLong("offset-x", 0&, True)
            .l_OffsetY = xmlEngine.GetLong("offset-y", 0&, True)
            .l_CanvasXModifier = xmlEngine.GetDouble("stretch-x", 1#, True)
            .l_CanvasYModifier = xmlEngine.GetDouble("stretch-y", 1#, True)
            .l_ShearX = xmlEngine.GetDouble("shear-x", 0#, True)
            .l_ShearY = xmlEngine.GetDouble("shear-y", 0#, True)
            .l_Angle = xmlEngine.GetDouble("angle", 0#, True)
            .l_RotateCenterX = xmlEngine.GetDouble("rotation-center-x", 0.5, True)
            .l_RotateCenterY = xmlEngine.GetDouble("rotation-center-y", 0.5, True)
            .l_ResizeQuality = Layers.GetLayerResizeIDFromString(xmlEngine.GetString("sample-quality", Layers.GetLayerResizeStringFromID(LRQ_Bilinear), True))
            .l_FrameTimeInMS = xmlEngine.GetLong("frame-time-ms", 0, True)
            .l_MaskExists = xmlEngine.GetBool("mask-exists", False, True)
            
            'If a mask exists, initialize the underlying mask class, but we need to wait for a later step
            ' to actually fill the mask (as its full data is not crammed into this XML packet)
            If .l_MaskExists Then
                .l_MaskActive = xmlEngine.GetBool("mask-active", False, True)
                Set m_LayerMask = New pdLayerMask
                'TODO: un-serialize mask?  IDK, serializing the full mask to a string is stupid
            Else
                .l_MaskActive = False
                Set m_LayerMask = Nothing
            End If
            
        End With
        
        'Next, we're going to do something kinda weird.  We're going to create a blank surface
        ' for the layer contents using data supplied by this XML packet, but we're *not* going to
        ' fill the surface yet!  An external function handles that step, because the actual surface
        ' pixel data is (likely) stored in a heavily compressed binary format.
        '
        'Note that this behavior can be overridden by setting the optional createNonDestructively
        ' parameter to TRUE.  This special case is used by the Undo/Redo engine to re-load an
        ' image's header data, without actually overwriting the current layer's backing surface.
        ' (This produces very small undo/redo files when e.g. layer visibility is toggled.)
        If (Not createNonDestructively) Then
        
            Set m_LayerDIB = New pdDIB
            With xmlEngine
                
                'Validate width/height and other layer parameters
                Dim srfColorDepth As Long, srfColorDetails As String, srfWidth As Long, srfHeight As Long, srfStride As Long
                srfColorDepth = .GetLong("px-color-depth", 0, True)
                srfColorDetails = .GetString("px-color-details", vbNullString, True)
                
                'While we allow 24-bpp surfaces here, PD has only ever used 32-bpp surfaces, and it will only
                ' ever write 32-bpp surfaces.  I'm not even sure that 24-bpp surfaces would work once loaded
                ' (as a lot of internal functions assume the presence of an alpha channel), so this check
                ' probably needs to be updated after actual testing.  TODO!
                If ((srfColorDepth = 32) Or (srfColorDepth = 24)) And (LCase$(Left$(srfColorDetails, 6)) = "b8g8r8") Then
                
                    'Retrieve dimensions (including stride) and validate accordingly
                    srfWidth = .GetLong("px-width", 0, True)
                    srfHeight = .GetLong("px-height", 0, True)
                    srfStride = .GetLong("px-stride", 0, True)
                    
                    If (srfWidth > 0) And (srfHeight > 0) And (srfStride > 0) Then
                    
                        'Also grab alpha premultiplication; at present, PD *always* writes premultiplied alpha,
                        ' but we cover the case of un-premultiplied alpha in case we decide to revert this
                        ' decision in the future.
                        Dim srfAlphaPremult As Boolean
                        srfAlphaPremult = .GetBool("px-alpha-premultiplied", True, True)
                        
                        'Create the surface!
                        m_LayerDIB.CreateBlank srfWidth, srfHeight, srfColorDepth, 0, 0
                        
                        'Mark alpha premultiplication in advance
                        m_LayerDIB.SetInitialAlphaPremultiplicationState srfAlphaPremult
                        
                        'Return success; note that we create a backing surface even for vector layers,
                        ' as it's required for features like adjustment layers that sit above this one.
                        CreateNewLayerFromXML = True
                        
                    Else
                        InternalError "CreateNewLayerFromXML", "bad dimensions: " & srfWidth & "x" & srfHeight & ", " & srfStride
                    End If
                    
                Else
                    InternalError "CreateNewLayerFromXML", "bad color-depth or details: " & srfColorDepth & ", " & srfColorDetails
                End If
                
            End With
            
            'Note that all internal layer caches are no longer valid
            Me.NotifyOfDestructiveChanges
            
        Else
            CreateNewLayerFromXML = True
        End If
        
    'If the layer XML data can't be validated, try the legacy serialization engine
    Else
        Set xmlEngine = Nothing
        PDDebug.LogAction "pdLayer.CreateNewLayerFromXML() is trying again using the legacy serialization engine..."
        CreateNewLayerFromXML = CreateNewLayerFromXML_Legacy(xmlString, useCustomLayerID, createNonDestructively)
    End If
    
End Function

'Create a copy of all critical layer header data in an XML-format string.  All data necessary to
' recreate this layer from scratch must be included in this string, EXCEPT FOR PIXEL DATA.
' (For obvious performance reasons, the pixel contents of this layer must be handled separately,
' in a dedicated binary format.)
Friend Function GetLayerHeaderAsXML() As String
    
    'Prep a fast XML param engine, which greatly simplifies the process of assembling XML data
    Dim xmlEngine As pdSerialize
    Set xmlEngine = New pdSerialize
    
    'Add a basic header and version info
    xmlEngine.AddXMLString "<photodemon-layer>"
    xmlEngine.AddParam "layer-version", 1&, True, True
    
    'Start by writing out the hard-coded layer ID.
    xmlEngine.AddParam "layer-id", m_layerID, True, True
    
    'Next, write out the layer type (raster, vector, etc).  Remember that this value is immutable,
    ' and cannot be changed for the life of the layer.  (If you want to change it, you must perform
    ' a (usually) one-way conversion to a new format.)
    xmlEngine.AddParam "type", Layers.GetLayerTypeStringFromID(m_LayerType), True, True
    
    'Next, write out all entries in this layer's LayerData header.  (NOTE: new layer properties must
    ' ALWAYS be manually added to this segment.)
    With myLayerData
        xmlEngine.AddParam "name", .l_Name, True, True
        xmlEngine.AddParam "group-id", .l_GroupID, True, True
        xmlEngine.AddParam "visible", .l_Visibility, True, True
        xmlEngine.AddParam "opacity", .l_Opacity, True, True
        xmlEngine.AddParam "blend-mode", Colors.GetBlendModeStringFromID(.l_BlendMode), True, True
        xmlEngine.AddParam "alpha-mode", Colors.GetAlphaModeStringFromID(.l_AlphaMode), True, True
        xmlEngine.AddParam "offset-x", .l_OffsetX, True, True
        xmlEngine.AddParam "offset-y", .l_OffsetY, True, True
        xmlEngine.AddParam "stretch-x", .l_CanvasXModifier, True, True
        xmlEngine.AddParam "stretch-y", .l_CanvasYModifier, True, True
        xmlEngine.AddParam "shear-x", .l_ShearX, True, True
        xmlEngine.AddParam "shear-y", .l_ShearY, True, True
        xmlEngine.AddParam "angle", .l_Angle, True, True
        xmlEngine.AddParam "rotation-center-x", .l_RotateCenterX, True, True
        xmlEngine.AddParam "rotation-center-y", .l_RotateCenterY, True, True
        xmlEngine.AddParam "sample-quality", Layers.GetLayerResizeStringFromID(.l_ResizeQuality), True, True
        xmlEngine.AddParam "frame-time-ms", .l_FrameTimeInMS, True, True
        xmlEngine.AddParam "mask-exists", .l_MaskExists, True, True
        xmlEngine.AddParam "mask-active", .l_MaskActive, True, True
    End With
    
    'Next, write out enough information for us to reconstruct the underlying pixel data for this layer
    ' (if any).  Note that we always write these values, even for a vector layer,
    With xmlEngine
        .AddParam "px-color-depth", m_LayerDIB.GetDIBColorDepth, True, True
        If (m_LayerDIB.GetDIBColorDepth = 32) Then
            .AddParam "px-color-details", "b8g8r8a8", True, True
        Else
            .AddParam "px-color-details", "b8g8r8", True, True
        End If
        .AddParam "px-width", m_LayerDIB.GetDIBWidth, True, True
        .AddParam "px-height", m_LayerDIB.GetDIBHeight, True, True
        .AddParam "px-stride", m_LayerDIB.GetDIBStride, True, True
        .AddParam "px-alpha-premultiplied", m_LayerDIB.GetAlphaPremultiplication, True, True
    End With
    
    'Close the initial pdLayer tag
    xmlEngine.AddXMLString "</photodemon-layer>"
    
    'The layer XML string is now complete.  Return it.
    GetLayerHeaderAsXML = xmlEngine.GetParamString()
    
End Function

'Write a copy of a vector layer's contents to an XML-format string.  All data necessary to recreate
' a vector layer from scratch must be included in this string.  For obvious performance reasons,
' the pixel contents of this layer are not part of this function.  PD regenerates the actual backing
' surface of vector layers from scratch, at load-time.
Friend Function GetVectorDataAsXML() As String
    
    'First, make sure this is not an image layer
    If Not Me.IsLayerVector() Then
        InternalError "GetVectorDataAsXML", "can't get vector data for raster layers"
        Exit Function
    End If
    
    'Prepare an XML engine, which greatly simplifies the process of assembling XML data
    Dim xmlEngine As pdSerialize
    Set xmlEngine = New pdSerialize
    xmlEngine.AddXMLString "<photodemon-vector-object>"
    xmlEngine.AddParam "vector-version", 1&, True, True
    
    'All vector layers store certain data attributes
    xmlEngine.AddParam "bounding-width", myVectorData.vd_Width, True, True
    xmlEngine.AddParam "bounding-height", myVectorData.vd_Height, True, True
    
    'Additional vector data varies by layer type
    Select Case m_LayerType
        
        'pdTextRenderer (used for both basic and advanced text layers) handles XML serialization on its own
        Case PDL_TextBasic, PDL_TextAdvanced
            xmlEngine.AddParam "vector-text-data", myTextRenderer.GetAllFontSettingsAsXML(), True
        
    End Select
    
    xmlEngine.AddXMLString "</photodemon-vector-object>"
        
    'The layer XML string is now complete.  Return it.
    GetVectorDataAsXML = xmlEngine.GetParamString()

End Function

'As the sister function to GetVectorDataAsXML(), this function initializes a
' vector layer's contents from an XML-format string.
Friend Function SetVectorDataFromXML(ByRef xmlString As String) As Boolean

    'Prepare an XML engine, which greatly simplifies the process of parsing XML data
    Dim xmlEngine As pdSerialize
    Set xmlEngine = New pdSerialize
    xmlEngine.SetParamString xmlString
    
    'If basic validation fails, revert to the legacy font serialization engine
    Dim useLegacyLoader As Boolean
    useLegacyLoader = (Not xmlEngine.DoesParamExist("photodemon-vector-object", True))
    If (Not useLegacyLoader) Then useLegacyLoader = (xmlEngine.GetLong("vector-version", 0, True) < 1)
    
    If useLegacyLoader Then
        Set xmlEngine = Nothing
        SetVectorDataFromXML = SetVectorDataFromXML_Legacy(xmlString)
    Else
        
        'All vector layers store some similar data
        With myVectorData
            .vd_Width = xmlEngine.GetLong("bounding-width", 1, True)
            .vd_Height = xmlEngine.GetLong("bounding-height", 1, True)
        End With
        
        'Next, the vector tags we look for vary by layer type
        Select Case m_LayerType
            
            'pdTextRenderer handles its own XML serialization, thankfully
            Case PDL_TextBasic, PDL_TextAdvanced
                myTextRenderer.SetAllFontSettingsFromXML xmlEngine.GetString("vector-text-data", vbNullString, True)
                
        End Select
        
        'Note that all internal layer caches are no longer valid; this will also cause the raster cache to be marked unclean
        Me.NotifyOfDestructiveChanges
        SetVectorDataFromXML = True
    
    End If
    
End Function

'Initialize a new layer.  Optional inputs include:
' - Layer name
' - For image-type layers, a DIB can be passed; the layer's raster contents will be initialized to match the DIB.
'   (By default, this function will simply set a reference to the passed DIB; if you want a full clone, specify it
'    via the final optional parameter.)
Friend Sub InitializeNewLayer(ByVal newLayerType As PD_LayerType, Optional ByVal newLayerName As String = vbNullString, Optional ByRef srcDIB As pdDIB = Nothing, Optional ByVal cloneSrcDIBLocally As Boolean = False)

    'Mark the layer type.
    m_LayerType = newLayerType
    
    'Copy the name locally, and set other values to their natural defaults
    With myLayerData
        If (LenB(newLayerName) <> 0) Then
            .l_Name = newLayerName
        Else
            .l_Name = g_Language.TranslateMessage("New layer")
        End If
        .l_GroupID = 0
        .l_Opacity = 100!
        .l_BlendMode = BM_Normal
        .l_Visibility = True
        .l_ResizeQuality = LRQ_Bilinear
        .l_Angle = 0#
        .l_RotateCenterX = 0.5
        .l_RotateCenterY = 0.5
        .l_ShearX = 0#
        .l_ShearY = 0#
        .l_AlphaMode = AM_Normal
        .l_FrameTimeInMS = 0
        .l_MaskExists = False
        .l_MaskActive = False
        Set m_LayerMask = Nothing
    End With
    
    'Reset any vector-specific data
    With myVectorData
        .vd_Width = 0
        .vd_Height = 0
    End With
    
    m_VectorDIBSynched = False
    
    'If passed, create a local copy of the passed DIB.
    If cloneSrcDIBLocally Then
        Set m_LayerDIB = New pdDIB
        If (Not srcDIB Is Nothing) Then m_LayerDIB.CreateFromExistingDIB srcDIB
    Else
        If (Not srcDIB Is Nothing) Then Set m_LayerDIB = srcDIB Else Set m_LayerDIB = New pdDIB
    End If
    
    'Keep GDI object count down as much as possible
    m_LayerDIB.FreeFromDC
    
    'Set default offsets
    myLayerData.l_OffsetX = 0
    myLayerData.l_OffsetY = 0
    
    'In the future, we might want to set an intial offset relative to the parent image's viewport settings...?
    
    'Note that all internal layer caches are no longer valid.  This is particularly important for vector layers, which generate a raster copy
    ' of their contents "on the fly"
    Me.NotifyOfDestructiveChanges

End Sub

'Copy an existing layer.  All layer contents will be copied manually, so make sure that new layer properties are
' manually added to this function!
Friend Sub CopyExistingLayer(ByRef srcLayer As pdLayer)

    'Copy all generic layer properties from the source layer.  (Note that canonical layer ID is *not* copied; that must always
    ' be unique for every created layer!)
    m_LayerType = srcLayer.GetLayerType
    
    With myLayerData
        .l_Name = srcLayer.GetLayerName()
        .l_GroupID = srcLayer.GetLayerGroup()
        .l_Opacity = srcLayer.GetLayerOpacity()
        .l_BlendMode = srcLayer.GetLayerBlendMode()
        .l_OffsetX = srcLayer.GetLayerOffsetX()
        .l_OffsetY = srcLayer.GetLayerOffsetY()
        .l_CanvasXModifier = srcLayer.GetLayerCanvasXModifier()
        .l_CanvasYModifier = srcLayer.GetLayerCanvasYModifier()
        .l_Angle = srcLayer.GetLayerAngle()
        .l_RotateCenterX = srcLayer.GetLayerRotateCenterX()
        .l_RotateCenterY = srcLayer.GetLayerRotateCenterY()
        .l_Visibility = srcLayer.GetLayerVisibility()
        .l_ResizeQuality = srcLayer.GetLayerResizeQuality()
        .l_ShearX = srcLayer.GetLayerShearX()
        .l_ShearY = srcLayer.GetLayerShearY()
        .l_AlphaMode = srcLayer.GetLayerAlphaMode()
        .l_FrameTimeInMS = srcLayer.GetLayerFrameTimeInMS()
        .l_MaskExists = srcLayer.GetGenericLayerProperty(pgp_MaskExists)
        
        'Only retrieve a mask if the source layer has one
        If .l_MaskExists Then
            .l_MaskActive = srcLayer.GetGenericLayerProperty(pgp_MaskActive)
            Set m_LayerMask = New pdLayerMask
            m_LayerMask.CloneExistingMask srcLayer.GetLayerMask
        Else
            .l_MaskActive = False
            Set m_LayerMask = Nothing
        End If
        
    End With
    
    'Copy the source layer's DIB, if one exists.
    If Not (srcLayer.GetLayerDIB Is Nothing) Then Me.GetLayerDIB.CreateFromExistingDIB srcLayer.GetLayerDIB
    
    'Depending on the type of layer we are copying, copy any extra layer data, or generate a new layer mask to
    ' match the layer's vector contents.
    Select Case srcLayer.GetLayerType
    
        Case PDL_Image
        
        'Vector layers copy additional vector-specific data
        Case PDL_TextBasic, PDL_TextAdvanced
            myVectorData.vd_Width = srcLayer.GetLayerWidth
            myVectorData.vd_Height = srcLayer.GetLayerHeight
            
            'Other vector-specific data can be transferred via XML, saving us the hassle of explicitly copying it here
            Me.SetVectorDataFromXML srcLayer.GetVectorDataAsXML()
        
        Case PDL_Adjustment
    
    End Select
    
    'Note that all internal layer caches are no longer valid
    Me.NotifyOfDestructiveChanges
    
End Sub

'Assuming the current layer DIB is null-padded (a term I use to describe a layer that has been forced to the size of the
' image by surrounding it with transparent pixels), intelligently crop the DIB by removing all fully transparent padding.
' When this is done, modify the current layer offsets so that the image stays in effectively the same place, but without
' all that blank padding.
'
'IMPORTANT NOTE: if you call this function without first null-padding the layer DIB, weird shit may happen.  That is not
'                 an intended use-case, so don't attempt it!
'
'Returns TRUE if layer was trimmed successfully, FALSE if it was not.  (False occurs if the entire layer is transparent.)
Friend Function CropNullPaddedLayer() As Boolean
    
    'Make sure the source DIB isn't empty; (this should never happen, but I'm trying to be better about catching
    ' edge error conditions).
    If (m_LayerDIB.GetDIBDC <> 0) And (m_LayerDIB.GetDIBWidth <> 0) And (m_LayerDIB.GetDIBHeight <> 0) And (m_LayerDIB.GetDIBColorDepth = 32) Then
    
        'Point an array at the layer's DIB data
        Dim srcImageData() As Byte, srcSA As SafeArray2D
        m_LayerDIB.WrapArrayAroundDIB srcImageData, srcSA
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = m_LayerDIB.GetDIBWidth - 1
        finalY = m_LayerDIB.GetDIBHeight - 1
        
        'Each edge will be scanned independently, and these variables will be calculated with the new bounding
        ' RECT for this layer.
        Dim newTop As Long, newBottom As Long, newLeft As Long, newRight As Long
        
        'Before doing anything else, perform a quick sanity check of the four image corners.  If none of them are
        ' fully transparent, we know the layer is untrimmable (because interesting pixel data extends to all sides),
        ' so we can exit now without actually having to perform a full scan.
        If (srcImageData(3, 0) <> 0) And (srcImageData(finalX * 4 + 3, 0) <> 0) And (srcImageData(3, finalY) <> 0) And (srcImageData(finalX * 4 + 3, finalY) <> 0) Then
            
            'Release our array pointer.  (This MUST be done for all function exits, or VB will crash.)
            m_LayerDIB.UnwrapArrayFromDIB srcImageData
            Debug.Print "Quick check of layer corners shows no trimmable areas; premature exit initialized."
            CropNullPaddedLayer = True
            Exit Function
        
        End If
        
        'First, scan the top of the image.
        
        'All edges follow the same formula, so I'm only providing detailed commenting in this first section.
        
        'Unlike PD's "AutoCrop" function, we're only looking to eliminate completely transparent edges in this function.
        ' That greatly simplifies (and accelerates) the process of determining if an edge row or column is "blank".
        
        'If a non-transparent pixel is found, this will be set to TRUE.  We must track failure state so that we can exit
        ' both the x and y For loops.
        Dim pixelFails As Boolean
        pixelFails = False
        
        'Scan the image, starting at the top-left and moving right
        For y = 0 To finalY
        For x = 0 To finalX
            
            'If this pixel is not 100% transparent, we don't want to crop it.  Exit now.
            If (srcImageData(x * 4 + 3, y) <> 0) Then
                pixelFails = True
                Exit For
            End If
            
        Next x
            If pixelFails Then Exit For
        Next y
    
        'We have now reached one of two conditions:
        '1) The entire layer DIB is transparent
        '2) The loop progressed part-way through the layer, and terminated once it found a non-transparent pixel
    
        'Check for case (1) and exit if it occurred
        If (Not pixelFails) Then
            m_LayerDIB.UnwrapArrayFromDIB srcImageData
            Debug.Print "Request for layer null padding trim denied - entire layer is transparent, so crop is impossible!"
            CropNullPaddedLayer = False
            Exit Function
        
        'Next, check for case (2)
        Else
            newTop = y
        End If
        
        'To save us a bit of time, subsequent scans will only be performed from the image's cropped top location.
        initY = newTop
        
        'Next, repeat the steps above for the bottom of the image.
        pixelFails = False
        
        For y = finalY To initY Step -1
        For x = 0 To finalX - 1
            
            If (srcImageData(x * 4 + 3, y) <> 0) Then
                pixelFails = True
                Exit For
            End If
            
        Next x
            If pixelFails Then Exit For
        Next y
        
        'As before, we only want to scan pixels that lie inside the detected trim rect.  (Efficiency is key for this function,
        ' as it is used so frequently throughout the program!)
        newBottom = y
        finalY = newBottom
        
        'Repeat the above steps, but tracking the left edge instead.  Note also that we will only be scanning from wherever
        ' the top and bottom border detections arrived (to save processing time).
        pixelFails = False
        Dim xStride As Long
        
        For x = 0 To finalX
            xStride = x * 4 + 3
        For y = initY To finalY
        
            If (srcImageData(xStride, y) <> 0) Then
                pixelFails = True
                Exit For
            End If
            
        Next y
            If pixelFails Then Exit For
        Next x
    
        newLeft = x
        
        'Repeat the above steps, but tracking the right edge instead.  Note also that we will only be scanning from wherever
        ' the top crop failed (to save processing time).
        pixelFails = False
        
        For x = finalX To newLeft Step -1
            xStride = x * 4 + 3
        For y = initY To finalY
        
            If (srcImageData(xStride, y) <> 0) Then
                pixelFails = True
                Exit For
            End If
            
        Next y
            If pixelFails Then Exit For
        Next x
        
        newRight = x
        
        'Safely deallocate imageData()
        m_LayerDIB.UnwrapArrayFromDIB srcImageData
        
        'If we made it all the way here, our search for null borders completed successfully.
        
        'If all border values are 0, return FALSE.  This basically means that the interesting bits of the layer
        ' extend to the full size of the image, so we can't trim it.
        If (newLeft = 0) And (newTop = 0) And (newRight = 0) And (newBottom = 0) Then
            CropNullPaddedLayer = False
        Else
            
            'Trimmable borders were found.  Use the calculated values to crop the image to its smallest relevant size.
            Dim tmpCropDIB As pdDIB
            Set tmpCropDIB = New pdDIB
            tmpCropDIB.CreateBlank newRight - newLeft + 1, newBottom - newTop + 1, 32, 0
            GDI.BitBltWrapper tmpCropDIB.GetDIBDC, 0, 0, newRight - newLeft + 1, newBottom - newTop + 1, m_LayerDIB.GetDIBDC, newLeft, newTop, vbSrcCopy
            tmpCropDIB.SetInitialAlphaPremultiplicationState m_LayerDIB.GetAlphaPremultiplication
            
            'Want to know what was cropped?  Use this line to report:
            'Debug.Print "Crop rect: (" & newLeft & "," & newTop & ") - (" & newRight & "," & newBottom & ")"
            
            'Replace the current layer DIB with our temporary DIB
            Set m_LayerDIB = tmpCropDIB
            
            'Modify our layer offsets to reflect our new size
            myLayerData.l_OffsetX = newLeft
            myLayerData.l_OffsetY = newTop
            
            'TODO: synchronize layer mask here, if one exists
            
            'Return TRUE, so that the calling function knows to redraw the screen with the new size and offsets
            CropNullPaddedLayer = True
            
        End If
    
    'Failsafe check for nonexistent layer DIB
    Else
        Debug.Print "Request for layer null padding trim denied - layer DIB does not currently exist!"
        CropNullPaddedLayer = False
    End If
    
End Function

'Assuming the current layer DIB is null-padded (a term I use to describe a layer that has been forced to the size of the
' image by surrounding it with transparent pixels), crop the DIB to some arbitrary rect.  Layer offsets and size will be
' automatically calculated.
'
'IMPORTANT NOTE: if you call this function without first null-padding the layer DIB, weird shit may happen.  That is not
'                 an intended use-case, so don't attempt it!
'
'Returns TRUE if layer was trimmed successfully, FALSE if it was not.  (False should only occur if an invalid rect
' is passed, or the layer wasn't null-padded first.)
Friend Function CropNullPaddedLayerToRect(ByRef srcRectF As RectF) As Boolean
    
    'Make sure the source DIB isn't empty; (this should never happen, but I'm trying to be better about catching
    ' edge error conditions).
    If (m_LayerDIB.GetDIBDC = 0) Or (m_LayerDIB.GetDIBWidth <= 0) Or (m_LayerDIB.GetDIBHeight <= 0) Or (m_LayerDIB.GetDIBColorDepth <> 32) Then
        PDDebug.LogAction "Request for layer null padding trim denied - layer DIB does not currently exist!"
        CropNullPaddedLayerToRect = False
        Exit Function
    End If
    
    'Validate the incoming rect
    If (srcRectF.Width <= 0!) Or (srcRectF.Height <= 0!) Then
        PDDebug.LogAction "Invalid rect passed to CropNullPaddedLayerToRect"
        CropNullPaddedLayerToRect = False
        Exit Function
    End If
    
    'The target rect appears valid.  Use its values to crop the image to the relevant size.
    Dim newWidth As Long, newHeight As Long
    newWidth = Int(srcRectF.Width + 0.5!)
    newHeight = Int(srcRectF.Height + 0.5!)
    
    Dim tmpCropDIB As pdDIB
    Set tmpCropDIB = New pdDIB
    tmpCropDIB.CreateBlank newWidth, newHeight, 32, 0, 0
    GDI.BitBltWrapper tmpCropDIB.GetDIBDC, 0, 0, newWidth, newHeight, m_LayerDIB.GetDIBDC, Int(srcRectF.Left + 0.5!), Int(srcRectF.Top + 0.5!), vbSrcCopy
    tmpCropDIB.SetInitialAlphaPremultiplicationState m_LayerDIB.GetAlphaPremultiplication
    
    'Replace the current layer DIB with our temporary DIB
    Set m_LayerDIB = tmpCropDIB
    
    'Modify our layer offsets to reflect our new size
    myLayerData.l_OffsetX = Int(srcRectF.Left + 0.5!)
    myLayerData.l_OffsetY = Int(srcRectF.Top + 0.5!)
    
    'TODO: synchronize layer mask here, if one exists
    
    'Return TRUE, so that the calling function knows to redraw the screen with the new size and offsets
    CropNullPaddedLayerToRect = True
    
End Function

'Null-pad this layer (with transparent pixels) against the size of our parent pdImage.  Note that the parent pdImage's dimensions
' must be explicitly passed, because layers do not maintain a link to their parent.
Friend Sub ConvertToNullPaddedLayer(ByVal parentImageWidth As Long, ByVal parentImageHeight As Long, Optional ByVal applyCanvasTransformsFirst As Boolean = True)

    'If this layer has transforms applied to it, we have to make them permanent before proceeding.
    '
    'Note that this behavior can be overridden if a function is ABSOLUTELY CERTAIN that the end result
    ' won't botch the transform, but this is only true for pixel transforms that operate strictly in
    ' separate x/y domains, such as resizing.  Note that this setting may be forcibly ignored if the
    ' layer has complex affine transforms active (like rotation), because those transforms change the
    ' (x, y) domain of the layer relative to the (x, y) domain of the image.
    Dim mustMakeCanvasTransformsPermanent As Boolean
    mustMakeCanvasTransformsPermanent = applyCanvasTransformsFirst Or Me.AffineTransformsActive(False)
    If mustMakeCanvasTransformsPermanent Then Me.MakeCanvasTransformsPermanent parentImageWidth, parentImageHeight
    
    'We can skip this step for layers that are already the size of the image.  This is common for background layers.
    Dim nullPaddingUnnecessary As Boolean
    nullPaddingUnnecessary = False
    
    With myLayerData
        nullPaddingUnnecessary = (.l_OffsetX = 0&) And (.l_OffsetY = 0&) And (Me.GetLayerWidth(True) = parentImageWidth) And (Me.GetLayerHeight(True) = parentImageHeight)
    End With
    
    'If the layer still requires specialized padding, carry on
    If (Not nullPaddingUnnecessary) Then
        
        'Create a blank destination DIB at the size of the parent image
        Dim tmpPaddingDIB As pdDIB
        Set tmpPaddingDIB = New pdDIB
        tmpPaddingDIB.CreateBlank parentImageWidth, parentImageHeight, 32, 0, 0
    
        'Copy the current layer into the temporary DIB, with proper offsets applied
        GDI.BitBltWrapper tmpPaddingDIB.GetDIBDC, myLayerData.l_OffsetX, myLayerData.l_OffsetY, m_LayerDIB.GetDIBWidth, m_LayerDIB.GetDIBHeight, m_LayerDIB.GetDIBDC, 0, 0, vbSrcCopy
    
        'Replace our current layer DIB with the temporary DIB
        tmpPaddingDIB.SetInitialAlphaPremultiplicationState m_LayerDIB.GetAlphaPremultiplication
        Set m_LayerDIB = tmpPaddingDIB
        
        'Reset layer offset values
        myLayerData.l_OffsetX = 0&
        myLayerData.l_OffsetY = 0&
        
        'TODO: synchronize layer mask here, if one exists
        
        'Mark the layer as dirty
        Me.NotifyOfDestructiveChanges
    
    End If
    
End Sub

'Return the current layer as a null-padded DIB.  This is preferable to ConvertToNullPaddedLayer, above, because you
' can pass an initialized DIB (perhaps from a previous step in a multi-layer conversion process) and this function
' will try to reuse it.  (Note that things like affine-transformed DIBs complicate this process and may allocate
' their own temporary memory - alas!)
'
'Note, however, that if you apply the returned DIB to this layer later on, you'll need to manually handle things
' like resetting offsets to [0, 0] and un-marking affine transform state.
'
'By design, this function is LOSSLESS.  It does not modify current layer state.
Friend Sub GetLayerDIB_NullPadded(ByRef dstDIB As pdDIB, ByVal parentImageWidth As Long, ByVal parentImageHeight As Long)
    
    'Ensure the destination DIB is the correct size
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    If (dstDIB.GetDIBWidth <> parentImageWidth) Or (dstDIB.GetDIBHeight <> parentImageHeight) Then
        dstDIB.CreateBlank parentImageWidth, parentImageHeight, 32, 0, 0
    Else
        dstDIB.ResetDIB 0
    End If
    
    dstDIB.SetInitialAlphaPremultiplicationState m_LayerDIB.GetAlphaPremultiplication()
    
    'If this layer has transforms applied to it, we will simply ask it to render directly into the
    ' destination DIB (instead of creating a new DIB, then passing that to us).
    If Me.AffineTransformsActive(True) Then
        Me.GetLayerCornerCoordinates m_PlgPoints
        GDI_Plus.GDIPlus_PlgBlt dstDIB, m_PlgPoints, m_LayerDIB, 0, 0, Me.GetLayerWidth(False), Me.GetLayerHeight(False), 1!, Me.GetLayerResizeQuality_GDIPlus
        
    'If this is a normal layer, we can just BitBlt into position
    Else
        GDI.BitBltWrapper dstDIB.GetDIBDC, myLayerData.l_OffsetX, myLayerData.l_OffsetY, m_LayerDIB.GetDIBWidth, m_LayerDIB.GetDIBHeight, m_LayerDIB.GetDIBDC, 0, 0, vbSrcCopy
    End If
    
End Sub

'Update the internal thumbnail.  Only call this function if the current thumbnail has been marked as DIRTY; otherwise, just use the
' existing thumbnail as-is.
Private Sub UpdateInternalThumbnail(Optional ByVal requiredSize As Long = 0)
    
    'If the layer has not been instantiated properly, reject the thumbnail request
    If (m_LayerDIB Is Nothing) Then Exit Sub
    If (m_LayerDIB.GetDIBWidth = 0) Or (m_LayerDIB.GetDIBHeight = 0) Then Exit Sub
    
    'To help us coalesce multiple thumbnail requests together, we simply cache a "large enough" copy
    ' and then produce smaller copies on-demand.
    If (requiredSize < LAYER_THUMB_SIZE) Then requiredSize = LAYER_THUMB_SIZE
    m_CurrentLayerThumbSize = requiredSize
    
    'Determine proper dimensions for the thumbnail image.
    Dim newThumbWidth As Long, newThumbHeight As Long
    PDMath.ConvertAspectRatio m_LayerDIB.GetDIBWidth, m_LayerDIB.GetDIBHeight, requiredSize, requiredSize, newThumbWidth, newThumbHeight
    
    'Prepare the thumbnail DIB
    If (m_LayerThumbnail.GetDIBWidth <> newThumbWidth) Or (m_LayerThumbnail.GetDIBHeight <> newThumbHeight) Then
        m_LayerThumbnail.CreateBlank newThumbWidth, newThumbHeight, 32, 0
    Else
        m_LayerThumbnail.ResetDIB 0
    End If
    
    'Retrieve a composited thumbnail.  (Note that the user's thumbnail performance setting affects the interpolation
    ' method used.)
    GDI_Plus.GDIPlus_StretchBlt m_LayerThumbnail, 0!, 0!, newThumbWidth, newThumbHeight, m_LayerDIB, 0!, 0!, m_LayerDIB.GetDIBWidth, m_LayerDIB.GetDIBHeight, , UserPrefs.GetThumbnailInterpolationPref(), , , , True
    m_LayerThumbnail.SetInitialAlphaPremultiplicationState True
    
    'TODO: consider layer mask here?
    
    'Before exiting, apply color-management.  This spares callers from needing to do it (and when wouldn't we want
    ' a color-managed thumbnail anyway?)
    ColorManagement.ApplyDisplayColorManagement m_LayerThumbnail, , True
    
    'Mark the thumbnail state as clean
    m_IsThumbClean = True
    
End Sub

'External functions can use this function to request a thumbnail version of the layer contents.
Friend Function RequestThumbnail(ByRef dstThumbnailDIB As pdDIB, Optional ByVal thumbnailSize As Long = LAYER_THUMB_SIZE, Optional ByVal thumbIsSquare As Boolean = True, Optional ByVal ptrToRectF As Long = 0) As Boolean
    
    If (dstThumbnailDIB Is Nothing) Then Set dstThumbnailDIB = New pdDIB
    
    'Check the (rare) special case where the requested thumbnail size equals the layer's size;
    ' if this happens we can do a simple copy and exit immediately.
    If (thumbnailSize = Me.GetLayerWidth(True)) And (thumbnailSize = Me.GetLayerHeight(True)) And (Not Me.AffineTransformsActive) And (ptrToRectF = 0) Then
        dstThumbnailDIB.CreateFromExistingDIB Me.GetLayerDIB
        RequestThumbnail = True
        Exit Function
    End If
    
    'Is the current thumbnail dirty?  If so, regenerate it.
    If (Not m_IsThumbClean) Or (thumbnailSize > m_CurrentLayerThumbSize) Then UpdateInternalThumbnail thumbnailSize
    
    'We also need to determine the thumbnail's actual width and height, and any x and y offset necessary to preserve the
    ' aspect ratio and center the image on the thumbnail.
    Dim thumbWidth As Long, thumbHeight As Long
    PDMath.ConvertAspectRatio m_LayerThumbnail.GetDIBWidth, m_LayerThumbnail.GetDIBHeight, thumbnailSize, thumbnailSize, thumbWidth, thumbHeight
    
    Dim tmpRectF As RectF
    
    'Thumbnails have some interesting requirements.  We always want them to be square, with the image set in the middle
    ' of the thumbnail (with aspect ratio preserved) and any empty edges made transparent.
    If thumbIsSquare Then
    
        'If the image is wider than it is tall, center the thumbnail vertically
        If (thumbWidth > thumbHeight) Then
            tmpRectF.Left = 0!
            tmpRectF.Top = (thumbnailSize - thumbHeight) * 0.5
        
        '...otherwise, center it horizontally
        Else
            tmpRectF.Top = 0!
            tmpRectF.Left = (thumbnailSize - thumbWidth) * 0.5
        End If
        
        tmpRectF.Width = thumbWidth
        tmpRectF.Height = thumbHeight
        
    'If this is a non-standard request (e.g. a square thumbnail isn't needed), paint out the thumbnail at whatever
    ' size the caller wants.  Note that they are responsible for preparing the target DIB in this scenario.
    Else
        
        If (ptrToRectF <> 0) Then
            CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
        Else
            With tmpRectF
                .Left = 0!
                .Top = 0!
                .Width = thumbWidth
                .Height = thumbHeight
            End With
        End If
        
    End If
    
    'If the caller has manually specified a destination rectangle, we do *not* want to prepare a DIB for them
    ' (as they will have prepared it themselves, at whatever size they require).
    Dim prepDstDIB As Boolean: prepDstDIB = True
    If (Not thumbIsSquare) And (ptrToRectF <> 0) Then prepDstDIB = False
    
    If prepDstDIB Then
        If thumbIsSquare Then
            dstThumbnailDIB.CreateBlank thumbnailSize, thumbnailSize, 32, 0, 0
        Else
            dstThumbnailDIB.CreateBlank thumbWidth, thumbHeight, 32, 0, 0
        End If
    End If
    
    'Paint the thumbnail into place!
    GDI_Plus.GDIPlus_StretchBlt dstThumbnailDIB, tmpRectF.Left, tmpRectF.Top, tmpRectF.Width, tmpRectF.Height, m_LayerThumbnail, 0!, 0!, m_LayerThumbnail.GetDIBWidth, m_LayerThumbnail.GetDIBHeight, , UserPrefs.GetThumbnailInterpolationPref(), , , , True
    dstThumbnailDIB.SetInitialAlphaPremultiplicationState True
    
    'NOTE: the thumbnail *has already been color-managed*, so the caller doesn't need to manage it further.
    RequestThumbnail = True
    
End Function

'Specialized thumbnail function; renders a layer thumbnail while taking into account offsets and
' non-destructive transforms.  This produces a thumbnail of how the layer appears *in the current image*,
' so for small layers, it may use a ton of dead space.  This is by design.
'
'Note that this specialized function cannot physically make use of the current thumbnail cache, as it
' requires specialized transforms.  As such, its performance will possibly be (much?) slower than the
' normal thumbnail function.
'
'TODO: color-management?
Friend Function RequestThumbnail_ImageCoords(ByRef dstThumbnail As pdDIB, ByRef srcImage As pdImage, Optional ByVal thumbnailSize As Long = LAYER_THUMB_SIZE, Optional ByVal thumbIsSquare As Boolean = True, Optional ByVal ptrToRectF As Long = 0) As Boolean
    
    RequestThumbnail_ImageCoords = False
    If (srcImage Is Nothing) Then Exit Function
    
    'We now need to figure out where the hell to position the layer inside the new DIB.  Rather than use manual
    ' coordinate calculations, we're simply going to use a transformation matrix.
    Dim cTransform As pd2DTransform
    Me.GetCopyOfLayerTransformationMatrix_Full cTransform
    
    'The default transformation matrix produces a transformation of the current layer (at full-size) into its
    ' parent image (at full size).  Unfortunately, we are not working at full-size - we are working at some
    ' (unpredictable) thumbnail size.  As such, we need to append additional scaling operations, based on what
    ' kind of thumbnail has been requested.
    Dim thumbWidth As Long, thumbHeight As Long
    PDMath.ConvertAspectRatio srcImage.Width, srcImage.Height, thumbnailSize, thumbnailSize, thumbWidth, thumbHeight
    
    Dim tmpRectF As RectF
    
    'Thumbnails have some interesting requirements.  We always want them to be square, with the image set in the middle
    ' of the thumbnail (with aspect ratio preserved) and any empty edges made transparent.
    If thumbIsSquare Then
    
        'If the image is wider than it is tall, center the thumbnail vertically
        If (thumbWidth > thumbHeight) Then
            tmpRectF.Left = 0!
            tmpRectF.Top = Int((thumbnailSize - thumbHeight) / 2 + 0.5)
        
        '...otherwise, center it horizontally
        Else
            tmpRectF.Top = 0!
            tmpRectF.Left = Int((thumbnailSize - thumbWidth) / 2 + 0.5)
        End If
        
        tmpRectF.Width = thumbWidth
        tmpRectF.Height = thumbHeight
        
    'If this is a non-standard request (e.g. a square thumbnail isn't needed), paint out the thumbnail at whatever
    ' size the caller wants.  Note that they are responsible for preparing the target DIB in this scenario.
    Else
        
        If (ptrToRectF <> 0) Then
            CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
        Else
            With tmpRectF
                .Left = 0!
                .Top = 0!
                .Width = thumbWidth
                .Height = thumbHeight
            End With
        End If
        
        'In this branch, we can perform another optimization.  If this layer is the same size as
        ' the target image (extremely common for photos, animations, etc), and this layer doesn't have
        ' any active affine transforms, we can simply paint it as-is directly into the target rect.
        ' This saves us a lot of time, and we can use a special paint operation with edge-fixes for
        ' some GDI+ weaknesses.
        If (Me.GetLayerWidth = srcImage.Width) And (Me.GetLayerHeight = srcImage.Height) And (Not Me.AffineTransformsActive(True)) Then
            GDI_Plus.GDIPlus_StretchBlt dstThumbnail, tmpRectF.Left, tmpRectF.Top, tmpRectF.Width, tmpRectF.Height, Me.GetLayerDIB, 0, 0, Me.GetLayerWidth, Me.GetLayerHeight, interpolationType:=GP_IM_Bilinear, dstCopyIsOkay:=True
            RequestThumbnail_ImageCoords = True
            Exit Function
        End If
        
    End If
    
    'We now have transform settings for the image as a whole.  Use these to calculate an x/y scaling ratio.
    Dim xScale As Single, yScale As Single
    xScale = tmpRectF.Width / srcImage.Width
    yScale = tmpRectF.Height / srcImage.Height
    
    'Append the scale settings to the transform matrix
    cTransform.ApplyScaling xScale, yScale
    
    'Finally, append transform settings using the current layer x/y offset
    cTransform.ApplyTranslation tmpRectF.Left, tmpRectF.Top
    
    'If the caller has manually specified a destination rectangle, we do *not* want to prepare a DIB for them
    ' (as they will have prepared it themselves, at whatever size they require).
    Dim prepDstDIB As Boolean: prepDstDIB = True
    If (Not thumbIsSquare) And (ptrToRectF <> 0) Then prepDstDIB = False
    
    If prepDstDIB Then
        If (dstThumbnail Is Nothing) Then Set dstThumbnail = New pdDIB
        If thumbIsSquare Then
            dstThumbnail.CreateBlank thumbnailSize, thumbnailSize, 32, 0, 0
        Else
            dstThumbnail.CreateBlank thumbWidth, thumbHeight, 32, 0, 0
        End If
    End If
    
    'Transform the layer's corner points
    Dim cPoints() As PointFloat
    ReDim cPoints(0 To 2) As PointFloat
    cPoints(0).x = 0
    cPoints(0).y = 0
    cPoints(1).x = Me.GetLayerDIB.GetDIBWidth - 1
    cPoints(1).y = 0
    cPoints(2).x = 0
    cPoints(2).y = Me.GetLayerDIB.GetDIBHeight - 1
    
    cTransform.ApplyTransformToPointFs VarPtr(cPoints(0)), 3
    
    'Prep both surfaces using pd2D (including a cropping region for transparent portions of the thumbnail-)
    Dim cSurfaceDst As pd2DSurface, cSurfaceSrc As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDIB cSurfaceDst, dstThumbnail, False
    Drawing2D.QuickCreateSurfaceFromDIB cSurfaceSrc, Me.GetLayerDIB, False
    cSurfaceDst.SetSurfaceClip_FromRectF tmpRectF
    cSurfaceDst.SetSurfaceCompositing P2_CM_Overwrite
    cSurfaceDst.SetSurfacePixelOffset P2_PO_Half
    
    'Convert the user's thumbnail performance preference into a usable resample mode
    Dim thumbPerf As PD_PerformanceSetting
    thumbPerf = UserPrefs.GetThumbnailPerformancePref
    If (thumbPerf = PD_PERF_FASTEST) Then
        cSurfaceDst.SetSurfaceResizeQuality P2_RQ_Fast
    ElseIf (thumbPerf = PD_PERF_BALANCED) Then
        cSurfaceDst.SetSurfaceResizeQuality P2_RQ_Bilinear
    ElseIf (thumbPerf = PD_PERF_BESTQUALITY) Then
        cSurfaceDst.SetSurfaceResizeQuality P2_RQ_Bicubic
    End If
    
    'Paint the thumbnail into place!
    PD2D.DrawSurfaceTransformedF cSurfaceDst, cSurfaceSrc, cTransform, 0, 0, Me.GetLayerDIB.GetDIBWidth, Me.GetLayerDIB.GetDIBHeight
    Set cSurfaceDst = Nothing: Set cSurfaceSrc = Nothing
    dstThumbnail.SetInitialAlphaPremultiplicationState True
    
    'NOTE: the thumbnail *has already been color-managed*, so the caller doesn't need to manage it further.
    RequestThumbnail_ImageCoords = True
    
End Function

'Whenever the user does something with the mouse - click, move, etc - the primary canvas will poll this function.  This function
' will check the passed mouse values (which have already been translated into the IMAGE coordinate space), and return a Long-type
' value specifying whether or not that coordinate is of interest to this layer.  This basically provides a flexible framework for
' supporting mouse events for any layer, because it's up to the layer to determine which points are interesting and which are not.
'
'INPUTS: current mouse coords, already translated into the IMAGE coordinate space (not the canvas coordinate space!)
'
'OUTPUT: index of the POI, if any.  If the mouse is not over a POI, -1 is returned.
Friend Function CheckForPointOfInterest(ByVal imgX As Long, ByVal imgY As Long, Optional ByVal checkEdgeCenters As Boolean = True) As PD_PointOfInterest
    
    'MouseAccuracy in PD is a global value, but because we are working in image coordinates, we must compensate for the
    ' current zoom value.  (Otherwise, when zoomed out the user would be forced to work with tighter accuracy!)
    ' (TODO: come up with a better solution for this.  Accuracy should *really* be handled in the canvas coordinate space,
    '        so perhaps the caller should specify an image x/y and a radius...?)
    Dim mouseAccuracy As Double
    mouseAccuracy = Drawing.ConvertCanvasSizeToImageSize(Interface.GetStandardInteractionDistance(), PDImages.GetActiveImage)
    
    'Find the smallest distance for this mouse position
    Dim minDistance As Double
    minDistance = mouseAccuracy
    
    'When a point within the accuracy limit is found, its index will be assigned to this variable
    Dim closestPoint As PD_PointOfInterest
    closestPoint = poi_Undefined
    
    'To facilitate the notion of "generalized point of interest" handling, coordinate checks are performed against
    ' an array of potential POI targets.  While arrays may not make as much sense as something like a RECT (when
    ' working with standard layers), it lets us generalize all coordinate checks into a single function, which is awesome!
    Dim poiList() As PointFloat
    
    'Points of interest are sorted by layer type.  In most cases, only the corners of a layer are interesting
    ' (as they control the layer's bounding box), but in the future, things like a Pen tool could have many
    ' points of interest.
    Select Case m_LayerType
    
        'Image and text layers consider their corners to be points of interest
        Case PDL_Image, PDL_TextBasic, PDL_TextAdvanced
            
            'Manually populate the point of interest array with the layer's corner coordinates.
            ReDim poiList(0 To 7) As PointFloat
            poiList(0).x = 0!
            poiList(0).y = 0!
            poiList(1).x = Me.GetLayerWidth(False)
            poiList(1).y = 0!
            poiList(2).x = 0! + Me.GetLayerWidth(False)
            poiList(2).y = 0! + Me.GetLayerHeight(False)
            poiList(3).x = 0!
            poiList(3).y = 0! + Me.GetLayerHeight(False)
            
            'As of PD v8.0, rotation handles are also provided.  Note that changes to these positions
            ' *must be mirrored* in the GetLayerRotationNodeCoordinates() function above!
            poiList(4).x = 0! + Me.GetLayerWidth(False)
            poiList(4).y = 0! + Me.GetLayerHeight(False) * 0.5
            
            poiList(5).x = 0! + Me.GetLayerWidth(False) * 0.5
            poiList(5).y = 0! + Me.GetLayerHeight(False)
            
            poiList(6).x = 0!
            poiList(6).y = 0! + Me.GetLayerHeight(False) * 0.5
            
            poiList(7).x = 0! + Me.GetLayerWidth(False) * 0.5
            poiList(7).y = 0!
            
            'Convert our list of POI coordinates into the image working space
            Dim cTransform As pd2DTransform
            Me.GetCopyOfLayerTransformationMatrix_Full cTransform
            cTransform.ApplyTransformToPointFs VarPtr(poiList(0)), 8
            
            'Check to see if one of the corners has been clicked
            closestPoint = PDMath.FindClosestPointInFloatArray(imgX, imgY, minDistance, poiList)
            
            'If a corner was clicked, return its value and exit
            If (closestPoint <> poi_Undefined) Then
            
                'Map the POI index to a human-friendly constant.
                If (closestPoint = 0) Then
                    CheckForPointOfInterest = poi_CornerNW
                ElseIf (closestPoint = 1) Then
                    CheckForPointOfInterest = poi_CornerNE
                ElseIf (closestPoint = 2) Then
                    CheckForPointOfInterest = poi_CornerSE
                ElseIf (closestPoint = 3) Then
                    CheckForPointOfInterest = poi_CornerSW
                
                'Detecting hover on rotation nodes can be disabled (the user can toggle this in the Move Tool settings);
                ' if it's disabled, we want to ignore hits on that point.
                ElseIf (closestPoint = 4) Then
                    If checkEdgeCenters Then CheckForPointOfInterest = poi_EdgeE Else CheckForPointOfInterest = poi_Undefined
                ElseIf (closestPoint = 5) Then
                    If checkEdgeCenters Then CheckForPointOfInterest = poi_EdgeS Else CheckForPointOfInterest = poi_Undefined
                ElseIf (closestPoint = 6) Then
                    If checkEdgeCenters Then CheckForPointOfInterest = poi_EdgeW Else CheckForPointOfInterest = poi_Undefined
                ElseIf (closestPoint = 7) Then
                    If checkEdgeCenters Then CheckForPointOfInterest = poi_EdgeN Else CheckForPointOfInterest = poi_Undefined
                End If
                
                If (CheckForPointOfInterest <> poi_Undefined) Then
                    Exit Function
                Else
                    closestPoint = poi_Undefined
                End If
                
            End If
            
            'If we're at this line of code, a POI was not clicked. Perform one final check to see if the mouse is within
            ' the layer's boundaries, and if it is, return the "move" ID, then exit.
            Dim tmpPath As pd2DPath
            Set tmpPath = New pd2DPath
            tmpPath.AddPolygon 4, VarPtr(poiList(0)), True, False
            
            If tmpPath.IsPointInsidePathF(imgX, imgY) Then
                CheckForPointOfInterest = poi_Interior
            Else
                CheckForPointOfInterest = poi_Undefined
            End If
    
    End Select

End Function

'Some operations may remove affine transforms from the current layer, typically by making them permanent.
' Call this function to reset all affine transform settings to their default (inactive) state.
Friend Sub ResetAffineTransformProperties()
    
    With myLayerData
        .l_CanvasXModifier = 1#
        .l_CanvasYModifier = 1#
        .l_Angle = 0#
        .l_RotateCenterX = 0.5
        .l_RotateCenterY = 0.5
        .l_ResizeQuality = LRQ_Bilinear
        .l_ShearX = 0#
        .l_ShearY = 0#
    End With
    
    'While here, also assign default vector layer values
    With myVectorData
        .vd_Width = 0
        .vd_Height = 0
    End With
    
End Sub

'Sometimes, it is helpful to reset all non-destructive parameters to their default values.
' (For example, after flattening an image, we want to reset the base layer.)
Friend Sub ResetLayerParameters()

    'Assign default values to this instance's header
    With myLayerData
        .l_GroupID = 0
        .l_Opacity = 100!
        .l_BlendMode = BM_Normal
        .l_OffsetX = 0&
        .l_OffsetY = 0&
        .l_CanvasXModifier = 1#
        .l_CanvasYModifier = 1#
        .l_Angle = 0#
        .l_RotateCenterX = 0.5
        .l_RotateCenterY = 0.5
        .l_Visibility = True
        .l_ResizeQuality = LRQ_Bilinear
        .l_ShearX = 0#
        .l_ShearY = 0#
        .l_AlphaMode = AM_Normal
        .l_MaskExists = False
        .l_MaskActive = False
        Set m_LayerMask = Nothing
    End With
    
    'Assign default vector layer values
    With myVectorData
        .vd_Width = 0
        .vd_Height = 0
    End With
    
End Sub

'Non-destructively suspend as many layer resources as we can to either disk or memory.
Friend Sub SuspendLayer(Optional ByVal suspendViewportToo As Boolean = False, Optional ByVal suspendToDisk As Boolean = False)
    
    'Layer DIB is the most obvious target for memory reduction
    If (Not m_LayerDIB Is Nothing) Then m_LayerDIB.SuspendDIB cf_Lz4, False, suspendToDisk
    
    'Thumbnail is an easy target too (since it's small, suspending is roughly instantaneous)
    If (Not m_LayerThumbnail Is Nothing) Then m_LayerThumbnail.SuspendDIB cf_Lz4, False
    
    'Finally, LOD-DIBs can be suspended too, but be careful with these because there are meaningful
    ' perf implications for viewport rendering if these have to be swapped back in frequently.
    If suspendViewportToo Then
        Dim i As Long
        For i = LBound(m_lodDIB) To UBound(m_lodDIB)
            If (Not m_lodDIB(i) Is Nothing) Then m_lodDIB(i).SuspendDIB cf_Lz4, False
        Next i
    End If
    
    'TODO: layer mask here?
    
End Sub

'If something needs to access the layer DIB directly (like PD's central compositor), this function should first be called for
' any vector layers.  If the DIB is already synched, there is no penalty to calling this function, but if it is *not* synched,
' this will guarantee the DIB is ready for access.
'
'For vector DIBs, returns TRUE if we had to sync the vector DIB; FALSE if the DIB was already synched
'For raster DIBs, returns TRUE if the image has changed since the last viewport hash was set; FALSE otherwise.
'
'If the return value is TRUE, any DIBs derived from this layer (e.g. viewport ones) must be recreated
Friend Function SyncInternalDIB(Optional ByVal levelOfDetail As PD_CompositorLOD = CLC_Generic) As Boolean

    'If this is not a vector layer, rely on the "DIB has changed since last hash request" marker.
    If (m_LayerType = PDL_Image) Then
        SyncInternalDIB = m_DIBChangedSinceLastHash(levelOfDetail)
        Exit Function
    End If
    
    'If the DIB is already synched, disregard this request
    If m_VectorDIBSynched Then
        SyncInternalDIB = False
        Exit Function
    End If
    
    'From here, the specific vector syncing process varies by vector layer type
    Select Case m_LayerType
    
        Case PDL_TextBasic, PDL_TextAdvanced
            
            'Set the text rendering engine according to the layer type (basic vs advanced)
            If (m_LayerType = PDL_TextBasic) Then
                myTextRenderer.SetGenericTextProperty ptp_RenderingEngine, te_WAPI
            Else
                myTextRenderer.SetGenericTextProperty ptp_RenderingEngine, te_PhotoDemon
            End If
            
            'Initialize the backer DIB as necessary
            If (m_LayerDIB Is Nothing) Then Set m_LayerDIB = New pdDIB
            If (m_LayerDIB.GetDIBWidth <> myVectorData.vd_Width) Or (m_LayerDIB.GetDIBHeight <> myVectorData.vd_Height) Then
                m_LayerDIB.CreateBlank Int(myVectorData.vd_Width + 0.9999), Int(myVectorData.vd_Height + 0.9999), 32, 0, 0
            Else
                m_LayerDIB.ResetDIB 0
            End If
            
            m_LayerDIB.SetInitialAlphaPremultiplicationState True
            
            'Use pdTextRenderer to paint the string onto the backer DIB
            Const REPORT_TEXT_RENDER_TIME As Boolean = False
            Dim startTime As Currency
            If REPORT_TEXT_RENDER_TIME Then VBHacks.GetHighResTime startTime
            myTextRenderer.RenderTextToDIB m_LayerDIB, 0, 0, myVectorData.vd_Width, myVectorData.vd_Height
            If REPORT_TEXT_RENDER_TIME Then PDDebug.LogAction "Text rendering took: " & VBHacks.GetTimeDiffNowAsString(startTime)
            
    End Select
    
    'Note that the DIB is ready to go!
    m_VectorDIBSynched = True
    SyncInternalDIB = True

End Function

'Estimate the RAM usage for this layer, based on known heavy memory users (layer DIB, in particular)
Friend Function EstimateRAMUsage() As Double

    'Calculate the raw memory size of both the layer and the cached thumbnail DIB, if any
    If Not (m_LayerDIB Is Nothing) Then EstimateRAMUsage = m_LayerDIB.GetDIBStride * m_LayerDIB.GetDIBHeight
    If Not (m_LayerThumbnail Is Nothing) Then EstimateRAMUsage = EstimateRAMUsage + m_LayerThumbnail.GetDIBStride * m_LayerThumbnail.GetDIBHeight
    
    Dim i As Long
    For i = 0 To NUM_OF_LOD_CACHES - 1
        If Not (m_lodDIB(i) Is Nothing) Then EstimateRAMUsage = EstimateRAMUsage + m_lodDIB(i).GetDIBStride * m_lodDIB(i).GetDIBHeight
    Next i
    
    'If a layer mask exists, add it to the total
    If (Not m_LayerMask Is Nothing) Then EstimateRAMUsage = EstimateRAMUsage + m_LayerMask.EstimateRAMUsage
    
    'Assume a 10% overhead for this class instance and other miscellaneous layer data
    EstimateRAMUsage = EstimateRAMUsage * 1.1

End Function

'Update the layer's current "viewport state hash".  This is used by the compositor, only.
Friend Sub SetViewportHash(ByVal cacheID As PD_CompositorLOD, ParamArray allParams() As Variant)
    m_ViewportStateHashes(cacheID) = GetViewportHash_Theoretical(allParams)
    m_DIBChangedSinceLastHash(cacheID) = False
End Sub

'If the viewport generates a hash via getViewportHash_Theoretical(), it can pass that hash directly to this function to improve performance
Friend Sub SetViewportHash_FromLong(ByVal cacheID As PD_CompositorLOD, ByVal newHash As Long)
    m_ViewportStateHashes(cacheID) = newHash
    m_DIBChangedSinceLastHash(cacheID) = False
End Sub

'Retrieve the layer's current viewport hash
Friend Function GetViewportHash(ByVal cacheID As PD_CompositorLOD) As Long
    GetViewportHash = m_ViewportStateHashes(cacheID)
End Function

'Given a set of caller-specified parameters, generate a matching viewport hash.
' This is used by the compositor to detect changes to the current layer.
' Caller-specified params are relevant because the compositor will also pass viewport settings,
' and if *those* change (but the layer is unchanged), we still need to regenerate the function.
Friend Function GetViewportHash_Theoretical(ParamArray allParams() As Variant) As Long
    
    On Error GoTo BadHash
    
    'First, generate a hash of relevant internals.
    Dim tmpHash As Long
    
    With myLayerData
        tmpHash = (.l_Angle * 359013) Xor (.l_CanvasXModifier * 8091) Xor (.l_CanvasYModifier * 4095)
        tmpHash = tmpHash Xor (.l_OffsetX * 13) Xor (.l_OffsetY * 29)
        tmpHash = tmpHash Xor (.l_ResizeQuality * 620)
        tmpHash = tmpHash Xor Int(.l_Opacity * 65535)
        tmpHash = tmpHash Xor (.l_ShearX * 524287#) Xor (.l_ShearY * 1048575#)
        tmpHash = tmpHash Xor ((.l_BlendMode + 1) * 63) Xor ((.l_AlphaMode + 1) * 511)
        tmpHash = tmpHash Xor ((.l_RotateCenterX + 1#) * 3730#) Xor ((.l_RotateCenterY + 1# * 5850#))
    End With
    
    'Finally, hash any parameters we were passed.
    If (UBound(allParams) >= LBound(allParams)) Then
    
        Dim i As Long, tmpLong As Long
        For i = LBound(allParams) To UBound(allParams)
            tmpLong = Int(allParams(i))
            If (tmpLong >= 100000000#) Then tmpLong = tmpLong \ 4 'Safety against overflow on unpredictable inputs
            tmpHash = tmpHash Xor tmpLong * ((i + 1) * 7)
        Next i
    
    End If
    
BadHash:
    GetViewportHash_Theoretical = tmpHash
    
End Function

'Access a specific level-of-detail (LOD) cache.
Friend Function TmpLODDIB(Optional ByVal cacheID As PD_CompositorLOD = CLC_Generic) As pdDIB
    Set TmpLODDIB = m_lodDIB(cacheID)
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String)
    If UserPrefs.GenerateDebugLogs Then
        PDDebug.LogAction "pdLayer." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdLayer." & funcName & "() reported an error: " & errDescription
    End If
End Sub




'**********************************************************************************************
'Legacy support functions follow.  Do NOT modify these functions except to fix security issues.

'Legacy create-from-XML function
Private Function CreateNewLayerFromXML_Legacy(ByRef xmlString As String, Optional ByVal useCustomLayerID As Long = -1, Optional ByVal createNonDestructively As Boolean = False) As Boolean

    'Prepare an XML engine, which greatly simplifies the process of parsing XML data
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Validate the XML header...
    If (xmlEngine.LoadXMLFromString(xmlString) And xmlEngine.IsPDDataType("pdLayer")) Then
        
        'If the caller has not specified their own layer ID, use the one from file.
        If (useCustomLayerID = -1) Then
            m_layerID = xmlEngine.GetUniqueTag_Long("ID")
        Else
            m_layerID = useCustomLayerID
        End If
    
        'Read in the layer type.  Remember that this value is immutable, and cannot be changed for the life of the layer.
        m_LayerType = xmlEngine.GetUniqueTag_Long("Type", PDL_Image)
        
        'Read in the contents of the LayerData header.  (NOTE: this must be manually modified if new entries are
        ' added to the LayerData type.
        With myLayerData
            .l_Name = xmlEngine.GetUniqueTag_String("Name")
            .l_GroupID = xmlEngine.GetUniqueTag_Long("GroupID")
            .l_Opacity = xmlEngine.GetUniqueTag_Long("Opacity")
            .l_BlendMode = xmlEngine.GetUniqueTag_Long("BlendMode", BM_Normal)
            .l_OffsetX = Int(xmlEngine.GetUniqueTag_Double("OffsetX") + 0.5)
            .l_OffsetY = Int(xmlEngine.GetUniqueTag_Double("OffsetY") + 0.5)
            .l_CanvasXModifier = xmlEngine.GetUniqueTag_Double("CanvasXModifier", 1)
            .l_CanvasYModifier = xmlEngine.GetUniqueTag_Double("CanvasYModifier", 1)
            .l_Angle = xmlEngine.GetUniqueTag_Double("Angle")
            .l_RotateCenterX = xmlEngine.GetUniqueTag_Double("RotateCenterX", 0.5)
            .l_RotateCenterY = xmlEngine.GetUniqueTag_Double("RotateCenterY", 0.5)
            .l_Visibility = xmlEngine.GetUniqueTag_Boolean("Visibility")
            .l_ResizeQuality = xmlEngine.GetUniqueTag_Long("ResizeQuality", LRQ_Bilinear)
            .l_ShearX = xmlEngine.GetUniqueTag_Double("ShearX", 0)
            .l_ShearY = xmlEngine.GetUniqueTag_Double("ShearY", 0)
            .l_AlphaMode = xmlEngine.GetUniqueTag_Long("AlphaMode", AM_Normal)
            .l_MaskExists = False
            .l_MaskActive = False
            Set m_LayerMask = Nothing
        End With
        
        'Next, we're going to do something kinda weird.  We're going to create a blank DIB for the layer contents
        ' using the DIB header supplied in the XML, but we're *not* going to fill the DIB!  An external function will
        ' handle that, because the actual DIB bits are stored in binary format.  (Serializing them to XML is not feasible
        ' for the image sizes PD deals with.)
        '
        'Note that this behavior can be overridden by setting the optional createNonDestructively parameter to TRUE.
        ' This is used by the Undo/Redo engine to re-load an image's header data, without actually overwriting the
        ' current layer DIB.
        If (Not createNonDestructively) Then
        
            Set m_LayerDIB = New pdDIB
            With xmlEngine
                m_LayerDIB.CreateBlank .GetUniqueTag_Long("DIB_Width", 1), .GetUniqueTag_Long("DIB_Height", 1), .GetUniqueTag_Long("DIB_ColorDepth", 32)
                
                'Mark alpha premultiplication in advance
                m_LayerDIB.SetInitialAlphaPremultiplicationState .GetUniqueTag_Boolean("DIB_AlphaPremultiplication", True)
            
            End With
            
            'Note that all internal layer caches are no longer valid
            Me.NotifyOfDestructiveChanges
            
        End If
        
        CreateNewLayerFromXML_Legacy = True
    
    'If the layer XML data can't be validated, there's nothing we can do but exit.
    Else
        InternalError "CreateNewLayerFromXML_Legacy", "layer could not be created: are you sure the supplied data was valid?"
        CreateNewLayerFromXML_Legacy = False
    End If
    
End Function

'Legacy vector data parser
Private Function SetVectorDataFromXML_Legacy(ByRef xmlString As String) As Boolean

    'Prepare an XML engine, which greatly simplifies the process of parsing XML data
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Validate the XML header...
    If xmlEngine.LoadXMLFromString(xmlString) And xmlEngine.IsPDDataType("pdVector") Then
        
        'All vector layers store some similar data
        With myVectorData
            .vd_Width = xmlEngine.GetUniqueTag_Long("GenericVectorWidth", 1)
            .vd_Height = xmlEngine.GetUniqueTag_Long("GenericVectorHeight", 1)
        End With
        
        'Next, the vector tags we look for vary by layer type
        Select Case m_LayerType
            
            Case PDL_TextBasic, PDL_TextAdvanced
                
                'pdTextRenderer serializes all its data from XML for us, making this step very simple
                myTextRenderer.SetAllFontSettingsFromXML xmlEngine.ReturnCurrentXMLString(True)
                
        End Select
                
        'Note that all internal layer caches are no longer valid; this will also cause the raster cache to be marked unclean
        Me.NotifyOfDestructiveChanges
        SetVectorDataFromXML_Legacy = True
    
    'If the layer XML data can't be validated, there's nothing we can do but exit.
    Else
        Debug.Print "Vector layer could not be initialized: are you sure the supplied data is valid?"
        SetVectorDataFromXML_Legacy = False
    End If
    
End Function
