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
'Copyright 2012-2015 by Tanner Helland
'Created: 29/August/12
'Last updated: 16/October/15
'Last update: implement LOD caching for use by pdCompositor
'
'Once upon a time, the pdLayer class was simply a DIB wrapper.  Once PhotoDemon moved to a fully layer-compatible
' image editor, however, it became necessary to separate Layer-specific functions from generic DIB functions.
'
'The current pdLayer class is a comprehensive wrapper for all types of PhotoDemon layers.  All layers share a few
' key components: position, blend mode, etc, and all (except adjustment layers) are expected to provide a copy of
' the layer's contents in DIB format.  The main viewport renderer doesn't care how this DIB is created - e.g.
' whether it comes from text or images or something fancy like vector rendering - it just cares that the DIB is
' present and appropriately sized for its contents.
'
'Because layers interact heavily with PDI files (PhotoDemon's custom file format), each layer must be capable of
' writing its relevant header data into an XML string, and recreating itself from such a string if necessary.
' Layer DIB data is always handled separately, as some layer types may need to recreate their DIBs on-the-fly.
'
'Layer blend modes are declared here, but the actual blending of layers happens in the pdCompositor class.  This was
' done to enable the blending of arbitrary DIBs not attached to layers, and to simplify this class (which is
' already quite large).  Note also that pdLayer objects have no knowledge of their parent pdImage object.  If
' interaction with the parent object is important, functions will explicitly request a pdImage reference as part
' of their function parameters.
'
'All source code in this file is licensed under a modified BSD license. This means you may use the code in your own
' projects IF you provide attribution. For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Public Enum LAYER_TYPE
    PDL_IMAGE = 0
    PDL_TEXT = 1
    PDL_TYPOGRAPHY = 2
    PDL_ADJUSTMENT = 3
End Enum

#If False Then
    Const PDL_IMAGE = 0, PDL_TEXT = 1, PDL_TYPOGRAPHY = 2, PDL_ADJUSTMENT = 3
#End If

'Layer type is set at creation time.  Once a layer has been created, its type should not be changed.
' The one exception to this rule is converting vector layers (text, etc) to image layers.  This can be done, but it is a one-way conversion.
' Once a layer is converted to PDL_IMAGE type, it cannot be reverted.
Private myLayerType As LAYER_TYPE

Public Enum LAYER_BLENDMODE
    BL_NORMAL = 0
    BL_DARKEN = 1
    BL_MULTIPLY = 2
    BL_COLORBURN = 3
    BL_LINEARBURN = 4
    BL_LIGHTEN = 5
    BL_SCREEN = 6
    BL_COLORDODGE = 7
    BL_LINEARDODGE = 8
    BL_OVERLAY = 9
    BL_SOFTLIGHT = 10
    BL_HARDLIGHT = 11
    BL_VIVIDLIGHT = 12
    BL_LINEARLIGHT = 13
    BL_PINLIGHT = 14
    BL_HARDMIX = 15
    BL_DIFFERENCE = 16
    BL_EXCLUSION = 17
    BL_SUBTRACT = 18
    BL_DIVIDE = 19
    BL_HUE = 20
    BL_SATURATION = 21
    BL_COLOR = 22
    BL_LUMINOSITY = 23
    BL_GRAINEXTRACT = 24
    BL_GRAINMERGE = 25
End Enum

#If False Then
    Const BL_NORMAL = 0, BL_DARKEN = 1, BL_MULTIPLY = 2, BL_COLORBURN = 3, BL_LINEARBURN = 4, BL_LIGHTEN = 5, BL_SCREEN = 6, BL_COLORDODGE = 7
    Const BL_LINEARDODGE = 8, BL_OVERLAY = 9, BL_SOFTLIGHT = 10, BL_HARDLIGHT = 11, BL_VIVIDLIGHT = 12, BL_LINEARLIGHT = 13, BL_PINLIGHT = 14
    Const BL_HARDMIX = 15, BL_DIFFERENCE = 16, BL_EXCLUSION = 17, BL_SUBTRACT = 18, BL_DIVIDE = 19
    Const BL_HUE = 20, BL_SATURATION = 21, BL_COLOR = 22, BL_LUMINOSITY = 23, BL_GRAINEXTRACT = 24, BL_GRAINMERGE = 25
#End If

'Layer resize quality is defined different from other resample places in the project.  (Only a subset of options are exposed, for performance reasons.)
Public Enum LAYER_RESIZE_QUALITY
    LRQ_NEAREST_NEIGHBOR = 0
    LRQ_BILINEAR_NORMAL = 1
    LRQ_BICUBIC_NORMAL = 2
End Enum

#If False Then
    Const LRQ_NEAREST_NEIGHBOR = 0, LRQ_BILINEAR_NORMAL = 1, LRQ_BICUBIC_NORMAL = 2
#End If

'PD supports the notion "alpha modes", including "inheritance," where a layer "inherits" the alpha of a layer beneath it,
' (https://userbase.kde.org/Krita/Tutorial_2#Inherit_Alpha_.28alpha_.3D_transparency.29).  I may or may not choose to
' address masking via this property as well... I'm still deciding the best way to do it.
Public Enum LAYER_ALPHAMODE
    LA_NORMAL = 0
    LA_INHERIT = 1
End Enum

#If False Then
    Const LA_NORMAL = 0, LA_INHERIT = 1
#End If

'Canonical ID value for this layer.  This value is set once, at layer creation time, and can never be changed again.  It is
' persistent for the life of the layer, the life of the parent pdImage object (meaning no other layer in that image will
' ever be assigned this ID value), and the life of any copies of the parent pdImage saved to file.
'
'Because this value cannot be changed once created, I keep it separate from the LayerData struct (whose data can be changed).
Private layerID As Long

'Layers store a lot of layer-type-agnostic metadata. **All of these entries can be changed by the user at run-time.**
' Default values will be set at creation time, but never assume that these values are constant, and never attempt to reference
'  a layer by one of these properties (INCLUDING NAME, as it can be modified without warning)
Private Type LayerData
    l_Name As String                    'Layer name, as entered by the user
    l_GroupID As Long                   'Layer group (0 if not assigned to a group)
    l_Opacity As Long                   'Layer opacity (defaults to 100)
    l_BlendMode As LAYER_BLENDMODE      'Layer blend mode; defaults to BL_NORMAL
    l_OffsetX As Double                 'X/Y offset of the layer's top-left corner
    l_OffsetY As Double
    l_CanvasXModifier As Double         'Width/height modifier for the image; by default, these are set to 1.0.
    l_CanvasYModifier As Double
    l_Angle As Double                   'Layer angle (implemented in PD 7.0)
    l_Visibility As Boolean             'Layer visibility
    l_NonDestructiveFXActive As Boolean 'Are any non-destructive edits or other FX active?
    l_ResizeQuality As LAYER_RESIZE_QUALITY   'Resampling algorithm used when non-destructive resizes are active
    l_ShearX As Double                  'Layer shear (added in PD 7.0)
    l_ShearY As Double
    l_AlphaMode As LAYER_ALPHAMODE      'Layer alpha mode; defaults to LA_NORMAL
End Type

'Local instance of layer data for this layer
Private myLayerData As LayerData

'This class supports getting/setting layer properties via generic Get/SetGenericLayerProperty functions.  This enum is
' used to differentiate between layer properties, and any additions to LayerData, above, should be mirrored here.
Public Enum PDLAYER_GENERIC_PROPERTY
    pgp_Name = 0
    pgp_GroupID = 1
    pgp_Opacity = 2
    pgp_BlendMode = 3
    pgp_OffsetX = 4
    pgp_OffsetY = 5
    pgp_CanvasXModifier = 6
    pgp_CanvasYModifier = 7
    pgp_Angle = 8
    pgp_Visibility = 9
    pgp_NonDestructiveFXActive = 10
    pgp_ResizeQuality = 11
    pgp_ShearX = 12
    pgp_ShearY = 13
    pgp_AlphaMode = 14
End Enum

#If False Then
    Const pgp_Name = 0, pgp_GroupID = 1, pgp_Opacity = 2, pgp_BlendMode = 3, pgp_OffsetX = 4, pgp_OffsetY = 5, pgp_CanvasXModifier = 6
    Const pgp_CanvasYModifier = 7, pgp_Angle = 8, pgp_Visibility = 9, pgp_NonDestructiveFXActive = 10, pgp_ResizeQuality = 11
    Const pgp_ShearX = 12, pgp_ShearY = 13, pgp_AlphaMode = 14
#End If

'Non-destructive edits currently supported by the program.
Public Enum LAYER_NONDESTRUCTIVE_FX
    NDFX_EXPOSURE = 0
    NDFX_CONTRAST = 1
    NDFX_CLARITY = 2
    NDFX_VIBRANCE = 3
    NDFX_TEMPERATURE = 4
    NDFX_TINT = 5
End Enum

#If False Then
    Const NDFX_EXPOSURE = 0, NDFX_CONTRAST = 1, NDFX_CLARITY = 2, NDFX_VIBRANCE = 3, NDFX_TEMPERATURE = 4, NDFX_TINT = 5
#End If

'Local copy of non-destructive FX
Private Type nonDestructiveFXSettings
    fx_Exposure As Double
    fx_Contrast As Long
    fx_Clarity As Long
    fx_Vibrance As Long
    fx_Temperature As Long
    fx_Tint As Long
End Type

Private myNonDestructiveFX As nonDestructiveFXSettings

'Vector-type layers store many things that image layers do not.  Generic data shared between all vector layer types are
' stored in this struct.
Private Type genericVectorData
    vd_Width As Long          'Unlike raster layers, vector layers must explicitly store their width and height
    vd_Height As Long
End Type

Private myVectorData As genericVectorData

'Vector layers must be converted into raster format for rendering on the screen.  If the DIB is synched against the
' the current vector contents, this will be set to TRUE.  If it is FALSE, the DIB must be generated prior to display.
Private m_VectorDIBSynched As Boolean

'Next comes various structs for vector data in a specific format.  These structs must be comprehensive, because their data
' is used to reconstruct the raster copy of the layer between sessions.

'TEXT LAYER STRUCTS

'A dedicated text renderer is used to render the backer layer for text layers.  It also manages all relevant text data.
Private myTextRenderer As pdTextRenderer

'A DIB that stores the contents of the layer, in DIB format.
' - For image layers, the layer contents will be stored in this DIB.
' - For non-image layers, a rendering of the layer contents will be stored in this DIB.  This DIB is used to render the layer
'    onto the viewport, rather than constantly recreating whatever the layer's contents may be (text, etc)
' - For adjustment and other non-data layers, this DIB will remain unused.
'
'For performance reasons, I have chosen to make layer DIBs publicly accessible.  This allows edit functions to directly modify
' them as necessary.  That said, any function that directly modifies the DIB must also take care of any management functions
' as well (e.g. redrawing the viewport), because layers have no way to propagate change notifications up to their parent image.
' This is typically handled by a called to the parent's .notifyImageChanged function, and passing the relevant layer index.
Public layerDIB As pdDIB

'Temporary level-of-detail (LOD) compositing DIB(s).  These DIBs are only used under certain circumstances; at present, they are
' restricted to use by the pdCompositor class.  Because it's often necessary to generate temporary DIBs on a per-layer basis,
' these level-of-detail DIBs can save us a ton of time during performance-crucial operations (e.g. painting).
Private lodDIB() As pdDIB

'Layer thumbnail.  This 256x256 DIB is cached locally, then passed to external functions only when necessary.  Note that requests
' for sizes larger than 256x256 (in either dimension) will force creation of a custom thumbnail from scratch, so try to avoid that
' if at all possible!
Private Const LAYER_THUMB_SIZE As Long = 256
Private layerThumbnail As pdDIB

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

'If a layer has non-destructive, non-standard transformations active (e.g. rotation, skew), we use a special, parallelogram-based rendering function.
' To cut down on variable declarations, we declare this array once, and initialize it when a compositor instance is initialized.
Private m_PlgPoints() As POINTFLOAT

'If an external function modifies this layer's DIB contents (or the vector data from which the DIB is generated),
' it must call this function afterward.
'
'IMPORTANT NOTE! As of v7.0, PD functions should not call this sub directly.  Instead, they must call notifyImageChanged on
' the parent pdImage object, so the master object knows to recomposite itself (which is necessary for some viewport modes).
'
'ANOTHER IMPORTANT NOTE!  This function only needs to be called if a change was *destructive*.  Non-destructive changes
' - per the name, lol - do not require us to rebuild internal caches, because the effects and/or resizing will be
' handled by a compositor object.
'
'Finally, this function will mark all layer caches as dirty, causing a lot of subsequent work, so use it only when
' absolutely necessary.
Public Sub notifyOfDestructiveChanges()

    'Mark the layer thumbnail as dirty
    m_IsThumbClean = False
    
    'Note that the layer has changed since a past hash request
    Dim i As Long
    For i = 0 To NUM_OF_LOD_CACHES - 1
        m_DIBChangedSinceLastHash(i) = True
    Next i
    
    'If this is a vector layer, note that the backing DIB is now out of sync
    m_VectorDIBSynched = False
    
End Sub

'Get/set this layer's canonical ID value.  Note that this value is valid for both the life of the layer, and the life of its
' parent image (including persistence when writing image data to/from file).
'
'Note that the set function is preferenced by "assign" - I do this on purpose, to help me remember that the function should only
' ever be called once, right after the layer is first created.  After that, it should never, ever be changed!
Public Function getLayerID() As Long
    getLayerID = layerID
End Function

Public Sub assignLayerID(ByVal thisLayerID As Long)
    
    'As a failsafe, warn me if this layer has already been assigned an ID.  (This should *never* happen, but I like to verify.)
    If Not g_IsProgramCompiled Then
        If layerID <> -1 Then Debug.Print "WARNING! This layer already has an ID!"
    End If
    
    layerID = thisLayerID
    
End Sub

'Get layer type.
Public Function getLayerType() As LAYER_TYPE
    getLayerType = myLayerType
End Function

'Set layer type.  Note that this only sets the layer type flag; changing the actual layer contents to match (e.g. rasterization)
' must be handled separately.
Public Function setLayerType(ByVal newLayerType As LAYER_TYPE)
    myLayerType = newLayerType
    m_VectorDIBSynched = False
End Function

'If a caller simply wants to determine vector vs raster layers, this helper function takes care of it.  Note that text layers are
' considered vector layers, because all changes are non-destructive.
Public Function isLayerVector() As Boolean

    Select Case myLayerType
    
        Case PDL_IMAGE
            isLayerVector = False
        
        Case PDL_TEXT, PDL_TYPOGRAPHY
            isLayerVector = True
        
        Case Else
            Debug.Print "Layer type unknown!  Returning VECTOR = FALSE by default."
    
    End Select

End Function

'Counterpart to isLayerVector(), above
Public Function isLayerRaster() As Boolean
    isLayerRaster = Not isLayerVector()
End Function

'If a caller simply wants to determine text vs other layers, this helper function takes care of it.  Text layers currently include text and
' typography layers, and may include additional types (e.g. "WordArt") in the future
Public Function isLayerText() As Boolean

    Select Case myLayerType
    
        Case PDL_TEXT, PDL_TYPOGRAPHY
            isLayerText = True
        
        Case Else
            isLayerText = False
    
    End Select

End Function

'Rasterize a layer.  In the future, it may be possible to "Undo" this (e.g. restore the original vector state, since that data is
' easy to retain), but I haven't implemented it at present.  Rasterization is currently immutable, short of the Undo/Redo menu.
Public Sub rasterizeVectorData()

    'If the layer is already rasterized, ignore this request
    If Me.isLayerRaster Then
        Debug.Print "WARNING!  You just requested rasterization of a raster layer.  Fix this!"
        Exit Sub
    End If
    
    'Rasterizing is a lot simpler than it looks.  Basically, we just change the layer type to PDL_IMAGE and discard any superfluous vector data.
    myLayerType = PDL_IMAGE
        
    'While possibly unnecessary for some vector layer types, rasterization is a serious enough change that it's wise to regenerate
    ' any/all internal caches.
    Me.notifyOfDestructiveChanges

End Sub

'Get/set name
Public Function getLayerName() As String
    getLayerName = myLayerData.l_Name
End Function

Public Sub setLayerName(ByRef newLayerName As String)
    myLayerData.l_Name = newLayerName
End Sub

'Get/set layer group
Public Function getLayerGroup() As Long
    getLayerGroup = myLayerData.l_GroupID
End Function

Public Sub setLayerGroup(ByVal newLayerGroup As Long)
    myLayerData.l_GroupID = newLayerGroup
End Sub

'Get/set opacity
Public Function getLayerOpacity() As Long
    getLayerOpacity = myLayerData.l_Opacity
End Function

Public Sub setLayerOpacity(ByVal newLayerOpacity As Long)
    
    If newLayerOpacity > 100 Then newLayerOpacity = 100
    If newLayerOpacity < 0 Then newLayerOpacity = 0
    myLayerData.l_Opacity = newLayerOpacity
    
End Sub

'Get/set alpha and blend modes
Public Function getLayerAlphaMode() As LAYER_ALPHAMODE
    getLayerAlphaMode = myLayerData.l_AlphaMode
End Function

Public Sub setLayerAlphaMode(ByVal newLayerAlphaMode As LAYER_ALPHAMODE)
    myLayerData.l_AlphaMode = newLayerAlphaMode
End Sub

Public Function getLayerBlendMode() As LAYER_BLENDMODE
    getLayerBlendMode = myLayerData.l_BlendMode
End Function

Public Sub setLayerBlendMode(ByVal newLayerBlendMode As LAYER_BLENDMODE)
    myLayerData.l_BlendMode = newLayerBlendMode
End Sub

'Get/set visibility
Public Function getLayerVisibility() As Boolean
    getLayerVisibility = myLayerData.l_Visibility
End Function

Public Sub setLayerVisibility(ByVal newVisibility As Boolean)
    myLayerData.l_Visibility = newVisibility
End Sub

'Get/set layer offsets
Public Function getLayerOffsetX() As Double
    getLayerOffsetX = myLayerData.l_OffsetX
End Function

Public Function getLayerOffsetY() As Double
    getLayerOffsetY = myLayerData.l_OffsetY
End Function

Public Sub setLayerOffsetX(ByVal newOffsetX As Double)
    myLayerData.l_OffsetX = newOffsetX
End Sub

Public Sub setLayerOffsetY(ByVal newOffsetY As Double)
    myLayerData.l_OffsetY = newOffsetY
End Sub

'Helper set function for setting layer offset and non-destructive width/height all at once.
Friend Sub setOffsetsAndModifiersTogether(ByVal topLeftX As Single, ByVal topLeftY As Single, ByVal bottomRightX As Single, ByVal bottomRightY As Single)
    
    With myLayerData
        
        .l_OffsetX = topLeftX
        .l_OffsetY = topLeftY
        
        If getLayerWidth(False) <> 0 Then
            setLayerCanvasXModifier (bottomRightX - topLeftX) / getLayerWidth(False)
        Else
            setLayerCanvasXModifier 1#
        End If
        
        If getLayerHeight(False) <> 0 Then
            setLayerCanvasYModifier (bottomRightY - topLeftY) / getLayerHeight(False)
        Else
            setLayerCanvasYModifier 1#
        End If
        
    End With
    
End Sub

'Get/set layer size.  Note that the corresponding "set size" functions are ONLY VALID FOR VECTOR LAYERS; because size changes are destructive
' for raster layers, they must be handled manually.  (The exception, of course, being the non-destructive X/Y canvas modifier values, which exist
' for this explicit purpose!)
'
'Also, at present, these values are not stored persistently.  They are simply generated on-the-fly by taking the layerDIB object's dimensions,
' and multiplying them by the current canvas x/y modifiers.  The caller can override the canvas modifier calculation by passing FALSE to the function.
Public Function getLayerWidth(Optional ByVal applyCanvasModifier As Boolean = True) As Double
    
    'The returned width varies by layer type
    Select Case myLayerType
    
        'Image layers do not store persistent width/height values.  Return the layer DIB's dimensions in its place
        Case PDL_IMAGE
            
            If Not (layerDIB Is Nothing) Then
        
                If applyCanvasModifier Then
                    getLayerWidth = layerDIB.getDIBWidth * getLayerCanvasXModifier
                    If getLayerWidth < 1 Then getLayerWidth = 1
                Else
                    getLayerWidth = layerDIB.getDIBWidth
                End If
                
            Else
                getLayerWidth = 0
            End If
        
        'Vector layers do actually store persistent width/height values.  This is important because these values may fall
        ' out of sync with the backer DIB, which is only updated on an as-needed basis
        Case PDL_TEXT, PDL_TYPOGRAPHY
            getLayerWidth = myVectorData.vd_Width
    
    End Select
    
End Function

Public Function getLayerHeight(Optional ByVal applyCanvasModifier As Boolean = True) As Double
    
    'The returned height varies by layer type
    Select Case myLayerType
    
        'Image layers do not store persistent width/height values.  Return the layer DIB's dimensions in its place
        Case PDL_IMAGE
            
            If Not (layerDIB Is Nothing) Then
        
                If applyCanvasModifier Then
                    getLayerHeight = layerDIB.getDIBHeight * getLayerCanvasYModifier
                    If getLayerHeight < 1 Then getLayerHeight = 1
                Else
                    getLayerHeight = layerDIB.getDIBHeight
                End If
                
            Else
                getLayerHeight = 0
            End If
        
        'Vector layers do actually store persistent width/height values.  This is important because these values may fall
        ' out of sync with the backer DIB, which is only updated on an as-needed basis
        Case PDL_TEXT, PDL_TYPOGRAPHY
            getLayerHeight = myVectorData.vd_Height
    
    End Select
    
End Function

Public Sub setLayerWidth(ByVal newSize As Single)
    
    'Force size in-bounds
    If newSize < 1 Then newSize = 1
    
    'Make sure the layer is a vector layer
    If (myLayerType <> PDL_IMAGE) And (newSize <> myVectorData.vd_Width) Then
        
        'Store the new size and mark the DIB cache as dirty
        myVectorData.vd_Width = newSize
        m_VectorDIBSynched = False
        
    Else
        If myLayerType = PDL_IMAGE Then Debug.Print "WARNING!  The setLayerWidth and setLayerHeight functions are only valid for vector layers!"
    End If
    
End Sub

Public Sub setLayerHeight(ByVal newSize As Single)
    
    'Force size in-bounds
    If newSize < 1 Then newSize = 1
    
    'Make sure the layer is a vector layer
    If (myLayerType <> PDL_IMAGE) And (newSize <> myVectorData.vd_Width) Then
        
        'Store the new size and mark the DIB cache as dirty
        myVectorData.vd_Height = newSize
        m_VectorDIBSynched = False
        
    Else
        If myLayerType = PDL_IMAGE Then Debug.Print "WARNING!  The setLayerWidth and setLayerHeight functions are only valid for vector layers!"
    End If
    
End Sub

'Get/set layer canvas modifiers (necessary for non-destructive on-canvas resizing).  For vector layers, these have no effect.
' Vector layers always maintain canvas x/y modifiers of 1, and they simply translate these requests into setLayerWidth/Height calls.
Public Function getLayerCanvasXModifier() As Double
    getLayerCanvasXModifier = myLayerData.l_CanvasXModifier
End Function

Public Function getLayerCanvasYModifier() As Double
    getLayerCanvasYModifier = myLayerData.l_CanvasYModifier
End Function

Public Sub setLayerCanvasXModifier(ByVal newXModifier As Double)
    
    'Check bounds
    If newXModifier <= 0 Then newXModifier = 0.00000001
    
    'Separate handling by layer type
    Select Case myLayerType
    
        Case PDL_IMAGE
            myLayerData.l_CanvasXModifier = newXModifier
            
        Case PDL_TEXT, PDL_TYPOGRAPHY
            myLayerData.l_CanvasXModifier = 1
            setLayerWidth newXModifier * getLayerWidth
    
    End Select
    
End Sub

Public Sub setLayerCanvasYModifier(ByVal newYModifier As Double)
    
    'Check bounds
    If newYModifier <= 0 Then newYModifier = 0.00000001
    
    'Separate handling by layer type
    Select Case myLayerType
    
        Case PDL_IMAGE
            myLayerData.l_CanvasYModifier = newYModifier
            
        Case PDL_TEXT, PDL_TYPOGRAPHY
            myLayerData.l_CanvasYModifier = 1
            setLayerHeight newYModifier * getLayerHeight
    
    End Select
    
End Sub

'Get/set non-destructive resize algorithm.  Note that there are two GET functions:
' - One that uses PD's internal LAYER_RESIZE_QUALITY enum
' - One that converts the internal LAYER_RESIZE_QUALITY value into a GDI+ resize constant
Public Function getLayerResizeQuality() As LAYER_RESIZE_QUALITY
    getLayerResizeQuality = myLayerData.l_ResizeQuality
End Function

Public Function getLayerResizeQuality_GDIPlus() As Long
    
    Select Case myLayerData.l_ResizeQuality
        
        Case LRQ_NEAREST_NEIGHBOR
            getLayerResizeQuality_GDIPlus = InterpolationModeNearestNeighbor
        
        Case LRQ_BILINEAR_NORMAL
            getLayerResizeQuality_GDIPlus = InterpolationModeHighQualityBilinear
        
        Case LRQ_BICUBIC_NORMAL
            getLayerResizeQuality_GDIPlus = InterpolationModeHighQualityBicubic
        
    End Select
    
End Function

Public Sub setLayerResizeQuality(ByVal newQuality As LAYER_RESIZE_QUALITY)
    myLayerData.l_ResizeQuality = newQuality
End Sub

'Get/set non-destructive FX status.  Note that these functions work a little differently than other get/set functions.  Get() works
' normally, but Set() requires the user to activate or deactivate a specific
Public Function getLayerNonDestructiveFXState() As Boolean
    getLayerNonDestructiveFXState = myLayerData.l_NonDestructiveFXActive
End Function

Public Function getLayerNonDestructiveFXValue(ByVal ndEffect As LAYER_NONDESTRUCTIVE_FX) As Double
    
    'Return a new value contigent on the supplied ndEffect value
    Select Case ndEffect
    
        'Exposure
        Case NDFX_EXPOSURE
            getLayerNonDestructiveFXValue = myNonDestructiveFX.fx_Exposure
            
        'Contrast
        Case NDFX_CONTRAST
            getLayerNonDestructiveFXValue = myNonDestructiveFX.fx_Contrast
            
        'Clarity
        Case NDFX_CLARITY
            getLayerNonDestructiveFXValue = myNonDestructiveFX.fx_Clarity
            
        'Vibrance
        Case NDFX_VIBRANCE
            getLayerNonDestructiveFXValue = myNonDestructiveFX.fx_Vibrance
            
        'Temperature
        Case NDFX_TEMPERATURE
            getLayerNonDestructiveFXValue = myNonDestructiveFX.fx_Temperature
            
        'Tint
        Case NDFX_TINT
            getLayerNonDestructiveFXValue = myNonDestructiveFX.fx_Tint
            
    End Select
    
End Function

Public Sub setLayerNonDestructiveFXState(ByVal ndEffect As LAYER_NONDESTRUCTIVE_FX, Optional ByVal newEffectValue As Double = 0#)
    
    'Assign the new value to the appropriate entry in this layer's non-destructive FX array
    Select Case ndEffect
    
        'Exposure
        Case NDFX_EXPOSURE
            myNonDestructiveFX.fx_Exposure = newEffectValue
            
        'Contrast
        Case NDFX_CONTRAST
            myNonDestructiveFX.fx_Contrast = newEffectValue
        
        'Clarity
        Case NDFX_CLARITY
            myNonDestructiveFX.fx_Clarity = newEffectValue
            
        'Vibrance
        Case NDFX_VIBRANCE
            myNonDestructiveFX.fx_Vibrance = newEffectValue
            
        'Temperature
        Case NDFX_TEMPERATURE
            myNonDestructiveFX.fx_Temperature = newEffectValue
            
        'Tint
        Case NDFX_TINT
            myNonDestructiveFX.fx_Tint = newEffectValue
            
    End Select
    
    'Is the newEffectValue 0?  If it is, check to see if any other FX have non-zero values; if all values are zero, we can mark
    ' the layer's non-destructive FX state as FALSE, but if any are non-zero, it must remain as TRUE.
    If newEffectValue = 0 Then
    
        myLayerData.l_NonDestructiveFXActive = False
        
        'Check each entry in the local non-destructive effects cache; if any still have non-zero values, note that the layer still
        ' has active non-destructive effects.
        With myNonDestructiveFX
        
            If .fx_Exposure <> 0 Then myLayerData.l_NonDestructiveFXActive = True
            If .fx_Contrast <> 0 Then myLayerData.l_NonDestructiveFXActive = True
            If .fx_Clarity <> 0 Then myLayerData.l_NonDestructiveFXActive = True
            If .fx_Vibrance <> 0 Then myLayerData.l_NonDestructiveFXActive = True
            If .fx_Temperature <> 0 Then myLayerData.l_NonDestructiveFXActive = True
            If .fx_Tint <> 0 Then myLayerData.l_NonDestructiveFXActive = True
        
        End With
    
    Else
        myLayerData.l_NonDestructiveFXActive = True
    End If
    
End Sub

'This shortcut function can be used to deactivate all NDFX settings at once
Public Sub deactivateAllNDFX()
    
    myLayerData.l_NonDestructiveFXActive = False
    
    With myNonDestructiveFX
        .fx_Clarity = 0
        .fx_Contrast = 0
        .fx_Exposure = 0
        .fx_Temperature = 0
        .fx_Tint = 0
        .fx_Vibrance = 0
    End With
    
End Sub

'Generic layer properties can be get/set via these singular functions.  Note that these functions get and return Variants,
' as layer properties include a lot of different types.  As such, these aren't suitable for use inside a performance-sensitive
' loop (like pdCompositor, for example).  Instead, these are primarily used by UI elements and the Undo/Redo processor, where
' it is much simpler to use a single generic function.
Friend Function getGenericLayerProperty(ByVal desiredProperty As PDLAYER_GENERIC_PROPERTY) As Variant

    Select Case desiredProperty
    
        Case pgp_Name
            getGenericLayerProperty = myLayerData.l_Name
            
        Case pgp_GroupID
            getGenericLayerProperty = myLayerData.l_GroupID
            
        Case pgp_Opacity
            getGenericLayerProperty = myLayerData.l_Opacity
            
        Case pgp_BlendMode
            getGenericLayerProperty = myLayerData.l_BlendMode
            
        Case pgp_OffsetX
            getGenericLayerProperty = myLayerData.l_OffsetX
            
        Case pgp_OffsetY
            getGenericLayerProperty = myLayerData.l_OffsetY
            
        Case pgp_CanvasXModifier
            getGenericLayerProperty = myLayerData.l_CanvasXModifier
            
        Case pgp_CanvasYModifier
            getGenericLayerProperty = myLayerData.l_CanvasYModifier
            
        Case pgp_Angle
            getGenericLayerProperty = myLayerData.l_Angle
            
        Case pgp_Visibility
            getGenericLayerProperty = myLayerData.l_Visibility
            
        Case pgp_NonDestructiveFXActive
            getGenericLayerProperty = myLayerData.l_NonDestructiveFXActive
            
        Case pgp_ResizeQuality
            getGenericLayerProperty = myLayerData.l_ResizeQuality
            
        Case pgp_ShearX
            getGenericLayerProperty = myLayerData.l_ShearX
            
        Case pgp_ShearY
            getGenericLayerProperty = myLayerData.l_ShearY
            
        Case pgp_AlphaMode
            getGenericLayerProperty = myLayerData.l_AlphaMode
    
    End Select

End Function

Friend Sub setGenericLayerProperty(ByVal desiredProperty As PDLAYER_GENERIC_PROPERTY, ByVal newValue As Variant)

    Select Case desiredProperty
        
        Case pgp_Name
            If myLayerData.l_Name <> CStr(newValue) Then myLayerData.l_Name = CStr(newValue)
            
        Case pgp_GroupID
            If myLayerData.l_GroupID <> CLng(newValue) Then myLayerData.l_GroupID = CLng(newValue)
            
        Case pgp_Opacity
            If myLayerData.l_Opacity <> CLng(newValue) Then Me.setLayerOpacity CLng(newValue)
            
        Case pgp_BlendMode
            If myLayerData.l_BlendMode <> CLng(newValue) Then myLayerData.l_BlendMode = CLng(newValue)
            
        Case pgp_OffsetX
            If myLayerData.l_OffsetX <> CDbl(newValue) Then myLayerData.l_OffsetX = CDbl(newValue)
            
        Case pgp_OffsetY
            If myLayerData.l_OffsetY <> CDbl(newValue) Then myLayerData.l_OffsetY = CDbl(newValue)
            
        Case pgp_CanvasXModifier
            If myLayerData.l_CanvasXModifier <> CDbl(newValue) Then Me.setLayerCanvasXModifier CDbl(newValue)
            
        Case pgp_CanvasYModifier
            If myLayerData.l_CanvasYModifier <> CDbl(newValue) Then Me.setLayerCanvasYModifier CDbl(newValue)
            
        Case pgp_Angle
            If myLayerData.l_Angle <> CDbl(newValue) Then myLayerData.l_Angle = CDbl(newValue)
            
        Case pgp_Visibility
            If myLayerData.l_Visibility <> CBool(newValue) Then myLayerData.l_Visibility = CBool(newValue)
            
        Case pgp_NonDestructiveFXActive
            If myLayerData.l_NonDestructiveFXActive <> CBool(newValue) Then myLayerData.l_NonDestructiveFXActive = CBool(newValue)
            
        Case pgp_ResizeQuality
            If myLayerData.l_ResizeQuality <> CLng(newValue) Then myLayerData.l_ResizeQuality = CLng(newValue)
            
        Case pgp_ShearX
            If myLayerData.l_ShearX <> CDbl(newValue) Then myLayerData.l_ShearX = CDbl(newValue)
            
        Case pgp_ShearY
            If myLayerData.l_ShearY <> CDbl(newValue) Then myLayerData.l_ShearY = CDbl(newValue)
            
        Case pgp_AlphaMode
            If myLayerData.l_AlphaMode <> CLng(newValue) Then myLayerData.l_AlphaMode = CLng(newValue)
        
    End Select
    
End Sub

'Text layer properties can be get/set via these singular functions.
Friend Function getTextLayerProperty(ByVal desiredProperty As PD_TEXT_PROPERTY) As Variant

    'All text properties are managed by the pdTextRenderer class.  This function is simply a passthrough.
    getTextLayerProperty = myTextRenderer.getGenericTextProperty(desiredProperty)
    
End Function

Friend Sub setTextLayerProperty(ByVal desiredProperty As PD_TEXT_PROPERTY, ByVal newValue As Variant)
    
    'All text properties are managed by the pdTextRenderer class.  This function is simply a passthrough.
    
    'Note that the setGenericTextProperty function will return TRUE if the requested change matters (e.g. if the text needs to be re-rendered
    ' due to this property change).  We modify our internal VectorDIBSynched tracker accordingly.
    m_VectorDIBSynched = m_VectorDIBSynched And (Not myTextRenderer.setGenericTextProperty(desiredProperty, newValue))
    
End Sub

'If this layer needs to have its on-canvas transformations made permanent (for example, when applying an action
' that involves a selection), this function can be used.
Public Sub makeCanvasTransformsPermanent(Optional ByVal maxWidth As Long = 0, Optional ByVal maxHeight As Long = 0)

    'If this layer does not have any active transforms, ignore this request
    If Me.affineTransformsActive(True) Then
        
        'Unfortunately, a temporary DIB will be required for the rescale operation
        Dim tmpTransformDIB As pdDIB
        Set tmpTransformDIB = New pdDIB
        
        'If complex affine transforms are active (e.g. rotation), this step becomes more complicated.
        If Me.affineTransformsActive(False) Then
        
            'Affine transforms can cause the layer's x/y offset to change
            Dim newOffsetX As Long, newOffsetY As Long
            
            'Retrieve a copy of the current DIB using the layer's automated affine-transformation function
            Me.getAffineTransformedDIB tmpTransformDIB, newOffsetX, newOffsetY, maxWidth, maxHeight
            
            'Update our internal DIB and offsets to match
            Set Me.layerDIB = tmpTransformDIB
            With myLayerData
                .l_OffsetX = newOffsetX
                .l_OffsetY = newOffsetY
                .l_CanvasXModifier = 1
                .l_CanvasYModifier = 1
                .l_Angle = 0
                .l_ShearX = 0
                .l_ShearY = 0
            End With
        
        'If rescaling is the only active transform, we can shortcut this function
        Else
            
            With tmpTransformDIB
                .createBlank getLayerWidth(True), getLayerHeight(True), 32, 0
                GDIPlusResizeDIB tmpTransformDIB, 0, 0, .getDIBWidth, .getDIBHeight, layerDIB, 0, 0, getLayerWidth(False), getLayerHeight(False), Me.getLayerResizeQuality_GDIPlus
            End With
            
            Set Me.layerDIB = tmpTransformDIB
            
            'Reset the width/height modifiers for this layer
            With myLayerData
                .l_CanvasXModifier = 1
                .l_CanvasYModifier = 1
            End With
        
        End If
                
        'Mark this layer as dirty
        Me.notifyOfDestructiveChanges
        
    End If

End Sub

'Get/set layer angle
Public Function getLayerAngle() As Double
    getLayerAngle = myLayerData.l_Angle
End Function

Public Sub setLayerAngle(ByVal newAngle As Double)
    myLayerData.l_Angle = newAngle
End Sub

'Get/set layer shear
Public Function getLayerShearX() As Double
    getLayerShearX = myLayerData.l_ShearX
End Function

Public Sub setLayerShearX(ByVal newShear As Double)
    myLayerData.l_ShearX = newShear
End Sub

Public Function getLayerShearY() As Double
    getLayerShearY = myLayerData.l_ShearY
End Function

Public Sub setLayerShearY(ByVal newShear As Double)
    myLayerData.l_ShearY = newShear
End Sub

'If one or more affine transformations are active (rotation, shear, etc), this function will return TRUE.
' This is used, for example, by pdCompositor to branch across several rendering paths.
Public Function affineTransformsActive(Optional ByVal includeScaling As Boolean = False) As Boolean
    
    'Most PD operations handle scaling separate from other affine operations, so we can typically skip this check
    If includeScaling Then
    
        If (myLayerData.l_CanvasXModifier <> 1) Or (myLayerData.l_CanvasYModifier <> 1) Then
            affineTransformsActive = True
            Exit Function
        End If
    
    End If
    
    'Rotate and shear are the only other relevant affine transforms
    If (myLayerData.l_Angle <> 0) Or (myLayerData.l_ShearX <> 0) Or (myLayerData.l_ShearY <> 0) Then
        affineTransformsActive = True
    Else
        affineTransformsActive = False
    End If
    
End Function

'If one or more affine transformations are active (rotation, skew, etc), you can call this function to retrieve three corner points
' of the layer *IN IMAGE COORDINATES*.
'
'dstPoints needs to have the bounds (0, 2) or (0, 3).  It will be filled with the following points in this order:
' 0) Top-Left
' 1) Top-Right
' 2) Bottom-Left
' 3) (optionally) Bottom-Right.  This point is optional because it can be inferred from the previous three
Friend Sub getLayerCornerCoordinates(ByRef dstPoints() As POINTFLOAT, Optional ByVal includeOffsets As Boolean = True, Optional ByVal normalizeToNewOrigin As Boolean = False, Optional ByVal normalizeX As Single = 0#, Optional ByVal normalizeY As Single = 0#, Optional ByVal includeNonDestructiveScaling As Boolean = True)
    
    'To simplify the work of translating layer boundaries by one or more affine transformations, we use a GraphicsMatrix object.
    Dim tmpMatrix As pdGraphicsMatrix
    
    'Retrieve a copy of the current layer transform matrix
    Me.getCopyOfLayerTransformationMatrix tmpMatrix, includeOffsets
    
    'Fill the dstPoints array with the current image corners *without* translation applied (but *with* scaling)
    dstPoints(0).x = 0
    dstPoints(0).y = 0
    
    dstPoints(1).x = Me.getLayerWidth(includeNonDestructiveScaling)
    dstPoints(1).y = 0
    
    dstPoints(2).x = 0
    dstPoints(2).y = Me.getLayerHeight(includeNonDestructiveScaling)
    
    'If the third point exists, fill it too
    If UBound(dstPoints) = 3 Then
        dstPoints(3).x = Me.getLayerWidth(includeNonDestructiveScaling)
        dstPoints(3).y = Me.getLayerHeight(includeNonDestructiveScaling)
    End If
    
    'Translate each point in turn
    tmpMatrix.applyMatrixToPointF dstPoints(0)
    tmpMatrix.applyMatrixToPointF dstPoints(1)
    tmpMatrix.applyMatrixToPointF dstPoints(2)
    If UBound(dstPoints) = 3 Then tmpMatrix.applyMatrixToPointF dstPoints(3)
    
    'Some compositing operations are easier if the layer is normalized to a standard origin of (0, 0).
    ' If the caller requested that, apply normalizing now.
    If normalizeToNewOrigin Then
        
        'Find the largest and smallest (x, y) coordinates
        Dim minX As Single, maxX As Single, MinY As Single, MaxY As Single
        findMaxMinOfPointFs dstPoints, UBound(dstPoints) + 1, minX, MinY, maxX, MaxY
        
        'Normalize all points against the minimum values
        Dim i As Long
        For i = 0 To UBound(dstPoints)
            dstPoints(i).x = dstPoints(i).x + (normalizeX - minX)
            dstPoints(i).y = dstPoints(i).y + (normalizeY - MinY)
        Next i
        
    End If
    
    'The temporary translation matrix will automatically be cleaned up when the function exits
    
    'Debug information can be displayed here:
    'Debug.Print "Layer corners: 0 - (" & dstPoints(0).x & ", " & dstPoints(0).y & ") 1 - (" & dstPoints(1).x & ", " & dstPoints(1).y & ") 1 - (" & dstPoints(2).x & ", " & dstPoints(2).y & ")"
    
End Sub

'As of PhotoDemon 7.0, non-destructive layer rotation is supported.  Call this function to retrieve the position of the layer's on-canvas
' rotation node *IN IMAGE COORDINATES*.  Note that an array is required, and it must be at least (0, 4) in size.  The first point contains
' the center point of the rotation node (e.g. the center of the current image), while subsequent points contain valid rotation nodes, which
' sit on the halfway point at each image boundary.
Friend Sub getLayerRotationNodeCoordinates(ByRef dstPoints() As POINTFLOAT, Optional ByVal includeOffsets As Boolean = True, Optional ByVal includeNonDestructiveScaling As Boolean = True)
    
    'Retrieve a copy of the current layer transform matrix
    Dim tmpMatrix As pdGraphicsMatrix
    Me.getCopyOfLayerTransformationMatrix tmpMatrix, includeOffsets
    
    'Fill the destination with two coordinates: the center if the image, and the position of the current rotation node.
    '  Note that both if these points are generated *without* translation applied (but *with* scaling).
    dstPoints(0).x = Me.getLayerWidth(includeNonDestructiveScaling) / 2
    dstPoints(0).y = Me.getLayerHeight(includeNonDestructiveScaling) / 2
    
    'Also, note that changes to the outer node position *must be mirrored* to the checkPointOfInterest() function below!
    dstPoints(1).x = Me.getLayerWidth(includeNonDestructiveScaling)
    dstPoints(1).y = Me.getLayerHeight(includeNonDestructiveScaling) / 2
    
    dstPoints(2).x = Me.getLayerWidth(includeNonDestructiveScaling) / 2
    dstPoints(2).y = Me.getLayerHeight(includeNonDestructiveScaling)
    
    dstPoints(3).x = 0
    dstPoints(3).y = Me.getLayerHeight(includeNonDestructiveScaling) / 2
    
    dstPoints(4).x = Me.getLayerWidth(includeNonDestructiveScaling) / 2
    dstPoints(4).y = 0
    
    'Apply the transformation matrix to each point
    Dim i As Long
    For i = 0 To 4
        tmpMatrix.applyMatrixToPointF dstPoints(i)
    Next i
    
    'The temporary translation matrix will automatically be cleaned up when the function exits
        
End Sub

'As of PhotoDemon 7.0, layers support one or more affine transformations (rotation, skew, etc).  If a caller needs a copy of the
' layer's current transformation matrix, it can use this function to retrieve it.
'
'Note that scaling is not applied, by design.  The caller should handle scaling manually.
Friend Sub getCopyOfLayerTransformationMatrix(ByRef dstMatrix As pdGraphicsMatrix, Optional ByVal applyTranslation As Boolean = True)

    'Create and/or reset the destination matrix, as necessary
    If dstMatrix Is Nothing Then
        Set dstMatrix = New pdGraphicsMatrix
    Else
        dstMatrix.ResetToIdentityMatrix
    End If
    
    With myLayerData
    
        'If an angle is active, rotate the matrix accordingly
        If .l_Angle <> 0 Then dstMatrix.RotateMatrix .l_Angle, Me.getLayerWidth(True) / 2, Me.getLayerHeight(True) / 2
        
        'If shearing is active, shear the matrix accordingly
        If (.l_ShearX <> 0) Or (.l_ShearY <> 0) Then dstMatrix.ShearMatrix .l_ShearX, .l_ShearY, Me.getLayerWidth(True) / 2, Me.getLayerHeight(True) / 2
        
        'Finally, if translation is desired, scale all points by the current layer offset
        If applyTranslation Then dstMatrix.TranslateMatrix .l_OffsetX, .l_OffsetY
        
    End With
    
End Sub

'As of PhotoDemon 7.0, layers support one or more affine transformations (rotation, skew, etc).  If the caller needs to determine
' layer boundaries with all affine transform operations considered, use this function.
Friend Sub getLayerBoundaryRect(ByRef dstRect As RECTF, Optional ByVal applyTranslation As Boolean = True, Optional ByVal normalizeToNewOrigin As Boolean = False, Optional ByVal normalizeX As Single = 0#, Optional ByVal normalizeY As Single = 0#)
    
    'Get a copy of the current layer coordinates
    Dim layerCorners() As POINTFLOAT
    ReDim layerCorners(0 To 3) As POINTFLOAT
    
    Me.getLayerCornerCoordinates layerCorners, applyTranslation, normalizeToNewOrigin, normalizeX, normalizeY
    
    'Find the largest and smallest (x, y) coordinates
    Dim minX As Single, maxX As Single, MinY As Single, MaxY As Single
    findMaxMinOfPointFs layerCorners, 4, minX, MinY, maxX, MaxY
    
    'The difference between max and min values define the boundary rect for this layer
    With dstRect
        .Left = minX
        .Top = MinY
        .Width = maxX - minX
        .Height = MaxY - MinY
    End With
    
End Sub

'Generally speaking, extracting the portion of this layer's DIB relevant to the image is really easy - just BitBlt it directly from
' the layer DIB itself!  Unfortunately, PD 7.0 greatly complicates this by introducing non-destructive affine transformations, which
' can make the layer occupy a non-rectangular portion of the image.
'
'This helper function can be used to retrieve a DIB - at 100% zoom - that contains an affine-transformed copy of the layer DIB and
' new (x, y) offsets (which are relevant for things like rotation, as the affine-transformed image will have negative offsets
' relative to the "pure" layer offsets).  Note that the destination (x, y) values will already have been transformed to integers
' for you; this function will automatically calculate offsets in a way that preserves sub-pixel positioning.
'
'I have optimized this function as much as I physically can.  Separate paths are used depending on the complexity of the transform
' (if the only active non-destructive transform is scaling, we can bypass certain steps, for example).
'
'RETURNS: TRUE if the destination DIB was filled, FALSE if the DIB lies outside the crop area (e.g. entirely off the image)
Friend Function getAffineTransformedDIB(ByRef dstDIB As pdDIB, ByRef dstX As Long, ByRef dstY As Long, Optional ByVal cropWidth As Long = 0, Optional ByVal cropHeight As Long = 0) As Boolean
    
    'Retrieve a copy of the layer corner points, in image coordinates (of course!)
    'Get a copy of the current layer coordinates
    Me.getLayerCornerCoordinates m_PlgPoints, True, False
    
    'From our boundary points, find the largest and smallest (x, y) values.  These give us a bounding box for the layer as a whole.
    Dim minX As Single, maxX As Single, MinY As Single, MaxY As Single
    findMaxMinOfPointFs m_PlgPoints, 4, minX, MinY, maxX, MaxY
    
    'Copy those values into a RECTF struct, for convenience
    Dim layerRect As RECTF
    With layerRect
        .Left = minX
        .Top = MinY
        .Width = maxX - minX
        .Height = MaxY - MinY
    End With
    
    'We now know two crucial pieces of data:
    ' 1) the final locations of the affine-transformed layer corner points, in image coordinate space
    ' 2) a bounding box for the affine-transformed layer, in image coordinate space
    
    'Next, we need to calculate the minimum boundary size required by the destination DIB.  This is the lesser of two things:
    ' 1) The crop width/height supplied by the caller (typically the size of the compositing surface itself; pixels outside
    '     this are irrelevant, so we can completely ignore them for a speed boost).
    ' 2) If the layer left/top values are negative, the max x/y values calculated above
    ' 3) If the layer left/top values are positive, the width of the layer boundary rect
    
    'During this step, we will also calculate fractional values I call "normalizeX/Y".  PD's internal compositor can only
    ' render to integer offsets, for performance reasons.  Affine transformations almost always result in an image with
    ' fractional offsets; to make sure these offsets are preserved, we "normalize" the coordinates by offseting them by
    ' their fractional amount.  This ensures that the final composite looks identical between GDI+ and PD's internal
    ' compositing code.
    
    'Start by calculating (2) or (3), and padding it to account for fractional offsets
    Dim intLayerWidth As Long, intLayerHeight As Long
    Dim normalizeX As Single, normalizeY As Single
    
    If minX < 0 Then
        dstX = 0
        normalizeX = 0
        intLayerWidth = Int(maxX + 0.999)
    Else
        dstX = Int(minX)
        normalizeX = minX - dstX
        intLayerWidth = Int(layerRect.Width + 0.999 + normalizeX)
    End If
    
    If MinY < 0 Then
        dstY = 0
        normalizeY = 0
        intLayerHeight = Int(MaxY + 0.999)
    Else
        dstY = Int(MinY)
        normalizeY = MinY - dstY
        intLayerHeight = Int(layerRect.Height + 0.999 + normalizeY)
    End If
    
    'Next, replace the results of (2) or (3) with (1), as relevant
    If cropWidth > 0 Then
        If dstX + intLayerWidth > cropWidth Then intLayerWidth = cropWidth - dstX
    End If
                    
    If cropHeight > 0 Then
        If dstY + intLayerHeight > cropHeight Then intLayerHeight = cropHeight - dstY
    End If
    
    'If the final calculated width and height are positive and non-zero, we can continue; otherwise, exit now.
    If (intLayerWidth > 0) And (intLayerHeight > 0) Then
    
        'Next, prepare the destination DIB.  If at all possible, we try to optimize this by reusing DIBs that are already
        ' an acceptable size.
        If (dstDIB.getDIBWidth <> intLayerWidth) Or (dstDIB.getDIBHeight <> intLayerHeight) Then
            dstDIB.createBlank intLayerWidth, intLayerHeight, 32, 0, 0
            dstDIB.setInitialAlphaPremultiplicationState True
        Else
            dstDIB.resetDIB
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
        If Me.affineTransformsActive(False) Then
            GDI_Plus.GDIPlus_PlgBlt dstDIB, m_PlgPoints, Me.layerDIB, 0, 0, Me.getLayerWidth(False), Me.getLayerHeight(False), 1, Me.getLayerResizeQuality_GDIPlus
        Else
            GDI_Plus.GDIPlus_StretchBlt dstDIB, normalizeX, normalizeY, Me.getLayerWidth(True), Me.getLayerHeight(True), Me.layerDIB, 0, 0, Me.getLayerWidth(False), Me.getLayerHeight(False), 1, Me.getLayerResizeQuality_GDIPlus
        End If
        
        dstDIB.setInitialAlphaPremultiplicationState True
        
        'Return success
        getAffineTransformedDIB = True
        
    Else
    
        'Return failure (e.g. the transformed DIB lies completely off the image, so there is no need for a temporary DIB)
        getAffineTransformedDIB = False
        
    End If
    
End Function

'If one or more affine transformations are active (rotation, skew, etc), you can call this function to retrieve three corner points
' of the layer *IN IMAGE COORDINATES*.

'Given an array with at least upper bound 2, find the max and min x/y values.
Private Sub findMaxMinOfPointFs(ByRef srcPoints() As POINTFLOAT, ByVal numOfPoints As Long, ByRef minX As Single, ByRef MinY As Single, ByRef maxX As Single, ByRef MaxY As Single)
    
    If srcPoints(0).x < srcPoints(1).x Then
        minX = srcPoints(0).x
        maxX = srcPoints(1).x
    Else
        minX = srcPoints(1).x
        maxX = srcPoints(0).x
    End If
    
    If srcPoints(0).y < srcPoints(1).y Then
        MinY = srcPoints(0).y
        MaxY = srcPoints(1).y
    Else
        MinY = srcPoints(1).y
        MaxY = srcPoints(0).y
    End If
    
    Dim i As Long
    For i = 2 To numOfPoints - 1
        
        If srcPoints(i).x < minX Then minX = srcPoints(i).x
        If srcPoints(i).y < MinY Then MinY = srcPoints(i).y
        
        If srcPoints(i).x > maxX Then maxX = srcPoints(i).x
        If srcPoints(i).y > MaxY Then MaxY = srcPoints(i).y
        
    Next i
    
End Sub

Private Sub Class_Initialize()

    'Assign default values to this instance
    With myLayerData
        .l_Name = g_Language.TranslateMessage("New Layer")
        .l_GroupID = 0
        .l_Opacity = 100
        .l_BlendMode = BL_NORMAL
        .l_OffsetX = 0
        .l_OffsetY = 0
        .l_CanvasXModifier = 1
        .l_CanvasYModifier = 1
        .l_Angle = 0
        .l_Visibility = True
        .l_NonDestructiveFXActive = False
        .l_ResizeQuality = LRQ_BICUBIC_NORMAL
        .l_ShearX = 0
        .l_ShearY = 0
        .l_AlphaMode = LA_NORMAL
    End With
    
    With myNonDestructiveFX
        .fx_Exposure = 0
        .fx_Contrast = 0
        .fx_Clarity = 0
        .fx_Vibrance = 0
        .fx_Temperature = 0
        .fx_Tint = 0
    End With
    
    With myVectorData
        .vd_Width = 0
        .vd_Height = 0
    End With
    
    'Initialize a default text renderer, regardless of layer type
    Set myTextRenderer = New pdTextRenderer
    
    'Initialize the layer's DIB
    Set layerDIB = New pdDIB
    
    'Initialize all temporary LOD compositing DIBs and corresponding hashes
    ReDim m_ViewportStateHashes(0 To NUM_OF_LOD_CACHES - 1) As Long
    ReDim m_DIBChangedSinceLastHash(0 To NUM_OF_LOD_CACHES - 1) As Boolean
    
    ReDim lodDIB(0 To NUM_OF_LOD_CACHES - 1) As pdDIB
    Dim i As Long
    For i = 0 To NUM_OF_LOD_CACHES - 1
        Set lodDIB(i) = New pdDIB
    Next i
    
    'Initialize the layer thumbnail and note that it is not clean, which will force a regeneration when the thumbnail is next requested.
    Set layerThumbnail = New pdDIB
    m_IsThumbClean = False
    
    'Set the canonical ID to -1.  This can be used to determine if the layer has been activated.
    layerID = -1
    
    'By default, vector layers do not have a backer DIB generated
    m_VectorDIBSynched = False
    
    'Prep a default set of parallelogram corner points, in case the user decides to activate non-destructive transforms
    ReDim m_PlgPoints(0 To 3) As POINTFLOAT
    
End Sub

'Write a copy of the layer header to an XML-format string.  All data necessary to recreate a layer from scratch will be
' included in this string, EXCEPT FOR DIB DATA.  For obvious performance reasons, the DIB contents of this layer must be
' handled separately, in binary format.
'
'If the optional parameter "avoidFluffEntries" is set, descriptive comments and pretty formatting will not be
' applied to the XML string.  This helpful for improving read/write performance of the data, such as when writing
' Undo/Redo entries.
Public Function getLayerHeaderAsXML(Optional ByVal avoidFluffEntries As Boolean = False) As String

    'Prepare an XML engine, which greatly simplifies the process of assembling XML data
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Add a basic header and explanatory comment
    xmlEngine.prepareNewXML "pdLayer"
    
    'Add a helpful comment, if requested.
    If Not avoidFluffEntries Then
        xmlEngine.writeBlankLine
        xmlEngine.writeComment "This file contains a summary of crucial data for a given pdLayer entry."
        xmlEngine.writeBlankLine
    End If
    
    'Start by writing out the layerID.  Note that functions making use of this XML data must be very careful about how they
    ' handle this value.  Depending on the context, it may not be advisable to re-use this data at load-time.
    xmlEngine.writeTag "ID", layerID
    
    'Next, write out the layer type.  Remember that this value is immutable, and cannot be changed for the life of the layer.
    xmlEngine.writeTag "Type", myLayerType
    
    'Write out the contents of our current LayerData header.  (NOTE: this must be manually modified if new entries are
    ' added to the LayerData type.
    With myLayerData
        xmlEngine.writeTag "Name", .l_Name
        xmlEngine.writeTag "GroupID", .l_GroupID
        xmlEngine.writeTag "Opacity", .l_Opacity
        xmlEngine.writeTag "BlendMode", .l_BlendMode
        xmlEngine.writeTag "OffsetX", .l_OffsetX
        xmlEngine.writeTag "OffsetY", .l_OffsetY
        xmlEngine.writeTag "CanvasXModifier", .l_CanvasXModifier
        xmlEngine.writeTag "CanvasYModifier", .l_CanvasYModifier
        xmlEngine.writeTag "Angle", .l_Angle
        xmlEngine.writeTag "Visibility", .l_Visibility
        xmlEngine.writeTag "NonDestructiveFXActive", .l_NonDestructiveFXActive
        xmlEngine.writeTag "ResizeQuality", .l_ResizeQuality
        xmlEngine.writeTag "ShearX", .l_ShearX
        xmlEngine.writeTag "ShearY", .l_ShearY
        xmlEngine.writeTag "AlphaMode", .l_AlphaMode
    End With
    
    'Next, write out any non-destructive FX parameters
    With myNonDestructiveFX
        xmlEngine.writeTag "NDFX_Exposure", .fx_Exposure
        xmlEngine.writeTag "NDFX_Contrast", .fx_Contrast
        xmlEngine.writeTag "NDFX_Clarity", .fx_Clarity
        xmlEngine.writeTag "NDFX_Vibrance", .fx_Vibrance
        xmlEngine.writeTag "NDFX_Temperature", .fx_Temperature
        xmlEngine.writeTag "NDFX_Tint", .fx_Tint
    End With
    
    'Next, write out the header of the layer's DIB contents.  These values are necessary for reconstructing the layer's DIB
    ' in the future.
    With xmlEngine
        .writeTag "DIB_ColorDepth", layerDIB.getDIBColorDepth
        .writeTag "DIB_Width", layerDIB.getDIBWidth
        .writeTag "DIB_Height", layerDIB.getDIBHeight
        .writeTag "DIB_ArrayWidth", layerDIB.getDIBArrayWidth
        .writeTag "DIB_AlphaPremultiplication", layerDIB.getAlphaPremultiplication
    End With
    
    'Note: we obviously do not write out the layer's DIB contents.  It must be handled via separate binary methods.
    
    If Not avoidFluffEntries Then xmlEngine.writeBlankLine
    
    'The layer XML string is now complete.  Return it.
    getLayerHeaderAsXML = xmlEngine.returnCurrentXMLString(avoidFluffEntries)

End Function

'Retrieve the layer's key data bits, but in PD's native param-string format.  This is used internally to check for
' non-destructive changes to layer data; if any are found, the Undo/Redo engine knows to create a new point on the stack.
Public Function getLayerHeaderAsParamString() As String

    'Start by compiling a layer header
    With myLayerData
        getLayerHeaderAsParamString = buildParams(.l_Name, .l_GroupID, .l_Opacity, .l_BlendMode, .l_OffsetX, .l_OffsetY, .l_CanvasXModifier, .l_CanvasYModifier, .l_Angle, .l_Visibility, .l_NonDestructiveFXActive, .l_ResizeQuality, .l_ShearX, .l_ShearY, .l_AlphaMode)
    End With
    
    'Next, add any non-destructive FX parameters
    With myNonDestructiveFX
        getLayerHeaderAsParamString = getLayerHeaderAsParamString & "|" & buildParams(.fx_Exposure, .fx_Contrast, .fx_Clarity, .fx_Vibrance, .fx_Temperature, .fx_Tint)
    End With
    
End Function

'Given a param string, recreate the layer header accordingly.  Note that BY DESIGN, immutable layer parameters - including
' layer ID and layer type - are never stored in param strings, and thus cannot be changed by this function.
Public Sub setLayerHeaderFromParamString(ByVal srcParamString As String)

    'Create a param string parser
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString srcParamString
    
    'Assign all params to the layer header
    With myLayerData
        .l_Name = cParams.GetString(1)
        .l_GroupID = cParams.GetLong(2)
        .l_Opacity = cParams.GetLong(3)
        .l_BlendMode = cParams.GetLong(4)
        .l_OffsetX = cParams.GetDouble(5)
        .l_OffsetY = cParams.GetDouble(6)
        .l_CanvasXModifier = cParams.GetDouble(7)
        .l_CanvasYModifier = cParams.GetDouble(8)
        .l_Angle = cParams.GetDouble(9)
        .l_Visibility = cParams.GetBool(10)
        .l_NonDestructiveFXActive = cParams.GetBool(11)
        .l_ResizeQuality = cParams.GetLong(12)
        .l_ShearX = cParams.GetDouble(13, 0)
        .l_ShearY = cParams.GetDouble(14, 0)
        .l_AlphaMode = cParams.GetLong(15, 0)
    End With
    
    'After the layer header comes non-destructive FX settings.  Make sure to adjust these indices if new values get added to
    ' myLayerData, above!
    With myNonDestructiveFX
        .fx_Exposure = cParams.GetDouble(13)
        .fx_Contrast = cParams.GetLong(14)
        .fx_Clarity = cParams.GetLong(15)
        .fx_Vibrance = cParams.GetLong(16)
        .fx_Temperature = cParams.GetLong(17)
        .fx_Tint = cParams.GetLong(18)
    End With
        
    'Per PD convention, this class has no way to notify outsiders of updates, so the calling function must handle any
    ' viewport redraws or other interface changes resulting from this assignment.

End Sub

'The sister function to "getLayerHeaderAsXML" above, this function will initialize a layer header from an XML-format string.
'
'Note that if the optional createNonDestructively parameter is FALSE (as it is by default), this function will create a
' blank DIB for the layer DIB, using the dimensions and color depth stored in the XML string.  However, it will obviously
' NOT apply any stored image data.  That must be handled separately, as the layer data varies by layer type.  (Raster layers
' store their data as a raw byte stream, while vector and text layers store XML data and we simply recreate the raster copy
' at run-time.)
Public Function CreateNewLayerFromXML(ByRef xmlString As String, Optional ByVal useCustomLayerID As Long = -1, Optional ByVal createNonDestructively As Boolean = False) As Boolean

    'Prepare an XML engine, which greatly simplifies the process of parsing XML data
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Validate the XML header...
    If xmlEngine.loadXMLFromString(xmlString) And xmlEngine.isPDDataType("pdLayer") Then
        
        'If the caller has not specified their own layerID, use the one from file.
        If useCustomLayerID = -1 Then
            layerID = xmlEngine.getUniqueTag_Long("ID")
        Else
            layerID = useCustomLayerID
        End If
    
        'Read in the layer type.  Remember that this value is immutable, and cannot be changed for the life of the layer.
        myLayerType = xmlEngine.getUniqueTag_Long("Type", PDL_IMAGE)
        
        'Read in the contents of the LayerData header.  (NOTE: this must be manually modified if new entries are
        ' added to the LayerData type.
        With myLayerData
            .l_Name = xmlEngine.getUniqueTag_String("Name")
            .l_GroupID = xmlEngine.getUniqueTag_Long("GroupID")
            .l_Opacity = xmlEngine.getUniqueTag_Long("Opacity")
            .l_BlendMode = xmlEngine.getUniqueTag_Long("BlendMode", BL_NORMAL)
            .l_OffsetX = xmlEngine.getUniqueTag_Double("OffsetX")
            .l_OffsetY = xmlEngine.getUniqueTag_Double("OffsetY")
            .l_CanvasXModifier = xmlEngine.getUniqueTag_Double("CanvasXModifier", 1)
            .l_CanvasYModifier = xmlEngine.getUniqueTag_Double("CanvasYModifier", 1)
            .l_Angle = xmlEngine.getUniqueTag_Double("Angle")
            .l_Visibility = xmlEngine.getUniqueTag_Boolean("Visibility")
            .l_NonDestructiveFXActive = xmlEngine.getUniqueTag_Boolean("NonDestructiveFXActive", False)
            .l_ResizeQuality = xmlEngine.getUniqueTag_Long("ResizeQuality", LRQ_BILINEAR_NORMAL)
            .l_ShearX = xmlEngine.getUniqueTag_Double("ShearX", 0)
            .l_ShearY = xmlEngine.getUniqueTag_Double("ShearY", 0)
            .l_AlphaMode = xmlEngine.getUniqueTag_Long("AlphaMode", LA_NORMAL)
        End With
        
        'Read in the contents of the non-destructive FX chunk
        With myNonDestructiveFX
            .fx_Exposure = xmlEngine.getUniqueTag_Double("NDFX_Exposure", 0)
            .fx_Contrast = xmlEngine.getUniqueTag_Long("NDFX_Contrast", 0)
            .fx_Clarity = xmlEngine.getUniqueTag_Double("NDFX_Clarity", 0)
            .fx_Vibrance = xmlEngine.getUniqueTag_Long("NDFX_Vibrance", 0)
            .fx_Temperature = xmlEngine.getUniqueTag_Long("NDFX_Temperature", 0)
            .fx_Tint = xmlEngine.getUniqueTag_Long("NDFX_Tint", 0)
        End With
        
        'Next, we're going to do something kinda weird.  We're going to create a blank DIB for the layer contents
        ' using the DIB header supplied in the XML, but we're *not* going to fill the DIB!  An external function will
        ' handle that, because the actual DIB bits are stored in binary format.  (Serializing them to XML is not feasible
        ' for the image sizes PD deals with.)
        '
        'Note that this behavior can be overridden by setting the optional createNonDestructively parameter to TRUE.
        ' This is used by the Undo/Redo engine to re-load an image's header data, without actually overwriting the
        ' current layer DIB.
        If Not createNonDestructively Then
        
            Set layerDIB = New pdDIB
            With xmlEngine
                layerDIB.createBlank .getUniqueTag_Long("DIB_Width", 1), .getUniqueTag_Long("DIB_Height", 1), .getUniqueTag_Long("DIB_ColorDepth", 32)
                
                'Mark alpha premultiplication in advance
                layerDIB.setInitialAlphaPremultiplicationState .getUniqueTag_Boolean("DIB_AlphaPremultiplication", True)
            
            End With
            
            'Note that all internal layer caches are no longer valid
            Me.notifyOfDestructiveChanges
            
        End If
        
        CreateNewLayerFromXML = True
    
    'If the layer XML data can't be validated, there's nothing we can do but exit.
    Else
        Debug.Print "Layer could not be created: are you sure the supplied data was valid?"
        CreateNewLayerFromXML = False
    End If
    
End Function

'Write a copy of a vector layer's contents to an XML-format string.  All data necessary to recreate a vector layer
' from scratch will be included in this string.  For obvious performance reasons, the DIB contents of this layer are not
' part of this function.  PD regenerates the DIB copy of vector layers from scratch, at load-time.
'
'If the optional parameter "avoidFluffEntries" is set, descriptive comments and pretty formatting will not be
' applied to the XML string.  This helpful for improving read/write performance of the data, such as when writing
' Undo/Redo entries.
Public Function getVectorDataAsXML(Optional ByVal avoidFluffEntries As Boolean = False) As String
    
    'First, make sure this is not an image layer
    If Not Me.isLayerVector() Then
        Debug.Print "WARNING!  The getVectorDataAsXML() function is only valid for vector layers!"
        Exit Function
    End If
    
    'Prepare an XML engine, which greatly simplifies the process of assembling XML data
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Add a basic header and explanatory comment
    xmlEngine.prepareNewXML "pdVector"
    
    'Add a helpful comment, if requested.
    If Not avoidFluffEntries Then
        xmlEngine.writeBlankLine
        xmlEngine.writeComment "This XML string contains a summary of crucial data for a pdVector object."
        xmlEngine.writeBlankLine
    End If
    
    'All vector layers store some similar data
    With xmlEngine
        .writeTag "GenericVectorWidth", myVectorData.vd_Width
        .writeTag "GenericVectorHeight", myVectorData.vd_Height
    End With
    
    'Additional vector data varies by layer type
    Select Case myLayerType
        
        Case PDL_TEXT, PDL_TYPOGRAPHY
            
            'pdTextRenderer serializes all its data to XML for us, making this step very simple
            xmlEngine.writeGenericText myTextRenderer.getAllFontSettingsAsXML()
        
    End Select
        
    If Not avoidFluffEntries Then xmlEngine.writeBlankLine
    
    'The layer XML string is now complete.  Return it.
    getVectorDataAsXML = xmlEngine.returnCurrentXMLString(avoidFluffEntries)

End Function

'The sister function to "getVectorDataAsXML" above, this function will initialize a vector layer's contents from an XML-format string.
Public Function CreateVectorDataFromXML(ByRef xmlString As String) As Boolean

    'Prepare an XML engine, which greatly simplifies the process of parsing XML data
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Validate the XML header...
    If xmlEngine.loadXMLFromString(xmlString) And xmlEngine.isPDDataType("pdVector") Then
        
        'All vector layers store some similar data
        With myVectorData
            .vd_Width = xmlEngine.getUniqueTag_Long("GenericVectorWidth", 1)
            .vd_Height = xmlEngine.getUniqueTag_Long("GenericVectorHeight", 1)
        End With
        
        'Next, the vector tags we look for vary by layer type
        Select Case myLayerType
            
            Case PDL_TEXT, PDL_TYPOGRAPHY
                
                'pdTextRenderer serializes all its data from XML for us, making this step very simple
                myTextRenderer.setAllFontSettingsFromXML xmlEngine.returnCurrentXMLString(True)
                
        End Select
                
        'Note that all internal layer caches are no longer valid; this will also cause the raster cache to be marked unclean
        Me.notifyOfDestructiveChanges
        
        CreateVectorDataFromXML = True
    
    'If the layer XML data can't be validated, there's nothing we can do but exit.
    Else
        Debug.Print "Vector layer could not be initialized: are you sure the supplied data is valid?"
        CreateVectorDataFromXML = False
    End If
    
End Function

'Initialize a new layer.  Optional inputs include:
' - Layer name
' - For image-type layers, a DIB can be passed; the layer's raster contents will be initialized to match the DIB
' - A parent pdImage object.  This is currently unused, but in the future, I may use it to center the new layer on the current viewport
Public Sub InitializeNewLayer(ByVal newLayerType As LAYER_TYPE, Optional ByVal newLayerName As String = "", Optional ByRef srcDIB As pdDIB = Nothing, Optional ByRef parentImage As pdImage = Nothing)

    'Mark the layer type.
    myLayerType = newLayerType
    
    'Copy the name locally, and set other values to their natural defaults
    With myLayerData
        If Len(newLayerName) <> 0 Then
            .l_Name = newLayerName
        Else
            .l_Name = g_Language.TranslateMessage("New layer")
        End If
        .l_GroupID = 0
        .l_Opacity = 100
        .l_BlendMode = BL_NORMAL
        .l_Visibility = True
        .l_NonDestructiveFXActive = False
        .l_ResizeQuality = LRQ_BILINEAR_NORMAL
        .l_ShearX = 0
        .l_ShearY = 0
        .l_AlphaMode = LA_NORMAL
    End With
    
    'Reset any non-destructive fx parameters
    With myNonDestructiveFX
        .fx_Exposure = 0
        .fx_Contrast = 0
        .fx_Clarity = 0
        .fx_Vibrance = 0
        .fx_Temperature = 0
        .fx_Tint = 0
    End With
    
    'Reset any vector-specific data
    With myVectorData
        .vd_Width = 0
        .vd_Height = 0
    End With
    
    m_VectorDIBSynched = False
    
    'If passed, create a local copy of the passed DIB.
    Set layerDIB = New pdDIB
    If Not (srcDIB Is Nothing) Then layerDIB.createFromExistingDIB srcDIB
    
    'Set default offsets
    myLayerData.l_OffsetX = 0
    myLayerData.l_OffsetY = 0
    
    'In the future, we might want to set an intial offset relative to the parent image's viewport settings...?
    
    'Note that all internal layer caches are no longer valid.  This is particularly important for vector layers, which generate a raster copy
    ' of their contents "on the fly"
    Me.notifyOfDestructiveChanges

End Sub

'Copy an existing layer.  All layer contents will be copied manually, so make sure that new layer properties are
' manually added to this function!
Public Sub CopyExistingLayer(ByRef srcLayer As pdLayer)

    'Copy all generic layer properties from the source layer.  (Note that canonical layer ID is *not* copied; that must always
    ' be unique for every created layer!)
    myLayerType = srcLayer.getLayerType
    
    With myLayerData
        .l_Name = incrementTrailingNumber(srcLayer.getLayerName)
        .l_GroupID = srcLayer.getLayerGroup
        .l_Opacity = srcLayer.getLayerOpacity
        .l_BlendMode = srcLayer.getLayerBlendMode
        .l_OffsetX = srcLayer.getLayerOffsetX
        .l_OffsetY = srcLayer.getLayerOffsetY
        .l_CanvasXModifier = srcLayer.getLayerCanvasXModifier
        .l_CanvasYModifier = srcLayer.getLayerCanvasYModifier
        .l_Angle = srcLayer.getLayerAngle
        .l_Visibility = srcLayer.getLayerVisibility
        .l_NonDestructiveFXActive = srcLayer.getLayerNonDestructiveFXState
        .l_ResizeQuality = srcLayer.getLayerResizeQuality
        .l_ShearX = srcLayer.getLayerShearX
        .l_ShearY = srcLayer.getLayerShearY
        .l_AlphaMode = srcLayer.getLayerAlphaMode
    End With
    
    'Next, copy any non-destructive FX parameters
    With myNonDestructiveFX
        .fx_Exposure = srcLayer.getLayerNonDestructiveFXValue(NDFX_EXPOSURE)
        .fx_Contrast = srcLayer.getLayerNonDestructiveFXValue(NDFX_CONTRAST)
        .fx_Clarity = srcLayer.getLayerNonDestructiveFXValue(NDFX_CLARITY)
        .fx_Vibrance = srcLayer.getLayerNonDestructiveFXValue(NDFX_VIBRANCE)
        .fx_Temperature = srcLayer.getLayerNonDestructiveFXValue(NDFX_TEMPERATURE)
        .fx_Tint = srcLayer.getLayerNonDestructiveFXValue(NDFX_TINT)
    End With
    
    'Copy the source layer's DIB, if one exists.
    If Not (srcLayer.layerDIB Is Nothing) Then layerDIB.createFromExistingDIB srcLayer.layerDIB
    
    'Depending on the type of layer we are copying, copy any extra layer data, or generate a new layer mask to
    ' match the layer's vector contents.
    Select Case srcLayer.getLayerType
    
        Case PDL_IMAGE
        
        'Vector layers copy additional vector-specific data
        Case PDL_TEXT, PDL_TYPOGRAPHY
            myVectorData.vd_Width = srcLayer.getLayerWidth
            myVectorData.vd_Height = srcLayer.getLayerHeight
            
            'Other vector-specific data can be transferred via XML, saving us the hassle of explicitly copying it here
            Me.CreateVectorDataFromXML srcLayer.getVectorDataAsXML(True)
        
        Case PDL_ADJUSTMENT
    
    End Select
    
    'Note that all internal layer caches are no longer valid
    Me.notifyOfDestructiveChanges
    
End Sub

'Assuming the current layer DIB is null-padded (a term I use to describe a layer that has been forced to the size of the
' image by surrounding it with transparent pixels), intelligently crop the DIB by removing all fully transparent padding.
' When this is done, modify the current layer offsets so that the image stays in effectively the same place, but without
' all that blank padding.
' IMPORTANT NOTE: if you call this function without first null-padding the layer DIB, weird shit may happen.  That is not
'                 an intended use-case, so don't attempt it!
'
'Returns TRUE if layer was trimmed successfully, FALSE if it was not.  (False occurs if the entire layer is transparent.)
Public Function cropNullPaddedLayer() As Boolean
    
    'Make sure the source DIB isn't empty; (this should never happen, but I'm trying to be better about catching
    ' edge error conditions).
    If (layerDIB.getDIBDC <> 0) And (layerDIB.getDIBWidth <> 0) And (layerDIB.getDIBHeight <> 0) And (layerDIB.getDIBColorDepth = 32) Then
    
        'Point an array at the layer's DIB data
        Dim srcImageData() As Byte
        Dim srcSA As SAFEARRAY2D
        prepInternalSafeArray srcSA
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
        Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
        initX = 0
        initY = 0
        finalX = layerDIB.getDIBWidth - 1
        finalY = layerDIB.getDIBHeight - 1
        
        'Each edge will be scanned independently, and these variables will be calculated with the new bounding
        ' RECT for this layer.
        Dim newTop As Long, newBottom As Long, newLeft As Long, newRight As Long
        
        'Before doing anything else, perform a quick sanity check of the four image corners.  If none of them are
        ' fully transparent, we know the layer is untrimmable (because interesting pixel data extends to all sides),
        ' so we can exit now without actually having to perform a full scan.
        If (srcImageData(3, 0) <> 0) And (srcImageData(finalX * 4 + 3, 0) <> 0) And (srcImageData(3, finalY) <> 0) And (srcImageData(finalX * 4 + 3, finalY) <> 0) Then
            
            'Release our array pointer.  (This MUST be done for all function exits, or VB will crash.)
            CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
            Debug.Print "Quick check of layer corners shows no trimmable areas; premature exit initialized."
            cropNullPaddedLayer = False
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
            If srcImageData(x * 4 + 3, y) <> 0 Then
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
        If Not pixelFails Then
            CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
            Debug.Print "Request for layer null padding trim denied - entire layer is transparent, so crop is impossible!"
            cropNullPaddedLayer = False
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
            
            If srcImageData(x * 4 + 3, y) <> 0 Then
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
        Dim QuickVal As Long
        
        For x = 0 To finalX
            QuickVal = x * 4 + 3
        For y = initY To finalY
        
            If srcImageData(QuickVal, y) <> 0 Then
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
            QuickVal = x * 4 + 3
        For y = initY To finalY
        
            If srcImageData(QuickVal, y) <> 0 Then
                pixelFails = True
                Exit For
            End If
            
        Next y
            If pixelFails Then Exit For
        Next x
        
        newRight = x
        
        'With our work complete, point ImageData() away from the DIB and deallocate it
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        Erase srcImageData
        
        'If we made it all the way here, our search for null borders completed successfully.
        
        'If all border values are 0, return FALSE.  This basically means that the interesting bits of the layer
        ' extend to the full size of the image, so we can't trim it.
        If (newLeft = 0) And (newTop = 0) And (newRight = 0) And (newBottom = 0) Then
            Debug.Print "Layer null padding scan completed successfully, but no trimmable edges were found."
            cropNullPaddedLayer = False
        Else
        
            'Trimmable borders were found.  Use the calculated values to crop the image to its smallest relevant size.
            Dim tmpCropDIB As pdDIB
            Set tmpCropDIB = New pdDIB
            tmpCropDIB.createBlank newRight - newLeft + 1, newBottom - newTop + 1, 32, 0
            Debug.Print "Crop rect: (" & newLeft & "," & newTop & ") - (" & newRight & "," & newBottom & ")"
            BitBlt tmpCropDIB.getDIBDC, 0, 0, newRight - newLeft + 1, newBottom - newTop + 1, layerDIB.getDIBDC, newLeft, newTop, vbSrcCopy
            
            'Replace the current layer DIB with our temporary DIB
            layerDIB.createFromExistingDIB tmpCropDIB
            
            'Modify our layer offsets to reflect our new size
            myLayerData.l_OffsetX = newLeft
            myLayerData.l_OffsetY = newTop
        
            'Return TRUE, so that the calling function knows to redraw the screen with the new size and offsets
            cropNullPaddedLayer = True
            
        End If
        
    Else
        Debug.Print "Request for layer null padding trim denied - layer DIB does not currently exist!"
        cropNullPaddedLayer = False
    End If
    
End Function

'Null-pad this layer (with transparent pixels) against the size of our parent pdImage.  Note that the parent pdImage's dimensions
' must be explicitly passed, because layers do not maintain a link to their parent.
' TODO: determine if off-image layer cropping is an acceptable solution in all scenarios.  (It really isn't, but there are some egregious performance
'       considerations with layers that lie off-canvas, so a solution needs to be carefully planned.)
Public Sub convertToNullPaddedLayer(ByVal parentImageWidth As Long, ByVal parentImageHeight As Long, Optional ByVal applyCanvasTransformsFirst As Boolean = True)

    'If this layer has transforms applied to it, we have to make them permanent before proceeding.  Otherwise, the resulting image may
    ' not look right.
    '
    'Note that this behavior can be overridden if a function is ABSOLUTELY CERTAIN that the end result won't botch the transform, but this
    ' is only true for pixel transforms that operate strictly in separate x/y domains, such as resizing.  Note that this setting may be
    ' forcibly ignored if the layer has complex affine transforms active (like rotation), because those transforms change the (x, y) domain
    ' of the layer relative to the (x, y) domain of the image.
    Dim mustMakeCanvasTransformsPermanent As Boolean
    mustMakeCanvasTransformsPermanent = applyCanvasTransformsFirst
    mustMakeCanvasTransformsPermanent = mustMakeCanvasTransformsPermanent Or Me.affineTransformsActive(False)
    
    If mustMakeCanvasTransformsPermanent Then makeCanvasTransformsPermanent parentImageWidth, parentImageHeight
    
    'We can skip this step for layers that are already the size of the image.  This is common for background layers.
    Dim nullPaddingUnnecessary As Boolean
    nullPaddingUnnecessary = False
    
    With myLayerData
        nullPaddingUnnecessary = (.l_OffsetX = 0) And (.l_OffsetY = 0) And (Me.getLayerWidth(True) = parentImageWidth) And (Me.getLayerHeight(True) = parentImageHeight)
    End With
    
    'If the layer still requires specialized padding, carry on
    If Not nullPaddingUnnecessary Then
        
        'Create a blank destination DIB at the size of the parent image
        Dim tmpPaddingDIB As pdDIB
        Set tmpPaddingDIB = New pdDIB
        tmpPaddingDIB.createBlank parentImageWidth, parentImageHeight, 32, 0
    
        'Copy the current layer into the temporary DIB, with proper offsets applied
        BitBlt tmpPaddingDIB.getDIBDC, myLayerData.l_OffsetX, myLayerData.l_OffsetY, layerDIB.getDIBWidth, layerDIB.getDIBHeight, layerDIB.getDIBDC, 0, 0, vbSrcCopy
    
        'Replace our current layer DIB with the temporary DIB
        Set Me.layerDIB = tmpPaddingDIB
        Me.layerDIB.setInitialAlphaPremultiplicationState True
    
        'Reset layer offset values
        myLayerData.l_OffsetX = 0
        myLayerData.l_OffsetY = 0
    
        'Mark the layer as dirty
        Me.notifyOfDestructiveChanges
    
    End If
    
End Sub

'Update the internal thumbnail.  Only call this function if the current thumbnail has been marked as DIRTY; otherwise, just use the
' existing thumbnail as-is.
Private Sub updateInternalThumbnail()
    
    'Uncomment the two lines below (and another at the end of the function) to profile different thumbnail resize strategies
    'Dim startTime As Double
    'startTime = Timer
    
    'If the layer has not been instantiated properly, reject the thumbnail request
    If (layerDIB Is Nothing) Or (layerDIB.getDIBWidth = 0) Or (layerDIB.getDIBHeight = 0) Then Exit Sub
    
    'Thumbnails have some interesting requirements.  We always want them to be square, with the image set in the middle
    ' of the thumbnail (with aspect ratio preserved) and any empty edges made transparent.
    
    'We also need to determine the thumbnail's actual width and height, and any x and y offset necessary to preserve the
    ' aspect ratio and center the image on the thumbnail.
    Dim tIcoWidth As Long, tIcoHeight As Long, tX As Double, tY As Double
    
    'Start by determining proper dimensions for the resized thumbnail image.
    convertAspectRatio layerDIB.getDIBWidth, layerDIB.getDIBHeight, LAYER_THUMB_SIZE, LAYER_THUMB_SIZE, tIcoWidth, tIcoHeight
    
    'If the form is wider than it is tall, center the thumbnail vertically
    If tIcoWidth > tIcoHeight Then
        tX = 0
        tY = (LAYER_THUMB_SIZE - tIcoHeight) / 2
    
    '...otherwise, center it horizontally
    Else
        tY = 0
        tX = (LAYER_THUMB_SIZE - tIcoWidth) / 2
    End If
    
    'Prepare the destination DIB
    layerThumbnail.createBlank LAYER_THUMB_SIZE, LAYER_THUMB_SIZE, 32, 0
    
    'There are two possible ways to create a thumbnail image.  If FreeImage is available, we prefer to use it, as it
    ' is faster (I've actually profiled it!), but if FreeImage is not available, GDI+ will suffice.
    If g_ImageFormats.FreeImageEnabled Then
        FreeImageResizeDIBFast layerThumbnail, tX, tY, tIcoWidth, tIcoHeight, layerDIB, IIf(g_ThumbnailPerformance = PD_PERF_FASTEST, FILTER_BOX, FILTER_BILINEAR), True
    Else
        GDIPlusResizeDIB layerThumbnail, tX, tY, tIcoWidth, tIcoHeight, layerDIB, 0, 0, layerDIB.getDIBWidth, layerDIB.getDIBHeight, IIf(g_ThumbnailPerformance = PD_PERF_FASTEST, InterpolationModeNearestNeighbor, InterpolationModeHighQualityBicubic)
    End If
    
    'Uncomment the line below to profile different thumbnail resize strategies
    'Debug.Print "Time taken for layer thumbnail generation: " & CStr((Timer - startTime) * 1000) & " ms"
    
    'Mark the thumbnail state as clean
    m_IsThumbClean = True
    
End Sub

'External functions can use this function to request a thumbnail version of the layer contents.
Public Function requestThumbnail(ByRef dstThumbnailDIB As pdDIB, Optional ByVal thumbnailSize As Long = 64) As Boolean
    
    'Is the current thumbnail dirty?  If so, regenerate it.
    If Not m_IsThumbClean Then updateInternalThumbnail
    
    'Clear the destination thumbnail
    If dstThumbnailDIB Is Nothing Then Set dstThumbnailDIB = New pdDIB
    
    'If the requested thumbnail size differs from our cached version, resize the thumbnail prior to passing it along
    If thumbnailSize <> LAYER_THUMB_SIZE Then
        
        If (dstThumbnailDIB.getDIBWidth <> thumbnailSize) Or (dstThumbnailDIB.getDIBHeight <> thumbnailSize) Then
            dstThumbnailDIB.createBlank thumbnailSize, thumbnailSize, 32, 0
        Else
            dstThumbnailDIB.resetDIB 0
        End If
        
        If g_ImageFormats.FreeImageEnabled Then
            FreeImageResizeDIBFast dstThumbnailDIB, 0, 0, thumbnailSize, thumbnailSize, layerThumbnail, IIf(g_ThumbnailPerformance = PD_PERF_FASTEST, FILTER_BOX, FILTER_BILINEAR), True
        Else
            GDIPlusResizeDIB dstThumbnailDIB, 0, 0, thumbnailSize, thumbnailSize, layerThumbnail, 0, 0, layerThumbnail.getDIBWidth, layerThumbnail.getDIBHeight, IIf(g_ThumbnailPerformance = PD_PERF_FASTEST, InterpolationModeNearestNeighbor, InterpolationModeHighQualityBicubic)
        End If
        
        dstThumbnailDIB.setInitialAlphaPremultiplicationState True
        
    Else
        dstThumbnailDIB.createFromExistingDIB layerThumbnail
    End If
    
    requestThumbnail = True
    
End Function

'This function can be used to populate a valid SAFEARRAY2D structure against the layer's current DIB
Private Sub prepInternalSafeArray(ByRef srcSA As SAFEARRAY2D)
    
    'Populate a relevant SafeArray variable for the supplied DIB
    With srcSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = layerDIB.getDIBHeight
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = layerDIB.getDIBArrayWidth
        .pvData = layerDIB.getActualDIBBits
    End With
    
End Sub

'Whenever the user does something with the mouse - click, move, etc - the primary canvas will poll this function.  This function
' will check the passed mouse values (which have already been translated into the LAYER coordinate space), and return a Long-type
' value specifying whether or not that coordinate is of interest to this layer.  This basically provides a flexible framework for
' supporting mouse events for any layer, because it's up to the layer to determine which points are interesting and which are not.
'
'INPUTS: current mouse coords, already translated into the LAYER coordinate space (not the canvas coordinate space!)
'
'OUTPUT: index of the POI, if any.  If the mouse is not over a POI, -1 is returned.
Public Function checkForPointOfInterest(ByVal layerX As Long, ByVal layerY As Long) As Long
    
    'MouseAccuracy in PD is a global value, but because we are working in image coordinates, we must compensate for the
    ' current zoom value.  (Otherwise, when zoomed out the user would be forced to work with tighter accuracy!)
    ' (TODO: come up with a better solution for this.  Accuracy should *really* be handled in the canvas coordinate space,
    '        so perhaps the caller should specify an image x/y and a radius...?)
    Dim mouseAccuracy As Double
    mouseAccuracy = g_MouseAccuracy * (1 / g_Zoom.getZoomValue(pdImages(g_CurrentImage).currentZoomValue))
    
    'Find the smallest distance for this mouse position
    Dim minDistance As Double
    minDistance = mouseAccuracy
    
    'When a point within the accuracy limit is found, its index will be assigned to this variable
    Dim closestPoint As Long
    closestPoint = -1
    
    'To facilitate the notion of "generalized point of interest" handling, coordinate checks are performed against
    ' an array of potential POI targets.  While arrays may not make as much sense as something like a RECT (when
    ' working with standard layers), it lets us generalize all coordinate checks into a single function, which is awesome!
    Dim poiList() As POINTAPI
    
    'Points of interest are sorted by layer type.  In most cases, only the corners of a layer are interesting (as they control
    ' the layer's bounding box), but in the future, things like a Pen tool could have tons of points of interest.
    Select Case myLayerType
    
        'Image and text layers consider their corners to be points of interest
        Case PDL_IMAGE, PDL_TEXT, PDL_TYPOGRAPHY
            
            'Manually populate the point of interest array with the layer's corner coordinates.
            ReDim poiList(0 To 7) As POINTAPI
            poiList(0).x = getLayerOffsetX
            poiList(0).y = getLayerOffsetY
            poiList(1).x = getLayerOffsetX + getLayerWidth(True)
            poiList(1).y = getLayerOffsetY
            poiList(2).x = getLayerOffsetX
            poiList(2).y = getLayerOffsetY + getLayerHeight(True)
            poiList(3).x = getLayerOffsetX + getLayerWidth(True)
            poiList(3).y = getLayerOffsetY + getLayerHeight(True)
            
            'As of PD 7.7, rotation handles are also provided.  Note that changes to these positions *must be mirrored* in the
            ' getLayerRotationNodeCoordinates() function above!
            poiList(4).x = getLayerOffsetX + getLayerWidth(True)
            poiList(4).y = getLayerOffsetY + getLayerHeight(True) \ 2
            
            poiList(5).x = getLayerOffsetX + getLayerWidth(True) \ 2
            poiList(5).y = getLayerOffsetY + getLayerHeight(True)
            
            poiList(6).x = getLayerOffsetX
            poiList(6).y = getLayerOffsetY + getLayerHeight(True) \ 2
            
            poiList(7).x = getLayerOffsetX + getLayerWidth(True) \ 2
            poiList(7).y = getLayerOffsetY
            
            'Check to see if one of the corners has been clicked
            closestPoint = findClosestPointInArray(layerX, layerY, minDistance, poiList)
            
            'If a corner was clicked, return its value and exit
            If closestPoint <> -1 Then
                checkForPointOfInterest = closestPoint
                Exit Function
            End If
            
            'If we're at this line of code, a POI was not clicked. Perform one final check to see if the mouse is within
            ' the layer's boundaries, and if it is, return the "move" ID, then exit.
            If (layerX > poiList(0).x) And (layerX < poiList(1).x) And (layerY > poiList(0).y) And (layerY < poiList(2).y) Then
                checkForPointOfInterest = 8
            Else
                checkForPointOfInterest = -1
            End If
    
    End Select

End Function

'Sometimes, it is helpful to reset all non-destructive parameters to their default values.  (For example, after flattening
' an image, we want to reset the base layer.)  Use this function to do so.
Public Sub resetLayerParameters()

    'Assign default values to this instance's header
    With myLayerData
        .l_GroupID = 0
        .l_Opacity = 100
        .l_BlendMode = BL_NORMAL
        .l_OffsetX = 0
        .l_OffsetY = 0
        .l_CanvasXModifier = 1
        .l_CanvasYModifier = 1
        .l_Angle = 0
        .l_Visibility = True
        .l_NonDestructiveFXActive = False
        .l_ResizeQuality = LRQ_BILINEAR_NORMAL
        .l_ShearX = 0
        .l_ShearY = 0
        .l_AlphaMode = LA_NORMAL
    End With
    
    'Assign default values to the non-destructive FX params
    With myNonDestructiveFX
        .fx_Exposure = 0
        .fx_Contrast = 0
        .fx_Clarity = 0
        .fx_Vibrance = 0
        .fx_Temperature = 0
        .fx_Tint = 0
    End With
    
    'Assign default vector layer values
    With myVectorData
        .vd_Width = 0
        .vd_Height = 0
    End With
    
End Sub

'If something needs to access the layer DIB directly (like PD's central compositor), this function should first be called for
' any vector layers.  If the DIB is already synched, there is no penalty to calling this function, but if it is *not* synched,
' this will guarantee the DIB is ready for access.
'
'For vector DIBs, returns TRUE if we had to sync the vector DIB; FALSE if the DIB was already synched
'For raster DIBs, returns TRUE if the image has changed since the last viewport hash was set; FALSE otherwise.
'
'If the return value is TRUE, any DIBs derived from this layer (e.g. viewport ones) must be recreated
Public Function syncInternalDIB(Optional ByVal levelOfDetail As COMPOSITOR_LOD_CLUE = CLC_Generic) As Boolean

    'If this is not a vector layer, rely on the "DIB has changed since last hash request" marker.
    If myLayerType = PDL_IMAGE Then
        syncInternalDIB = m_DIBChangedSinceLastHash(levelOfDetail)
        Exit Function
    End If
    
    'If the DIB is already synched, disregard this request
    If m_VectorDIBSynched Then
        syncInternalDIB = False
        Exit Function
    End If
    
    'From here, the specific vector syncing process varies by vector layer type
    Select Case myLayerType
    
        Case PDL_TEXT, PDL_TYPOGRAPHY
            
            'Set the text rendering engine according to the layer type (text vs typography)
            If myLayerType = PDL_TEXT Then
                myTextRenderer.setGenericTextProperty ptp_RenderingEngine, tre_WAPI
            Else
                myTextRenderer.setGenericTextProperty ptp_RenderingEngine, tre_PHOTODEMON
            End If
            
            'Initialize the backer DIB as necessary
            If (layerDIB Is Nothing) Then Set layerDIB = New pdDIB
            
            If (layerDIB.getDIBWidth <> myVectorData.vd_Width) Or (layerDIB.getDIBHeight <> myVectorData.vd_Height) Then
                layerDIB.createBlank myVectorData.vd_Width, myVectorData.vd_Height, 32, 0, 0
            Else
                layerDIB.resetDIB
            End If
            
            layerDIB.setInitialAlphaPremultiplicationState True
            
            'Use pdTextRenderer to paint the string onto the backer DIB
            myTextRenderer.renderTextToDIB layerDIB, 0, 0, myVectorData.vd_Width, myVectorData.vd_Height
            
    End Select
    
    'Note that the DIB is ready to go!
    m_VectorDIBSynched = True
    syncInternalDIB = True

End Function

'Estimate the RAM usage for this layer, based on known heavy memory users (layer DIB, in particular)
Public Function estimateRAMUsage() As Double

    'Calculate the raw memory size of both the layer and the cached thumbnail DIB, if any
    If Not (layerDIB Is Nothing) Then estimateRAMUsage = layerDIB.getDIBArrayWidth * layerDIB.getDIBHeight
    If Not (layerThumbnail Is Nothing) Then estimateRAMUsage = estimateRAMUsage + layerThumbnail.getDIBArrayWidth * layerThumbnail.getDIBHeight
    
    Dim i As Long
    For i = 0 To NUM_OF_LOD_CACHES - 1
        If Not (lodDIB(i) Is Nothing) Then estimateRAMUsage = estimateRAMUsage + lodDIB(i).getDIBArrayWidth * lodDIB(i).getDIBHeight
    Next i
    
    'Assume a 10% overhead for this class instance and other miscellaneous layer data
    estimateRAMUsage = estimateRAMUsage * 1.1

End Function

'Update the layer's current "viewport state hash".  This is used by the compositor, only.
Public Sub setViewportHash(ByVal cacheID As COMPOSITOR_LOD_CLUE, ParamArray allParams() As Variant)
    m_ViewportStateHashes(cacheID) = getViewportHash_Theoretical(allParams)
    m_DIBChangedSinceLastHash(cacheID) = False
End Sub

'If the viewport generates a hash via getViewportHash_Theoretical(), it can pass that hash directly to this function to improve performance
Public Sub setViewportHash_FromLong(ByVal cacheID As COMPOSITOR_LOD_CLUE, ByVal newHash As Long)
    m_ViewportStateHashes(cacheID) = newHash
    m_DIBChangedSinceLastHash(cacheID) = False
End Sub

'Retrieve the layer's current viewport hash
Public Function getViewportHash(ByVal cacheID As COMPOSITOR_LOD_CLUE) As Long
    getViewportHash = m_ViewportStateHashes(cacheID)
End Function

'Given a set of caller-specified parameters, generate a matching viewport hash.  This is used by the compositor to detect changes to the current layer.
' Caller-specified params are relevant because the compositor will also pass viewport settings, and if *those* change (but the layer is unchanged),
' we still need to regenerate the function.
Public Function getViewportHash_Theoretical(ParamArray allParams() As Variant) As Long
    
    'First, generate a hash of relevant internals.
    Dim tmpHash As Long
    
    With myLayerData
        tmpHash = (.l_Angle * 359013) Xor (.l_CanvasXModifier * 8091) Xor (.l_CanvasYModifier * 4095)
        tmpHash = tmpHash Xor (.l_OffsetX * 10) Xor (.l_OffsetY * 25)
        tmpHash = tmpHash Xor (.l_ResizeQuality * 620)
        tmpHash = tmpHash Xor (.l_Opacity * 65535)
        tmpHash = tmpHash Xor (.l_ShearX * 524287) Xor (.l_ShearY * 1048575)
        tmpHash = tmpHash Xor ((.l_BlendMode + 1) * 63) Xor ((.l_AlphaMode + 1) * 511)
    End With
    
    'Next, add any non-destructive FX parameters to the hash
    With myNonDestructiveFX
        tmpHash = tmpHash Xor (.fx_Clarity * 7) Xor (.fx_Contrast * 15)
        tmpHash = tmpHash Xor (.fx_Temperature * 63) Xor (.fx_Tint * 127) Xor (.fx_Vibrance * 255)
        tmpHash = tmpHash Xor (.fx_Exposure * 2047)
    End With
    
    'Finally, hash any parameters we were passed.
    If UBound(allParams) >= LBound(allParams) Then
    
        Dim i As Long
        For i = LBound(allParams) To UBound(allParams)
            tmpHash = tmpHash Xor allParams(i) * ((i + 1) * 127)
        Next i
    
    End If
    
    getViewportHash_Theoretical = tmpHash
    
End Function

'Access a specific level-of-detail (LOD) cache.
Friend Function tmpLODDIB(Optional ByVal cacheID As COMPOSITOR_LOD_CLUE = CLC_Generic) As pdDIB
    Set tmpLODDIB = lodDIB(cacheID)
End Function

Private Sub Class_Terminate()
    Set layerDIB = Nothing
    Set layerThumbnail = Nothing
    Set myTextRenderer = Nothing
    
    Dim i As Long
    For i = 0 To NUM_OF_LOD_CACHES - 1
        If Not (lodDIB(i) Is Nothing) Then
            lodDIB(i).eraseDIB
            Set lodDIB(i) = Nothing
        End If
    Next i
End Sub
