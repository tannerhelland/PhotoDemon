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
'Copyright ©2012-2014 by Tanner Helland
'Created: 29/August/12
'Last updated: 20/June/14
'Last update: add comprehensive support for non-destructive layer FX
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
    PDL_ADJUSTMENT = 2
End Enum

#If False Then
    Const PDL_IMAGE = 0, PDL_TEXT = 1, PDL_ADJUSTMENT = 2
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
End Enum

#If False Then
    Const BL_NORMAL = 0, BL_DARKEN = 1, BL_MULTIPLY = 2, BL_COLORBURN = 3, BL_LINEARBURN = 4, BL_LIGHTEN = 5, BL_SCREEN = 6, BL_COLORDODGE = 7
    Const BL_LINEARDODGE = 8, BL_OVERLAY = 9, BL_SOFTLIGHT = 10, BL_HARDLIGHT = 11, BL_VIVIDLIGHT = 12, BL_LINEARLIGHT = 13, BL_PINLIGHT = 14
    Const BL_HARDMIX = 15, BL_DIFFERENCE = 16, BL_EXCLUSION = 17, BL_SUBTRACT = 18, BL_DIVIDE = 19
    Const BL_HUE = 20, BL_SATURATION = 21, BL_COLOR = 22, BL_LUMINOSITY = 23
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
    l_Angle As Double                   'Layer angle.  (A long-term goal is to support non-destructive rotation of layers.)
    l_Visibility As Boolean             'Layer visibility
    l_NonDestructiveFXActive As Boolean 'Are any non-destructive edits or other FX active?
End Type

'Local instance of layer data for this layer
Private myLayerData As LayerData

'Non-destructive edits currently supported by the program.
Public Enum LAYER_NONDESTRUCTIVE_FX
    NDFX_EXPOSURE = 0
    NDFX_CONTRAST = 1
End Enum

#If False Then
    Const NDFX_EXPOSURE = 0, NDFX_CONTRAST = 1
#End If

'Local copy of non-destructive FX
Private Type nonDestructiveFXSettings
    fx_Exposure As Long
    fx_Contrast As Long
End Type

Private myNonDestructiveFX As nonDestructiveFXSettings

'Layer width, height.  Note that this values are automatically updated as necessary, depending on the contents of the layer
'Private layerWidth As Long, layerHeight As Long

'A DIB that stores the contents of the layer, in DIB format.
' - For image layers, the layer contents will be stored in this DIB.
' - For non-image layers, a rendering of the layer contents will be stored in this DIB.  This DIB is used to render the layer
'    onto the viewport, rather than constantly recreating whatever the layer's contents may be (text, etc)
' - For adjustment and other non-data layers, this DIB will remain unused.
'
'For performance reasons, I have chosen to make layer DIBs publicly accessible.  This allows edit functions to directly modify
' them as necessary.  That said, any function that directly modifies the DIB must also take care of any management functions
' as well (e.g. redrawing the viewport), because layers have no way to propagate change notifications up to their parent image.
Public layerDIB As pdDIB

'Layer thumbnail.  This 256x256 DIB is cached locally, then passed to external functions only when necessary.  Note that requests
' for sizes larger than 256x256 (in either dimension) will force creation of a custom thumbnail from scratch, so try to avoid that
' if at all possible!
Private Const LAYER_THUMB_SIZE As Long = 256
Private layerThumbnail As pdDIB

'If the current thumbnail is clean or dirty.  Clean means that to the best of this class's knowledge, the current thumbnail accurately
' reflects the layer's contents.  Dirty means that we know the current thumbnail is invalid, so regenerate it before continuing.
Private m_IsThumbClean As Boolean


'If an external function modifies this layer's DIB contents (or the vector data from which the DIB is generated),
' *IT MUST CALL THIS FUNCTION AFTERWARD*.
'
'As part of new program-wide optimizations in v6.4, I am trying to limit the amount of data that PD has to generate on the fly.
' For layers, we can reduce churn by caching things like the layer thumbnail, which should only be updated when the layer's
' contents are updated (destructively).
'
'This notifyLayerModified() function will rebuild all caches for this layer, so use it only when absolutely necessary.
Public Sub notifyLayerModified()

    'Mark the layer thumbnail as dirty
    m_IsThumbClean = False

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

'Get layer type.  Note that there is no SET layer type function; layer type is set at creation, and cannot currently be changed.
Public Function getLayerType() As LAYER_TYPE
    getLayerType = myLayerType
End Function

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

'Get/set blend mode
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

'Get/set layer canvas modifiers (necessary for non-destructive on-canvas resizing)
Public Function getLayerCanvasXModifier() As Double
    getLayerCanvasXModifier = myLayerData.l_CanvasXModifier
End Function

Public Function getLayerCanvasYModifier() As Double
    getLayerCanvasYModifier = myLayerData.l_CanvasYModifier
End Function

Public Sub setLayerCanvasXModifier(ByVal newXModifier As Double)
    If newXModifier <= 0 Then
        myLayerData.l_CanvasXModifier = 0.00000001
    Else
        myLayerData.l_CanvasXModifier = newXModifier
    End If
End Sub

Public Sub setLayerCanvasYModifier(ByVal newYModifier As Double)
    If newYModifier <= 0 Then
        myLayerData.l_CanvasYModifier = 0.00000001
    Else
        myLayerData.l_CanvasYModifier = newYModifier
    End If
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
        
        End With
    
    Else
        myLayerData.l_NonDestructiveFXActive = True
    End If
    
End Sub

'As a convenience to the user, modified width/height values can be retrieved using this function.  Note that
' error handling is automatically performed to prevent over/underflow.
Public Function getLayerCanvasWidthModified() As Long
    getLayerCanvasWidthModified = layerDIB.getDIBWidth * myLayerData.l_CanvasXModifier
    If getLayerCanvasWidthModified < 1 Then getLayerCanvasWidthModified = 1
End Function

Public Function getLayerCanvasHeightModified() As Long
    getLayerCanvasHeightModified = layerDIB.getDIBHeight * myLayerData.l_CanvasYModifier
    If getLayerCanvasHeightModified < 1 Then getLayerCanvasHeightModified = 1
End Function

'If this layer needs to have its on-canvas transformations made permanent (for example, when applying an action
' that involves a selection), this function can be used.
Public Sub makeCanvasTransformsPermanent()

    'If this layer does not have any active transforms, ignore this request
    If (myLayerData.l_CanvasXModifier = 1) And (myLayerData.l_CanvasYModifier = 1) Then Exit Sub

    'Resize the current layer DIB to match its on-canvas transform values
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    tmpDIB.createBlank getLayerCanvasWidthModified, getLayerCanvasHeightModified, 32, 0
    GDIPlusResizeDIB tmpDIB, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, layerDIB, 0, 0, layerDIB.getDIBWidth, layerDIB.getDIBHeight, InterpolationModeHighQualityBicubic
    
    layerDIB.createFromExistingDIB tmpDIB
    Set tmpDIB = Nothing
    
    'Reset the width/height modifiers for this layer
    myLayerData.l_CanvasXModifier = 1
    myLayerData.l_CanvasYModifier = 1

End Sub

'Get/set layer angle
Public Function getLayerAngle() As Double
    getLayerAngle = myLayerData.l_Angle
End Function

Public Sub setLayerAngle(ByVal newAngle As Double)
    myLayerData.l_Angle = newAngle
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
    End With
    
    With myNonDestructiveFX
        .fx_Exposure = 0
        .fx_Contrast = 0
    End With
    
    'Initialize the layer's DIB
    Set layerDIB = New pdDIB
    
    'Initialize the layer thumbnail and note that it is not clean, which will force a regeneration when the thumbnail is next requested.
    Set layerThumbnail = New pdDIB
    m_IsThumbClean = False
    
    'Set the canonical ID to -1.  This can be used to determine if the layer has been activated.
    layerID = -1
    
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
    End With
    
    'Next, write out any non-destructive FX parameters
    With myNonDestructiveFX
        xmlEngine.writeTag "NDFX_Exposure", .fx_Exposure
        xmlEngine.writeTag "NDFX_Contrast", .fx_Contrast
    End With
    
    'Next, write out the header of the layer's DIB contents.  These values are necessary for reconstructing the layer's DIB
    ' in the future.
    With xmlEngine
        .writeTag "DIB_ColorDepth", layerDIB.getDIBColorDepth
        .writeTag "DIB_Width", layerDIB.getDIBWidth
        .writeTag "DIB_Height", layerDIB.getDIBHeight
        .writeTag "DIB_ArrayWidth", layerDIB.getDIBArrayWidth
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
        getLayerHeaderAsParamString = buildParams(.l_Name, .l_GroupID, .l_Opacity, .l_BlendMode, .l_OffsetX, .l_OffsetY, .l_CanvasXModifier, .l_CanvasYModifier, .l_Angle, .l_Visibility, .l_NonDestructiveFXActive)
    End With
    
    'Next, add any non-destructive FX parameters
    With myNonDestructiveFX
        getLayerHeaderAsParamString = getLayerHeaderAsParamString & "|" & buildParams(.fx_Exposure, .fx_Contrast)
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
    End With
    
    'After the layer header comes non-destructive FX settings.  Make sure to adjust these indices if new values get added to
    ' myLayerData, above!
    With myNonDestructiveFX
        .fx_Exposure = cParams.GetLong(12)
        .fx_Contrast = cParams.GetLong(13)
    End With
        
    'Per PD convention, this class has no way to notify outsiders of updates, so the calling function must handle any
    ' viewport redraws or other interface changes resulting from this assignment.

End Sub

'The sister function to "getLayerHeaderAsXML" above, this function will initialize a layer header from an XML-format string.
'
'Note that if the optional createNonDestructively parameter is FALSE (as it is by default), this function will create a
' blank DIB for the layer DIB, using the dimensions and color depth stored in the XML string.  However, it will obviously
' NOT apply any stored image data.  That must be handled separately, via binary data interactions.
Public Function CreateNewImageLayerFromXML(ByRef xmlString As String, Optional ByVal useCustomLayerID As Long = -1, Optional ByVal createNonDestructively As Boolean = False) As Boolean

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
            .l_BlendMode = xmlEngine.getUniqueTag_Long("BlendMode")
            .l_OffsetX = xmlEngine.getUniqueTag_Double("OffsetX")
            .l_OffsetY = xmlEngine.getUniqueTag_Double("OffsetY")
            .l_CanvasXModifier = xmlEngine.getUniqueTag_Double("CanvasXModifier", 1)
            .l_CanvasYModifier = xmlEngine.getUniqueTag_Double("CanvasYModifier", 1)
            .l_Angle = xmlEngine.getUniqueTag_Double("Angle")
            .l_Visibility = xmlEngine.getUniqueTag_Boolean("Visibility")
            .l_NonDestructiveFXActive = xmlEngine.getUniqueTag_Boolean("NonDestructiveFXActive", False)
        End With
        
        'Read in the contents of the non-destructive FX chunk
        With myNonDestructiveFX
            .fx_Exposure = xmlEngine.getUniqueTag_Long("NDFX_Exposure", 0)
            .fx_Contrast = xmlEngine.getUniqueTag_Long("NDFX_Contrast", 0)
        End With
        
        'Next, we're going to do something kinda weird.  We're going to create a blank DIB for the layer contents
        ' using the DIB header supplied in the XML, but we're *not* going to fill the DIB!  An external function will
        ' handle that, because the actual DIB bits are stored in binary format.
        '
        'Note that this behavior can be overridden by setting the optional createNonDestructively parameter to TRUE.
        ' This is used by the Undo/Redo engine to re-load an image's header data, without actually overwriting the
        ' current layer DIB.
        If Not createNonDestructively Then
        
            Set layerDIB = New pdDIB
            With xmlEngine
                layerDIB.createBlank .getUniqueTag_Long("DIB_Width", 1), .getUniqueTag_Long("DIB_Height", 1), .getUniqueTag_Long("DIB_ColorDepth", 32)
            End With
            
            'Note that the layer thumbnail is no longer valid
            m_IsThumbClean = False
            
        End If
        
        CreateNewImageLayerFromXML = True
    
    'If the layer XML data can't be validated, there's nothing we can do but exit.
    Else
        Debug.Print "Layer could not be created: are you sure the supplied data was valid?"
        CreateNewImageLayerFromXML = False
    End If
    
End Function

'Create a new image-type layer, using a DIB as the reference.  Note that this will automatically set the layer type to PDL_IMAGE.
' If the parentImage parameter is supplied, the layer will automatically center itself on the parent image.
Public Sub CreateNewImageLayer(ByRef srcDIB As pdDIB, Optional ByRef parentImage As pdImage, Optional ByVal newLayerName As String = "")

    'Mark the layer as being IMAGE type.
    myLayerType = PDL_IMAGE
    
    'Copy the name locally, and set other values to their natural defaults
    With myLayerData
        If Len(newLayerName) > 0 Then
            .l_Name = newLayerName
        Else
            .l_Name = g_Language.TranslateMessage("New Image Layer")
        End If
        .l_GroupID = 0
        .l_Opacity = 100
        .l_BlendMode = BL_NORMAL
        .l_Visibility = True
        .l_NonDestructiveFXActive = False
    End With
    
    'Reset any non-destructive fx parameters
    With myNonDestructiveFX
        .fx_Exposure = 0
        .fx_Contrast = 0
    End With
    
    'Create a local copy of the passed DIB.
    Set layerDIB = New pdDIB
    layerDIB.createFromExistingDIB srcDIB
    
    'Note that the layer thumbnail is not valid
    m_IsThumbClean = False
    
    'Set the initial offset so that the image is centered on the parent canvas
    'If Not (parentImage Is Nothing) Then
    '
    '    Dim layerDiff As Double
    '    layerDiff = parentImage.Width - layerDIB.getDIBWidth
    '    myLayerData.l_OffsetX = layerDiff / 2
    '
    '    layerDiff = parentImage.Height - layerDIB.getDIBHeight
    '    myLayerData.l_OffsetY = layerDiff / 2
    '
    'Else
        myLayerData.l_OffsetX = 0
        myLayerData.l_OffsetY = 0
    'End If
    
End Sub

'Copy an existing layer.  All layer contents will be copied manually, so make sure that new layer properties are
' manually added to this function!
Public Sub CopyExistingLayer(ByRef srcLayer As pdLayer)

    'Copy all relevant layer properties from the source layer.  (Note that canonical layer ID is *not* copied; that must always
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
    End With
    
    'Next, copy any non-destructive FX parameters
    With myNonDestructiveFX
        .fx_Exposure = srcLayer.getLayerNonDestructiveFXValue(NDFX_EXPOSURE)
        .fx_Contrast = srcLayer.getLayerNonDestructiveFXValue(NDFX_CONTRAST)
    End With
    
    'Finally, copy the source layer's DIB.
    layerDIB.createFromExistingDIB srcLayer.layerDIB
    
    'Note that the layer thumbnail is no longer valid
    m_IsThumbClean = False
    
    'Depending on the type of layer we are copying, copy any extra layer data, or generate a new layer mask to
    ' match the layer's vector contents.
    Select Case srcLayer.getLayerType
    
        Case PDL_IMAGE
        
        Case PDL_TEXT
        
        Case PDL_ADJUSTMENT
    
    End Select

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
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createBlank newRight - newLeft + 1, newBottom - newTop + 1, 32, 0
            Debug.Print "Crop rect: (" & newLeft & "," & newTop & ") - (" & newRight & "," & newBottom & ")"
            BitBlt tmpDIB.getDIBDC, 0, 0, newRight - newLeft + 1, newBottom - newTop + 1, layerDIB.getDIBDC, newLeft, newTop, vbSrcCopy
            
            'Replace the current layer DIB with our temporary DIB
            layerDIB.createFromExistingDIB tmpDIB
            
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
' TODO: determine if off-image layer cropping is an acceptable solution in all scenarios.
Public Sub convertToNullPaddedLayer(ByVal parentImageWidth As Long, ByVal parentImageHeight As Long, Optional ByVal applyCanvasTransformsFirst As Boolean = True)

    'If this layer has transforms applied to it, we have to make them permanent before proceeding.
    ' Otherwise, the resulting image may not look right.  (Note that this behavior can be overridden if a function
    ' is ABSOLUTELY CERTAIN that the end result won't botch the transform - this is only true for pixels transforms
    ' that operate strictly in separate x/y domains, such as resizing).
    If applyCanvasTransformsFirst Then makeCanvasTransformsPermanent
    
    'Create a blank destination DIB at the size of the image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createBlank parentImageWidth, parentImageHeight, 32, 0
    
    'Copy our current layer into the temporary DIB, with proper offsets applied
    BitBlt tmpDIB.getDIBDC, myLayerData.l_OffsetX, myLayerData.l_OffsetY, layerDIB.getDIBWidth, layerDIB.getDIBHeight, layerDIB.getDIBDC, 0, 0, vbSrcCopy
    
    'Replace our current layer DIB with the temporary DIB
    layerDIB.createFromExistingDIB tmpDIB
    
    'Reset our offset values
    myLayerData.l_OffsetX = 0
    myLayerData.l_OffsetY = 0
    
    'Release the temporary DIB
    Set tmpDIB = Nothing
    
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
        FreeImageResizeDIBFast layerThumbnail, tX, tY, tIcoWidth, tIcoHeight, layerDIB, FILTER_BILINEAR, True
    Else
        GDIPlusResizeDIB layerThumbnail, tX, tY, tIcoWidth, tIcoHeight, layerDIB, 0, 0, layerDIB.getDIBWidth, layerDIB.getDIBHeight, InterpolationModeHighQualityBicubic
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
    Set dstThumbnailDIB = New pdDIB
    
    'If the requested thumbnail size differs from our cached version, resize the thumbnail prior to passing it along
    If thumbnailSize <> LAYER_THUMB_SIZE Then
        
        dstThumbnailDIB.createBlank thumbnailSize, thumbnailSize, 32, 0
        
        If g_ImageFormats.FreeImageEnabled Then
            FreeImageResizeDIBFast dstThumbnailDIB, 0, 0, thumbnailSize, thumbnailSize, layerThumbnail, FILTER_BILINEAR, True
        Else
            GDIPlusResizeDIB dstThumbnailDIB, 0, 0, thumbnailSize, thumbnailSize, layerThumbnail, 0, 0, layerThumbnail.getDIBWidth, layerThumbnail.getDIBHeight, InterpolationModeHighQualityBicubic
        End If
        
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
' will check the passed mouse values (which have already been translated into the IMAGE coordinate space), and return a Long-type
' value specifying whether or not that coordinate is of interest to this layer.  This basically provides a flexible framework for
' supporting mouse events for any layer, because it's up to the layer to determine which points are interesting and which are not.
'
'INPUTS: current mouse coords, already translated into the IMAGE coordinate space (not the canvas coordinate space!)
'
'OUTPUT: index of the point of interest, if any.  If the mouse is not over a point of interest, -1 is returned.
Public Function checkForPointOfInterest(ByVal imgX As Long, ByVal imgY As Long) As Long
    
    'MouseAccuracy in PD is a global value, but because we are working in image coordinates, we must compensate for the
    ' current zoom value.  (Otherwise, when zoomed out the user would be forced to work with tighter accuracy!)
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
        Case PDL_IMAGE, PDL_TEXT
            
            'Manually populate the point of interest array with the layer's corner coordinates.
            ReDim poiList(0 To 3) As POINTAPI
            poiList(0).x = getLayerOffsetX
            poiList(0).y = getLayerOffsetY
            poiList(1).x = getLayerOffsetX + (layerDIB.getDIBWidth * getLayerCanvasXModifier)
            poiList(1).y = getLayerOffsetY
            poiList(2).x = getLayerOffsetX + (layerDIB.getDIBWidth * getLayerCanvasXModifier)
            poiList(2).y = getLayerOffsetY + (layerDIB.getDIBHeight * getLayerCanvasYModifier)
            poiList(3).x = getLayerOffsetX
            poiList(3).y = getLayerOffsetY + (layerDIB.getDIBHeight * getLayerCanvasYModifier)
            
            'Check to see if one of the corners has been clicked
            closestPoint = findClosestPointInArray(imgX, imgY, minDistance, poiList)
            
            'If a corner was clicked, return its value and exit
            If closestPoint <> -1 Then
                checkForPointOfInterest = closestPoint
                Exit Function
            End If
            
            'If we're at this line of code, a corner was not clicked. Perform one final check to see if the mouse is within
            ' the layer's boundaries, and if it is, return the "move selection" ID, then exit.
            If (imgX > poiList(0).x) And (imgX < poiList(1).x) And (imgY > poiList(0).y) And (imgY < poiList(2).y) Then
                checkForPointOfInterest = 4
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
    End With
    
    'Assign default values to the non-destructive FX params
    With myNonDestructiveFX
        .fx_Exposure = 0
        .fx_Contrast = 0
    End With

End Sub
