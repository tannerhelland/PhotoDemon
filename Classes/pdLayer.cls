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
'Copyright ©2013-2014 by Tanner Helland
'Created: 29/August/12
'Last updated: 18/March/14
'Last update: moved all DIB-centric code to the pdDIB class, and started rewriting this class as a source-agnostic layer
'              wrapper.  Going forward, PD will support many different types of layers, and it will be up to this class
'              to manage those types in a relatively OOP way.
'
'Description forthcoming once more features have been implemented...
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

'Layer type is set at creation time.  Once a layer has been created, its type cannot be changed.
Private myLayerType As LAYER_TYPE

Public Enum LAYER_BLENDMODE
    BL_NORMAL = 0
    BL_MULTIPLY = 1
    BL_SCREEN = 2
    BL_OVERLAY = 3
    BL_HARDLIGHT = 4
    BL_SOFTLIGHT = 5
End Enum

#If False Then
    Const BL_NORMAL = 0, BL_MULTIPLY = 1, BL_SCREEN = 2, BL_OVERLAY = 3, BL_HARDLIGHT = 4, BL_SOFTLIGHT = 5
#End If

'Canonical ID value for this layer.  This value is set once, at layer creation time, and can never be changed again.  It is
' persistent for the life of the layer, the life of the parent pdImage object (meaning no other layer in that image will
' ever be assigned this ID value), and the life of any copies of the parent pdImage saved to file.
'
'Because this value cannot be changed once created, I have separated it out from the LayerData enum.
Private layerID As Long

'Layers store a lot of layer-type-agnostic metadata. **All of these entries can be changed by the user at run-time.**
' Default values will be set at creation time, but never assume that those values are constant, and never attempt to reference
'  a layer by one of these properties (INCLUDING NAME, as it can be modified without warning)
Private Type LayerData
    l_Name As String                'Layer name, as entered by the user
    l_GroupID As Long               'Layer group (0 if not assigned to a group)
    l_Opacity As Long               'Layer opacity (defaults to 100)
    l_BlendMode As LAYER_BLENDMODE  'Layer blend mode; defaults to BL_NORMAL
    l_OffsetX As Double             'X/Y offset of the layer's top-left corner
    l_OffsetY As Double
    l_Angle As Double               'Layer angle.  (A long-term goal is to support non-destructive rotation of layers.)
    l_Visibility As Boolean         'Layer visibility
End Type

'Local instance of layer data for this layer
Private myLayerData As LayerData

'Layer width, height.  Note that this values are automatically updated as necessary, depending on the contents of the layer
'Private layerWidth As Long, layerHeight As Long

'A DIB that stores the contents of the layer, in DIB format.
' - For image layers, the layer contents will be stored in this DIB.
' - For non-image layers, a rendering of the layer contents will be stored in this DIB.  This DIB is used to render the layer,
'    rather than constantly redrawing whatever the layer's contents may be (text, etc)
' - For adjustment and other non-data layers, this DIB will remain unused.
'
'For performance reasons, I have made this DIB publicly accessible.  This allows edit functions to directly modify it as necessary.
' That said, any function that directly modifies the DIB must also take care of any management functions as well (e.g. redrawing
' the viewport), as layers have no way tp propagate commands up to their parent image.
Public layerDIB As pdDIB

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
        .l_Angle = 0
        .l_Visibility = True
    End With
    
    'Initialize the layer's DIB
    Set layerDIB = New pdDIB
    
    'Set the canonical ID to -1.  This can be used to determine if the layer has been activated.
    layerID = -1
    
End Sub

'Create a new image-type layer, using a DIB as the reference.  Note that this will automatically set the layer type to PDL_IMAGE.
Public Sub CreateNewImageLayer(ByVal newLayerType As LAYER_TYPE, ByRef srcDIB As pdDIB, Optional ByRef parentImage As pdImage, Optional ByVal newLayerName As String = "")

    'Mark the layer as being IMAGE type.
    myLayerType = PDL_IMAGE
    
    'Copy the name locally, and set other values to their natural defaults
    With myLayerData
        If Len(newLayerName) > 0 Then .l_Name = g_Language.TranslateMessage("New Image Layer")
        .l_GroupID = 0
        .l_Opacity = 100
        .l_BlendMode = BL_NORMAL
        .l_Visibility = True
    End With
    
    'Create a local copy of the passed DIB.
    Set layerDIB = New pdDIB
    layerDIB.createFromExistingDIB srcDIB
    
    'Set the initial offset so that the image is centered on the parent canvas
    If Not (parentImage Is Nothing) Then
        
        Dim layerDiff As Double
        layerDiff = parentImage.Width - layerDIB.getDIBWidth
        myLayerData.l_OffsetX = layerDiff / 2
        
        layerDiff = parentImage.Height - layerDIB.getDIBHeight
        myLayerData.l_OffsetY = layerDiff / 2
        
    End If
    
End Sub

'Copy an existing layer.  All layer contents will be copied manually, so make sure that new layer properties are
' manually added to this function!
Public Sub CopyExistingLayer(ByRef srcLayer As pdLayer)

    'Copy all relevant layer properties from the source layer
    myLayerType = srcLayer.getLayerType
    
    With myLayerData
        .l_Name = srcLayer.getLayerName
        .l_GroupID = srcLayer.getLayerGroup
        .l_Opacity = srcLayer.getLayerOpacity
        .l_BlendMode = srcLayer.getLayerBlendMode
        .l_OffsetX = srcLayer.getLayerOffsetX
        .l_OffsetY = srcLayer.getLayerOffsetY
        .l_Angle = srcLayer.getLayerAngle
        .l_Visibility = srcLayer.getLayerVisibility
    End With
    
    'Finally, copy the source layer's DIB.
    layerDIB.createFromExistingDIB srcLayer.layerDIB
    
    'Depending on the type of layer we are copying, copy any extra layer data, or generate a new layer mask to
    ' match the layer's vector contents.
    Select Case srcLayer.getLayerType
    
        Case PDL_IMAGE
        
        Case PDL_TEXT
        
        Case PDL_ADJUSTMENT
    
    End Select

End Sub
