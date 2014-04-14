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
' the viewport), as layers have no way to propagate commands up to their parent image.
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

'Write a copy of the layer header to an XML-format string.  All data necessary to recreate a layer from scratch will be
' included in this string, EXCEPT FOR DIB DATA.  For obvious performance reasons, the DIB contents of this layer must be
' handled separately, in binary format.
'
'If the optional parameter "avoidFluffEntries" is set, descriptive comments and pretty formatting will not be
' applied to the XML string.  This helpful for improving read/write performance for the data.
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
        xmlEngine.writeTag "Angle", .l_Angle
        xmlEngine.writeTag "Visibility", .l_Visibility
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

'The sister function to "getLayerHeaderAsXML" above, this function will initialize a layer header from an XML-format string.
'
'Note that this function will create a blank DIB for the layer DIB, using the dimensions and color depth stored in the
' XML string.  However, it will obviously NOT apply any stored image data.  That must be handled separately, via binary
' data interactions.
Public Function CreateNewImageLayerFromXML(ByRef xmlString As String, Optional ByVal useCustomLayerID As Long = -1) As Boolean

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
            .l_Angle = xmlEngine.getUniqueTag_Double("Angle")
            .l_Visibility = xmlEngine.getUniqueTag_Boolean("Visibility")
        End With
        
        'Next, we're going to do something kinda weird.  We're going to create a blank DIB for the layer contents
        ' using the DIB header supplied in the XML, but we're *not* going to fill the DIB!  An external function will
        ' handle that, because the actual DIB bits are stored in binary format.
        Set layerDIB = New pdDIB
        With xmlEngine
            layerDIB.createBlank .getUniqueTag_Long("DIB_Width", 1), .getUniqueTag_Long("DIB_Height", 1), .getUniqueTag_Long("DIB_ColorDepth", 32)
        End With
        
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
    End With
    
    'Create a local copy of the passed DIB.
    Set layerDIB = New pdDIB
    layerDIB.createFromExistingDIB srcDIB
    
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

'Assuming the current layer DIB is null-padded (a term I use to describe a layer that has been forced to the size of the
' image by surrounding it with transparent pixels), intelligently crop the DIB by removing all fully transparent padding.
' When this is done, modify the current layer offsets so that the image stays in effectively the same place, but without
' all that blank padding.
' IMPORTANT NOTE: if you call this function without first null-padding the layer DIB, weird shit may happen.  That is not
'                 an intended use-case, so don't attempt it!
'
'Returns TRUE if layer was trimmed successfully, FALSE if it was not.  (False occurs if the entire layer is transparent.)
Public Function cropNullPaddedLayerDIB() As Boolean
    
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
            cropNullPaddedLayerDIB = False
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
            cropNullPaddedLayerDIB = False
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
        Dim quickVal As Long
        
        For x = 0 To finalX
            quickVal = x * 4 + 3
        For y = initY To finalY
        
            If srcImageData(quickVal, y) <> 0 Then
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
            quickVal = x * 4 + 3
        For y = initY To finalY
        
            If srcImageData(quickVal, y) <> 0 Then
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
            cropNullPaddedLayerDIB = False
        Else
        
            'Trimmable borders were found.  Use the calculated values to crop the image to its smallest relevant size.
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createBlank newRight - newLeft, newBottom - newTop, 32, 0
            
            BitBlt tmpDIB.getDIBDC, 0, 0, newRight - newLeft, newBottom - newTop, layerDIB.getDIBDC, newRight, newTop, vbSrcCopy
            
            'Replace the current layer DIB with our temporary DIB
            layerDIB.createFromExistingDIB tmpDIB
            
            'Modify our layer offsets to reflect our new size
            myLayerData.l_OffsetX = newRight
            myLayerData.l_OffsetY = newTop
        
            'Return TRUE, so that the calling function knows to redraw the screen with the new size and offsets
            cropNullPaddedLayerDIB = True
            
        End If
        
    Else
        Debug.Print "Request for layer null padding trim denied - layer DIB does not currently exist!"
        cropNullPaddedLayerDIB = False
    End If
    
End Function

'External functions can use this function to request a thumbnail version of the layer contents.
'
'FreeImage is preferred for thumbnail resizing, but if it's not available, GDI+ will be used.
Public Function requestThumbnail(ByRef dstThumbnailDIB As pdDIB, Optional ByVal thumbnailSize As Long = 64) As Boolean
    
    'If the layer has not been instantiated properly, reject the thumbnail request
    If (layerDIB Is Nothing) Or (layerDIB.getDIBWidth = 0) Or (layerDIB.getDIBHeight = 0) Then
        requestThumbnail = False
        Exit Function
    End If
    
    'Thumbnails have some interesting requirements.  We always want them to be square, with the image set in the middle
    ' of the thumbnail (with aspect ratio preserved) and any empty edges made transparent.
    
    'Start by determining an aspect ratio for the current image.
    Dim aspectRatio As Double
    aspectRatio = CDbl(layerDIB.getDIBWidth) / CDbl(layerDIB.getDIBHeight)
    
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
        fi_DIB = FreeImage_CreateFromDC(layerDIB.getDIBDC)
        
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
        requestThumbnail = GDIPlusResizeDIB(dstThumbnailDIB, tX, tY, tIcoWidth, tIcoHeight, layerDIB, 0, 0, layerDIB.getDIBWidth, layerDIB.getDIBHeight, InterpolationModeHighQualityBilinear)
        
    End If
    
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
