Attribute VB_Name = "Layers"
'***************************************************************************
'Layer Interface
'Copyright 2014-2026 by Tanner Helland
'Created: 24/March/14
'Last updated: 14/November/23
'Last update: new support for adding multiple layers from file at once (Layers > Add > from File...)
'
'This module provides all layer-related functions that interact with PhotoDemon's central processor.  Most of these
' functions are triggered by either the Layer menu, or the Layer toolbox.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Enum PD_LayerType
    PDL_Image = 0
    PDL_TextBasic = 1
    PDL_TextAdvanced = 2
    PDL_Adjustment = 3
End Enum

#If False Then
    Const PDL_Image = 0, PDL_TextBasic = 1, PDL_TextAdvanced = 2, PDL_Adjustment = 3
#End If

'This class supports getting/setting layer properties via generic Get/SetGenericLayerProperty functions.  This enum is
' used to differentiate between layer properties, and any additions to LayerData, above, should be mirrored here.
Public Enum PD_LayerGenericProperty
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
    pgp_RotateCenterX = 15
    pgp_RotateCenterY = 16
    pgp_FrameTime = 17
    pgp_MaskExists = 18
    pgp_MaskActive = 19
End Enum

#If False Then
    Private Const pgp_Name = 0, pgp_GroupID = 1, pgp_Opacity = 2, pgp_BlendMode = 3, pgp_OffsetX = 4, pgp_OffsetY = 5, pgp_CanvasXModifier = 6, pgp_CanvasYModifier = 7, pgp_Angle = 8, pgp_Visibility = 9, pgp_NonDestructiveFXActive = 10
    Private Const pgp_ResizeQuality = 11, pgp_ShearX = 12, pgp_ShearY = 13, pgp_AlphaMode = 14, pgp_RotateCenterX = 15, pgp_RotateCenterY = 16, pgp_FrameTime = 17, pgp_MaskExists = 18, pgp_MaskActive = 19
#End If

'Layer resize quality is defined different from other resampling options in the project.
' (Only a subset of options are exposed, for performance reasons.)
Public Enum PD_LayerResizeQuality
    LRQ_NearestNeighbor = 0
    LRQ_Bilinear = 1
    LRQ_Bicubic = 2
End Enum

#If False Then
    Const LRQ_NearestNeighbor = 0, LRQ_Bilinear = 1, LRQ_Bicubic = 2
#End If

'Used when converting layers to standalone images and vice-versa
Private Type LayerConvertCache
    id As Long
    mustConvert As Boolean
    srcLayerName As String
    srcImageWidth As Long
    srcImageHeight As Long
End Type
    
'XML-based wrapper for AddBlankLayer(), below
Public Sub AddBlankLayer_XML(ByRef processParameters As String)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    Layers.AddBlankLayer cParams.GetLong("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex), cParams.GetLong("layertype", PDL_Image)
End Sub

'Add a blank 32bpp layer above the specified layer index (typically the currently active layer).
' RETURNS: layer ID (*not* index!) of the newly created layer.  PD doesn't always make use of this value,
' but it's there if you need it.
Public Function AddBlankLayer(ByVal dLayerIndex As Long, Optional ByVal newLayerType As PD_LayerType = PDL_Image) As Long

    'Validate the requested layer index
    If (dLayerIndex < 0) Then dLayerIndex = 0
    If (dLayerIndex > PDImages.GetActiveImage.GetNumOfLayers - 1) Then dLayerIndex = PDImages.GetActiveImage.GetNumOfLayers - 1
    
    'Ask the parent pdImage to create a new layer object
    Dim newLayerID As Long
    newLayerID = PDImages.GetActiveImage.CreateBlankLayer(dLayerIndex)
    
    'Until vector layers are implemented, let's just assign the newly created layer the IMAGE type,
    ' and initialize it to the size of the image.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateBlank PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, 32, 0, 0
    tmpDIB.SetInitialAlphaPremultiplicationState True
    PDImages.GetActiveImage.GetLayerByID(newLayerID).InitializeNewLayer newLayerType, g_Language.TranslateMessage("Blank layer"), tmpDIB
    
    'Make the blank layer the new active layer
    PDImages.GetActiveImage.SetActiveLayerByID newLayerID
    
    'Notify the parent of the change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.NotifyLayerChange
    
    'Render the new image to screen (not technically necessary, but doesn't hurt)
    Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
            
    'Synchronize the interface to the new image
    SyncInterfaceToCurrentImage
    
    AddBlankLayer = newLayerID
    
End Function

'XML-based wrapper for AddNewLayer(), below
Public Sub AddNewLayer_XML(ByRef processParameters As String)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    
    'NEW (cheat) for 8.0.
    ' (A better, "proper" implementation of this is TODO post-8.0.)
    
    'PD's 8.0 release shipped with a new PSD import/export engine.  When writing the PSD engine,
    ' I had to make some hard choices about what Photoshop features to support, especially in cases
    ' where PD doesn't support a direct analog of a PS feature.
    
    'One such compromise was layer groups.  PD doesn't support layer groups.  PS does.
    ' My workaround (inspired by the 3rd-party Paint.NET PSD plugin) is to add "dummy layers"
    ' as group start/end markers.  These "dummy" layers use a specific naming scheme, and when found,
    ' PhotoDemon automatically maps them to/from PSD layers at export/import time.  This is a
    ' temporary solution until PhotoDemon natively supports layer groups.
    
    'This implementation obviously creates some annoyances when trying to add layer groups to an
    ' image destined for cross-support with Photoshop.  As an additional hacky workaround, I've
    ' modified PD's "Add new layer" wrapper (this function!) so it scans for layer group names when
    ' adding new layers.  If it finds a group-compatible name, it will automatically add a *pair* of
    ' group-compatible layers so that you don't have to manually add a second "dummy" layer.
    Dim useGroupWorkaround As Boolean
    useGroupWorkaround = False
    
    Dim targetStartName As String, targetEndName As String
    targetStartName = g_Language.TranslateMessage("Group start:")
    targetEndName = g_Language.TranslateMessage("Group end:")
    
    'If the passed layer name starts with "Group start:" or "Group end:", activate the layer
    ' group workaround.
    Dim srcLayerName As String, srcLayerNameNoGroup As String
    srcLayerName = Trim$(cParams.GetString("layername", vbNullString, True))
    If (LenB(srcLayerName) >= LenB(targetStartName)) Then
        useGroupWorkaround = Strings.StringsEqual(targetStartName, Left$(srcLayerName, Len(targetStartName)), True)
        If useGroupWorkaround And (LenB(srcLayerName) > LenB(targetStartName)) Then srcLayerNameNoGroup = Right$(srcLayerName, Len(srcLayerName) - Len(targetStartName))
    End If
    If (Not useGroupWorkaround) Then
        If (LenB(srcLayerName) >= LenB(targetEndName)) Then
            useGroupWorkaround = Strings.StringsEqual(targetEndName, Left$(srcLayerName, Len(targetEndName)), True)
            If useGroupWorkaround And (LenB(srcLayerName) > LenB(targetEndName)) Then srcLayerNameNoGroup = Right$(srcLayerName, Len(srcLayerName) - Len(targetEndName))
        End If
    End If
    
    'If the user supplied a group-compatible name, add *two* layers (one for group start, one for group end)
    Dim newLayerID As Long
    If useGroupWorkaround Then
        With cParams
            newLayerID = Layers.AddNewLayer(.GetLong("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex), .GetLong("layertype", PDL_Image), .GetLong("layersubtype", 0), .GetLong("layercolor", vbBlack), .GetLong("layerposition", 0), .GetBool("activatelayer", True), targetStartName & " " & srcLayerNameNoGroup, suspendRedraws:=True)
            newLayerID = Layers.AddNewLayer(.GetLong("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex), .GetLong("layertype", PDL_Image), .GetLong("layersubtype", 0), .GetLong("layercolor", vbBlack), .GetLong("layerposition", 0), .GetBool("activatelayer", True), targetEndName & " " & srcLayerNameNoGroup)
        End With
    Else
        With cParams
            newLayerID = Layers.AddNewLayer(.GetLong("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex), .GetLong("layertype", PDL_Image), .GetLong("layersubtype", 0), .GetLong("layercolor", vbBlack), .GetLong("layerposition", 0), .GetBool("activatelayer", True), .GetString("layername"))
        End With
    End If
    
End Sub

'Add a non-blank 32bpp layer to the image.  (This function is used by the Add New Layer button on the layer box.)
' RETURNS: layer ID (*not* index!) of the newly created layer.  PD doesn't always make use of this value,
' but it's there if you need it.
Public Function AddNewLayer(ByVal dLayerIndex As Long, ByVal dLayerType As PD_LayerType, ByVal dLayerSubType As Long, ByVal dLayerColor As Long, ByVal dLayerPosition As Long, ByVal dLayerAutoSelect As Boolean, Optional ByVal dLayerName As String = vbNullString, Optional ByVal initialXOffset As Single = 0!, Optional ByVal initialYOffset As Single = 0!, Optional ByVal suspendRedraws As Boolean = False) As Long

    'Before making any changes, make a note of the currently active layer
    Dim prevActiveLayerID As Long
    prevActiveLayerID = PDImages.GetActiveImage.GetActiveLayerID
    
    'Validate the requested layer index
    If (dLayerIndex < 0) Then dLayerIndex = 0
    If (dLayerIndex > PDImages.GetActiveImage.GetNumOfLayers - 1) Then dLayerIndex = PDImages.GetActiveImage.GetNumOfLayers - 1
    
    'Ask the parent pdImage to create a new layer object
    Dim newLayerID As Long
    newLayerID = PDImages.GetActiveImage.CreateBlankLayer(dLayerIndex)
    
    'Assign the newly created layer the IMAGE type, and initialize it to the size of the image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'The parameters passed to the new DIB vary according to layer type.  Use the specified type to determine how we
    ' initialize the new layer.  (Note that this is only relevant for raster layers.)
    If (dLayerType = PDL_Image) Then
    
        Select Case dLayerSubType
        
            'Transparent (blank)
            Case 0
                tmpDIB.CreateBlank PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, 32, 0, 0
            
            'Black
            Case 1
                tmpDIB.CreateBlank PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, 32, vbBlack, 255
            
            'White
            Case 2
                tmpDIB.CreateBlank PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, 32, vbWhite, 255
            
            'Custom color
            Case 3
                tmpDIB.CreateBlank PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, 32, dLayerColor, 255
            
        End Select
        
    Else
    
        'Create a 1x1 transparent DIB to avoid errors; subsequent functions will resize the DIB as required
        tmpDIB.CreateBlank 1, 1, 32, 0, 0
    
    End If
    
    'Layers always start with premultiplied alpha
    tmpDIB.SetInitialAlphaPremultiplicationState True
    
    'Set the layer name
    If (LenB(Trim$(dLayerName)) = 0) Then
    
        Select Case dLayerType
        
            Case PDL_Image
                dLayerName = g_Language.TranslateMessage("Blank layer")
                
            Case PDL_TextBasic
                dLayerName = g_Language.TranslateMessage("Basic text layer")
                
            Case PDL_TextAdvanced
                dLayerName = g_Language.TranslateMessage("Advanced text layer")
        
        End Select
        
    End If
    
    'Assign the newly created DIB and layer name to the layer object
    PDImages.GetActiveImage.GetLayerByID(newLayerID).InitializeNewLayer dLayerType, dLayerName, tmpDIB
    
    'Apply initial layer offsets
    PDImages.GetActiveImage.GetLayerByID(newLayerID).SetLayerOffsetX initialXOffset
    PDImages.GetActiveImage.GetLayerByID(newLayerID).SetLayerOffsetY initialYOffset
    
    'Some layer types may require extra initialization steps in the future
    Select Case dLayerType
        
        Case PDL_Image
        
        'Set an initial width/height of 1x1
        Case PDL_TextBasic, PDL_TextAdvanced
            PDImages.GetActiveImage.GetLayerByID(newLayerID).SetLayerWidth 1!
            PDImages.GetActiveImage.GetLayerByID(newLayerID).SetLayerHeight 1!
        
    End Select
        
    'Activate the new layer
    PDImages.GetActiveImage.SetActiveLayerByID prevActiveLayerID
    
    'Move the layer into position as necessary.
    If (dLayerPosition <> 0) Then
    
        Select Case dLayerPosition
        
            'Place below current layer
            Case 1
                MoveLayerAdjacent PDImages.GetActiveImage.GetLayerIndexFromID(newLayerID), False, False
            
            'Move to top of stack
            Case 2
                MoveLayerToEndOfStack PDImages.GetActiveImage.GetLayerIndexFromID(newLayerID), True, False
            
            'Move to bottom of stack
            Case 3
                MoveLayerToEndOfStack PDImages.GetActiveImage.GetLayerIndexFromID(newLayerID), False, False
        
        End Select
        
        'Note that each of the movement functions, above, will call the necessary interface refresh functions,
        ' so we don't need to manually do it here.
        
    End If
    
    'Make the newly created layer the active layer
    If dLayerAutoSelect Then
        Layers.SetActiveLayerByID newLayerID, False, Not suspendRedraws
    Else
        Layers.SetActiveLayerByID prevActiveLayerID, False, Not suspendRedraws
    End If
    
    'Notify the parent of the change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
    
    'Redraw the main viewport (if requested)
    If (Not suspendRedraws) Then
        
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange
        
    End If
    
    AddNewLayer = newLayerID
    
End Function

'Create a new layer from the current composite image, and place it at the top of the layer stack.
' (If the replaceActiveLayerInstead parameter is TRUE, the contents of the active layer will be replaced
' instead - this is the difference between the `Layer > Add from visible layers` and `Layer > Replace
' from visible layers` commands.)
'RETURNS: layer ID (*not* index!) of the newly created layer.  PD doesn't always make use of this value,
' but it's there if you need it.
Public Function AddLayerFromVisibleLayers(Optional replaceActiveLayerInstead As Boolean = False) As Long
    
    Dim targetLayerID As Long, tmpDIB As pdDIB
    
    'Retrieve a composite of the current image
    PDImages.GetActiveImage.GetCompositedImage tmpDIB, True
        
    'This function can either 1) create a new layer, or 2) replace an existing layer.
    ' (Replacing is much simpler, since we can use most of the existing layer as-is - only the pixel surface
    ' needs to be modified.)
    If replaceActiveLayerInstead Then
    
        'Updating the layer is as simple as replacing its existing surface reference with tmpDIB.
        ' (Just make sure not to mess with or forcibly free tmpDIB after this point!)
        PDImages.GetActiveImage.GetActiveLayer.SetLayerDIB tmpDIB
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
        
    Else
    
        'Figure out where the top of the layer stack sits
        Dim topLayerIndex As Long
        topLayerIndex = PDImages.GetActiveImage.GetNumOfLayers - 1
        
        'Ask the parent pdImage to create a new layer object at the top of its stack,
        ' then initialize that layer using the composite image we retrieved earlier.
        targetLayerID = PDImages.GetActiveImage.CreateBlankLayer(topLayerIndex)
        PDImages.GetActiveImage.GetLayerByID(targetLayerID).InitializeNewLayer PDL_Image, g_Language.TranslateMessage("Visible"), tmpDIB
        
        'Make the blank layer the new active layer
        PDImages.GetActiveImage.SetActiveLayerByID targetLayerID
    
        'Notify the parent of the change
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
    
    End If
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.NotifyLayerChange
    
    'Render the new image to screen (not technically necessary, but doesn't hurt)
    Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
            
    'Synchronize the interface to the new image
    Interface.SyncInterfaceToCurrentImage
    
    AddLayerFromVisibleLayers = targetLayerID
    
End Function

'Shortcut function to mimic Ctrl+C, Ctrl+V, Remove selection
Public Sub AddLayerViaCopy()
    g_Clipboard.ClipboardCopy False, False, pdcf_InternalPD
    g_Clipboard.ClipboardPaste True
    Selections.RemoveCurrentSelection True
End Sub

'Shortcut function to mimic Ctrl+X, Ctrl+V, Remove selection
Public Sub AddLayerViaCut()
    g_Clipboard.ClipboardCut False, pdcf_InternalPD
    g_Clipboard.ClipboardPaste True
    Selections.RemoveCurrentSelection True
End Sub

'Shortcut function to create a new layer (sort of like AddLayerViaCopy/Cut above, but this
' will NOT touch the clipboard).  Requires an active selection.  You MUST validate selection state
' before calling this function.
Public Function AddLayerViaSelection(Optional ByVal preMultipliedAlphaState As Boolean = False, Optional ByVal useMergedImage As Boolean = False, Optional ByVal eraseOriginalPixels As Boolean = False) As Boolean
        
    'Failsafe check
    If (Not PDImages.IsImageActive) Then Exit Function
    If (Not PDImages.GetActiveImage.IsSelectionActive) Then Exit Function
    
    'Start by retrieving the selected pixels into a temporary DIB
    Dim tmpDIB As pdDIB
    If PDImages.GetActiveImage.RetrieveProcessedSelection(tmpDIB, preMultipliedAlphaState, useMergedImage) Then
        
        Dim i As Long
        
        'Before going any further, cache the ID of the currently active layer.
        ' (We need to refer to it later, and the *active* layer ID will change after we add our new layer.)
        Dim origLayerID As Long
        origLayerID = PDImages.GetActiveImage.GetActiveLayerID()
        
        'With the selection successfully retrieved, optionally erase their original values
        If eraseOriginalPixels Then
            
            'This step is simple if we are only cutting from a single layer...
            If (Not useMergedImage) Then
                PDImages.GetActiveImage.EraseProcessedSelection PDImages.GetActiveImage.GetActiveLayerIndex
            
            'If we're cutting from the merged image, we need to iterate all *visible* layers
            ' (that's important - we will ignore invisible layers) and erase any pixels that
            ' overlap the selected area.
            Else
                
                'Iterate all layers but IGNORE INVISIBLE layers when erasing
                For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
                    If PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility Then
                        PDImages.GetActiveImage.EraseProcessedSelection i
                    End If
                Next i
                
            End If
            
        End If
        
        'Create a name for the new layer.
        ' (This varies depending on whether the source was a layer or the merged image.)
        ' (Note also that we need to do this *before* creating a blank layer; otherwise the
        ' active layer will be the blank one we create!)
        Dim newLayerName As String
        If useMergedImage Then
            newLayerName = g_Language.TranslateMessage("New layer from selection")
        Else
            newLayerName = TextSupport.IncrementTrailingNumber(PDImages.GetActiveImage.GetActiveLayer.GetLayerName())
        End If
        
        'Ask the current image to prepare a blank layer for us
        Dim newLayerID As Long
        newLayerID = PDImages.GetActiveImage.CreateBlankLayer()
        
        'Convert the layer to an IMAGE-type layer and copy the newly loaded DIB's contents into it
        PDImages.GetActiveImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, newLayerName, tmpDIB
        
        'Note that tmpDIB now belongs to the parent image - do NOT erase it or modify pixels beyond this point.
        
        'Set the layer's initial (x, y) position to the current selection's top-left corner.
        PDImages.GetActiveImage.GetLayerByID(newLayerID).SetLayerOffsetX PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect.Left
        PDImages.GetActiveImage.GetLayerByID(newLayerID).SetLayerOffsetY PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect.Top
        
        'Set the new layer as the active layer
        PDImages.GetActiveImage.SetActiveLayerByID newLayerID
        
        'Notify the parent image that the entire image now needs to be recomposited, and note that
        ' the changes we've made are vector-safe if...
        ' 1) we're using copy (not cut), or...
        ' 2) we're using cut and the source layer(s) are not vector layers
        Dim undoRequest As PD_UndoType
        undoRequest = UNDO_Image_VectorSafe
        
        'If "cut" was used, we need to see if any vector layers were affected
        If eraseOriginalPixels Then
            If useMergedImage Then
                
                'Do a quick iterate, looking for visible vector layers.  If any exist, we need to cache
                ' a non-vector-safe Undo collection (which takes more time, so we avoid it unless necessary).
                For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
                    If PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility Then
                        If PDImages.GetActiveImage.GetLayerByIndex(i).IsLayerVector Then
                            undoRequest = UNDO_Image
                            Exit For
                        End If
                    End If
                Next i
                
            Else
                If PDImages.GetActiveImage.GetLayerByID(origLayerID).IsLayerVector Then undoRequest = UNDO_Image
            End If
        End If
        
        'Notify the parent image of the change, and make sure to pass UNDO_IMAGE (*not* the vector-safe variety)
        ' if a vector layer was chopped up.
        PDImages.GetActiveImage.NotifyImageChanged undoRequest
        
    End If

End Function

'Load an image file, and add it to the current image as a new layer.
' NOTE: this function is called from a lot of places!  Drag/drop, Edit > Paste, Layers > Replace layer, Layers > Add from file.
' Each of these paths has slightyl different considerations (e.g. Layers > Add from File can add multiple layers at once).
Public Function LoadImageAsNewLayer(ByVal raiseDialog As Boolean, Optional ByVal imagePath As String = vbNullString, Optional ByVal customLayerName As String = vbNullString, Optional ByVal createUndo As Boolean = False, Optional ByVal refreshUI As Boolean = True, Optional ByVal xOffset As Long = LONG_MAX, Optional ByVal yOffset As Long = LONG_MAX, Optional ByVal replaceActiveLayerInstead As Boolean = False) As Boolean

    'This function handles two cases: retrieving the filename from a common dialog box, and actually
    ' loading the image file and applying it to the current pdImage as a new layer.
    Dim listOfFiles As pdStringStack
    
    'If raiseDialog is TRUE, we need to get a file path from the user
    If raiseDialog Then
    
        'Retrieve one (or more) files from a common dialog.  Note that "Replace layer from file"
        ' only allows a single image, while "New layer from file" allows multi-select.
        Dim imgListCondensed As String
        If replaceActiveLayerInstead Then
            
            'Ask for a single file
            LoadImageAsNewLayer = FileMenu.PhotoDemon_OpenImageDialog_SingleFile(imgListCondensed, FormMain.hWnd)
            If LoadImageAsNewLayer Then
                Process "Replace layer from file", False, imgListCondensed, UNDO_Layer
            Else
                Exit Function
            End If
        
        'Multi-select version
        Else
            
            'Ask for as many files as the user wants
            LoadImageAsNewLayer = FileMenu.PhotoDemon_OpenImageDialog(listOfFiles, FormMain.hWnd)
            If LoadImageAsNewLayer Then
            
                'Serialize the stack to a single string object, then continue processing
                imgListCondensed = listOfFiles.SerializeStackToSingleString()
                Process "New layer from file", False, imgListCondensed, UNDO_Image_VectorSafe
                
            Else
                Exit Function
            End If
            
        End If
        
    'If raiseDialog is FALSE, the user has already selected a file, and we just need to load it
    Else
    
        'Prepare a temporary DIB
        Dim tmpDIB As pdDIB
        
        'We now have two branches:
        ' 1) "New layer from file"
        '     - This supports loading a whole bunch of files at once
        ' 2) "Replace layer from file"
        '     - This supports loading just ONE file at a time
        
        'REPLACE active layer
        If replaceActiveLayerInstead Then
            
            'Load the image file, and treat it as a single layer (if the source is multi-layer)
            Set tmpDIB = New pdDIB
            LoadImageAsNewLayer = Loading.QuickLoadImageToDIB(imagePath, tmpDIB)
            If LoadImageAsNewLayer Then
                
                'Forcibly convert the new layer to 32bpp
                ' (failsafe only; it should already be in 32-bpp mode from the loader)
                If (tmpDIB.GetDIBColorDepth <> 32) Then tmpDIB.ConvertTo32bpp
                
                'Easy-peasy - replace the current layer's backing surface with the contents of the newly loaded file
                PDImages.GetActiveImage.GetActiveLayer.SetLayerDIB tmpDIB
                PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
                
            'Failed to load the source image; the load function will have raised any relevant error UI
            Else
                PDDebug.LogAction "Image file could not be loaded as new layer.  (User cancellation is one possible outcome, FYI.)"
            End If
        
        'CREATE one or more new layers
        Else
            
            'The user may be loading multiple images.  Parse the source file string into a string stack.
            ' (This step will look for delimiters and safely convert the source string into a stack of filenames.)
            Set listOfFiles = New pdStringStack
            listOfFiles.RecreateStackFromSerializedString imagePath
            
            Dim i As Long
            For i = 0 To listOfFiles.GetNumOfStrings - 1
                
                imagePath = listOfFiles.GetString(i)
                
                'Load the image file, and treat it as a single layer (if the source is multi-layer)
                Set tmpDIB = New pdDIB
                LoadImageAsNewLayer = Loading.QuickLoadImageToDIB(imagePath, tmpDIB)
                If LoadImageAsNewLayer Then
                    
                    'Forcibly convert the new layer to 32bpp
                    ' (failsafe only; it should already be in 32-bpp mode from the loader)
                    If (tmpDIB.GetDIBColorDepth <> 32) Then tmpDIB.ConvertTo32bpp
                    
                    'Ask the current image to prepare a blank layer for us
                    Dim newLayerID As Long
                    newLayerID = PDImages.GetActiveImage.CreateBlankLayer()
                    
                    'Convert the layer to an IMAGE-type layer and copy the newly loaded DIB's contents into it
                    If (LenB(customLayerName) = 0) Then
                        PDImages.GetActiveImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, Trim$(Files.FileGetName(imagePath, True)), tmpDIB, True
                    Else
                        PDImages.GetActiveImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, customLayerName, tmpDIB, True
                    End If
                    
                    'With the layer successfully created, we now want to position it on-screen.  Rather than dump the layer
                    ' at (0, 0), let's be polite and place it at the top-left corner of the current viewport.
                    ' (Note that the caller can override this by supplying their own coordinates; if they do this,
                    ' we'll still validate the requested position to ensure it fits nicely in the current viewport.)
                    
                    'Start by retrieving the current top-left position of the canvas, in *image* coordinates
                    Dim newX As Double, newY As Double, safeX As Double, safeY As Double
                    If (Not PDImages.GetActiveImage Is Nothing) Then
                        Drawing.ConvertCanvasCoordsToImageCoords FormMain.MainCanvas(0), PDImages.GetActiveImage, 0#, 0#, newX, newY, True
                        safeX = newX
                        safeY = newY
                    Else
                        newX = 0
                        newY = 0
                    End If
                    
                    'If the caller passed in their own x and/or y value, validate each value before replacing
                    ' our auto-calculated position
                    If (xOffset <> LONG_MAX) Then
                        
                        newX = xOffset
                        
                        'Make sure the newly placed layer doesn't lie off-canvas.  (This is a risk with drag+drop,
                        ' as the use may drop the layer into blank areas around the image.)
                        If (newX > PDImages.GetActiveImage.Width - tmpDIB.GetDIBWidth) Then newX = PDImages.GetActiveImage.Width - tmpDIB.GetDIBWidth
                        If (newX < 0) Then newX = 0
                        
                        'Finally, if our calculated position lies to the left of the current viewport,
                        ' reposition it accordingly.  (This provides the most intuitive behavior, IMO, as it
                        ' guarantees the newly created layer will *always* be visible within the current
                        ' viewport, regardless of where the caller attempted to position it.)
                        If (newX < safeX) Then newX = safeX
                        
                    End If
                    
                    'Repeat all the above steps for y
                    If (yOffset <> LONG_MAX) Then
                        newY = yOffset
                        If (newY > PDImages.GetActiveImage.Height - tmpDIB.GetDIBHeight) Then newY = PDImages.GetActiveImage.Height - tmpDIB.GetDIBHeight
                        If (newY < 0) Then newY = 0
                        If (newY < safeY) Then newY = safeY
                    End If
                    
                    'Assign the new coordinates to our layer!
                    PDImages.GetActiveImage.GetLayerByID(newLayerID).SetLayerOffsetX newX
                    PDImages.GetActiveImage.GetLayerByID(newLayerID).SetLayerOffsetY newY
                    
                    'Notify the parent image that the entire image now needs to be recomposited
                    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
                    
                    'If loading multiple files at once, suspend this layer to free up working space for
                    ' future images.  (Images typically require a surplus of memory at load-time.)
                    If (listOfFiles.GetNumOfStrings > 1) Then PDImages.GetActiveImage.GetLayerByID(newLayerID).SuspendLayer True
                    
                'Failed to load the source image; the load function will have raised any relevant error UI
                Else
                    PDDebug.LogAction "Image file could not be loaded as new layer.  (User cancellation is one possible outcome, FYI.)"
                End If
                
            'Continue with any other files the user requested
            Next i
            
        '/END replace layer vs create new layer
        End If
        
        'If either layer path was successful, we now need to refresh the UI and create an Undo point
        If LoadImageAsNewLayer Then
            
            'If the caller wants us to manually create an Undo point, do so now.
            ' (In PD, this is only required when adding images via Drag/Drop, which require a custom pathway.
            ' Anything initiated by menus, hotkeys, or other "normal" interactions will generate Undo data
            ' automatically via PD's central processor.)
            If createUndo Then
                
                Dim tmpProcCall As PD_ProcessCall
                With tmpProcCall
                    .pcParameters = vbNullString
                    .pcRaiseDialog = False
                    .pcRecorded = True
                    
                    'Remaining parameters depend on new layer vs replacing existing.  (Note that the replace
                    ' pathway is not actually used at present; drag+dropped layers are always added "as new".)
                    If replaceActiveLayerInstead Then
                        .pcID = g_Language.TranslateMessage("Replace layer from file")
                        .pcUndoType = UNDO_Layer
                    Else
                        .pcID = g_Language.TranslateMessage("New layer from file")
                        .pcUndoType = UNDO_Image_VectorSafe     'Because layer count has changed, we must generate a full-image undo
                    End If
                    
                End With
                
                If replaceActiveLayerInstead Then
                    PDImages.GetActiveImage.UndoManager.CreateUndoData tmpProcCall, PDImages.GetActiveImage.GetActiveLayerID
                Else
                    PDImages.GetActiveImage.UndoManager.CreateUndoData tmpProcCall
                End If
                
            '/END create undo point
            End If
            
            'If requested, synchronize the interface to the new image
            If refreshUI Then
                Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
                SyncInterfaceToCurrentImage
            End If
            
            If replaceActiveLayerInstead Then
                Message "Layer updated successfully."
            Else
                Message "New layer added successfully."
            End If
        
        '/END image data loaded successfully from file failsafe check
        End If
    
    '/END raiseDialog true/false
    End If

End Function

'Make a given layer fully transparent.  This is used by the Edit > Cut menu at present, if the user cuts without first making a selection.
Public Sub EraseLayerByIndex(ByVal layerIndex As Long)

    If PDImages.IsImageActive() Then
    
        'How we "clear" the layer varies by layer type
        Select Case PDImages.GetActiveImage.GetLayerByIndex(layerIndex).GetLayerType
        
            'For image layers, force the layer DIB to all zeroes
            Case PDL_Image
                With PDImages.GetActiveImage.GetLayerByIndex(layerIndex)
                    .GetLayerDIB.CreateBlank .GetLayerWidth(False), .GetLayerHeight(False), 32, 0, 0
                End With
            
            'For text layers, simply erase the current text.  (This has the effect of making the layer fully transparent,
            ' while retaining all text settings... I'm not sure of a better solution at present.)
            Case PDL_TextBasic, PDL_TextAdvanced
                With PDImages.GetActiveImage.GetLayerByIndex(layerIndex)
                    .SetTextLayerProperty ptp_Text, vbNullString
                End With
        
        End Select
        
        'Notify the parent object of the change
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, layerIndex
    
    End If

End Sub

'Reverse the order of layers in this image
Public Sub ReverseLayerOrder()
    If PDImages.IsImageNonNull() Then PDImages.GetActiveImage.ReverseLayerOrder
End Sub

'Select a neighboring layer (up or down)
Public Sub SelectLayerAdjacent(ByVal layerDirectionIsUp As Boolean)

    Dim curLayerIndex As Long
    curLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex
    
    'Determine a new, valid layer index (with wrapping around top/bottom)
    If layerDirectionIsUp Then curLayerIndex = curLayerIndex + 1 Else curLayerIndex = curLayerIndex - 1
    If (curLayerIndex < 0) Then curLayerIndex = PDImages.GetActiveImage.GetNumOfLayers - 1
    If (curLayerIndex >= PDImages.GetActiveImage.GetNumOfLayers) Then curLayerIndex = 0
    
    'Select the new layer
    Layers.SetActiveLayerByIndex curLayerIndex, True, False

End Sub

'Select the top or bottom layer in this image
Public Sub SelectLayerTopBottom(ByVal topIsWanted As Boolean)

    'Determine a new, valid layer index (with wrapping around top/bottom)
    Dim curLayerIndex As Long
    If topIsWanted Then curLayerIndex = PDImages.GetActiveImage.GetNumOfLayers - 1 Else curLayerIndex = 0
    
    'Select the new layer
    Layers.SetActiveLayerByIndex curLayerIndex, True, False

End Sub

'Activate a layer.  Use this instead of directly calling the pdImage.setActiveLayer function if you want to also
' synchronize the UI to match.
Public Sub SetActiveLayerByID(ByVal newLayerID As Long, Optional ByVal alsoRedrawViewport As Boolean = False, Optional ByVal alsoSyncInterface As Boolean = True)

    'If this layer is already active, ignore the request
    If (PDImages.GetActiveImage.GetActiveLayerID <> newLayerID) Then
        
        'Check for any non-destructive property changes to the previously active layer
        'Processor.FlagFinalNDFXState_Generic pgp_Visibility, PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility, PDImages.GetActiveImage.GetActiveLayerID
        
        'Notify the parent PD image of the change
        PDImages.GetActiveImage.SetActiveLayerByID newLayerID
        
        'Notify the Undo/Redo engine of all non-destructive property values for the newly activated layer.
        Processor.SyncAllGenericLayerProperties PDImages.GetActiveImage.GetActiveLayer
        If PDImages.GetActiveImage.GetActiveLayer.IsLayerText Then Processor.SyncAllTextLayerProperties PDImages.GetActiveImage.GetActiveLayer
        
        'Sync the interface to the new layer
        If alsoSyncInterface Then SyncInterfaceToCurrentImage
        
        'Redraw the viewport, but only if requested
        If alsoRedrawViewport Then Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
    End If
        
End Sub

'Same idea as setActiveLayerByID, above
Public Sub SetActiveLayerByIndex(ByVal newLayerIndex As Long, Optional ByVal alsoRedrawViewport As Boolean = False, Optional ByVal alsoSyncInterface As Boolean = True)
    
    'If this layer is already active, ignore the request
    If (PDImages.GetActiveImage.GetActiveLayerID <> PDImages.GetActiveImage.GetLayerByIndex(newLayerIndex).GetLayerID) Then
        
        'Check for any non-destructive property changes to the previously active layer
        'Processor.FlagFinalNDFXState_Generic pgp_Visibility, PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility, PDImages.GetActiveImage.GetActiveLayerID
        
        'Notify the parent PD image of the change
        PDImages.GetActiveImage.SetActiveLayerByIndex newLayerIndex
        
        'Notify the Undo/Redo engine of all non-destructive property values for the newly activated layer.
        Processor.SyncAllGenericLayerProperties PDImages.GetActiveImage.GetActiveLayer
        If PDImages.GetActiveImage.GetActiveLayer.IsLayerText Then Processor.SyncAllTextLayerProperties PDImages.GetActiveImage.GetActiveLayer
        
        'Sync the interface to the new layer
        If alsoSyncInterface Then SyncInterfaceToCurrentImage
            
        'Redraw the viewport, but only if requested
        If alsoRedrawViewport Then Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
    End If
        
End Sub

'Make all layers visible or hidden
Public Sub SetLayerVisibility_AllLayers(Optional ByVal isLayerVisible As Boolean = True)
    
    Dim i As Long
    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
        PDImages.GetActiveImage.GetLayerByIndex(i).SetLayerVisibility isLayerVisible
    Next i
    
    PDImages.GetActiveImage.NotifyImageChanged UNDO_ImageHeader
    
End Sub

'Make only one layer visible; all others will be hidden
Public Sub MakeJustOneLayerHidden(ByVal dLayerIndex As Long)
    
    Dim i As Long
    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
        PDImages.GetActiveImage.GetLayerByIndex(i).SetLayerVisibility (i <> dLayerIndex)
    Next i
    
    PDImages.GetActiveImage.NotifyImageChanged UNDO_ImageHeader
    
End Sub

'Make only one layer visible; all others will be hidden
Public Sub MakeJustOneLayerVisible(ByVal dLayerIndex As Long)
    
    Dim i As Long
    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
        PDImages.GetActiveImage.GetLayerByIndex(i).SetLayerVisibility (i = dLayerIndex)
    Next i
    
    PDImages.GetActiveImage.NotifyImageChanged UNDO_ImageHeader
    
End Sub

'Toggle visibility of a target layer (e.g. visible will be made invisible, invisible made visible)
Public Sub ToggleLayerVisibility(ByVal dLayerIndex As Long)
    PDImages.GetActiveImage.GetLayerByIndex(dLayerIndex).SetLayerVisibility Not PDImages.GetActiveImage.GetLayerByIndex(dLayerIndex).GetLayerVisibility()
    PDImages.GetActiveImage.NotifyImageChanged UNDO_LayerHeader, dLayerIndex
End Sub

'XML-based wrapper for DuplicateLayerByIndex(), below
Public Sub DuplicateLayerByIndex_XML(ByRef processParameters As String)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    Layers.DuplicateLayerByIndex cParams.GetLong("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex)
End Sub

'Duplicate a given layer (note: it doesn't have to be the active layer)
Public Sub DuplicateLayerByIndex(ByVal dLayerIndex As Long)

    'Validate the requested layer index
    If (dLayerIndex < 0) Then dLayerIndex = 0
    If (dLayerIndex > PDImages.GetActiveImage.GetNumOfLayers - 1) Then dLayerIndex = PDImages.GetActiveImage.GetNumOfLayers - 1
    
    'Before doing anything else, make a copy of the current active layer ID.  We will use this to restore the same
    ' active layer after the creation is complete.
    Dim activeLayerID As Long
    activeLayerID = PDImages.GetActiveImage.GetActiveLayerID
    
    'Also copy the ID of the layer we are creating.
    Dim dupedLayerID As Long
    dupedLayerID = PDImages.GetActiveImage.GetLayerByIndex(dLayerIndex).GetLayerID
    
    'Ask the parent pdImage to create a new layer object
    Dim newLayerID As Long
    newLayerID = PDImages.GetActiveImage.CreateBlankLayer(dLayerIndex)
    
    'Ask the new layer to copy the contents of the layer we are duplicating
    PDImages.GetActiveImage.GetLayerByID(newLayerID).CopyExistingLayer PDImages.GetActiveImage.GetLayerByID(dupedLayerID)
    
    'Make the duplicate layer the active layer
    PDImages.GetActiveImage.SetActiveLayerByID newLayerID
    
    'Notify the parent image that the entire image now needs to be recomposited
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.NotifyLayerChange
    
    'Render the new image to screen
    Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
            
    'Synchronize the interface to the new image
    Interface.SyncInterfaceToCurrentImage
    
End Sub

'Ensure a given layer index is inbounds its parent image.  We use this when pasting layers to a new image.
' Returns TRUE if one or more layer positions are modified by the in-bounding.
Public Function EnsureLayerInbounds(ByVal srcLayerID As Long) As Boolean
    
    Dim tmpLayer As pdLayer
    Set tmpLayer = PDImages.GetActiveImage.GetLayerByID(srcLayerID)
    
    Dim layerBoundsF As RectF
    tmpLayer.GetLayerBoundaryRect layerBoundsF
    
    If (layerBoundsF.Left >= PDImages.GetActiveImage.Width) Then
        tmpLayer.SetLayerOffsetX 0
        EnsureLayerInbounds = True
    End If
    
    If (layerBoundsF.Top >= PDImages.GetActiveImage.Height) Then
        tmpLayer.SetLayerOffsetY 0
        EnsureLayerInbounds = True
    End If
    
    If (layerBoundsF.Left + layerBoundsF.Width <= 0) Then
        tmpLayer.SetLayerOffsetX 0
        EnsureLayerInbounds = True
    End If
    
    If (layerBoundsF.Top + layerBoundsF.Height <= 0) Then
        tmpLayer.SetLayerOffsetY 0
        EnsureLayerInbounds = True
    End If
    
End Function

'Merge the layer at layerIndex up or down.
Public Sub MergeLayerAdjacent(ByVal dLayerIndex As Long, ByVal mergeDown As Boolean)
    
    'Look for a valid target layer to merge with in the requested direction.
    Dim mergeTarget As Long
    mergeTarget = IsLayerAllowedToMergeAdjacent(dLayerIndex, mergeDown)
    
    'If we've been given a valid merge target, apply it now!
    If (mergeTarget >= 0) Then
    
        If mergeDown Then
        
            With PDImages.GetActiveImage()
                
                'Request a merge from the parent pdImage
                .MergeTwoLayers .GetLayerByIndex(dLayerIndex), .GetLayerByIndex(mergeTarget)
                
                'Delete the now-merged layer
                .DeleteLayerByIndex dLayerIndex
                
                'Notify the parent of the change
                .NotifyImageChanged UNDO_Layer, mergeTarget
                
                'Set the newly merged layer as the active layer
                .SetActiveLayerByIndex mergeTarget
            
            End With
            
        Else
        
            With PDImages.GetActiveImage()
            
                'Request a merge from the parent pdImage
                .MergeTwoLayers .GetLayerByIndex(mergeTarget), .GetLayerByIndex(dLayerIndex)
                
                'Delete the now-merged layer
                .DeleteLayerByIndex mergeTarget
                
                'Notify the parent of the change
                .NotifyImageChanged UNDO_Layer, dLayerIndex
                
                'Set the newly merged layer as the active layer
                .SetActiveLayerByIndex dLayerIndex
                
            End With
        
        End If
                
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange
        
        'Redraw the viewport
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
    End If

End Sub

'Is this layer allowed to merge up or down?  Note that invisible layers are not generally considered suitable
' for merging, so a layer will typically be merged with the next VISIBLE layer.  If none are available, merging
' is disallowed.
'
'Note that the return value for this function is a little wonky.  This function will return the TARGET MERGE LAYER
' INDEX if the function is successful.  This value will always be >= 0.  If no valid layer can be found, -1 will be
' returned (which obviously isn't a valid index, but IS true, so it's a little confusing - handle accordingly!)
'
'It should be obvious, but the parameter srcLayerIndex is the index of the layer the caller wants to merge.
Public Function IsLayerAllowedToMergeAdjacent(ByVal srcLayerIndex As Long, ByVal moveDown As Boolean) As Long
    
    Dim i As Long
    
    'First, make sure the layer in question exists
    If (Not PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex) Is Nothing) Then
    
        'Check MERGE DOWN
        If moveDown Then
        
            'As an easy check, make sure this layer is visible, and not already at the bottom.
            If (srcLayerIndex <= 0) Or (Not PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).GetLayerVisibility) Then
                IsLayerAllowedToMergeAdjacent = -1
                Exit Function
            End If
            
            'Search for the nearest valid layer beneath this one.
            For i = srcLayerIndex - 1 To 0 Step -1
                If PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility Then
                    IsLayerAllowedToMergeAdjacent = i
                    Exit Function
                End If
            Next i
            
            'If we made it all the way here, no valid merge target was found.  Return failure (-1).
            IsLayerAllowedToMergeAdjacent = -1
        
        'Check MERGE UP
        Else
        
            'As an easy check, make sure this layer isn't already at the top.
            If (srcLayerIndex >= PDImages.GetActiveImage.GetNumOfLayers - 1) Or (Not PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).GetLayerVisibility) Then
                IsLayerAllowedToMergeAdjacent = -1
                Exit Function
            End If
            
            'Search for the nearest valid layer above this one.
            For i = srcLayerIndex + 1 To PDImages.GetActiveImage.GetNumOfLayers - 1
                If PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility Then
                    IsLayerAllowedToMergeAdjacent = i
                    Exit Function
                End If
            Next i
            
            'If we made it all the way here, no valid merge target was found.  Return failure (-1).
            IsLayerAllowedToMergeAdjacent = -1
        
        End If
        
    End If

End Function

'Take various layers in an image, and split them out into their own standalone images
Public Function SplitLayerToImage(Optional ByRef processParameters As String) As Boolean
    
    SplitLayerToImage = False
    
    'Retrieve any conversion parameters
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    
    Dim targetIndex As Long
    targetIndex = cParams.GetLong("target-layer", -1)
    
    'Process parameters specify *which* layer(s) should be converted to a standalone image.
    ' Layers are identified by index.  We want to handle the conversion in two steps:
    ' 1) Convert all layer(s) to standalone images
    ' 2) Remove all converted layer(s) from the current image (typically, all but the active layer)
    Dim i As Long
    
    'Make a safe local reference to the currently active image - because the active image will
    ' change as we load other images.
    Dim srcImage As pdImage
    Set srcImage = PDImages.GetActiveImage
    
    'To simplify this process, construct an array that identifies all layers by their ID
    ' (which is immutable, and will not change - unlike layer indices)
    With srcImage
        
        Dim listOfLayers() As LayerConvertCache
        ReDim listOfLayers(0 To .GetNumOfLayers - 1) As LayerConvertCache
        
        For i = 0 To .GetNumOfLayers - 1
            
            listOfLayers(i).id = .GetLayerByIndex(i).GetLayerID
            
            If (targetIndex = -1) Then
                listOfLayers(i).mustConvert = True
            Else
                listOfLayers(i).mustConvert = (i = targetIndex)
            End If
            
        Next i
        
    End With
    
    'We now have a list which layers require converting.  Iterate through each layer,
    ' convert it to a null-padded layer (which greatly simplifies re-assembly later),
    ' split it into a separate image, then remove it from the image.
    For i = 0 To UBound(listOfLayers)
    
        If listOfLayers(i).mustConvert Then
            
            Message "Copying layer ""%1"" to standalone image...", srcImage.GetLayerByID(listOfLayers(i).id).GetLayerName()
            
            'Load said layer as a separate image
            Dim tmpLayerFile As String
            tmpLayerFile = UserPrefs.GetTempPath & "LayerConvert.pdi"
            
            Dim tmpImage As pdImage
            Set tmpImage = New pdImage
            
            'In the temporary pdImage object, create a blank layer; this will receive the processed DIB
            Dim newLayerID As Long
            newLayerID = tmpImage.CreateBlankLayer
            tmpImage.GetLayerByID(newLayerID).CopyExistingLayer srcImage.GetLayerByID(listOfLayers(i).id)
            
            'Force the layer to visible
            tmpImage.GetLayerByID(newLayerID).SetLayerVisibility True
            
            'Convert the layer to a null-padded layer (a layer at the same size as the current image)
            tmpImage.GetLayerByID(newLayerID).ConvertToNullPaddedLayer srcImage.Width, srcImage.Height
            tmpImage.UpdateSize
            
            'Write the image out to file, then free its associated memory
            Saving.SavePDI_Image tmpImage, tmpLayerFile, True, cf_Lz4, cf_Lz4
            Set tmpImage = Nothing
            
            'Construct a title (name) for the new image, and insert the original layer index.
            ' (This is helpful if the user decides to reconstruct the layers into an image later.)
            Dim sTitle As String
            sTitle = srcImage.GetLayerByID(listOfLayers(i).id).GetLayerName()
            If (LenB(sTitle) = 0) Then sTitle = g_Language.TranslateMessage("[untitled image]")
            
            'We can now use the standard image load routine to import the temporary file.
            ' (Note that we explicitly suspend warnings as we load, because we may be loading a *lot* of images.)
            Dim importDialogResults As VbMsgBoxResult
            importDialogResults = vbNo
            Loading.LoadFileAsNewImage tmpLayerFile, sTitle, False, importDialogResults, False
            
            'Be polite and remove the temporary file
            Files.FileDeleteIfExists tmpLayerFile
            
        End If
    
    Next i
    
    'PDImages.SetActiveImageID srcImage.imageID
    CanvasManager.ActivatePDImage srcImage.imageID
    
    SplitLayerToImage = True
    Message "Conversion complete."
            
End Function

'XML-based wrapper to DeleteLayer(), below
Public Sub DeleteLayer_XML(ByRef processParameters As String)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    Layers.DeleteLayer cParams.GetLong("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex)
End Sub

'Delete a given layer
Public Sub DeleteLayer(ByVal dLayerIndex As Long, Optional ByVal updateUI As Boolean = True)

    'Cache the current layer index
    Dim curLayerIndex As Long
    curLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex - 1

    PDImages.GetActiveImage.DeleteLayerByIndex dLayerIndex
    
    If updateUI Then
        
        'Set a new active layer
        If (curLayerIndex > PDImages.GetActiveImage.GetNumOfLayers - 1) Then curLayerIndex = PDImages.GetActiveImage.GetNumOfLayers - 1
        If (curLayerIndex < 0) Then curLayerIndex = 0
        SetActiveLayerByIndex curLayerIndex, False
        
        'Notify the parent image that the entire image now needs to be recomposited
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
        
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange
        
        'Redraw the viewport
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
    End If

End Sub

'Delete all hidden layers
Public Sub DeleteHiddenLayers()

    'Perform a couple fail-safe checks.  These should not be a problem, as calling functions should have safeguards
    ' against bad requests, but better safe than sorry.
    
    'If there are no hidden layers, exit
    If PDImages.GetActiveImage.GetNumOfHiddenLayers = 0 Then Exit Sub
    
    'If all layers are hidden, exit
    If PDImages.GetActiveImage.GetNumOfHiddenLayers = PDImages.GetActiveImage.GetNumOfLayers Then Exit Sub
    
    'We can now assume that the image in question has at least one visible layer, and at least one hidden layer.
    
    'Cache the currently active layerID - IF the current layer is visible.  If it isn't, it's going to be deleted,
    ' so we must pick a new arbitrary layer (why not the bottom layer?).
    Dim activeLayerID As Long
    
    If PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility Then
        activeLayerID = PDImages.GetActiveImage.GetActiveLayerID
    Else
        activeLayerID = -1
    End If
    
    'Starting at the top and moving down, delete all hidden layers.
    Dim i As Long
    For i = PDImages.GetActiveImage.GetNumOfLayers - 1 To 0 Step -1
        If Not PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility Then
            PDImages.GetActiveImage.DeleteLayerByIndex i
        End If
    Next i
    
    'Set a new active layer
    If (activeLayerID = -1) Then
        SetActiveLayerByIndex 0, False
    Else
        SetActiveLayerByID activeLayerID
    End If
    
    'Notify the parent image that the entire image now needs to be recomposited
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.NotifyLayerChange
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

'Given all open images (besides the current one), assemble them as layers into the current image.
Public Sub MergeImagesToLayers(Optional ByVal processParameters As String = vbNullString)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    
    Dim i As Long, j As Long
    Dim openImageIDs As pdStack
    
    'This function is sort of an odd case for PD, as operations are typically strictly limited
    ' to only affecting the active image.
    
    'First, we want to make a list of open images, and try to figure out where (if any) each
    ' image should be inserted.  This is most relevant when the images were originally split using
    ' the "Split layers into images" command - by parsing layer indices out of their titles, we can
    ' split each image back into its original parent image *at its original location*.
    If (Not PDImages.GetListOfActiveImageIDs(openImageIDs)) Then Exit Sub
    
    Dim listOfImages() As LayerConvertCache
    ReDim listOfImages(0 To openImageIDs.GetNumOfInts - 1) As LayerConvertCache
    
    Dim maxWidth As Long, maxHeight As Long
    
    For i = 0 To UBound(listOfImages)
        
        'We want to convert all images *except* the currently active one
        listOfImages(i).mustConvert = (PDImages.GetImageByID(openImageIDs.GetInt(i)).imageID <> PDImages.GetActiveImageID)
        
        'For each image-to-be-converted...
        If listOfImages(i).mustConvert Then
            
            'Make a note of the image's ID and size
            listOfImages(i).id = openImageIDs.GetInt(i)
            listOfImages(i).srcImageWidth = PDImages.GetImageByID(listOfImages(i).id).Width
            listOfImages(i).srcImageHeight = PDImages.GetImageByID(listOfImages(i).id).Height
            
            'Track max width/height; we may need these to resize the image
            If (listOfImages(i).srcImageWidth > maxWidth) Then maxWidth = listOfImages(i).srcImageWidth
            If (listOfImages(i).srcImageHeight > maxHeight) Then maxHeight = listOfImages(i).srcImageHeight
            
            'Next, pull the name of the base layer.  This is the layer name we want to match
            ' against the layer names in our existing image, to try and identify matches.
            listOfImages(i).srcLayerName = PDImages.GetImageByID(listOfImages(i).id).GetLayerByIndex(0).GetLayerName()
            
        End If
        
    Next i

    'Make a note of the currently active layer index
    Dim activeLayerIndex As Long
    activeLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex()
    
    'Make a safe local reference to the currently active image - because the active image may
    ' change as we access other images.
    Dim srcImage As pdImage
    Set srcImage = PDImages.GetActiveImage()
    
    'If the user wants us to resize the image to fit imported layers, do so now
    If cParams.GetBool("resize-canvas-fit", False) Then
        If (srcImage.Width > maxWidth) Then maxWidth = srcImage.Width
        If (srcImage.Height > maxHeight) Then maxHeight = srcImage.Height
        srcImage.UpdateSize False, maxWidth, maxHeight
    End If
    
    'We may need to set a custom anchor position; pull the relevant parameter value now
    Dim anchorPosition As Long
    anchorPosition = cParams.GetLong("layer-anchor", 0)
    
    'The number of layers in the base image may change as other images are migrated in - to avoid
    ' our tracking array having fewer indices than the base image (whose layer count is actively
    ' increasing), we need to cache the search limit in advance.
    Dim ubLayers As Long
    ubLayers = srcImage.GetNumOfLayers - 1
    
    'Next, we want to make a bool array to track which layer names we have matched so far
    ' (in the active image).  On the off chance that there are 2+ layers with identical names,
    ' we want to match the layers in-order (instead of overwriting the same one twice).
    Dim layerMatched() As Boolean
    ReDim layerMatched(0 To ubLayers) As Boolean
    
    Dim overwriteMatchingLayers As Boolean
    overwriteMatchingLayers = cParams.GetBool("overwrite-layers", False)
    
    'Next, we basically want to iterate through all images in the collection, and add each one
    ' to this image - as a unique layer - in turn.
    For i = 0 To UBound(listOfImages)
        
        If listOfImages(i).mustConvert Then
            
            Message "Adding image ""%1"" as layer...", listOfImages(i).srcLayerName
            
            'Ask the target file to write itself out to a temp PDI file
            Dim tmpLayerFile As String
            tmpLayerFile = UserPrefs.GetTempPath & "LayerConvert.pdi"
            If Saving.SavePDI_Image(PDImages.GetImageByID(listOfImages(i).id), tmpLayerFile, True, cf_Lz4, cf_Lz4) Then
                
                'We now want to load the resulting image as a standalone layer.  We use a convenient
                ' wrapper function that ensures the image is loaded as a single layer, even if it
                ' contains multiple layers.  (This is by design, to allow the user to do things like
                ' overlay text on a single layer, then merge that layer back into a parent image.)
                Dim tmpDIB As pdDIB
                Set tmpDIB = New pdDIB
                If Loading.QuickLoadImageToDIB(tmpLayerFile, tmpDIB, False, False) Then
                    
                    Dim targetIndex As Long
                    targetIndex = -1
                    
                    'Next, try to find a layer with this name in the current image.
                    If overwriteMatchingLayers Then
                        
                        For j = 0 To ubLayers
                            If Strings.StringsEqual(srcImage.GetLayerByIndex(j).GetLayerName, listOfImages(i).srcLayerName, False) Then
                                
                                'Make sure we haven't matched this layer already
                                If (Not layerMatched(j)) Then
                                    layerMatched(j) = True
                                    targetIndex = j
                                    Exit For
                                End If
                                    
                            End If
                        Next j
                        
                    End If
                    
                    'Add the new layer to this image in one of two ways:
                    ' 1) If a matching layer name was found in the current image, overwrite that layer
                    '    with the one we've imported from file.
                    Dim mustCreateNewLayer As Boolean
                    mustCreateNewLayer = True
                    If (targetIndex >= 0) Then
                        mustCreateNewLayer = Not layerMatched(targetIndex)
                    End If
                    
                    'Because layers are automatically null-padded when they're split into separate images,
                    ' we *always* create a new layer during import, then crop the null padding to determine
                    ' the new layer's position.  This provides maximum flexibility for the user.
                    Dim newLayerID As Long
                    newLayerID = srcImage.CreateBlankLayer(targetIndex)
                    srcImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, listOfImages(i).srcLayerName, tmpDIB, True
                    Set tmpDIB = Nothing
                    
                    If (Not mustCreateNewLayer) Then Layers.DeleteLayer targetIndex, False
                    
                    'Based on the anchor position, determine x and y locations for the new layer
                    Dim dstX As Long, dstY As Long
                    Dim imgWidth As Long, imgHeight As Long, layWidth As Long, layHeight As Long
                    imgWidth = srcImage.Width
                    imgHeight = srcImage.Height
                    layWidth = srcImage.GetLayerByID(newLayerID).GetLayerWidth
                    layHeight = srcImage.GetLayerByID(newLayerID).GetLayerHeight
                    
                    Select Case anchorPosition
                    
                        'Top-left
                        Case 0
                            dstX = 0
                            dstY = 0
                        
                        'Top-center
                        Case 1
                            dstX = (imgWidth - layWidth) \ 2
                            dstY = 0
                        
                        'Top-right
                        Case 2
                            dstX = (imgWidth - layWidth)
                            dstY = 0
                        
                        'Middle-left
                        Case 3
                            dstX = 0
                            dstY = (imgHeight - layHeight) \ 2
                        
                        'Middle-center
                        Case 4
                            dstX = (imgWidth - layWidth) \ 2
                            dstY = (imgHeight - layHeight) \ 2
                        
                        'Middle-right
                        Case 5
                            dstX = (imgWidth - layWidth)
                            dstY = (imgHeight - layHeight) \ 2
                        
                        'Bottom-left
                        Case 6
                            dstX = 0
                            dstY = (imgHeight - layHeight)
                        
                        'Bottom-center
                        Case 7
                            dstX = (imgWidth - layWidth) \ 2
                            dstY = (imgHeight - layHeight)
                        
                        'Bottom right
                        Case 8
                            dstX = (imgWidth - layWidth)
                            dstY = (imgHeight - layHeight)
                    
                    End Select
                    
                    srcImage.GetLayerByID(newLayerID).SetLayerOffsetX dstX
                    srcImage.GetLayerByID(newLayerID).SetLayerOffsetY dstY
                    
                    'Finally, auto-crop the layer, as it will have been null-padded by a previous step
                    srcImage.GetLayerByID(newLayerID).CropNullPaddedLayer
                    
                End If
                
                'Delete the temp file
                Files.FileDeleteIfExists tmpLayerFile
            
            End If
            
        End If
        
    Next i
    
    'Restore the currently active layer index
    PDImages.GetActiveImage.SetActiveLayerByIndex activeLayerIndex
    
    'Make sure the original image is notified of the new layer arrangement (which prompts it
    ' to update things like its internal thumbnail cache)
    srcImage.NotifyImageChanged UNDO_Image
    
    'If the user requested it, unload each merged image in turn
    Dim unloadSourceImages As Long
    unloadSourceImages = cParams.GetLong("close-source-images", 0)
    If (unloadSourceImages <> 0) Then
        
        For i = 0 To UBound(listOfImages)
        
            If listOfImages(i).mustConvert Then
            
                'Prompt before closing
                If (unloadSourceImages = 1) Then
                    CanvasManager.FullPDImageUnload listOfImages(i).id
                
                'Close without prompting
                ElseIf (unloadSourceImages = 2) Then
                    CanvasManager.UnloadPDImage listOfImages(i).id
                    
                End If
            
            End If
            
        Next i
        
    End If
    
    'Restore the originally active image as the image with focus.  (By default, newly loaded images
    ' "steal" focus - this is a rare case where we don't want that.)
    CanvasManager.ActivatePDImage srcImage.imageID, "Split images into layers", True, , True
    
    Message "Conversion complete."
     
End Sub

'Move a layer up or down in the stack (referred to as "raise" and "lower" in the menus)
Public Sub MoveLayerAdjacent(ByVal dLayerIndex As Long, ByVal directionIsUp As Boolean, Optional ByVal updateInterface As Boolean = True)

    'Make a copy of the currently active layer's ID
    Dim curActiveLayerID As Long
    curActiveLayerID = PDImages.GetActiveImage.GetActiveLayerID
    
    'Ask the parent pdImage to move the layer for us
    PDImages.GetActiveImage.MoveLayerByIndex dLayerIndex, directionIsUp
    
    'Restore the active layer
    SetActiveLayerByID curActiveLayerID, False
    
    'Notify the parent image that the entire image now needs to be recomposited
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
    
    If updateInterface Then
        
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange
        
        'Redraw the viewport
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
    End If

End Sub

'Move a layer to the top or bottom of the stack (referred to as "raise to top" and "lower to bottom" in the menus)
Public Sub MoveLayerToEndOfStack(ByVal dLayerIndex As Long, ByVal moveToTopOfStack As Boolean, Optional ByVal updateInterface As Boolean = True)

    'Make a copy of the currently active layer's ID
    Dim curActiveLayerID As Long
    curActiveLayerID = PDImages.GetActiveImage.GetActiveLayerID
    
    Dim i As Long
    
    'Until this layer is at the desired end of the stack, ask the parent to keep moving it for us!
    If moveToTopOfStack Then
    
        For i = dLayerIndex To PDImages.GetActiveImage.GetNumOfLayers - 1
            
            'Ask the parent pdImage to move the layer up for us
            PDImages.GetActiveImage.MoveLayerByIndex i, True
            
        Next i
    
    Else
    
        For i = dLayerIndex To 0 Step -1
            
            'Ask the parent pdImage to move the layer up for us
            PDImages.GetActiveImage.MoveLayerByIndex i, False
            
        Next i
    
    End If
    
    'Restore the active layer.  (This will also re-synchronize the interface against the new image.)
    SetActiveLayerByID curActiveLayerID, False
    
    'Notify the parent image that the entire image now needs to be recomposited
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
    
    If updateInterface Then
    
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange
        
        'Redraw the viewport
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
    End If

End Sub

'Given a multi-layered image, flatten it.  Note that flattening does *not* remove alpha!  It simply merges all layers,
' including discarding invisible ones.
Public Sub FlattenImage(Optional ByVal functionParams As String = vbNullString)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString functionParams
    
    Dim removeTransparency As Boolean, newBackgroundColor As Long
    removeTransparency = cParams.GetBool("removetransparency", False)
    newBackgroundColor = cParams.GetLong("backgroundcolor", vbWhite)
    
    'Start by retrieving a copy of the composite image
    Dim compositeDIB As pdDIB
    Set compositeDIB = New pdDIB
    
    PDImages.GetActiveImage.GetCompositedImage compositeDIB
    
    'If the caller wants the flattened image to *not* have transparency, remove said transparency now
    If removeTransparency Then compositeDIB.CompositeBackgroundColor Colors.ExtractRed(newBackgroundColor), Colors.ExtractGreen(newBackgroundColor), Colors.ExtractBlue(newBackgroundColor)
    
    'Also, grab the name of the bottom-most layer.  This will be used as the name of our only layer in the flattened image.
    Dim flattenedName As String
    flattenedName = PDImages.GetActiveImage.GetLayerByIndex(0).GetLayerName
    
    'With this information, we can now delete all image layers.
    Do
        PDImages.GetActiveImage.DeleteLayerByIndex 0
    Loop While (PDImages.GetActiveImage.GetNumOfLayers > 1)
    
    'Note that the delete operation does not allow us to delete all layers.  (If there is only one layer present,
    ' it will exit without modifying the image.)  Because of that, the image will still retain one layer, which
    ' we will have to manually overwrite.
        
    'Reset any optional layer parameters to their default state
    PDImages.GetActiveImage.GetLayerByIndex(0).ResetLayerParameters
    
    'Overwrite the final layer with the composite DIB.
    PDImages.GetActiveImage.GetLayerByIndex(0).InitializeNewLayer PDL_Image, flattenedName, compositeDIB
    
    'Mark the only layer present as the active one.  (This will also re-synchronize the interface against the new image.)
    SetActiveLayerByIndex 0, False
    
    'Notify the parent of the change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, 0
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.NotifyLayerChange
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

'Is a given coordinate - in IMAGE coordinate space - over a specified layer?  All possible affine transforms
' are handled automatically by this function.
Public Function IsCoordinateOverLayer(ByVal idxLayer As Long, ByVal xCoordImg As Single, ByVal yCoordImg As Single) As Boolean
    
    If (idxLayer < 0) Or (idxLayer >= PDImages.GetActiveImage.GetNumOfLayers) Then Exit Function
    
    Dim layerCorners(0 To 3) As PointFloat
    PDImages.GetActiveImage.GetLayerByIndex(idxLayer, False).GetLayerCornerCoordinates layerCorners, True
    
    Dim tmpPath As pd2DPath
    Set tmpPath = New pd2DPath
    tmpPath.AddPolygon 4, VarPtr(layerCorners(0)), True, False
    
    IsCoordinateOverLayer = tmpPath.IsPointInsidePathF(xCoordImg, yCoordImg)
    
End Function

'Flattening an image with transparency should raise a dialog, so the user can decide whether to
' 1) Keep transparency in the final flatten, or...
' 2) Replace transparency with a new background color.
'
'If an image does NOT have exposed transparency, such a dialog is irrelevant.  Call this function to
' decide whether to suppress or raise the flatten dialog prior to flattening.  (TRUE means raise the
' Flatten dialog, FALSE means suppress it.)
Public Function IsFlattenDialogRelevant() As Boolean
    
    IsFlattenDialogRelevant = True
    
    'Grab a small copy of the composite image, and if it doesn't contain meaningful transparency,
    ' recommend suppressing the flatten dialog.
    If PDImages.IsImageActive() Then
        
        Dim newWidth As Long, newHeight As Long
        PDMath.ConvertAspectRatio PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, 256, 256, newWidth, newHeight
        
        Dim tmpRectF As RectF
        With tmpRectF
            .Left = 0
            .Top = 0
            .Width = newWidth
            .Height = newHeight
        End With
        
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateBlank newWidth, newHeight, 32, 0, 0
        PDImages.GetActiveImage.RequestThumbnail tmpDIB, 256, False, VarPtr(tmpRectF)
        
        IsFlattenDialogRelevant = DIBs.IsDIBTransparent(tmpDIB)
        
    End If
    
End Function

'Given a multi-layered image, merge all visible layers, while ignoring any hidden ones.  Note that flattening does *not*
' remove alpha!  It simply merges all visible layers.
Public Sub MergeVisibleLayers()
    
    'If there's only one layer, this function should not be called - but just in case, exit in advance.
    If (PDImages.GetActiveImage.GetNumOfLayers = 1) Then Exit Sub
    
    'Similarly, if there's only one *visible* layer, this function should not be called - but just in case, exit in advance.
    If (PDImages.GetActiveImage.GetNumOfVisibleLayers = 1) Then Exit Sub
    
    'By this point, we can assume there are at least two visible layers in the image.  Rather than deal with the messiness
    ' of finding the lowest base layer and gradually merging everything into it, we're going to just create a new blank
    ' layer at the base of the image, then merge everything with it until finally all visible layers have been merged.
    
    'Insert a new layer at the bottom of the layer stack.
    PDImages.GetActiveImage.CreateBlankLayer 0
    
    'Technically, the command above does not actually insert a new layer at the base of the image.  Per convention,
    ' it always inserts the requested layer at the spot one *above* the requested spot.  To work around this, swap
    ' our newly created layer with the layer at position 0.
    PDImages.GetActiveImage.SwapTwoLayers 0, 1
    
    'Fill that new layer with a blank DIB at the dimensions of the image.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateBlank PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, 32, 0, 0
    tmpDIB.SetInitialAlphaPremultiplicationState True
    PDImages.GetActiveImage.GetLayerByIndex(0).InitializeNewLayer PDL_Image, g_Language.TranslateMessage("Merged layers"), tmpDIB
    
    'With that done, merging visible layers is actually not that hard.  Loop through the layer collection,
    ' merging visible layers with the base layer, until all visible layers have been merged.
    Dim i As Long
    For i = 1 To PDImages.GetActiveImage.GetNumOfLayers - 1
    
        'If this layer is visible, merge it with the base layer
        If PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility Then
            PDImages.GetActiveImage.MergeTwoLayers PDImages.GetActiveImage.GetLayerByIndex(i), PDImages.GetActiveImage.GetLayerByIndex(0)
        End If
    
    Next i
    
    'Now that our base layer contains the result of merging all visible layers, we can now delete all
    ' other visible layers.
    For i = PDImages.GetActiveImage.GetNumOfLayers - 1 To 1 Step -1
        If PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility Then
            PDImages.GetActiveImage.DeleteLayerByIndex i
        End If
    Next i
    
    'Mark the new merged layer as the active one.  (This will also re-synchronize the interface against the new image.)
    Layers.SetActiveLayerByIndex 0, False
    
    'Notify the parent image that the entire image now needs to be recomposited
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, 0
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.NotifyLayerChange
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

Public Function ReplaceLayerWithClipboard() As Boolean
        
    'See if there's anything useable on the clipboard
    Dim tmpDIB As pdDIB: Set tmpDIB = New pdDIB
    ReplaceLayerWithClipboard = g_Clipboard.ClipboardPaste(True, tmpDIB)
    
    If ReplaceLayerWithClipboard And (Not tmpDIB Is Nothing) Then
    
        'The paste operation succeeded.  Overwrite the active layer's contents with whatever we retrieved
        ' from the clipboard, and notify the parent image of the change.
        PDImages.GetActiveImage.GetActiveLayer.SetLayerDIB tmpDIB
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
        
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange PDImages.GetActiveImage.GetActiveLayerID
        
        'Render the new image to screen (not technically necessary, but doesn't hurt)
        Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
        'Synchronize the interface to the new image
        Interface.SyncInterfaceToCurrentImage
        
    '/if the clipboard doesn't contain useable data, do nothing
    End If
    
End Function

'If a layer has been transformed using the on-canvas tools, this will reset it to its default size.
Public Sub ResetLayerSize(ByVal srcLayerIndex As Long)

    PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).SetLayerCanvasXModifier 1#
    PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).SetLayerCanvasYModifier 1#
    
    'Notify the parent image of the change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_LayerHeader, srcLayerIndex
    
    'Re-sync the interface
    Interface.SyncInterfaceToCurrentImage
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

'Add the necessary non-destructive transform parameters to force the current layer to its
' parent image's size.
Public Sub FitLayerToImageSize(ByVal srcLayerIndex As Long)
    
    If (srcLayerIndex < 0) Then srcLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex
    
    'If the move/size tool is currently active, forcibly disable the "lock aspect ratio"
    ' setting, as it may prevent us from sizing the layer to match the current image
    If (g_CurrentTool = NAV_MOVE) Then toolpanel_MoveSize.chkAspectRatio.Value = False
    
    'Reset to position (0, 0)
    PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).SetLayerOffsetX 0#
    PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).SetLayerOffsetY 0#
    
    'Set x/y size modifiers that result in a full-image stretch
    Dim lyrWidth As Double, lyrHeight As Double, imgWidth As Double, imgHeight As Double
    imgWidth = PDImages.GetActiveImage.Width
    imgHeight = PDImages.GetActiveImage.Height
    lyrWidth = PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).GetLayerWidth(False)
    lyrHeight = PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).GetLayerHeight(False)
    
    'Failsafe check only; PD doesn't allow 0-width/height layers at present
    If (lyrWidth > 0#) And (lyrHeight > 0#) Then
        PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).SetLayerCanvasXModifier imgWidth / lyrWidth
        PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).SetLayerCanvasYModifier imgHeight / lyrHeight
    End If
    
    'Notify the parent image of the change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_LayerHeader, srcLayerIndex
    
    'Re-sync the interface
    Interface.SyncInterfaceToCurrentImage
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

'If a layer has been transformed using the on-canvas tools, this will make those transforms permanent.
Public Sub MakeLayerAffineTransformsPermanent(ByVal srcLayerIndex As Long)
    
    'Layers are capable of making this change internally
    PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).MakeCanvasTransformsPermanent
    
    'Notify the parent object of this change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, srcLayerIndex
    
    'Re-sync the interface
    Interface.SyncInterfaceToCurrentImage
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

'Resize a layer non-destructively, e.g. by only changing its position and on-canvas x/y modifiers
Public Sub ResizeLayerNonDestructive(ByVal srcLayerIndex As Long, ByRef resizeParams As String)

    'Create a parameter parser to help us interpret the passed param string
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString resizeParams
    
    'Apply the passed parameters to the specified layer
    With PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex)
        .SetLayerOffsetX cParams.GetDouble("layer-offsetx")
        .SetLayerOffsetY cParams.GetDouble("layer-offsety")
        
        'Raster and vector layers use different size descriptors.  (Vector layers use an absolute size; raster layers use the
        ' underlying DIB size, plus a fractional modifier.)
        If (.GetLayerType = PDL_Image) Then
            .SetLayerCanvasXModifier cParams.GetDouble("layer-modifierx")
            .SetLayerCanvasYModifier cParams.GetDouble("layer-modifiery")
        Else
            .SetLayerWidth cParams.GetLong("layer-sizex")
            .SetLayerHeight cParams.GetLong("layer-sizey")
        End If
        
    End With
    
    'Notify the parent image of the change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_LayerHeader, srcLayerIndex
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

'Rotate a layer non-destructively, e.g. by only changing its header angle value
Public Sub RotateLayerNonDestructive(ByVal srcLayerIndex As Long, ByRef resizeParams As String)

    'Create a parameter parser to help us interpret the passed param string
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString resizeParams
    
    'Apply the passed parameter to the specified layer
    With PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex)
        .SetLayerAngle cParams.GetDouble("layer-angle", 0#)
    End With
    
    'Notify the parent image of the change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_LayerHeader, srcLayerIndex
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

'Move a layer to a new x/y position on the canvas
Public Sub MoveLayerOnCanvas(ByVal srcLayerIndex As Long, ByRef resizeParams As String)

    'Create a parameter parser to help us interpret the passed param string
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString resizeParams
    
    'Apply the passed parameters to the specified layer
    With PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex)
        .SetLayerOffsetX cParams.GetDouble("layer-offsetx")
        .SetLayerOffsetY cParams.GetDouble("layer-offsety")
    End With
    
    'Notify the parent of the change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_LayerHeader, srcLayerIndex
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

'Given a layer, populate a rect with its coordinates (relative to the main image coordinates, always).
' As of PD 7.0, an additional "includeAffineTransforms" parameter is available.  This will return the bounds of the layer, after any/all
' affine transforms (rotate, etc) have been processed.
Public Sub FillRectForLayerF(ByRef srcLayer As pdLayer, ByRef dstRect As RectF, Optional ByVal useCanvasModifiers As Boolean = False, Optional ByVal includeAffineTransforms As Boolean = True)

    With srcLayer
        
        If includeAffineTransforms Then
            .GetLayerBoundaryRect dstRect
        Else
            dstRect.Left = .GetLayerOffsetX
            dstRect.Width = .GetLayerWidth(useCanvasModifiers)
            dstRect.Top = .GetLayerOffsetY
            dstRect.Height = .GetLayerHeight(useCanvasModifiers)
        End If
        
    End With

End Sub

'Given a PD layer resize enum, return a corresponding string representation.  PD layer resize mode
' strings are ALWAYS 4-chars long; append spaces as necessary.
Public Function GetLayerResizeIDFromString(ByRef srcString As String) As PD_LayerResizeQuality
    Select Case srcString
        Case "near"
            GetLayerResizeIDFromString = LRQ_NearestNeighbor
        Case "blnr"
            GetLayerResizeIDFromString = LRQ_Bilinear
        Case "bcub"
            GetLayerResizeIDFromString = LRQ_Bicubic
        Case Else
            GetLayerResizeIDFromString = LRQ_NearestNeighbor
            PDDebug.LogAction "WARNING! Layers.GetLayerResizeIDFromString received a bad value: " & srcString
    End Select
End Function

Public Function GetLayerResizeStringFromID(ByVal srcID As PD_LayerResizeQuality) As String
    Select Case srcID
        Case LRQ_NearestNeighbor
            GetLayerResizeStringFromID = "near"
        Case LRQ_Bilinear
            GetLayerResizeStringFromID = "blnr"
        Case LRQ_Bicubic
            GetLayerResizeStringFromID = "bcub"
        Case Else
            GetLayerResizeStringFromID = "near"
            PDDebug.LogAction "WARNING! Colors.GetLayerResizeStringFromID received a bad value: " & srcID
    End Select
End Function

'Given a PD layer type enum, return a corresponding string representation.  PD layer type
' strings are ALWAYS 4-chars long; append spaces as necessary.
Public Function GetLayerTypeIDFromString(ByRef srcString As String) As PD_LayerType
    Select Case srcString
        Case "rast"
            GetLayerTypeIDFromString = PDL_Image
        Case "txtb"
            GetLayerTypeIDFromString = PDL_TextBasic
        Case "txta"
            GetLayerTypeIDFromString = PDL_TextAdvanced
        Case "adjs"
            GetLayerTypeIDFromString = PDL_Adjustment
        Case Else
            GetLayerTypeIDFromString = PDL_Image
            PDDebug.LogAction "WARNING! Layers.GetLayerTypeIDFromString received a bad value: " & srcString
    End Select
End Function

Public Function GetLayerTypeStringFromID(ByVal srcID As PD_LayerType) As String
    Select Case srcID
        Case PDL_Image
            GetLayerTypeStringFromID = "rast"
        Case PDL_TextBasic
            GetLayerTypeStringFromID = "txtb"
        Case PDL_TextAdvanced
            GetLayerTypeStringFromID = "txta"
        Case PDL_Adjustment
            GetLayerTypeStringFromID = "adjs"
        Case Else
            GetLayerTypeStringFromID = "rast"
            PDDebug.LogAction "WARNING! Colors.GetLayerTypeStringFromID received a bad value: " & srcID
    End Select
End Function

'Given a layer index and an x/y position (ALREADY CONVERTED TO LAYER COORDINATE SPACE!), return an RGBQUAD for the pixel
' at that location.  Note that the returned result is unprocessed; e.g. it will be in premultipled format.
'
'If the pixel lies outside the layer boundaries, the function will return FALSE.  Make sure to check this before evaluating
' the RGBQUAD.
Public Function GetRGBAPixelFromLayer(ByVal layerIndex As Long, ByVal layerX As Long, ByVal layerY As Long, ByRef dstQuad As RGBQuad) As Boolean

    'Before doing anything else, check to see if the x/y coordinate even lies inside the image
    Dim tmpLayerRef As pdLayer
    Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(layerIndex)
        
    If (layerX >= 0) And (layerY >= 0) And (layerX < tmpLayerRef.GetLayerDIB.GetDIBWidth) And (layerY < tmpLayerRef.GetLayerDIB.GetDIBHeight) Then
    
        'The point lies inside the layer, which means we need to figure out the color at this position
        GetRGBAPixelFromLayer = True
        
        'X and Y now represent the passed coordinate, but translated into the specified layer's coordinate space.
        ' Retrieve the color (and alpha, if relevant) at that point.
        Dim tmpData() As Byte, tSA As SafeArray2D
        tmpLayerRef.GetLayerDIB.WrapArrayAroundDIB tmpData, tSA
        
        Dim xStride As Long
        xStride = layerX * (tmpLayerRef.GetLayerDIB.GetDIBColorDepth \ 8)
        
        'Failsafe bounds check
        If ((xStride + 3) < tmpLayerRef.GetLayerDIB.GetDIBStride) And (layerY < tmpLayerRef.GetLayerDIB.GetDIBHeight) Then
        
            With dstQuad
                .Blue = tmpData(xStride, layerY)
                .Green = tmpData(xStride + 1, layerY)
                .Red = tmpData(xStride + 2, layerY)
                If (tmpLayerRef.GetLayerDIB.GetDIBColorDepth = 32) Then .Alpha = tmpData(xStride + 3, layerY)
            End With
            
        End If
        
        tmpLayerRef.GetLayerDIB.UnwrapArrayFromDIB tmpData
        
    'This coordinate does not lie inside the layer.
    Else
        GetRGBAPixelFromLayer = False
    End If

End Function

'Given an x/y pair (in IMAGE COORDINATES), return the top-most layer under that position, if any.
' The long-named optional parameter, "givePreferenceToCurrentLayer", will check the currently active layer
' before checking any others.  If the mouse is over one of the current layer's points-of-interest
' (e.g. a resize node), the function will return that layer instead of others that lay atop it.
' This allows the user to move and resize the current layer preferentially, and only if the current layer
' is completely out of the picture will other layers become activated.
Public Function GetLayerUnderMouse(ByVal imgX As Single, ByVal imgY As Single, Optional ByVal givePreferenceToCurrentLayer As Boolean = True) As Long

    Dim tmpRGBA As RGBQuad
    Dim curPOI As PD_PointOfInterest
    
    'Note that the caller passes us an (x, y) coordinate pair in the IMAGE coordinate space.  We will be using these coordinates to
    ' generate various new coordinate pairs in individual LAYER coordinate spaces  (This became necessary in PD 7.0, as layers
    ' may have non-destructive affine transforms active, which means we can't blindly switch between image and layer coordinate spaces!)
    Dim layerX As Single, layerY As Single
    
    'If givePreferenceToCurrentLayer is selected, check the current layer first.  If the mouse is over one of the layer's POIs, return
    ' the active layer without even checking other layers.
    If givePreferenceToCurrentLayer Then
    
        'Convert the passed image (x, y) coordinates into the active layer's coordinate space
        Drawing.ConvertImageCoordsToLayerCoords_Full PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, imgX, imgY, layerX, layerY
    
        'See if the mouse is over a POI for the current layer (which may extend outside a layer's boundaries, because the clickable
        ' nodes have a radius greater than 0).  If the mouse is over a POI, return the active layer index immediately.
        curPOI = PDImages.GetActiveImage.GetActiveLayer.CheckForPointOfInterest(imgX, imgY)
        
        'If the mouse is over a point of interest, return this layer and immediately exit
        If (curPOI <> poi_Undefined) And (curPOI <> poi_Interior) Then
            GetLayerUnderMouse = PDImages.GetActiveImage.GetActiveLayerIndex
            Exit Function
        End If
        
    End If

    'With the active layer out of the way, iterate through all image layers in reverse (e.g. top-to-bottom).  If one is located
    ' beneath the mouse, and the hovered image section is non-transparent (pending the user's preference for this), return it.
    Dim i As Long
    For i = PDImages.GetActiveImage.GetNumOfLayers - 1 To 0 Step -1
    
        'Only evaluate the current layer if it is visible
        If PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility Then
        
            'Convert the image (x, y) coordinate into the layer's coordinate space
            Drawing.ConvertImageCoordsToLayerCoords_Full PDImages.GetActiveImage(), PDImages.GetActiveImage.GetLayerByIndex(i), imgX, imgY, layerX, layerY
            
            'Only evaluate the current layer if the mouse is over it
            If Layers.GetRGBAPixelFromLayer(i, layerX, layerY, tmpRGBA) Then
            
                'A layer was identified beneath the mouse!  If the pixel is non-transparent,
                ' return this layer as the selected one.
                If (Not toolpanel_MoveSize.chkIgnoreTransparent.Value) Then
                    GetLayerUnderMouse = i
                    Exit Function
                Else
                
                    If (tmpRGBA.Alpha > 0) Then
                        GetLayerUnderMouse = i
                        Exit Function
                    End If
                
                End If
                            
            End If
        
        End If
    
    Next i
    
    'If we made it all the way here, there is no layer under this position.  Return -1 to signify failure.
    GetLayerUnderMouse = -1

End Function

'If a function must rasterize a vector or text layer, it needs to call this function first.  This function will display a dialog
' asking the user for permission to rasterize the layer(s) in question.  Note that CANCEL is a valid return, so any callers need
' to handle that case gracefully!
Public Function AskIfOkayToRasterizeLayer(Optional ByVal srcLayerType As PD_LayerType = PDL_TextBasic, Optional ByVal questionID As String = "RasterizeLayer", Optional ByVal multipleLayersInvolved As Boolean = False) As VbMsgBoxResult
    
    Dim questionText As String, yesText As String, noText As String, cancelText As String, rememberText As String, dialogTitle As String
    
    'If multiple layers are involved, we don't care about the current layer type
    If multipleLayersInvolved Then
    
        questionText = g_Language.TranslateMessage("This action will convert text and vector layers to image (raster) layers, meaning you can no longer modify layer-specific settings like text, font, color or shape.")
        questionText = questionText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Are you sure you want to continue?")
        yesText = g_Language.TranslateMessage("Yes.  Convert text and vector layers to image (raster) layers.")
        noText = g_Language.TranslateMessage("No.  Leave text and vector layers as they are.")
    
    'If a single layer is involved, we'll further customize the prompt on a per-layer-type basis
    Else
    
        'Generate customized question text based on layer type
        Select Case srcLayerType
    
            Case PDL_TextBasic, PDL_TextAdvanced
                questionText = g_Language.TranslateMessage("This text layer will be changed to an image (raster) layer, meaning you can no longer modify its text or font settings.")
                questionText = questionText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Are you sure you want to continue?")
                yesText = g_Language.TranslateMessage("Yes.  Please convert this text layer.")
                noText = g_Language.TranslateMessage("No.  Leave this text layer as it is.")
            
            Case Else
                PDDebug.LogAction "WARNING!  Unknown or invalid layer type passed to Layers.AskIfOkayToRasterizeLayer!"
    
        End Select
    
    End If
    
    'Cancel text, "remember in the future" check box text, and dialog title are universal
    cancelText = g_Language.TranslateMessage("I can't decide.  Cancel this action.")
    rememberText = g_Language.TranslateMessage("in the future, do this without asking me")
    dialogTitle = "Rasterization required"
    
    'Display the dialog and return the result
    AskIfOkayToRasterizeLayer = Dialogs.PromptGenericYesNoDialog_SingleOutcome(questionID, questionText, yesText, noText, cancelText, rememberText, dialogTitle, vbYes, IDI_EXCLAMATION, vbYes)

End Function

'Rasterize a given layer.  Pass -1 to rasterize all vector layers.
Public Sub RasterizeLayer(Optional ByVal srcLayerIndex As Long = -1)

    '-1 tells us to rasterize all vector layers
    If (srcLayerIndex = -1) Then
    
        Dim i As Long
        For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
            If PDImages.GetActiveImage.GetLayerByIndex(i).IsLayerVector Then
            
                'Rasterize this layer, and notify the parent image of the change
                PDImages.GetActiveImage.GetLayerByIndex(i).RasterizeVectorData
                PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
            
            End If
        Next i
    
    Else
        
        'Rasterize just this one layer, and notify the parent image of the change
        If PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).IsLayerVector() Then
            PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).RasterizeVectorData
            PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, srcLayerIndex
        End If
        
    End If
    
    'Re-sync the interface
    SyncInterfaceToCurrentImage
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

'Commit the current scratch layer onto the active layer.  The caller *must* supply a Processor.Process
' ID to use, and they are responsible for localizing this string as well (since it won't be auto-detected
' by PD's translation file generator).  Also required is the boundary rectangle to commit; for performance
' reasons, this should obviously be the smallest size you can get away with.
Public Sub CommitScratchLayer(ByRef processNameToUse As String, ByRef srcRectF As RectF, Optional ByRef srcParamString As String = vbNullString)

    With srcRectF
        If (.Left < 0) Then .Left = 0
        If (.Top < 0) Then .Top = 0
        If (.Width > PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.GetDIBWidth) Then .Width = PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.GetDIBWidth
        If (.Height > PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.GetDIBHeight) Then .Height = PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.GetDIBHeight
    End With
    
    Dim undoType As PD_UndoType
    
    'Committing brush results is actually pretty easy!
    
    'First, if the layer beneath the paint stroke is a raster layer, we simply want to merge the scratch
    ' layer onto it.
    If PDImages.GetActiveImage.GetActiveLayer.IsLayerRaster Then
        
        PDImages.GetActiveImage.MergeTwoLayers PDImages.GetActiveImage.ScratchLayer, PDImages.GetActiveImage.GetActiveLayer, True, VarPtr(srcRectF)
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, PDImages.GetActiveImage.GetActiveLayerIndex
        
        'Ask the central processor to create Undo/Redo data for us.
        ' (IMPORTANT NOTE: when playing back a macro, the macro operation itself will take care of
        ' Undo/Redo generation - we simply need to flag the processor to update the screen and perform
        ' any maintenance tasks, but we explicitly request that the processor does *not* generate undo
        ' data on this internal call!)
        If ((Macros.GetMacroStatus = MacroPLAYBACK) Or (Macros.GetMacroStatus = MacroBATCH)) Then undoType = UNDO_Nothing Else undoType = UNDO_Layer
        Processor.Process processNameToUse, False, srcParamString, undoType, g_CurrentTool
        
        'Reset the scratch layer (if it hasn't already been freed)
        If (Not PDImages.GetActiveImage.ScratchLayer Is Nothing) Then
            If (Not PDImages.GetActiveImage.ScratchLayer.GetLayerDIB Is Nothing) Then PDImages.GetActiveImage.ScratchLayer.GetLayerDIB.ResetDIB 0
        End If
        
    'If the layer beneath this one is *not* a raster layer, let's add the stroke as a new layer, instead.
    Else
        
        'Before creating the new layer, check for an active selection.  If one exists, we need to pre-process
        ' the paint layer against it.
        If PDImages.GetActiveImage.IsSelectionActive Then
            
            'A selection is active.  Pre-mask the paint scratch layer against it.
            Dim cBlender As pdPixelBlender
            Set cBlender = New pdPixelBlender
            cBlender.ApplyMaskToTopDIB PDImages.GetActiveImage.ScratchLayer.GetLayerDIB, PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB, VarPtr(srcRectF)
            
        End If
        
        Dim newLayerID As Long
        newLayerID = PDImages.GetActiveImage.CreateBlankLayer(PDImages.GetActiveImage.GetActiveLayerIndex)
        
        'Point the new layer index at our scratch layer
        PDImages.GetActiveImage.PointLayerAtNewObject newLayerID, PDImages.GetActiveImage.ScratchLayer
        PDImages.GetActiveImage.GetLayerByID(newLayerID).SetLayerName g_Language.TranslateMessage("Paint layer")
        Set PDImages.GetActiveImage.ScratchLayer = Nothing
        
        'Activate the new layer
        PDImages.GetActiveImage.SetActiveLayerByID newLayerID
        
        'Crop any dead space from the scratch layer
        PDImages.GetActiveImage.GetActiveLayer.CropNullPaddedLayer
        
        'Notify the parent image of the new layer
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
        
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.NotifyLayerChange
        
        'Ask the central processor to create Undo/Redo data for us.  (See IMPORTANT NOTE on the
        ' previous section for details on why the undo status is set manually.)
        If ((Macros.GetMacroStatus = MacroPLAYBACK) Or (Macros.GetMacroStatus = MacroBATCH)) Then undoType = UNDO_Nothing Else undoType = UNDO_Image_VectorSafe
        Processor.Process processNameToUse, False, srcParamString, undoType, g_CurrentTool
        
        'Create a new scratch layer
        Tools.InitializeToolsDependentOnImage
        
    End If
    
    'Before exiting, forcibly clear the temporary viewport DIB for the scratch layer
    ' (as its contents are no longer needed).
    If (Not PDImages.GetActiveImage.ScratchLayer Is Nothing) Then
        If (Not PDImages.GetActiveImage.ScratchLayer.TmpLODDIB(CLC_Viewport) Is Nothing) Then
            PDImages.GetActiveImage.ScratchLayer.TmpLODDIB(CLC_Viewport).ResetDIB 0
        End If
    End If
    
End Sub

'When a non-layered image is first loaded, the image itself is created as the base layer.  Unlike other software
' (which just assigns a stupid "Background" label), PD tries to generate a meaningful name for this layer.
' IMPORTANT NOTE: if passing a page index, note that the value is 0-BASED, so page "1" should be passed as "0".
'  (We do this to simplify interactions with the FreeImage plugin, which handles the bulk of our multipage interface.)
Public Function GenerateInitialLayerName(ByRef srcFile As String, Optional ByVal suggestedFilename As String = vbNullString, Optional ByVal imageHasMultiplePages As Boolean = False, Optional ByRef srcImage As pdImage, Optional ByRef srcDIB As pdDIB, Optional ByVal currentPageIndex As Long = 0) As String
    
    'If a multipage image is loaded as individual layers, each layer will receive a custom name to reflect its position in the
    ' original file.  (For example, when loading .ICO files with multiple icons inside, PD will automatically add the name and
    ' original bit-depth to each layer, as relevant.)
    If imageHasMultiplePages Or (srcImage.GetOriginalFileFormat = FIF_ICO) Then
        
        Select Case srcImage.GetOriginalFileFormat
        
            'Animations will be called "frames" instead of pages
            Case PDIF_GIF, PDIF_JXL, PDIF_PNG, PDIF_WEBP
                GenerateInitialLayerName = g_Language.TranslateMessage("Frame %1", CStr(currentPageIndex + 1))
                
            'Any other format is treated as "pages" (0-based index)
            Case Else
                GenerateInitialLayerName = g_Language.TranslateMessage("Page %1", CStr(currentPageIndex + 1))
                
        End Select
    
    'The first layer of single-layer images use a simpler naming system
    Else
        If (LenB(suggestedFilename) = 0) Then
            GenerateInitialLayerName = Files.FileGetName(srcFile, True)
        Else
            GenerateInitialLayerName = suggestedFilename
        End If
    End If
    
End Function

Public Sub PadToImageSize(ByRef srcImage As pdImage, Optional ByVal srcLayerIndex As Long = -1)
    
    If (srcImage Is Nothing) Then Exit Sub
    If (srcLayerIndex = -1) Then srcLayerIndex = srcImage.GetActiveLayerIndex
    
    PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).ConvertToNullPaddedLayer srcImage.Width, srcImage.Height, True
    
End Sub

Public Sub TrimEmptyBorders(ByRef srcImage As pdImage, Optional ByVal srcLayerIndex As Long = -1)
    
    If (srcImage Is Nothing) Then Exit Sub
    If (srcLayerIndex = -1) Then srcLayerIndex = srcImage.GetActiveLayerIndex
    
    'Make sure the layer is null-padded before trimming
    PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).ConvertToNullPaddedLayer srcImage.Width, srcImage.Height, True
    PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).CropNullPaddedLayer
    
End Sub

Public Sub SetGenericLayerProperty(ByVal propID As PD_LayerGenericProperty, ByVal propValue As Variant, Optional ByVal srcLayerIndex As Long = -1)
    
    'Failsafe checks
    If (Not PDImages.IsImageNonNull(PDImages.GetActiveImageID)) Then Exit Sub
    
    'Validate layer index and translate into a fixed ID >= 0
    If (srcLayerIndex < 0) Then srcLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex
    If (srcLayerIndex >= PDImages.GetActiveImage.GetNumOfLayers()) Then srcLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex
    
    'Apply the new property.  (It's up to the layer object to validate any inputs.)
    PDImages.GetActiveImage.GetLayerByIndex(srcLayerIndex).SetGenericLayerProperty propID, propValue
    
End Sub
