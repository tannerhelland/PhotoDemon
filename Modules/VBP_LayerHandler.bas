Attribute VB_Name = "Layer_Handler"
'***************************************************************************
'Layer Interface
'Copyright ©2013-2014 by Tanner Helland
'Created: 24/March/14
'Last updated: 24/May/14
'Last update: add support for auto-activating the layer beneath a given mouse position
'
'This module provides all layer-related functions that interact with PhotoDemon's central processor.  Most of these
' functions are triggered by either the Layer menu, or the Layer toolbox.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Helper API for enlarging a given rect
Private Declare Function InflateRect Lib "user32" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

'Add a blank 32bpp layer above the specified layer index (typically the currently active layer)
Public Sub addBlankLayer(ByVal dLayerIndex As Long)

    'Validate the requested layer index
    If dLayerIndex < 0 Then dLayerIndex = 0
    If dLayerIndex > pdImages(g_CurrentImage).getNumOfLayers - 1 Then dLayerIndex = pdImages(g_CurrentImage).getNumOfLayers - 1
    
    'Ask the parent pdImage to create a new layer object
    Dim newLayerID As Long
    newLayerID = pdImages(g_CurrentImage).createBlankLayer(dLayerIndex)
    
    'Assign the newly created layer the IMAGE type, and initialize it to the size of the image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createBlank pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, 32, 0, 0
    pdImages(g_CurrentImage).getLayerByID(newLayerID).CreateNewImageLayer tmpDIB, , g_Language.TranslateMessage("Blank layer")
    
    'Make the blank layer the new active layer
    pdImages(g_CurrentImage).setActiveLayerByID newLayerID
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.forceRedraw True
    
    'Render the new image to screen (not technically necessary, but doesn't hurt)
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "New layer added"
            
    'Synchronize the interface to the new image
    syncInterfaceToCurrentImage
    
End Sub

'Allow the user to load an image file as a layer
Public Sub loadImageAsNewLayer(ByVal showDialog As Boolean, Optional ByVal imagePath As String = "", Optional ByVal customLayerName As String = "")

    'This function handles two cases: retrieving the filename from a common dialog box, and actually
    ' loading the image file and applying it to the current pdImage as a new layer.
    
    'If showDialog is TRUE, we need to get a file path from the user
    If showDialog Then
    
        'Retrieve a filepath
        Dim imgFilePath As String
        If File_Menu.PhotoDemon_OpenImageDialog_Simple(imgFilePath, FormMain.hWnd) Then
            Process "New Layer from File", False, imgFilePath, UNDO_IMAGE
        End If
    
    'If showDialog is FALSE, the user has already selected a file, and we just need to load it
    Else
    
        'Prepare a temporary DIB
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        
        'Load the file in question
        If Loading.QuickLoadImageToDIB(imagePath, tmpDIB) Then
            
            'Forcibly convert the new layer to 32bpp
            If tmpDIB.getDIBColorDepth = 24 Then tmpDIB.convertTo32bpp
            
            'Ask the current image to prepare a blank layer for us
            Dim newLayerID As Long
            newLayerID = pdImages(g_CurrentImage).createBlankLayer()
            
            'Convert the layer to an IMAGE-type layer and copy the newly loaded DIB's contents into it
            If Len(customLayerName) = 0 Then
                pdImages(g_CurrentImage).getLayerByID(newLayerID).CreateNewImageLayer tmpDIB, pdImages(g_CurrentImage), Trim$(getFilenameWithoutExtension(imagePath))
            Else
                pdImages(g_CurrentImage).getLayerByID(newLayerID).CreateNewImageLayer tmpDIB, pdImages(g_CurrentImage), customLayerName
            End If
            
            Debug.Print "Layer created successfully (ID# " & pdImages(g_CurrentImage).getLayerByID(newLayerID).getLayerName & ")"
            
            'Render the new image to screen
            PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "New layer added"
            
            'Synchronize the interface to the new image
            syncInterfaceToCurrentImage
            
            Message "New layer added successfully."
        
        Else
            Message "Image file could not be loaded (unknown error occurred)."
        End If
    
    End If

End Sub

'Activate a layer.  Use this instead of directly calling the pdImage.setActiveLayer function if you want to also
' synchronize the UI to match.
Public Sub setActiveLayerByID(ByVal newLayerID As Long, Optional ByVal alsoRedrawViewport As Boolean = False)

    'If this layer is already active, ignore the request
    If pdImages(g_CurrentImage).getActiveLayerID = newLayerID Then Exit Sub
    
    'Before changing to the new active layer, see if the previously active layer has had any non-destructive changes made.
    Processor.evaluateImageCheckpoint

    'Notify the parent PD image of the change
    pdImages(g_CurrentImage).setActiveLayerByID newLayerID
    
    'Sync the interface to the new layer
    syncInterfaceToCurrentImage
    
    'Set a new image checkpoint (necessary to do this manually, as we haven't invoked PD's central processor)
    Processor.setImageCheckpoint
    
    'Redraw the viewport, but only if requested
    If alsoRedrawViewport Then ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Same idea as setActiveLayerByID, above
Public Sub setActiveLayerByIndex(ByVal newLayerIndex As Long, Optional ByVal alsoRedrawViewport As Boolean = False)

    'If this layer is already active, ignore the request
    If pdImages(g_CurrentImage).getActiveLayerID = pdImages(g_CurrentImage).getLayerByIndex(newLayerIndex).getLayerID Then Exit Sub
    
    'Before changing to the new active layer, see if the previously active layer has had any non-destructive changes made.
    Processor.evaluateImageCheckpoint
    
    'Notify the parent PD image of the change
    pdImages(g_CurrentImage).setActiveLayerByIndex newLayerIndex
    
    'Sync the interface to the new layer
    syncInterfaceToCurrentImage
    
    'Set a new image checkpoint (necessary to do this manually, as we haven't invoked PD's central processor)
    Processor.setImageCheckpoint
    
    'Redraw the viewport, but only if requested
    If alsoRedrawViewport Then ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Set layer visibility.  Note that the layer's visibility state must be explicitly noted, e.g. there is no "toggle" option.
Public Sub setLayerVisibilityByIndex(ByVal dLayerIndex As Long, ByVal layerVisibility As Boolean, Optional ByVal alsoRedrawViewport As Boolean = False)
    
    'Store the new visibility setting in the parent pdImage object
    pdImages(g_CurrentImage).getLayerByIndex(dLayerIndex).setLayerVisibility layerVisibility
    
    'Redraw the layer box, but note that thumbnails don't need to be re-cached
    toolbar_Layers.forceRedraw False
    
    'Synchronize the interface to the new image
    syncInterfaceToCurrentImage
    
    'Redraw the viewport, but only if requested
    If alsoRedrawViewport Then ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Duplicate a given layer (note: it doesn't have to be the active layer)
Public Sub duplicateLayerByIndex(ByVal dLayerIndex As Long)

    'Validate the requested layer index
    If dLayerIndex < 0 Then dLayerIndex = 0
    If dLayerIndex > pdImages(g_CurrentImage).getNumOfLayers - 1 Then dLayerIndex = pdImages(g_CurrentImage).getNumOfLayers - 1
    
    'Before doing anything else, make a copy of the current active layer ID.  We will use this to restore the same
    ' active layer after the creation is complete.
    Dim activeLayerID As Long
    activeLayerID = pdImages(g_CurrentImage).getActiveLayerID
    
    'Also copy the ID of the layer we are creating.
    Dim dupedLayerID As Long
    dupedLayerID = pdImages(g_CurrentImage).getLayerByIndex(dLayerIndex).getLayerID
    
    'Ask the parent pdImage to create a new layer object
    Dim newLayerID As Long
    newLayerID = pdImages(g_CurrentImage).createBlankLayer(dLayerIndex)
            
    'Ask the new layer to copy the contents of the layer we are duplicating
    pdImages(g_CurrentImage).getLayerByID(newLayerID).CopyExistingLayer pdImages(g_CurrentImage).getLayerByID(dupedLayerID)
    
    'Make the duplicate layer the active layer
    pdImages(g_CurrentImage).setActiveLayerByID newLayerID
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.forceRedraw True
    
    'Render the new image to screen
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "New layer added"
            
    'Synchronize the interface to the new image
    syncInterfaceToCurrentImage
    
End Sub

'Merge the layer at layerIndex up or down.
Public Sub mergeLayerAdjacent(ByVal dLayerIndex As Long, ByVal mergeDown As Boolean)

    'Look for a valid target layer to merge with in the requested direction.
    Dim mergeTarget As Long
    mergeTarget = isLayerAllowedToMergeAdjacent(dLayerIndex, mergeDown)
    
    'If we've been given a valid merge target, apply it now!
    If mergeTarget >= 0 Then
    
        If mergeDown Then
        
            With pdImages(g_CurrentImage)
                
                'Request a merge from the parent pdImage
                .mergeTwoLayers .getLayerByIndex(dLayerIndex), .getLayerByIndex(mergeTarget), False
                
                'Delete the now-merged layer
                .deleteLayerByIndex dLayerIndex
                
                'Set the newly merged layer as the active layer
                .setActiveLayerByIndex mergeTarget
            
            End With
            
        Else
        
            With pdImages(g_CurrentImage)
            
                'Request a merge from the parent pdImage
                .mergeTwoLayers .getLayerByIndex(mergeTarget), .getLayerByIndex(dLayerIndex), False
                
                'Delete the now-merged layer
                .deleteLayerByIndex mergeTarget
                
                'Set the newly merged layer as the active layer
                .setActiveLayerByIndex dLayerIndex
                
            End With
        
        End If
        
        'Redraw the layer box, and note that thumbnails need to be re-cached
        toolbar_Layers.forceRedraw True
    
        'Redraw the viewport
        ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
    End If

End Sub

'Is this layer allowed to merge up or down?  Note that invisible layers are not generally considered suitable
' for merging, so a layer will typically be merged with the next VISIBLE layer.  If none are available, merging
' is disallowed.
'
'Note that the return value for this function is a little wonky.  This function will return the TARGET MERGE LAYER
' INDEX if the function is true.  This value will always be >= 0.  If no valid layer can be found, -1 will be
' returned (which obviously isn't a valid index, but IS true, so it's a little confusing - handle accordingly!)
Public Function isLayerAllowedToMergeAdjacent(ByVal dLayerIndex As Long, ByVal moveDown As Boolean) As Long

    Dim i As Long
    
    'Check MERGE DOWN
    If moveDown Then
    
        'As an easy check, make sure this layer is visible, and not already at the bottom.
        If (dLayerIndex <= 0) Or (Not pdImages(g_CurrentImage).getLayerByIndex(dLayerIndex).getLayerVisibility) Then
            isLayerAllowedToMergeAdjacent = -1
            Exit Function
        End If
        
        'Search for the nearest valid layer beneath this one.
        For i = dLayerIndex - 1 To 0 Step -1
            If pdImages(g_CurrentImage).getLayerByIndex(i).getLayerVisibility Then
                isLayerAllowedToMergeAdjacent = i
                Exit Function
            End If
        Next i
        
        'If we made it all the way here, no valid merge target was found.  Return failure (-1).
        isLayerAllowedToMergeAdjacent = -1
    
    'Check MERGE UP
    Else
    
        'As an easy check, make sure this layer isn't already at the top.
        If (dLayerIndex >= pdImages(g_CurrentImage).getNumOfLayers - 1) Or (Not pdImages(g_CurrentImage).getLayerByIndex(dLayerIndex).getLayerVisibility) Then
            isLayerAllowedToMergeAdjacent = -1
            Exit Function
        End If
        
        'Search for the nearest valid layer above this one.
        For i = dLayerIndex + 1 To pdImages(g_CurrentImage).getNumOfLayers - 1
            If pdImages(g_CurrentImage).getLayerByIndex(i).getLayerVisibility Then
                isLayerAllowedToMergeAdjacent = i
                Exit Function
            End If
        Next i
        
        'If we made it all the way here, no valid merge target was found.  Return failure (-1).
        isLayerAllowedToMergeAdjacent = -1
    
    End If

End Function

'Delete a given layer
Public Sub deleteLayer(ByVal dLayerIndex As Long)

    'Cache the current layer index
    Dim curLayerIndex As Long
    curLayerIndex = pdImages(g_CurrentImage).getActiveLayerIndex - 1

    pdImages(g_CurrentImage).deleteLayerByIndex dLayerIndex
    
    'Set a new active layer
    If curLayerIndex > pdImages(g_CurrentImage).getNumOfLayers - 1 Then curLayerIndex = pdImages(g_CurrentImage).getNumOfLayers - 1
    If curLayerIndex < 0 Then curLayerIndex = 0
    setActiveLayerByIndex curLayerIndex, False
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.forceRedraw True
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'Delete all hidden layers
Public Sub deleteHiddenLayers()

    'Perform a couple fail-safe checks.  These should not be a problem, as calling functions should have safeguards
    ' against bad requests, but better safe than sorry.
    
    'If there are no hidden layers, exit
    If pdImages(g_CurrentImage).getNumOfHiddenLayers = 0 Then Exit Sub
    
    'If all layers are hidden, exit
    If pdImages(g_CurrentImage).getNumOfHiddenLayers = pdImages(g_CurrentImage).getNumOfLayers Then Exit Sub
    
    'We can now assume that the image in question has at least one visible layer, and at least one hidden layer.
    
    'Cache the currently active layerID - IF the current layer is visible.  If it isn't, it's going to be deleted,
    ' so we must pick a new arbitrary layer (why not the bottom layer?).
    Dim activeLayerID As Long
    
    If pdImages(g_CurrentImage).getActiveLayer.getLayerVisibility Then
        activeLayerID = pdImages(g_CurrentImage).getActiveLayerID
    Else
        activeLayerID = -1
    End If
    
    'Starting at the top and moving down, delete all hidden layers.
    Dim i As Long
    For i = pdImages(g_CurrentImage).getNumOfLayers - 1 To 0 Step -1
    
        If Not pdImages(g_CurrentImage).getLayerByIndex(i).getLayerVisibility Then
            pdImages(g_CurrentImage).deleteLayerByIndex i
        End If
    Next i
    
    'Set a new active layer
    If activeLayerID = -1 Then
        setActiveLayerByIndex 0, False
    Else
        setActiveLayerByID activeLayerID
    End If
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.forceRedraw True
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'Move a layer up or down in the stack (referred to as "raise" and "lower" in the menus)
Public Sub moveLayerAdjacent(ByVal dLayerIndex As Long, ByVal directionIsUp As Boolean)

    'Make a copy of the currently active layer's ID
    Dim curActiveLayerID As Long
    curActiveLayerID = pdImages(g_CurrentImage).getActiveLayerID
    
    'Ask the parent pdImage to move the layer for us
    pdImages(g_CurrentImage).moveLayerByIndex pdImages(g_CurrentImage).getActiveLayerIndex, directionIsUp
    
    'Restore the active layer
    setActiveLayerByID curActiveLayerID, False
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.forceRedraw True
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'Move a layer to the top or bottom of the stack (referred to as "raise to top" and "lower to bottom" in the menus)
Public Sub moveLayerToEndOfStack(ByVal dLayerIndex As Long, ByVal moveToTopOfStack As Boolean)

    'Make a copy of the currently active layer's ID
    Dim curActiveLayerID As Long
    curActiveLayerID = pdImages(g_CurrentImage).getActiveLayerID
    
    'Until this layer is at the desired end of the stack, ask the parent to keep moving it for us!
    If moveToTopOfStack Then
    
        Do While pdImages(g_CurrentImage).getLayerIndexFromID(curActiveLayerID) < pdImages(g_CurrentImage).getNumOfLayers - 1
            
            'Ask the parent pdImage to move the layer up for us
            pdImages(g_CurrentImage).moveLayerByIndex pdImages(g_CurrentImage).getLayerIndexFromID(curActiveLayerID), True
            
        Loop
    
    Else
    
        Do While pdImages(g_CurrentImage).getLayerIndexFromID(curActiveLayerID) > 0
            
            'Ask the parent pdImage to move the layer up for us
            pdImages(g_CurrentImage).moveLayerByIndex pdImages(g_CurrentImage).getLayerIndexFromID(curActiveLayerID), False
            
        Loop
    
    End If
    
    'Restore the active layer.  (This will also re-synchronize the interface against the new image.)
    setActiveLayerByID curActiveLayerID, False
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.forceRedraw True
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'Given a multi-layered image, flatten it.  Note that flattening does *not* remove alpha!  It simply merges all layers,
' including discarding invisible ones.
Public Sub flattenImage()

    'Start by retrieving a copy of the composite image
    Dim compositeDIB As pdDIB
    Set compositeDIB = New pdDIB
    
    pdImages(g_CurrentImage).getCompositedImage compositeDIB
    
    'Also, grab the name of the bottom-most layer.  This will be used as the name of our only layer in the flattened image.
    Dim flattenedName As String
    flattenedName = pdImages(g_CurrentImage).getLayerByIndex(0).getLayerName
    
    'With this information, we can now delete all image layers.
    Do
        pdImages(g_CurrentImage).deleteLayerByIndex 0
    Loop While pdImages(g_CurrentImage).getNumOfLayers > 1
    
    'Note that the delete operation does not allow us to delete all layers.  (If there is only one layer present,
    ' it will exit without modifying the image.)  Because of that, the image will still retain one layer, which
    ' we will have to manually overwrite.
        
    'Reset any optional layer parameters to their default state
    pdImages(g_CurrentImage).getLayerByIndex(0).resetLayerParameters
    
    'Overwrite the final layer with the composite DIB.
    pdImages(g_CurrentImage).getLayerByIndex(0).CreateNewImageLayer compositeDIB, , flattenedName
    
    'Mark the only layer present as the active one.  (This will also re-synchronize the interface against the new image.)
    setActiveLayerByIndex 0, False
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.forceRedraw True
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'Given a multi-layered image, merge all visible layers, while ignoring any hidden ones.  Note that flattening does *not*
' remove alpha!  It simply merges all visible layers.
Public Sub mergeVisibleLayers()
    
    'If there's only one visible layer, this function should not be called - but just in case, exit in advance.
    If pdImages(g_CurrentImage).getNumOfLayers = 1 Then Exit Sub
    
    'SIf there's only one visible layer, this function should not be called - but just in case, exit in advance.
    If pdImages(g_CurrentImage).getNumOfVisibleLayers = 1 Then Exit Sub
    
    'By this point, we can assume there are at least two visible layers in the image.  Rather than deal with the messiness
    ' of finding the lowest base layer and gradually merging everything into it, we're going to just create a new blank
    ' layer at the base of the image, then merge everything with it until finally all visible layers have been merged.
    
    'Insert a new layer at the bottom of the layer stack.
    pdImages(g_CurrentImage).createBlankLayer 0
    
    'Technically, the command above does not actually insert a new layer at the base of the image.  Per convention,
    ' it always inserts the requested layer at the spot one *above* the requested spot.  To work around this, swap
    ' our newly created layer with the layer at position 0.
    pdImages(g_CurrentImage).swapTwoLayers 0, 1
    
    'Fill that new layer with a blank DIB at the dimensions of the image.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createBlank pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, 32, 0
    pdImages(g_CurrentImage).getLayerByIndex(0).CreateNewImageLayer tmpDIB, , g_Language.TranslateMessage("Merged layers")
    
    'With that done, merging visible layers is actually not that hard.  Loop through the layer collection,
    ' merging visible layers with the base layer, until all visible layers have been merged.
    Dim i As Long
    For i = 1 To pdImages(g_CurrentImage).getNumOfLayers - 1
    
        'If this layer is visible, merge it with the base layer
        If pdImages(g_CurrentImage).getLayerByIndex(i).getLayerVisibility Then
            pdImages(g_CurrentImage).mergeTwoLayers pdImages(g_CurrentImage).getLayerByIndex(i), pdImages(g_CurrentImage).getLayerByIndex(0), True
        End If
    
    Next i
    
    'Now that our base layer contains the result of merging all visible layers, we can now delete all
    ' other visible layers.
    For i = pdImages(g_CurrentImage).getNumOfLayers - 1 To 1 Step -1
        If pdImages(g_CurrentImage).getLayerByIndex(i).getLayerVisibility Then
            pdImages(g_CurrentImage).deleteLayerByIndex i
        End If
    Next i
    
    'Mark the new merged layer as the active one.  (This will also re-synchronize the interface against the new image.)
    setActiveLayerByIndex 0, False
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.forceRedraw True
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'If a layer has been transformed using the on-canvas tools, this will reset it to its default size.
Public Sub resetLayerSize(ByVal srcLayerIndex As Long)

    pdImages(g_CurrentImage).getLayerByIndex(srcLayerIndex).setLayerCanvasXModifier 1
    pdImages(g_CurrentImage).getLayerByIndex(srcLayerIndex).setLayerCanvasYModifier 1
    
    'Re-sync the interface
    syncInterfaceToCurrentImage
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'Resize a layer non-destructively, e.g. by only changing its position and on-canvas x/y modifiers
Public Sub resizeLayerNonDestructive(ByVal srcLayerIndex As Long, ByVal resizeParams As String)

    'Create a parameter parser to help us interpret the passed param string
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString resizeParams
    
    'Apply the passed parameters to the specified layer
    With pdImages(g_CurrentImage).getLayerByIndex(srcLayerIndex)
        .setLayerOffsetX cParams.GetDouble(1)
        .setLayerOffsetY cParams.GetDouble(2)
        .setLayerCanvasXModifier cParams.GetDouble(3)
        .setLayerCanvasYModifier cParams.GetDouble(4)
    End With
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'Move a layer to a new x/y position on the canvas
Public Sub moveLayerOnCanvas(ByVal srcLayerIndex As Long, ByVal resizeParams As String)

    'Create a parameter parser to help us interpret the passed param string
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString resizeParams
    
    'Apply the passed parameters to the specified layer
    With pdImages(g_CurrentImage).getLayerByIndex(srcLayerIndex)
        .setLayerOffsetX cParams.GetDouble(1)
        .setLayerOffsetY cParams.GetDouble(2)
    End With
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

'Given a layer, populate a rect with its coordinates (relative to the main image coordinates, always)
Public Sub fillRectForLayer(ByRef srcLayer As pdLayer, ByRef dstRect As RECT, Optional ByVal useCanvasModifiers As Boolean = False)

    With srcLayer
        dstRect.Left = .getLayerOffsetX
        If useCanvasModifiers Then
            dstRect.Right = .getLayerOffsetX + (.getLayerCanvasWidthModified)
        Else
            dstRect.Right = .getLayerOffsetX + .layerDIB.getDIBWidth
        End If
        dstRect.Top = .getLayerOffsetY
        If useCanvasModifiers Then
            dstRect.Bottom = .getLayerOffsetY + (.getLayerCanvasHeightModified)
        Else
            dstRect.Bottom = .getLayerOffsetY + .layerDIB.getDIBHeight
        End If
    End With

End Sub

'Given a param string (where the first entry denotes the target layer, and the subsequent parameters were set by
' a pdLayer object's getLayerHeaderAsParamString function), apply any changes to the specified layer.
Public Sub modifyLayerByParamString(ByVal pString As String)

    'Create a param string parser
    Dim cParams As pdParamString
    Set cParams = New pdParamString
    cParams.setParamString pString
    
    'Retrieve the ID of the layer in question
    Dim curLayerID As Long
    curLayerID = cParams.GetLong(1)
    
    'Remove that initial entry from the param string, then forward the rest of the string on to the specified layer class
    cParams.removeParamAtPosition 1
    pdImages(g_CurrentImage).getLayerByID(curLayerID).setLayerHeaderFromParamString cParams.getParamString

End Sub

'Given a layer index and an x/y position (IMAGE COORDINATE SPACE - necessary because we have to adjust the coordinates if the
' current layer has non-destructive resize modifiers applied), return an RGBQUAD for the pixel at that location.
'
'If the pixel lies outside the layer boundaries, the function will return FALSE.  Make sure to check this before evaluating
' the RGBQUAD.
Public Function getRGBAPixelFromLayer(ByVal layerIndex As Long, ByVal x As Long, ByVal y As Long, ByRef dstQuad As RGBQUAD, Optional ByVal enlargeForInteractionPadding As Boolean = True) As Boolean

    'Before doing anything else, check to see if the x/y coordinate even lies inside the image
    Dim tmpLayerRef As pdLayer
    Set tmpLayerRef = pdImages(g_CurrentImage).getLayerByIndex(layerIndex)
    
    Dim layerRect As RECT
    fillRectForLayer tmpLayerRef, layerRect, True
    
    If isPointInRect(x, y, layerRect) Then
        
        'The point lies inside the layer, which means we need to figure out the color at this position
        getRGBAPixelFromLayer = True
        
        'Re-calculate x and y to layer coordinates
        x = x - tmpLayerRef.getLayerOffsetX
        y = y - tmpLayerRef.getLayerOffsetY
        
        'If a non-destructive resize is active, remap the x/y coordinates to match
        If tmpLayerRef.getLayerCanvasXModifier <> 1 Then x = x / tmpLayerRef.getLayerCanvasXModifier
        If tmpLayerRef.getLayerCanvasYModifier <> 1 Then y = y / tmpLayerRef.getLayerCanvasYModifier
        
        'X and Y now represent the passed coordinate, but translated into the specified layer's coordinate space.
        ' Retrieve the color (and alpha, if relevant) at that point.
        Dim tmpData() As Byte
        Dim tSA As SAFEARRAY2D
        prepSafeArray tSA, tmpLayerRef.layerDIB
        CopyMemory ByVal VarPtrArray(tmpData()), VarPtr(tSA), 4
        
        Dim QuickX As Long
        QuickX = x * (tmpLayerRef.layerDIB.getDIBColorDepth \ 8)
        
        With dstQuad
            .Red = tmpData(QuickX + 2, y)
            .Green = tmpData(QuickX + 1, y)
            .Blue = tmpData(QuickX, y)
            If tmpLayerRef.layerDIB.getDIBColorDepth = 32 Then .Alpha = tmpData(QuickX + 3, y)
        End With
        
        CopyMemory ByVal VarPtrArray(tmpData), 0&, 4
    
    'This coordinate does not lie inside the layer.
    Else
    
        'If the "enlarge for interaction padding" option is set, make our rect a bit larger and then check again for validity.
        ' If this check succeeds, return TRUE, despite us not having a valid RGB coord for that location.
        If enlargeForInteractionPadding Then
        
            'Calculate PD's global mouse accuracy value, per the current image's zoom
            Dim mouseAccuracy As Double
            mouseAccuracy = g_MouseAccuracy * (1 / g_Zoom.getZoomValue(pdImages(g_CurrentImage).currentZoomValue))
            
            'Inflate the rect we were passed
            InflateRect layerRect, mouseAccuracy, mouseAccuracy
            
            'Check the point again
            If isPointInRect(x, y, layerRect) Then
            
                'Return TRUE, but the caller should know that the rgbQuad value is *not necessarily accurate*!
                getRGBAPixelFromLayer = True
            
            Else
                getRGBAPixelFromLayer = False
            End If
        
        Else
            getRGBAPixelFromLayer = False
        End If
    
    End If

End Function

'Given an x/y pair (in IMAGE COORDINATES), return the top-most layer under that position, if any.
' The long-named optional parameter, "givePreferenceToCurrentLayer", will check the currently active layer before checking any others.
' If the mouse is over one of the current layer's points-of-interest (e.g. a resize node), the function will return that layer instead
' of others that lay atop it.  This allows the user to move and resize the current layer preferentially, and only if the current layer
' is completely out of the picture will other layers become activated.
Public Function getLayerUnderMouse(ByVal curX As Long, ByVal curY As Long, Optional ByVal givePreferenceToCurrentLayer As Boolean = True) As Long

    Dim tmpRGBA As RGBQUAD
    Dim curPOI As Long
    
    'If givePreferenceToCurrentLayer is selected, check the current layer first.  If the mouse is over one of the layer's POIs, return
    ' the active layer without even checking other layers.
    If givePreferenceToCurrentLayer Then
    
        'See if the mouse is over a POI for the current layer (which may extend outside a layer's boundaries, because the clickable
        ' nodes have a radius greater than 0).  If the mouse is over a POI, return the active layer index immediately.
        curPOI = pdImages(g_CurrentImage).getActiveLayer.checkForPointOfInterest(curX, curY)
        
        'If the mouse is over a point of interest, return this layer and immediately exit
        If curPOI >= 0 And curPOI <= 3 Then
            getLayerUnderMouse = pdImages(g_CurrentImage).getActiveLayerIndex
            Exit Function
        End If
        
    End If

    'With the active layer out of the way, iterate through all image layers in reverse (e.g. top-to-bottom).  If one is located
    ' beneath the mouse,
    Dim i As Long
    For i = pdImages(g_CurrentImage).getNumOfLayers - 1 To 0 Step -1
    
        'Only evaluate the current layer if it is visible
        If pdImages(g_CurrentImage).getLayerByIndex(i).getLayerVisibility Then
        
            'Only evaluate the current layer if the mouse is over it
            If getRGBAPixelFromLayer(i, curX, curY, tmpRGBA) Then
            
                'A layer was identified beneath the mouse!  Now things get a bit tricky.
                ' NOTE: I have disabled this behavior until I can come up with a good solution on how to handle layer resizing
                '       while this "feature" is active.  If a layer has transparent pixels around its edges, resizing it becomes
                '       nearly impossible, because you have to blindly guess where the resize nodes are located, as they don't
                '       appear until the mouse is actually over them.  This is stupid.  As such, I'm disabling the transparent
                '       pixel modifier for now, and simply returning the layer beneath the mouse regardless of its transparency
                '       at the current point.
                
                'If the user doesn't want us to ignore transparent pixels, this is easy - just return true.
                'If Not CBool(toolbar_Tools.chkIgnoreTransparent) Then
                    'getLayerUnderMouse = i
                    'Exit Function
                
                'The user wants to ignore transparent pixels when auto-activating layers.  Before doing this, stop to see if
                ' the mouse is currently over a POI for this layer.  If it is, return the layer regardless of pixel transparency.
                'Else
                '
                '    curPOI = pdImages(g_CurrentImage).getLayerByIndex(i).checkForPointOfInterest(curX, curY)
                '
                '    'If the mouse is over a point of interest, return this layer regardless of underlying alpha
                '    If curPOI >= 0 And curPOI <= 3 Then
                '        getLayerUnderMouse = i
                '        Exit Function
                '
                '    'If the mouse is not over a point of interest, actually check alpha, and return this layer conditionally
                '    ' if the layer is NOT fully transparent at this point.
                '    Else
                '
                        If tmpRGBA.Alpha > 0 Then
                            getLayerUnderMouse = i
                            Exit Function
                        End If
                '
                '    End If
                '
                'End If
                            
            End If
        
        End If
    
    Next i
    
    'If we made it all the way here, there is no layer under this position.  Return -1 to signify failure.
    getLayerUnderMouse = -1

End Function
