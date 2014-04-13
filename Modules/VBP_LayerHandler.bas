Attribute VB_Name = "Layer_Handler"
'***************************************************************************
'Layer Interface
'Copyright ©2013-2014 by Tanner Helland
'Created: 24/March/14
'Last updated: 24/March/14
'Last update: initial build
'
'This module provides all layer-related functions that interact with PhotoDemon's central processor.  Most of these
' functions are triggered by either the Layer menu, or the Layer toolbox.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Allow the user to load an image file as a layer
Public Sub loadImageAsNewLayer(ByVal showDialog As Boolean, Optional ByVal imagePath As String = "", Optional ByVal customLayerName As String = "")

    'This function handles two cases: retrieving the filename from a common dialog box, and actually
    ' loading the image file and applying it to the current pdImage as a new layer.
    
    'If showDialog is TRUE, we need to get a file path from the user
    If showDialog Then
    
        'Retrieve a filepath
        Dim imgFilePath As String
        If File_Menu.PhotoDemon_OpenImageDialog_Simple(imgFilePath, FormMain.hWnd) Then
            Process "New Layer from File", False, imgFilePath
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

    'Notify the parent PD image of the change
    pdImages(g_CurrentImage).setActiveLayerByID newLayerID
    
    'Sync the interface to the new layer
    syncInterfaceToCurrentImage
    
    'Redraw the viewport, but only if requested
    If alsoRedrawViewport Then ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Same idea as setActiveLayerByID, above
Public Sub setActiveLayerByIndex(ByVal newLayerIndex As Long, Optional ByVal alsoRedrawViewport As Boolean = False)

    'Notify the parent PD image of the change
    pdImages(g_CurrentImage).setActiveLayerByIndex newLayerIndex
    
    'Sync the interface to the new layer
    syncInterfaceToCurrentImage
    
    'Redraw the viewport, but only if requested
    If alsoRedrawViewport Then ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Set layer visibility.  Note that the layer's visibility state must be explicitly noted, e.g. there is no "toggle" option.
Public Sub setLayerVisibilityByIndex(ByVal dLayerIndex As Long, ByVal layerVisibility As Boolean, Optional ByVal alsoRedrawViewport As Boolean = False)
    
    'Store the new visibility setting in the parent pdImage object
    pdImages(g_CurrentImage).getLayerByIndex(dLayerIndex).setLayerVisibility layerVisibility
    
    'Redraw the layer box, but note that thumbnails don't need to be re-cached
    toolbar_Layers.forceRedraw False
    
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
    
    'Restore the original active layer
    pdImages(g_CurrentImage).setActiveLayerByID activeLayerID
    
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

    pdImages(g_CurrentImage).deleteLayerByIndex dLayerIndex
    
    'Set a new active layer
    setActiveLayerByIndex pdImages(g_CurrentImage).getActiveLayerIndex, False
    
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
