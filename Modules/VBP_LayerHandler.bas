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
Public Sub loadImageAsNewLayer(ByVal showDialog As Boolean, Optional ByVal imagePath As String = "")

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
            pdImages(g_CurrentImage).getLayerByID(newLayerID).CreateNewImageLayer tmpDIB, pdImages(g_CurrentImage), Trim$(getFilename(imagePath))
            
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
    
    'Redraw the layer box, but note that thumbnails don't need to be re-cached
    toolbar_Layers.forceRedraw False
    
    'Redraw the viewport, but only if requested
    If alsoRedrawViewport Then ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Same idea as setActiveLayerByID, above
Public Sub setActiveLayerByIndex(ByVal newLayerIndex As Long, Optional ByVal alsoRedrawViewport As Boolean = False)

    'Notify the parent PD image of the change
    pdImages(g_CurrentImage).setActiveLayerByIndex newLayerIndex
    
    'Redraw the layer box, but note that thumbnails don't need to be re-cached
    toolbar_Layers.forceRedraw False
    
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
