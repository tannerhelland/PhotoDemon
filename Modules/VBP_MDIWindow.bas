Attribute VB_Name = "Image_Canvas_Handler"
'***************************************************************************
'Image Canvas Handler (formerly Image Window Handler)
'Copyright 2002-2016 by Tanner Helland
'Created: 11/29/02
'Last updated: 04/February/14
'Last update: rework all code to operate on Canvas user controls instead of standalone forms
'
'This module contains functions relating to the creation, sizing, and maintenance of the windows ("canvases") associated
' with each image loaded by the user.  At present, PD uses only a single canvas, on the main form.  This could change
' in the future.  This canvas is used to display the image on-screen, and its size is used to determine a number of
' viewport characteristics (such as whether or not scroll bars are needed to move around the image).
'
'Previously this module relied on internal VB measurements (like Form.ScaleWidth) when making viewport decisions.  With
' the fall '13 addition of floatable/dockable windows, and the full removal of MDI, it became necessary to rewrite much
' of this code against the program's new window manager (pdWindowManager class).  All calls to g_WindowManager relate
' to that class, which instead uses WAPI to return various window measurements.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Add an already-created pdImage object to the master pdImages() collection.  Do not pass empty objects!
Public Function AddImageToMasterCollection(ByRef srcImage As pdImage) As Boolean
    
    If Not (srcImage Is Nothing) Then
        
        If g_NumOfImagesLoaded > UBound(pdImages) Then
            ReDim Preserve pdImages(0 To g_NumOfImagesLoaded * 2 - 1) As pdImage
        End If
        
        Set pdImages(g_NumOfImagesLoaded) = srcImage
        
        'Activate the image and assign it a unique ID.  (IMPORTANT: at present, the ID always correlates to the
        ' image's position in the collection.  Do not change this behavior.)
        pdImages(g_NumOfImagesLoaded).IsActive = True
        pdImages(g_NumOfImagesLoaded).imageID = g_NumOfImagesLoaded
        
        'Newly loaded images are always auto-activated.
        g_CurrentImage = g_NumOfImagesLoaded
    
        'Track how many images we've loaded and/or currently have open
        g_NumOfImagesLoaded = g_NumOfImagesLoaded + 1
        g_OpenImageCount = g_OpenImageCount + 1
        
        AddImageToMasterCollection = True
        
    Else
        AddImageToMasterCollection = False
    End If
    
End Function

'Pass this function to obtain a default pdImage object, instantiated to match current UI settings and user preferences.
' Note that this function *does not touch* the main pdImages object, and as such, the created image will not yet have
' an imageID value.  That values is assigned when the object is added to the main pdImages() collection.
Public Sub GetDefaultPDImageObject(ByRef dstImage As pdImage)
    If (dstImage Is Nothing) Then Set dstImage = New pdImage
    dstImage.currentZoomValue = g_Zoom.GetZoom100Index
End Sub

'When loading an image file, there's a chance we won't be able to load the image correctly.  Because of that, we start
' with a "provisional" ID value for the image.  If the image fails to load, we can reuse this value on the next image.
Public Function GetProvisionalImageID() As Long
    GetProvisionalImageID = g_NumOfImagesLoaded
End Function

'Fit the current image onscreen at as large a size as possible (but never larger than 100% zoom)
Public Sub FitImageToViewport(Optional ByVal suppressRendering As Boolean = False)
    
    If g_OpenImageCount = 0 Then Exit Sub
    
    'Disable AutoScroll, because that messes with our calculations
    g_AllowViewportRendering = False
        
    'If the "fit all" zoom value is greater than 100%, use 100%.  Otherwise, use the "fit all" value as-is.
    Dim newZoomIndex As Long
    newZoomIndex = g_Zoom.GetZoomFitAllIndex
    
    If g_Zoom.GetZoomValue(newZoomIndex) > 1 Then newZoomIndex = g_Zoom.GetZoom100Index
    
    'Update the main canvas zoom drop-down, and the pdImage container for this image (so that zoom is restored properly when
    ' the user switches between loaded images).
    FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex = newZoomIndex
    pdImages(g_CurrentImage).currentZoomValue = newZoomIndex
    
    'Re-enable scrolling
    g_AllowViewportRendering = True
        
    'Now fix scrollbars and everything
    If (Not suppressRendering) Then Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_ResetToZero
    
    'Notify external UI elements of the change
    FormMain.mainCanvas(0).RelayViewportChanges
    
End Sub

'Fit the current image onscreen at as large a size as possible (including possibility of zoomed-in)
Public Sub FitOnScreen()
    
    If g_OpenImageCount = 0 Then Exit Sub
        
    'Disable AutoScroll, because that messes with our calculations
    g_AllowViewportRendering = False
    
    'Set zoom to the "fit whole" index
    FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex = g_Zoom.GetZoomFitAllIndex
    pdImages(g_CurrentImage).currentZoomValue = g_Zoom.GetZoomFitAllIndex
    
    'Re-enable scrolling
    g_AllowViewportRendering = True
        
    'Now fix scrollbars and everything
    Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_ResetToZero
    
    'Notify external UI elements of the change
    FormMain.mainCanvas(0).RelayViewportChanges
    
End Sub

'Center the current image onscreen without changing zoom
Public Sub CenterOnScreen(Optional ByVal suspendImmediateRedraw As Boolean = False)
    
    If g_OpenImageCount = 0 Then Exit Sub
        
    'Prevent the viewport from auto-updating on scroll bar events
    FormMain.mainCanvas(0).SetRedrawSuspension True
    
    'Set both canvas scrollbars to their midpoint
    FormMain.mainCanvas(0).SetScrollValue PD_HORIZONTAL, (FormMain.mainCanvas(0).GetScrollMin(PD_HORIZONTAL) + FormMain.mainCanvas(0).GetScrollMax(PD_HORIZONTAL)) / 2
    FormMain.mainCanvas(0).SetScrollValue PD_VERTICAL, (FormMain.mainCanvas(0).GetScrollMin(PD_VERTICAL) + FormMain.mainCanvas(0).GetScrollMax(PD_VERTICAL)) / 2
    
    'Re-enable scrolling
    FormMain.mainCanvas(0).SetRedrawSuspension False
        
    'Now fix scrollbars and everything
    If Not suspendImmediateRedraw Then Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    'Notify external UI elements of the change
    FormMain.mainCanvas(0).RelayViewportChanges
    
End Sub

'Previously, we could unload images by just unloading their containing form.  As image canvases are all custom-drawn now, this shortcut
' is no longer possible , so we must unload images using our own functions.
' (Note that this function simply wraps the imitation QueryUnload and Unload functions, below.)
'
'This function returns TRUE if the image was unloaded, FALSE if it was canceled.
Public Function FullPDImageUnload(ByVal imageID As Long, Optional ByVal redrawScreen As Boolean = True) As Boolean

    Dim toCancel As Integer
    Dim tmpUnloadMode As Integer
    
    'Perform a query unload on the image.  This will raise required warnings (e.g. unsaved changes) per the user's preferences.
    QueryUnloadPDImage toCancel, tmpUnloadMode, imageID
    
    If CBool(toCancel) Then
        FullPDImageUnload = False
        Exit Function
    End If
    
    UnloadPDImage toCancel, imageID, redrawScreen
    
    If CBool(toCancel) Then
        FullPDImageUnload = False
    Else
        
        'Redraw the screen
        If redrawScreen Then
        
            If g_OpenImageCount > 0 Then
                Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_ResetToCustom, pdImages(g_CurrentImage).imgViewport.getHScrollValue, pdImages(g_CurrentImage).imgViewport.getVScrollValue
            Else
                FormMain.mainCanvas(0).ClearCanvas
            End If
            
        End If
        
        FullPDImageUnload = True
    End If
    
    'If no images are open, take additional steps to free memory
    If g_OpenImageCount = 0 Then
        
        'Unload the backbuffer of the primary canvas
        Viewport_Engine.eraseViewportBuffers
        
        'Allow any tool panels to redraw themselves.  (Some tool panels dynamically change their contents based on the current image, so if no
        ' images are loaded, their contents may shift.)
        Tool_Support.syncToolOptionsUIToCurrentLayer
        
    End If
    
End Function

'Previously, we could unload images by just unloading their containing form.  Since PhotoDemon moved away from an MDI interface,
' this is no longer possible, so we must query unload images using this custom function.
Public Function QueryUnloadPDImage(ByRef Cancel As Integer, ByRef UnloadMode As Integer, ByVal imageID As Long) As Boolean

    Debug.Print "(Image #" & imageID & " received a Query_Unload trigger)"
    
    'Failsafe to make sure the image was properly initialized; if it wasn't, ignore this request entirely.
    If (imageID >= 0) And (imageID <= UBound(pdImages)) Then
        If pdImages(imageID) Is Nothing Then Exit Function
    Else
        Exit Function
    End If
    
    'If the user wants to be prompted about unsaved images, do it now
    If g_ConfirmClosingUnsaved And pdImages(imageID).IsActive Then
    
        'Check the .HasBeenSaved property of the image associated with this form
        If Not pdImages(imageID).getSaveState(pdSE_AnySave) Then
            
            'If the user hasn't already told us to deal with all unsaved images in the same fashion, run some checks
            If Not g_DealWithAllUnsavedImages Then
            
                g_NumOfUnsavedImages = 0
                                
                'Loop through all images to count how many unsaved images there are in total.
                ' NOTE: we only need to do this if the entire program is being shut down or if the user has selected "close all";
                ' otherwise, this close action only affects the current image, so we shouldn't present a "repeat for all images" option
                If g_ProgramShuttingDown Or g_ClosingAllImages Then
                    
                    Dim i As Long
                    For i = LBound(pdImages) To UBound(pdImages)
                        If Not (pdImages(i) Is Nothing) Then
                            If pdImages(i).IsActive And (Not pdImages(i).getSaveState(pdSE_AnySave)) Then
                                g_NumOfUnsavedImages = g_NumOfUnsavedImages + 1
                            End If
                        End If
                    Next i
                    
                End If
            
                'Before displaying the "do you want to save this image?" dialog, bring the image in question to the foreground.
                If imageID <> g_CurrentImage Then
                    If FormMain.Enabled Then ActivatePDImage imageID, "unsaved changes dialog required", True
                End If
                
                'Show the "do you want to save this image?" dialog. On that form, the number of unsaved images will be
                ' displayed and the user will be given an option to apply their choice to all unsaved images.
                Dim confirmReturn As VbMsgBoxResult
                confirmReturn = ConfirmClose(imageID)
                
            Else
                confirmReturn = g_HowToDealWithAllUnsavedImages
            End If
        
            'There are now three possible courses of action:
            ' 1) The user canceled. Quit and abandon all notion of closing.
            ' 2) The user asked us to save this image. Pass control to MenuSave (which will in turn call SaveAs if necessary)
            ' 3) The user doesn't give a shit. Exit without saving.
            
            'Cancel the close operation
            If confirmReturn = vbCancel Then
                
                Cancel = True
                If g_ProgramShuttingDown Then g_ProgramShuttingDown = False
                If g_ClosingAllImages Then g_ClosingAllImages = False
                g_DealWithAllUnsavedImages = False
                
            'Save the image
            ElseIf confirmReturn = vbYes Then
                
                'If the form being saved is enabled, bring that image to the foreground. (If a "Save As" is required, this
                ' helps show the user which image the Save As form is referencing.)
                If FormMain.Enabled Then ActivatePDImage imageID, "image being saved during shutdown", True
                
                'Attempt to save. Note that the user can still cancel at this point, and we want to honor their cancellation
                Dim saveSuccessful As Boolean
                saveSuccessful = MenuSave(pdImages(imageID))
                
                'If something went wrong, or the user canceled the save dialog, stop the unload process
                Cancel = Not saveSuccessful
 
                'If we make it here and the save was successful, force an immediate unload
                If Cancel Then
                    If g_ProgramShuttingDown Then g_ProgramShuttingDown = False
                    If g_ClosingAllImages Then g_ClosingAllImages = False
                    g_DealWithAllUnsavedImages = False
                End If
            
            'Do not save the image
            ElseIf confirmReturn = vbNo Then
                
                'No action is required here, because subsequent functions will take care of the rest of the unload process!
                
            End If
        
        End If
    
    End If

End Function

'Previously, we could unload images by just unloading their containing form.  This is no longer possible, so we must
' unload images using this special function.
Public Function UnloadPDImage(Cancel As Integer, ByVal imageIndex As Long, Optional ByVal resyncInterface As Boolean = True)

    'Failsafe to make sure the image was properly initialized
    If (pdImages(imageIndex) Is Nothing) Then Exit Function
    If pdImages(imageIndex).loadedSuccessfully And resyncInterface Then Message "Closing image..."
    
    'Decrease the open image count
    g_OpenImageCount = g_OpenImageCount - 1
    
    'Deactivate this DIB (note that this will take care of additional actions, like clearing the Undo/Redo cache
    ' for this image)
    pdImages(imageIndex).DeactivateImage
    
    'Remove this image from the thumbnail toolbar
    Interface.NotifyImageRemoved imageIndex, resyncInterface
    
    'Before exiting, restore focus to the next child window in line.  (But only if this image was the active window!)
    If g_CurrentImage = CLng(imageIndex) Then
    
        If g_OpenImageCount > 0 Then
        
            Dim i As Long
            i = Val(imageIndex) + 1
            If i > UBound(pdImages) Then i = i - 2
            
            Dim directionAscending As Boolean
            directionAscending = True
            
            Do While i >= 0
            
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).IsActive Then
                        ActivatePDImage i, "previous image unloaded", resyncInterface
                        Exit Do
                    End If
                End If
                
                If directionAscending Then
                    i = i + 1
                    If i > UBound(pdImages) Then
                        directionAscending = False
                        i = imageIndex
                    End If
                Else
                    i = i - 1
                End If
            
            Loop
        
        End If
        
    End If
    
    'Sync the interface to match the settings of whichever image is active (or disable a bunch of items if no images are active)
    If resyncInterface Then
        FormMain.mainCanvas(0).AlignCanvasView
        SyncInterfaceToCurrentImage
        Message "Finished."
    End If
    
End Function

'Previously, images could be activated by clicking on their window.  Now that all images are rendered to a single
' user control on the main form, we must activate them manually.
Public Sub ActivatePDImage(ByVal imageID As Long, Optional ByRef reasonForActivation As String = "", Optional ByVal refreshScreen As Boolean = True)

    'If this form is already the active image, don't waste time re-activating it
    If g_CurrentImage <> imageID Then
        
        'Update the current form variable
        g_CurrentImage = imageID
    
        'Because activation is an expensive process (requiring viewport redraws and more), I track the calls that access it.  This is used
        ' to minimize repeat calls as much as possible.
        Debug.Print "(Image #" & g_CurrentImage & " was activated because " & reasonForActivation & ")"
        
        'Double-check which monitor we are appearing on (for color management reasons)
        CheckParentMonitor True
        
        'Before displaying the form, redraw it, just in case something changed while it was deactivated (e.g. form resize)
        If Not (pdImages(g_CurrentImage) Is Nothing) Then
            
            If refreshScreen Then
            
                Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_ResetToCustom, pdImages(g_CurrentImage).imgViewport.getHScrollValue, pdImages(g_CurrentImage).imgViewport.getVScrollValue
                
                'This is ugly, but I'm working on a fix.  We need to restore the original scroll bar values, which we should
                ' really do by passing the values to the viewport in the previous step.  But I need to rework the whole
                ' way that damn function accepts parameters, so in the meantime, force the new values now.
                
                'TODO: fix this!
                
                'Reflow any image-window-specific chrome (status bar, rulers, etc)
                FormMain.mainCanvas(0).AlignCanvasView
            
                'Notify the thumbnail bar that a new image has been selected
                Interface.NotifyNewActiveImage g_CurrentImage
                
            End If
            
        End If
        
    End If
    
End Sub

'Find out whether the mouse pointer is over image contents or just the viewport
Public Function IsMouseOverImage(ByVal x1 As Long, ByVal y1 As Long, ByRef srcImage As pdImage) As Boolean

    If srcImage.imgViewport Is Nothing Then
        IsMouseOverImage = False
        Exit Function
    End If
    
    'Make sure the image is currently visible in the viewport
    If srcImage.imgViewport.getIntersectState Then
        
        'Remember: the imgViewport's intersection rect contains the intersection of the canvas and the image.
        ' If the target point lies inside this, it's over the image!
        Dim intRect As RECTF
        srcImage.imgViewport.getIntersectRectCanvas intRect
        IsMouseOverImage = Math_Functions.IsPointInRectF(x1, y1, intRect)
        
    Else
        IsMouseOverImage = False
    End If

End Function

'Find out whether the mouse pointer is over a given layer in an image
Public Function IsMouseOverLayer(ByVal imgX As Long, ByVal imgY As Long, ByRef srcImage As pdImage, ByRef srcLayerIndex As Long) As Boolean

    If srcImage.imgViewport Is Nothing Then
        IsMouseOverLayer = False
        Exit Function
    End If
    
    With srcImage.GetLayerByIndex(srcLayerIndex)
    
        If (imgX >= .getLayerOffsetX) And (imgX <= .getLayerOffsetX + .getLayerWidth(False)) Then
            If (imgY >= .getLayerOffsetY) And (imgY <= .getLayerOffsetY + .getLayerHeight(False)) Then
                IsMouseOverLayer = True
                Exit Function
            Else
                IsMouseOverLayer = False
            End If
            IsMouseOverLayer = False
        End If
    
    End With
    
End Function
