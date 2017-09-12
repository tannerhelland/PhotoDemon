Attribute VB_Name = "CanvasManager"
'***************************************************************************
'Image Canvas Handler (formerly Image Window Handler)
'Copyright 2002-2017 by Tanner Helland
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
    
    If (Not srcImage Is Nothing) Then
        
        Set pdImages(g_NumOfImagesLoaded) = srcImage
        
        'Activate the image and assign it a unique ID.  (IMPORTANT: at present, the ID always correlates to the
        ' image's position in the collection.  Do not change this behavior.)
        pdImages(g_NumOfImagesLoaded).ChangeActiveState True
        pdImages(g_NumOfImagesLoaded).imageID = g_NumOfImagesLoaded
        
        'Newly loaded images are always auto-activated.
        g_CurrentImage = g_NumOfImagesLoaded
    
        'Track how many images we've loaded and/or currently have open
        g_NumOfImagesLoaded = g_NumOfImagesLoaded + 1
        g_OpenImageCount = g_OpenImageCount + 1
        
        If (g_NumOfImagesLoaded > UBound(pdImages)) Then
            ReDim Preserve pdImages(0 To g_NumOfImagesLoaded * 2 - 1) As pdImage
        End If
        
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
    dstImage.SetZoom g_Zoom.GetZoom100Index
End Sub

'When loading an image file, there's a chance we won't be able to load the image correctly.  Because of that, we start
' with a "provisional" ID value for the image.  If the image fails to load, we can reuse this value on the next image.
Public Function GetProvisionalImageID() As Long
    GetProvisionalImageID = g_NumOfImagesLoaded
End Function

'Fit the current image onscreen at as large a size as possible (but never larger than 100% zoom)
Public Sub FitImageToViewport(Optional ByVal suppressRendering As Boolean = False)
    
    If (g_OpenImageCount <> 0) Then
    
        ViewportEngine.DisableRendering
            
        'If the "fit all" zoom value is greater than 100%, use 100%.  Otherwise, use the "fit all" value as-is.
        Dim newZoomIndex As Long
        newZoomIndex = g_Zoom.GetZoomFitAllIndex
        
        If (g_Zoom.GetZoomValue(newZoomIndex) > 1) Then newZoomIndex = g_Zoom.GetZoom100Index
        
        'Update the main canvas zoom drop-down, and the pdImage container for this image (so that zoom is restored properly when
        ' the user switches between loaded images).
        FormMain.mainCanvas(0).SetZoomDropDownIndex newZoomIndex
        pdImages(g_CurrentImage).SetZoom newZoomIndex
        
        'Re-enable scrolling
        ViewportEngine.EnableRendering
            
        'Now fix scrollbars and everything
        If (Not suppressRendering) Then ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_ResetToZero
        
        'Notify external UI elements of the change
        FormMain.mainCanvas(0).RelayViewportChanges
    
    End If

End Sub

'Fit the current image onscreen at as large a size as possible (including possibility of zoomed-in)
Public Sub FitOnScreen()
    
    If (g_OpenImageCount <> 0) Then
        
        ViewportEngine.DisableRendering
        
        'Set zoom to the "fit whole" index
        FormMain.mainCanvas(0).SetZoomDropDownIndex g_Zoom.GetZoomFitAllIndex
        pdImages(g_CurrentImage).SetZoom g_Zoom.GetZoomFitAllIndex
        
        'Re-enable scrolling
        ViewportEngine.EnableRendering
            
        'Now fix scrollbars and everything
        ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_ResetToZero
        
        'Notify external UI elements of the change
        FormMain.mainCanvas(0).RelayViewportChanges
        
    End If
    
End Sub

'Center the current image onscreen without changing zoom
Public Sub CenterOnScreen(Optional ByVal suspendImmediateRedraw As Boolean = False)
    
    If (g_OpenImageCount <> 0) Then
            
        'Prevent the viewport from auto-updating on scroll bar events
        FormMain.mainCanvas(0).SetRedrawSuspension True
        
        'Set both canvas scrollbars to their midpoint
        FormMain.mainCanvas(0).SetScrollValue PD_HORIZONTAL, (FormMain.mainCanvas(0).GetScrollMin(PD_HORIZONTAL) + FormMain.mainCanvas(0).GetScrollMax(PD_HORIZONTAL)) / 2
        FormMain.mainCanvas(0).SetScrollValue PD_VERTICAL, (FormMain.mainCanvas(0).GetScrollMin(PD_VERTICAL) + FormMain.mainCanvas(0).GetScrollMax(PD_VERTICAL)) / 2
        
        'Re-enable scrolling
        FormMain.mainCanvas(0).SetRedrawSuspension False
            
        'Now fix scrollbars and everything
        If (Not suspendImmediateRedraw) Then ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        'Notify external UI elements of the change
        FormMain.mainCanvas(0).RelayViewportChanges
        
    End If
        
End Sub

'Previously, we could unload images by just unloading their containing form.  As image canvases are all custom-drawn now, this shortcut
' is no longer possible , so we must unload images using our own functions.
' (Note that this function simply wraps the imitation QueryUnload and Unload functions, below.)
'
'This function returns TRUE if the image was unloaded, FALSE if it was canceled.
Public Function FullPDImageUnload(ByVal imageID As Long, Optional ByVal redrawScreen As Boolean = True) As Boolean

    Dim userCanceledUnload As Boolean: userCanceledUnload = False
    
    'Perform a query unload on the image.  This will raise required warnings (e.g. unsaved changes) per the user's preferences.
    CanvasManager.QueryUnloadPDImage userCanceledUnload, imageID
    
    'If the "save unsaved changes" dialog was canceled by the user, abandon any further unloading
    If userCanceledUnload Then
        FullPDImageUnload = False
    
    'The user is allowing the unload to proceed
    Else
    
        CanvasManager.UnloadPDImage imageID, redrawScreen
            
        'Redraw the screen to reflect any newly initialized images
        If redrawScreen Then
        
            If (g_OpenImageCount > 0) Then
                ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_ResetToCustom, pdImages(g_CurrentImage).ImgViewport.GetHScrollValue, pdImages(g_CurrentImage).ImgViewport.GetVScrollValue
            Else
                FormMain.mainCanvas(0).ClearCanvas
            End If
            
        End If
        
        FullPDImageUnload = True
        
        'If no images are open, take additional steps to free memory
        If (g_OpenImageCount = 0) And (Macros.GetMacroStatus <> MacroBATCH) Then
            
            'Unload the backbuffer of the primary canvas
            ViewportEngine.EraseViewportBuffers
            
            'Allow any tool panels to redraw themselves.  (Some tool panels dynamically change their contents based on the current image, so if no
            ' images are loaded, their contents may shift.)
            Tools.SyncToolOptionsUIToCurrentLayer
            
            'Release any cached Undo/Redo writers
            Saving.FreeUpMemory
            
        End If
    
    End If
    
End Function

'Previously, we could unload images by just unloading their containing form.  Since PhotoDemon moved away from an MDI interface,
' this is no longer possible, so we must query unload images using this custom function.
Public Function QueryUnloadPDImage(ByRef userCanceledUnload As Boolean, ByVal imageID As Long) As Boolean
    
    'Perform a few failsafe checks to make sure the current image was properly initialized
    Dim okayToQueryUnload As Boolean: okayToQueryUnload = True
    If (imageID < 0) Then okayToQueryUnload = False
    If (imageID > UBound(pdImages)) Then okayToQueryUnload = False
    If (pdImages(imageID) Is Nothing) Then okayToQueryUnload = False
    
    'Also, disable save prompts during batch processes
    If (Macros.GetMacroStatus = MacroBATCH) Then okayToQueryUnload = False
    
    If okayToQueryUnload Then
    
        'If the user wants to be prompted about unsaved images, do it now
        If (g_ConfirmClosingUnsaved And pdImages(imageID).IsActive) Then
       
           'Check the .HasBeenSaved property of the image associated with this form
           If (Not pdImages(imageID).GetSaveState(pdSE_AnySave)) Then
               
               'If we reach this line, the image in question has unsaved changes.
               
               'If the user hasn't already told us to deal with all unsaved images in the same fashion, run some checks
               If (Not g_DealWithAllUnsavedImages) Then
               
                   g_NumOfUnsavedImages = 0
                                   
                   'Loop through all open images to count how many unsaved images there are in total.
                   ' NOTE: we only need to do this if the entire program is being shut down or if the user has selected "close all";
                   ' otherwise, this close action only affects the current image, so we shouldn't present a "repeat for all images" option
                   If (g_ProgramShuttingDown Or g_ClosingAllImages) Then
                       
                       Dim i As Long
                       For i = LBound(pdImages) To UBound(pdImages)
                           If (Not pdImages(i) Is Nothing) Then
                               If pdImages(i).IsActive And (Not pdImages(i).GetSaveState(pdSE_AnySave)) Then
                                   g_NumOfUnsavedImages = g_NumOfUnsavedImages + 1
                               End If
                           End If
                       Next i
                       
                   End If
                   
                   'Before displaying the "do you want to save this image?" dialog, bring the image in question to the foreground.
                   If (imageID <> g_CurrentImage) Then
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
               ' 1) The user canceled the "unsaved changes" dialog.  Abandon all notion of closing this image (or the program).
               ' 2) The user asked us to save before exiting. Pass control to MenuSave (which will in turn call SaveAs if necessary).
               ' 3) The user doesn't care about saving changes.  Exit as-is.
               
               'Cancel the close operation
               If (confirmReturn = vbCancel) Then
                   
                   userCanceledUnload = True
                   If g_ProgramShuttingDown Then g_ProgramShuttingDown = False
                   If g_ClosingAllImages Then g_ClosingAllImages = False
                   g_DealWithAllUnsavedImages = False
                   
               'Save all unsaved images
               ElseIf (confirmReturn = vbYes) Then
                   
                   'If the form being saved is enabled, bring that image to the foreground. (If a "Save As" is required, this
                   ' helps show the user which image the Save As form is referencing.)
                   If FormMain.Enabled Then ActivatePDImage imageID, "image being saved during shutdown", True
                   
                   'Attempt to save. Note that the user can still cancel at this point, and we want to honor their cancellation
                   Dim saveSuccessful As Boolean
                   saveSuccessful = MenuSave(pdImages(imageID))
                   
                   'If something went wrong, or the user canceled the save dialog, stop the unload process
                   userCanceledUnload = (Not saveSuccessful)
    
                   'If we make it here and the save was successful, force an immediate unload
                   If userCanceledUnload Then
                       If g_ProgramShuttingDown Then g_ProgramShuttingDown = False
                       If g_ClosingAllImages Then g_ClosingAllImages = False
                       g_DealWithAllUnsavedImages = False
                   End If
               
               'Do not save the image
               ElseIf (confirmReturn = vbNo) Then
                   
                   'No action is required here, because subsequent functions will take care of the rest of the unload process!
                   
               End If
           
           End If
       
       End If
       
    End If

End Function

'Previously, we could unload images by just unloading their containing form.  This is no longer possible, so we must
' unload images using this special function.
Public Function UnloadPDImage(ByVal imageIndex As Long, Optional ByVal resyncInterface As Boolean = True)

    'Failsafe to make sure the image was properly initialized
    If (pdImages(imageIndex) Is Nothing) Then Exit Function
    If (pdImages(imageIndex).IsActive And resyncInterface) Then Message "Closing image..."
    
    'Decrease the open image count
    g_OpenImageCount = g_OpenImageCount - 1
    
    'Deactivate this DIB (note that this will take care of additional actions, like clearing the Undo/Redo cache
    ' for this image)
    pdImages(imageIndex).FreeAllImageResources
    
    'Remove this image from the thumbnail toolbar.  Note that we deliberately ignore the "resyncInterface" setting here,
    ' because the tabstrip needs to always reflect the currently active image.  (Also, it redraws quite quickly, so there's
    ' a minimal performance impact either way.)
    '
    'If we don't do this, it gets pretty confusing when a "Save" dialog is raised, but the filename doesn't match the
    ' currently active image in the tabstrip.
    Interface.NotifyImageRemoved imageIndex, True
    
    'Before exiting, restore focus to the next child window in line.  (But only if this image was the active window!)
    If (g_CurrentImage = CLng(imageIndex)) Then
    
        If (g_OpenImageCount > 0) Then
            
            'Figure out the next image that should receive focus.  If the image we're closing is the last one in line, move to
            ' the next-to-last one in line (instead of advancing forward, which is obviously not possible).
            Dim i As Long
            i = imageIndex + 1
            If (i > UBound(pdImages)) Then i = i - 2
            
            'Search through the image list until we find a valid image candidate to receive focus
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
                    If (i > UBound(pdImages)) Then
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
        Interface.SyncInterfaceToCurrentImage
        Message "Finished."
    End If
    
End Function

'Previously, images could be activated by clicking on their window.  Now that all images are rendered to a single
' user control on the main form, we must activate them manually.
Public Sub ActivatePDImage(ByVal imageID As Long, Optional ByRef reasonForActivation As String = "", Optional ByVal refreshScreen As Boolean = True, Optional ByVal associatedUndoType As PD_UndoType = UNDO_EVERYTHING)

    Dim startTime As Currency
    VBHacks.GetHighResTime startTime

    'If this form is already the active image, don't waste time re-activating it
    If (g_CurrentImage <> imageID) Then
        
        'Release some temporary resources on the old image, if we can
        pdImages(g_CurrentImage).DeactivateImage
        
        'Update the current form variable
        g_CurrentImage = imageID
    
        'Because activation is an expensive process (requiring viewport redraws and more), I track the calls that access it.  This is used
        ' to minimize repeat calls as much as possible.
        Debug.Print "(Image #" & g_CurrentImage & " was activated because " & reasonForActivation & ")"
        
        'Double-check which monitor we are appearing on (for color management reasons)
        ColorManagement.CheckParentMonitor True
        
    End If
    
    'Before displaying the form, redraw it, just in case something changed while it was deactivated (e.g. form resize)
    If (Not pdImages(g_CurrentImage) Is Nothing) And refreshScreen Then
        
        If (associatedUndoType = UNDO_EVERYTHING) Or (associatedUndoType = UNDO_IMAGE) Or (associatedUndoType = UNDO_IMAGE_VECTORSAFE) Or (associatedUndoType = UNDO_IMAGEHEADER) Then
            
            ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_ResetToCustom, pdImages(g_CurrentImage).ImgViewport.GetHScrollValue, pdImages(g_CurrentImage).ImgViewport.GetVScrollValue
            
            'Reflow any image-window-specific chrome (status bar, rulers, etc)
            FormMain.mainCanvas(0).AlignCanvasView
            
        Else
            ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0), poi_ReuseLast
        End If
        
        'Run the main SyncInterfaceToImage function, and notify a few peripheral functions of the updated image
        ' (e.g. updating thumbnails, window captions, etc)
        Interface.NotifyNewActiveImage g_CurrentImage
        
    End If
    
    'Make sure any tool initializations that vary by image are up-to-date.  (This includes things like
    ' making sure a scratch layer exists, and that it matches the current image's size.)
    Tools.InitializeToolsDependentOnImage
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "CanvasManager.ActivatePDImage finished in " & VBHacks.GetTimeDiffNowAsString(startTime)
    #End If
    
End Sub

'Find out whether the mouse pointer is over image contents or just the viewport
Public Function IsMouseOverImage(ByVal x1 As Long, ByVal y1 As Long, ByRef srcImage As pdImage) As Boolean

    If (srcImage.ImgViewport Is Nothing) Then
        IsMouseOverImage = False
        Exit Function
    End If
    
    'Make sure the image is currently visible in the viewport
    If srcImage.ImgViewport.GetIntersectState Then
        
        'Remember: the imgViewport's intersection rect contains the intersection of the canvas and the image.
        ' If the target point lies inside this, it's over the image!
        Dim intRect As RECTF
        srcImage.ImgViewport.GetIntersectRectCanvas intRect
        IsMouseOverImage = PDMath.IsPointInRectF(x1, y1, intRect)
        
    Else
        IsMouseOverImage = False
    End If

End Function

'Find out whether the mouse pointer is over a given layer in an image
Public Function IsMouseOverLayer(ByVal imgX As Long, ByVal imgY As Long, ByRef srcImage As pdImage, ByRef srcLayerIndex As Long) As Boolean

    If srcImage.ImgViewport Is Nothing Then
        IsMouseOverLayer = False
        Exit Function
    End If
    
    With srcImage.GetLayerByIndex(srcLayerIndex)
    
        If (imgX >= .GetLayerOffsetX) And (imgX <= .GetLayerOffsetX + .GetLayerWidth(False)) Then
            If (imgY >= .GetLayerOffsetY) And (imgY <= .GetLayerOffsetY + .GetLayerHeight(False)) Then
                IsMouseOverLayer = True
                Exit Function
            Else
                IsMouseOverLayer = False
            End If
            IsMouseOverLayer = False
        End If
    
    End With
    
End Function
