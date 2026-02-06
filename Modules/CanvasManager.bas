Attribute VB_Name = "CanvasManager"
'***************************************************************************
'Image Canvas Handler (formerly Image Window Handler)
'Copyright 2002-2026 by Tanner Helland
'Created: 11/29/02
'Last updated: 18/October/21
'Last update: rework image unload wrappers to support new "automatic session restore after reboot" preference
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Fit the current image onscreen at as large a size as possible (but never larger than 100% zoom)
Public Sub FitImageToViewport(Optional ByVal suppressRendering As Boolean = False)
    
    If PDImages.IsImageActive() Then
    
        Viewport.DisableRendering
            
        'If the "fit all" zoom value is greater than 100%, use 100%.  Otherwise, use the "fit all" value as-is.
        Dim newZoomIndex As Long
        newZoomIndex = Zoom.GetZoomFitAllIndex
        
        If (Zoom.GetZoomRatioFromIndex(newZoomIndex) > 1#) Then newZoomIndex = Zoom.GetZoom100Index
        
        'Update the main canvas zoom drop-down, and the pdImage container for this image (so that zoom is restored properly when
        ' the user switches between loaded images).
        FormMain.MainCanvas(0).SetZoomDropDownIndex newZoomIndex
        PDImages.GetActiveImage.SetZoomIndex newZoomIndex
        
        'Re-enable scrolling
        Viewport.EnableRendering
            
        'Now fix scrollbars and everything
        If (Not suppressRendering) Then Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0), VSR_ResetToZero
        
        'Notify external UI elements of the change
        Viewport.NotifyEveryoneOfViewportChanges
    
    End If

End Sub

'Fit the current image onscreen at as large a size as possible (including possibility of zoomed-in)
Public Sub FitOnScreen()
    
    If PDImages.IsImageActive() Then
        
        Viewport.DisableRendering
        
        'Set zoom to the "fit whole" index
        FormMain.MainCanvas(0).SetZoomDropDownIndex Zoom.GetZoomFitAllIndex
        PDImages.GetActiveImage.SetZoomIndex Zoom.GetZoomFitAllIndex
        
        'Re-enable scrolling
        Viewport.EnableRendering
            
        'Now fix scrollbars and everything
        Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0), VSR_ResetToZero
        
        'Notify external UI elements of the change
        Viewport.NotifyEveryoneOfViewportChanges
        
    End If
    
End Sub

'Center the current image onscreen without changing zoom
Public Sub CenterOnScreen(Optional ByVal suspendImmediateRedraw As Boolean = False)
    
    If PDImages.IsImageActive() Then
            
        'Prevent the viewport from auto-updating on scroll bar events
        FormMain.MainCanvas(0).SetRedrawSuspension True
        
        'Set both canvas scrollbars to their midpoint
        FormMain.MainCanvas(0).SetScrollValue pdo_Horizontal, (FormMain.MainCanvas(0).GetScrollMin(pdo_Horizontal) + FormMain.MainCanvas(0).GetScrollMax(pdo_Horizontal)) / 2
        FormMain.MainCanvas(0).SetScrollValue pdo_Vertical, (FormMain.MainCanvas(0).GetScrollMin(pdo_Vertical) + FormMain.MainCanvas(0).GetScrollMax(pdo_Vertical)) / 2
        
        'Re-enable scrolling
        FormMain.MainCanvas(0).SetRedrawSuspension False
            
        'Now fix scrollbars and everything
        If (Not suspendImmediateRedraw) Then Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
        'Notify external UI elements of the change
        Viewport.NotifyEveryoneOfViewportChanges
        
    End If
        
End Sub

'Attempt to close all open images.  This is inovked by File > Close All, or by exiting PD while multiple images are open.
' IMPORTANT NOTE: you need to deal with the return value of this function.  If the return value is TRUE, all images were
' unloaded successfully - but if this returns FALSE, it means the user interrupted the unload process.  (This can happen
' if an image has unsaved changes, but the user cancels the Save dialog.)  You may need to sync on-screen UI elements
' afetr the process terminates; this is deliberately *not* handled by this function, as we don't care about syncing
' on-screen elements under certain circumstances (e.g. PD is shutting down).
Public Function CloseAllImages() As Boolean
    
    'Assume success; specific failure states will change this to FALSE before exiting
    CloseAllImages = True
    
    'Note that we are attempting to close all images
    g_ClosingAllImages = True
    
    'We now branch according to an interesting condition.  If the user initiated this shutdown,
    ' we'll close normally - first by closing all images with saved changes, then by prompting
    ' for any unsaved changes (and continuing/canceling accordingly).
    '
    '*But* if this shutdown request came from the system (e.g. from Windows rebooting due to a
    ' system update or similar), then we need to check a user preference.  If the user allows us
    ' to automatically pick up where we left off after a reboot, we don't need to prompt -
    ' just shut down and rely on the AutoSave engine to restore everything for us.  But if that
    ' option is disabled, we need to (attempt to) prevent shutdown by raising a modal dialog
    ' for unsaved changes.
    Dim useNormalShutdownBehavior As Boolean: useNormalShutdownBehavior = True
    If (Not g_ThunderMain Is Nothing) Then
        
        If g_ThunderMain.WasEndSessionReceived(True) Then
            
            'This is a system-initiated shutdown.  Check the auto-start-after-reboot preference.
            If UserPrefs.GetPref_Boolean("Loading", "RestoreAfterReboot", False) Then
                
                'Auto-start-after-reboot is enabled.  Skip shutdown dialogs and leave the
                ' entire session as-is, so the AutoSave engine can flag it accordingly.
                useNormalShutdownBehavior = False
                
            End If
            
        End If
        
    End If
    
    Dim listOfOpenImages As pdStack
    Dim i As Long, tmpImageID As Long
    
    If PDImages.GetListOfActiveImageIDs(listOfOpenImages) Then
        
        If useNormalShutdownBehavior Then
            
            'We are now going to close images in a somewhat strange fashion (but one that improves performance).
            
            'First, figure out how many images need to be closed.  (We need this number so that we can display
            ' progress reports to the user.)
            Dim numImagesToClose As Long
            numImagesToClose = listOfOpenImages.GetNumOfInts()
            
            Dim numImagesClosed As Long
            numImagesClosed = 0
            
            'Next, unload all images *without* unsaved changes.  These images don't require shutdown prompts,
            ' so we can unload them without consequence or user intervention.
            For i = 0 To listOfOpenImages.GetNumOfInts - 1
                tmpImageID = listOfOpenImages.GetInt(i)
                If PDImages.GetImageByID(tmpImageID).GetSaveState(pdSE_AnySave) Then
                    numImagesClosed = numImagesClosed + 1
                    Message "Unloading image %1 of %2", numImagesClosed, numImagesToClose
                    CanvasManager.FullPDImageUnload tmpImageID, False
                End If
            Next i
            
            'If the above step unloaded one or more images, we need to forcibly redraw the image tabstrip.
            ' (If we don't, and the user cancels the "unsaved changes" dialog we are about to raise,
            ' the tabpstrip will display an out-of-date list of open images.)
            If (numImagesClosed > 0) Then Interface.RequestTabstripRedraw False
            
            'The only images still open (if any) are ones with unsaved changes.  Starting with the currently
            ' active image, unload each remaining image in turn.
            Do While (PDImages.GetNumOpenImages() > 0)
                
                numImagesClosed = numImagesClosed + 1
                Message "Unloading image %1 of %2", numImagesClosed, numImagesToClose
                
                'Attempt to unload the currently active image.
                ' (NOTE: this function returns a boolean saying whether the image was successfully unloaded,
                '        but for this fringe case, we ignore it in favor of checking g_ProgramShuttingDown.)
                CanvasManager.FullPDImageUnload PDImages.GetActiveImageID(), False
                
                'If the "unsaved changes" prompt canceled shut down for some reason, it will reset the
                ' g_ClosingAllImages variable.  Read that variable and use it to determine whether we
                ' are allowed to continue closing images.
                If (Not g_ClosingAllImages) Then
                    CloseAllImages = False
                    Exit Function
                End If
                
            Loop
            
        'Non-standard closing behavior.  Immediately clear every pdImage object.
        Else
            
            PDDebug.LogAction "Session restart pending; dumping pdImage objects ASAP..."
            
            For i = 0 To listOfOpenImages.GetNumOfInts - 1
                tmpImageID = listOfOpenImages.GetInt(i)
                CanvasManager.UnloadPDImage tmpImageID, False
            Next i
            
            CloseAllImages = True
            
        End If
        
    'No open images
    Else
        CloseAllImages = True
    End If
    
End Function

'Previously, we could unload images by just unloading their containing form.  As image canvases are all custom-drawn now, this shortcut
' is no longer possible , so we must unload images using our own functions.
' (Note that this function simply wraps the imitation QueryUnload and Unload functions, below.)
'
'This function returns TRUE if the image was unloaded, FALSE if it was canceled.
Public Function FullPDImageUnload(ByVal imageID As Long, Optional ByVal displayMessages As Boolean = True) As Boolean

    'Perform a query unload on the image.  This will raise required warnings (e.g. unsaved changes) per the user's preferences.
    If CanvasManager.QueryUnloadPDImage(imageID) Then
        
        'The user is allowing the unload to proceed.  Unload the image in question, and note that this function will
        ' also handle the messy process of updating the UI to match (including activating the next image in line).
        CanvasManager.UnloadPDImage imageID, displayMessages
        
        'If we have just closed the final open image in the program, take additional steps to free memory
        If (PDImages.GetNumOpenImages() = 0) And (Macros.GetMacroStatus <> MacroBATCH) Then
            
            'Unload the backbuffer of the primary canvas
            Viewport.EraseViewportBuffers
            
            'Allow any tool panels to redraw themselves.  (Some tool panels dynamically change their contents based
            ' on the current image, so if no images are loaded, their contents may shift.)
            Tools.SyncToolOptionsUIToCurrentLayer
            
            'Release any cached Undo/Redo writers
            Saving.FreeUpMemory
            
        End If
        
        FullPDImageUnload = True
    
    'The "save unsaved changes" dialog was canceled by the user.  Abandon any further unloading.
    Else
        FullPDImageUnload = False
    End If
    
End Function

'Previously, we could unload images by just unloading their containing form.  Since PhotoDemon moved away from an MDI interface,
' this is no longer possible, so we must query unload images using this custom function.
'RETURNS: TRUE if the image was unloaded successfully; FALSE if the unload process was interrupted by the user
Public Function QueryUnloadPDImage(ByVal imageID As Long) As Boolean
    
    QueryUnloadPDImage = True
    
    'Perform a few failsafe checks to make sure the current image was properly initialized
    Dim okayToQueryUnload As Boolean: okayToQueryUnload = True
    If (Not PDImages.IsImageActive(imageID)) Then okayToQueryUnload = False
    
    'Also, disable save prompts during batch processes
    If okayToQueryUnload Then okayToQueryUnload = Not (Macros.GetMacroStatus = MacroBATCH)
    
    'If we are allowed to present a UI to the user, do so now
    If okayToQueryUnload Then
    
        'If the user wants to be prompted about unsaved images, do it now
        If (g_ConfirmClosingUnsaved And PDImages.GetImageByID(imageID).IsActive) Then
        
            'Check the .HasBeenSaved property of the image associated with this form
            If (Not PDImages.GetImageByID(imageID).GetSaveState(pdSE_AnySave)) Then
               
                'If we reach this line, the image in question has unsaved changes.
               
                'If the user hasn't already told us to deal with all unsaved images in the same fashion, run some checks
                If (Not g_DealWithAllUnsavedImages) Then
                    
                    Dim numOfUnsavedImages As Long
                    numOfUnsavedImages = 0
                    
                    Dim imageIndices As pdStack
                    Set imageIndices = New pdStack
                    
                    'We also want to record a list of the unsaved image's ID values.  (This is required
                    ' to generate an interactive "images with unsaved changes" window, where the user can
                    ' browse through a list of ALL images with unsaved changes.)
                    Dim listOfOpenImages As pdStack
                    If PDImages.GetListOfActiveImageIDs(listOfOpenImages) Then
                    
                        'Next, from the list of "open images", we want to we want to pare down the list
                        ' to remove any images *without* unsaved changes.  (Such images can be blindly
                        ' closed without consequence.)
                        
                        'NOTE: we only do this if the entire program is being shut down or if the user has
                        ' selected "close all"; otherwise, this close request only affects the current image,
                        ' so we shouldn't present a "repeat this action for all images" dialog option.
                        If (g_ProgramShuttingDown Or g_ClosingAllImages) Then
                            
                            Dim tmpImageID As Long
                            Do While listOfOpenImages.PopInt(tmpImageID)
                                If (Not PDImages.GetImageByID(tmpImageID).GetSaveState(pdSE_AnySave)) Then
                                    numOfUnsavedImages = numOfUnsavedImages + 1
                                    imageIndices.AddInt tmpImageID
                                End If
                            Loop
                            
                        End If
                    
                    'We do not need an (else) branch here, as this block of code requires there to always
                    ' be at least one image with unsaved changes (the current image); otherwise, we would
                    ' not be inside this block in the first place.  Said another way, this If() branch is
                    ' entirely overkill at present.
                    End If
                    
                    'Show the "do you want to save this image?" dialog. On that form, the number of unsaved images will be
                    ' displayed and the user will be given an option to apply their choice to all unsaved images.
                    Dim confirmReturn As VbMsgBoxResult
                    confirmReturn = Dialogs.ConfirmClose(imageID, numOfUnsavedImages, imageIndices)
                    
                Else
                    confirmReturn = g_HowToDealWithAllUnsavedImages
                End If
                
                'There are now three possible courses of action:
                ' 1) The user canceled the "unsaved changes" dialog.  Abandon all notion of closing this image (or the program).
                ' 2) The user asked us to save before exiting. Pass control to MenuSave (which will in turn call SaveAs if necessary).
                ' 3) The user doesn't care about saving changes.  Exit as-is.
                
                'Cancel the close operation
                If (confirmReturn = vbCancel) Then
                    
                    QueryUnloadPDImage = False
                    g_ProgramShuttingDown = False
                    g_ClosingAllImages = False
                    g_DealWithAllUnsavedImages = False
                    
                    'As a failsafe, also reset flags for any system-originating shutdown messages
                    If (Not g_ThunderMain Is Nothing) Then g_ThunderMain.ResetEndSessionFlags
                       
                'Save all unsaved images
                ElseIf (confirmReturn = vbYes) Then
                   
                    'If the form being saved is enabled, bring that image to the foreground. (If a "Save As" is required, this
                    ' helps show the user which image the Save As form is referencing.)
                    If FormMain.Enabled Then CanvasManager.ActivatePDImage imageID, "image being saved during shutdown", True
                    
                    'Attempt to save. Note that the user can still cancel at this point, and we want to honor their cancellation
                    QueryUnloadPDImage = FileMenu.MenuSave(PDImages.GetImageByID(imageID))
                    
                    'If something went wrong, or the user canceled the save dialog, stop the unload process
                    If (Not QueryUnloadPDImage) Then
                        g_ProgramShuttingDown = False
                        g_ClosingAllImages = False
                        g_DealWithAllUnsavedImages = False
                    End If
               
                'Do not save the image.
                ElseIf (confirmReturn = vbNo) Then
                    'No action is required here, because subsequent functions will take care of the rest of the unload process!
                End If
           
           'If the image does not have any unsaved changes, we can always close it successfully!
           End If
       
       End If
       
    End If

End Function

'Unload a user image and perform all required UI updates (e.g. selecting a new "active" image if multiple
' images have been loaded).  Do not call this function directly; instead, call QueryUnloadPDImage(), above,
' which will prompt the user for things like unsaved changes.
Public Sub UnloadPDImage(ByVal imageIndex As Long, Optional ByVal displayMessages As Boolean = True)

    'Failsafes to make sure the image was properly initialized
    If (Not PDImages.IsImageActive(imageIndex)) Then Exit Sub
    
    If displayMessages Then Message "Closing image..."
    
    'Remove this image from the thumbnail toolbar, and explicitly ask it to *not* repaint itself.  (It will repaint
    ' automatically later in this function.)
    Interface.NotifyImageRemoved imageIndex, False
    
    'Ask the central image collection to free resources associated with the target image
    PDImages.RemovePDImageFromCollection imageIndex
    
    'If this image was also the active canvas, activate the next image in line (if any others exist).
    If (PDImages.GetNumOpenImages > 0) And (PDImages.GetActiveImageID() = imageIndex) Then
        
        Dim imgCollectionSize As Long
        imgCollectionSize = PDImages.GetImageCollectionSize()
        
        'Figure out the next image that should receive focus.  If the image we're closing is the last one in line, move to
        ' the next-to-last one in line (instead of advancing forward, which is obviously not possible).
        Dim i As Long
        i = imageIndex + 1
        If (i > imgCollectionSize) Then i = i - 2
        
        'Search through the image list until we find a valid image candidate to receive focus
        Dim directionAscending As Boolean
        directionAscending = True
        
        Do While (i >= 0)
        
            If PDImages.IsImageActive(i) Then
                CanvasManager.ActivatePDImage i, "previous image unloaded", True
                Exit Do
            End If
            
            If directionAscending Then
                i = i + 1
                If (i > imgCollectionSize) Then
                    directionAscending = False
                    i = imageIndex
                End If
            Else
                i = i - 1
            End If
        
        Loop
        
    'If no more images are open, clear the UI to match
    Else
        
        'Sync the interface to match the settings of whichever image is active (or disable a bunch of items if no images are active)
        FormMain.MainCanvas(0).AlignCanvasView
        Interface.SyncInterfaceToCurrentImage
        
    End If
    
    If displayMessages Then Message "Finished."
    
End Sub

'Previously, images could be activated by clicking on their window.  Now that all images are rendered to a single
' user control on the main form, we must activate them manually.
Public Sub ActivatePDImage(ByVal imageID As Long, Optional ByRef reasonForActivation As String = vbNullString, Optional ByVal refreshScreen As Boolean = True, Optional ByVal associatedUndoType As PD_UndoType = UNDO_Everything, Optional ByVal newImageJustLoaded As Boolean = False)

    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'If this form is already the active image, don't waste time re-activating it
    Dim activeImageChanging As Boolean
    activeImageChanging = (PDImages.GetActiveImageID() <> imageID) Or newImageJustLoaded
    
    If activeImageChanging Then
        
        'Release some temporary resources on the old image, if we can
        PDImages.GetActiveImage.DeactivateImage True
        UIImages.FreeSharedCompressBuffer
        
        'Update the current form variable
        PDImages.SetActiveImageID imageID
        
        'Double-check which monitor we are appearing on (for color management reasons)
        ColorManagement.CheckParentMonitor True
        
    End If
    
    'Before displaying the form, redraw it, just in case something changed while it was deactivated (e.g. form resize)
    If (PDImages.IsImageActive() And refreshScreen) Then
        
        If (associatedUndoType = UNDO_Everything) Or (associatedUndoType = UNDO_Image) Or (associatedUndoType = UNDO_Image_VectorSafe) Or (associatedUndoType = UNDO_ImageHeader) Then
            
            Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0), VSR_ResetToCustom, PDImages.GetActiveImage.ImgViewport.GetHScrollValue, PDImages.GetActiveImage.ImgViewport.GetVScrollValue
            
            'Reflow any image-window-specific chrome (status bar, rulers, etc)
            FormMain.MainCanvas(0).AlignCanvasView
            
        Else
            Dim tmpViewportParams As PD_ViewportParams
            tmpViewportParams = Viewport.GetDefaultParamObject()
            tmpViewportParams.curPOI = poi_ReuseLast
            Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0), VarPtr(tmpViewportParams)
        End If
        
        'Run the main SyncInterfaceToImage function, and notify a few peripheral functions of the updated image
        ' (e.g. updating thumbnails, window captions, etc)
        Interface.NotifyNewActiveImage PDImages.GetActiveImageID()
        
    End If
    
    'Make sure any tool initializations that vary by image are up-to-date.  (This includes things like
    ' making sure a scratch layer exists, and that it matches the current image's size.)
    Tools.InitializeToolsDependentOnImage activeImageChanging
    PDDebug.LogAction "CanvasManager.ActivatePDImage says: image #" & PDImages.GetActiveImageID() & " - " & Interface.GetWindowCaption(PDImages.GetActiveImage(), False) & " - was activated because " & reasonForActivation
    PDDebug.LogAction "CanvasManager.ActivatePDImage finished in " & VBHacks.GetTimeDiffNowAsString(startTime)
        
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
        Dim intRect As RectF
        srcImage.ImgViewport.GetIntersectRectCanvas intRect
        IsMouseOverImage = PDMath.IsPointInRectF(x1, y1, intRect)
        
    Else
        IsMouseOverImage = False
    End If

End Function

