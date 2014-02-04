Attribute VB_Name = "Image_Canvas_Handler"
'***************************************************************************
'Image Canvas Handler (formerly Image Window Handler)
'Copyright ©2002-2014 by Tanner Helland
'Created: 11/29/02
'Last updated: 01/February/14
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

'Create a new, blank MDI child
Public Sub CreateNewPDImage(Optional ByVal forInternalUse As Boolean = False)

    'The viewport will automatically attempt to update whenever a form is resized.  We forcibly disable such updating by setting
    ' this value to FALSE prior to working with the viewport.  When we are finished, we must set it to TRUE for the viewport to work!
    g_AllowViewportRendering = False

    'Increase the number of images we're tracking
    g_NumOfImagesLoaded = g_NumOfImagesLoaded + 1
    ReDim Preserve pdImages(0 To g_NumOfImagesLoaded) As pdImage
    
    Set pdImages(g_NumOfImagesLoaded) = New pdImage
    
    'Assign it a 32bpp icon matching the PD one; this will be overwritten momentarily, but as the user may sneak a peak
    ' if the image takes a long time to load (e.g. RAW photos), it's nice to display something other than the crappy
    ' stock VB form icon.
    'SetIcon newImageForm.hWnd, "AAA", False
    
    'IMPORTANT: the form tag is the only way we can keep track of separate forms
    'DO NOT CHANGE THIS TAG VALUE!
    'newImageForm.Tag = g_NumOfImagesLoaded
    
    'Remember this ID in the associated image class
    pdImages(g_NumOfImagesLoaded).IsActive = True
    pdImages(g_NumOfImagesLoaded).imageID = g_NumOfImagesLoaded
    
    'If this image wasn't loaded by the user (e.g. it's an internal PhotoDemon process), mark is as such
    pdImages(g_NumOfImagesLoaded).forInternalUseOnly = forInternalUse
        
    'Note the current vertical offset of the viewport
    pdImages(g_NumOfImagesLoaded).imgViewport.setBottomOffset FormMain.mainCanvas(0).getStatusBarHeight
    
    'Set a default zoom of 100% (this is likely to change, assuming the user has auto-zoom enabled)
    pdImages(g_NumOfImagesLoaded).currentZoomValue = g_Zoom.getZoom100Index
    
    'Hide the form off-screen while the loading takes place, but remember its location so we can restore it post-load.
    Dim mainClientRect As winRect
    g_WindowManager.getActualMainFormClientRect mainClientRect
    pdImages(g_NumOfImagesLoaded).WindowLeft = mainClientRect.x1
    pdImages(g_NumOfImagesLoaded).WindowTop = mainClientRect.y1
    'newImageForm.Left = 0
    'newImageForm.Top = g_cMonitors.DesktopHeight * Screen.TwipsPerPixelY
    
    'Previously we used the .Show event here to display the form, but we now rely on the window manager to handle this
    ' later in the load process.  (This reduces flicker while loading images.)
    'If MacroStatus <> MacroBATCH Then newImageForm.Show vbModeless, FormMain
    
    'Use the window manager to properly assign the main form ownership over this window
    'g_WindowManager.requestNewOwner newImageForm.hWnd, FormMain.hWnd
    
    'Supply a temporary caption (again, only necessary if the image takes a long time to load)
    'newImageForm.Caption = g_Language.TranslateMessage("Loading image...")
    
    'Set this image as the current one
    g_CurrentImage = g_NumOfImagesLoaded
    
    'Track how many windows we currently have open
    g_OpenImageCount = g_OpenImageCount + 1
    
    'Run a separate subroutine to enable/disable menus (important primarily if this is the first image to be loaded)
    syncInterfaceToCurrentImage
    
    'Re-enable viewport adjustments
    g_AllowViewportRendering = True
    
End Sub

'Fit the current image onscreen at as large a size as possible (but never larger than 100% zoom)
Public Sub FitImageToViewport(Optional ByVal suppressRendering As Boolean = False)
    
    If g_OpenImageCount = 0 Then Exit Sub
    
    'Disable AutoScroll, because that messes with our calculations
    g_AllowViewportRendering = False
    
    'Note that all window dimension and position information comes from PD's window manager.  Because we modify window borders on-the-fly,
    ' VB's internal measurements are not accurate, so we must rely on the window manager's measurements.
    
    'In order to properly calculate auto-zoom, we need to know the largest possible area we have to work with. Ask the window manager
    ' for that value now (which is calculated based on the main form's area, minus toolbar sizes if docked).
    Dim maxWidth As Long, maxHeight As Long
    maxWidth = g_WindowManager.requestActualMainFormClientWidth
    maxHeight = g_WindowManager.requestActualMainFormClientHeight
    
    'Remove any additional per-window chrome from the available space (rulers, status bar, etc)
    maxHeight = maxHeight - pdImages(g_CurrentImage).imgViewport.getVerticalOffset
        
    'Use this to track the zoom value required to fit the image on-screen; we will start at 100%, then move downward until we find an ideal zoom.
    Dim zVal As Long
    zVal = g_Zoom.getZoom100Index
    
    Dim i As Long
    
    'First, let's check to see if we need to adjust zoom because the width is too big
    If pdImages(g_CurrentImage).Width > maxWidth Then
        
        'The image is larger than the maximum available area.  Loop backwards through all possible zoom values until we find one that fits.
        For i = g_Zoom.getZoom100Index To g_Zoom.getZoomCount Step 1
        
            If (pdImages(g_CurrentImage).Width * g_Zoom.getZoomValue(i)) < maxWidth Then
                zVal = i
                Exit For
            End If
        Next i
        
    End If
    
    'Repeat the above step, but for height.  Note that we start our "find best zoom" search from whatever zoom the horizontal search found.
    If (pdImages(g_CurrentImage).Height * g_Zoom.getZoomValue(zVal)) > maxHeight Then
    
        For i = zVal To g_Zoom.getZoomCount Step 1
            If (pdImages(g_CurrentImage).Height * g_Zoom.getZoomValue(i)) < maxHeight Then
                zVal = i
                Exit For
            End If
        Next i
        
    End If
    
    'Change the zoom combo box to reflect the new zoom value
    toolbar_File.CmbZoom.ListIndex = zVal
    pdImages(g_CurrentImage).currentZoomValue = zVal
    
    'Re-enable scrolling
    g_AllowViewportRendering = True
    
    'Now fix scrollbars and everything
    If Not suppressRendering Then PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "FitImageToViewport"
    
End Sub

'Fit the current image onscreen at as large a size as possible (including possibility of zoomed-in)
Public Sub FitOnScreen()
    
    If g_OpenImageCount = 0 Then Exit Sub
        
    'Disable AutoScroll, because that messes with our calculations
    g_AllowViewportRendering = False
    
    'Note that all window dimension and position information comes from PD's window manager.  Because we modify window borders on-the-fly,
    ' VB's internal measurements are not accurate, so we must rely on the window manager's measurements.
    
    'In order to properly calculate auto-zoom, we need to know the largest possible area we have to work with. Ask the window manager
    ' for that value now (which is calculated based on the main form's area, minus toolbar sizes if docked).
    Dim maxWidth As Long, maxHeight As Long
    maxWidth = g_WindowManager.requestActualMainFormClientWidth
    maxHeight = g_WindowManager.requestActualMainFormClientHeight
    
    'Remove any additional per-window chrome from the available space (rulers, status bar, etc)
    maxHeight = maxHeight - pdImages(g_CurrentImage).imgViewport.getVerticalOffset
        
    'Use this to track zoom
    Dim zVal As Long
    zVal = 0
    
    Dim i As Long
    
    'Run a loop backwards through the possible zoom values, until we find one that fits the current image
    For i = 0 To g_Zoom.getZoomCount Step 1
        If (pdImages(g_CurrentImage).Width * g_Zoom.getZoomValue(i)) < maxWidth Then
            zVal = i
            Exit For
        End If
    Next i
    
    'Now do the same thing for the height, starting at whatever zoom value we previously found
    For i = zVal To g_Zoom.getZoomCount Step 1
        If (pdImages(g_CurrentImage).Height * g_Zoom.getZoomValue(i)) < maxHeight Then
            zVal = i
            Exit For
        End If
    Next i
    
    toolbar_File.CmbZoom.ListIndex = zVal
    pdImages(g_CurrentImage).currentZoomValue = zVal
    
    'Re-enable scrolling
    g_AllowViewportRendering = True
        
    'Now fix scrollbars and everything
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "FitOnScreen"
    
End Sub

'Restore the main form to the window coordinates saved in the preferences file
Public Sub restoreMainWindowLocation()

    Exit Sub

    'First, check which state the window was in previously.
    Dim lWindowState As Long
    lWindowState = g_UserPreferences.GetPref_Long("Core", "Last Window State", 0)
        
    Dim lWindowLeft As Long, lWindowTop As Long
    Dim lWindowWidth As Long, lWindowHeight As Long
        
    'If the window state was "minimized", reset it to "normal"
    If lWindowState = vbMinimized Then lWindowState = 0
        
    'If the window state was "maximized", set that and ignore the saved width/height values
    If lWindowState = vbMaximized Then
            
        lWindowLeft = g_UserPreferences.GetPref_Long("Core", "Last Window Left", 1)
        lWindowTop = g_UserPreferences.GetPref_Long("Core", "Last Window Top", 1)
        FormMain.Left = lWindowLeft * Screen.TwipsPerPixelX
        FormMain.Top = lWindowTop * Screen.TwipsPerPixelY
        FormMain.WindowState = vbMaximized
        
    'If the window state is normal, attempt to restore the last-used values
    Else
            
        'Start by pulling the last left/top/width/height values from the preferences file
        lWindowLeft = g_UserPreferences.GetPref_Long("Core", "Last Window Left", 1)
        lWindowTop = g_UserPreferences.GetPref_Long("Core", "Last Window Top", 1)
        lWindowWidth = g_UserPreferences.GetPref_Long("Core", "Last Window Width", 1)
        lWindowHeight = g_UserPreferences.GetPref_Long("Core", "Last Window Height", 1)
            
        'If the left/top/width/height values all equal "1" (the default value), then a previous location has never
        ' been saved.  Center the form on the current screen at its default size.
        If (lWindowLeft = 1) And (lWindowTop = 1) And (lWindowWidth = 1) And (lWindowHeight = 1) Then
            
            g_cMonitors.CenterFormOnMonitor FormMain, FormMain
            
        'If values have been saved, perform a sanity check and then restore them
        Else
                                
            'Make sure the location values will result in an on-screen form.  If they will not (for example, if the user
            ' detached a second monitor that was previously attached and PhotoDemon was being used on that monitor),
            ' change the values to ensure that the program window appears on-screen.
            If (lWindowLeft + lWindowWidth) < g_cMonitors.DesktopLeft Then lWindowLeft = g_cMonitors.DesktopLeft
            If lWindowLeft > g_cMonitors.DesktopLeft + g_cMonitors.DesktopWidth Then lWindowLeft = g_cMonitors.DesktopWidth - lWindowWidth
            If lWindowTop < g_cMonitors.DesktopTop Then lWindowTop = g_cMonitors.DesktopTop
            If lWindowTop > g_cMonitors.DesktopHeight Then lWindowTop = g_cMonitors.DesktopHeight - lWindowHeight
                
            'Perform a similar sanity check for width and height using arbitrary values (200 pixels at present)
            If lWindowWidth < 200 Then lWindowWidth = 200
            'If lWindowHeight < 200 Then lWindowHeight = 200
                
            'With all values now set to guaranteed-safe values, set the main window's location
            FormMain.Left = lWindowLeft * Screen.TwipsPerPixelX
            FormMain.Top = lWindowTop * Screen.TwipsPerPixelY
            FormMain.Width = lWindowWidth * Screen.TwipsPerPixelX
            FormMain.Height = lWindowHeight * Screen.TwipsPerPixelY
            
        End If
            
        
    End If
        
    'Store the current window location to file (in case it hasn't been saved before, or we had to move it from
    ' an unavailable monitor to an available one)
    g_UserPreferences.SetPref_Long "Core", "Last Window Left", FormMain.Left / Screen.TwipsPerPixelX
    g_UserPreferences.SetPref_Long "Core", "Last Window Top", FormMain.Top / Screen.TwipsPerPixelY
    g_UserPreferences.SetPref_Long "Core", "Last Window Width", FormMain.Width / Screen.TwipsPerPixelX
    g_UserPreferences.SetPref_Long "Core", "Last Window Height", FormMain.Height / Screen.TwipsPerPixelY
    
End Sub

'Previously, we could unload images by just unloading their containing form.  This is no longer possible, so we must
' unload images using dedicated custom functions.  Note that this function simply wraps the QueryUnload and Unload
' functions, below.
'This function returns TRUE if the image was unloaded, FALSE if it was canceled.
Public Function fullPDImageUnload(ByVal imageID As Long) As Boolean

    Dim toCancel As Integer
    Dim tmpUnloadMode As Integer
    
    QueryUnloadPDImage toCancel, tmpUnloadMode, imageID
    
    If CBool(toCancel) Then
        fullPDImageUnload = False
        Exit Function
    End If
    
    UnloadPDImage toCancel, imageID
    
    If CBool(toCancel) Then
        fullPDImageUnload = False
    Else
        
        'Redraw the screen
        If g_OpenImageCount > 0 Then
            PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "another image closed"
        Else
            FormMain.mainCanvas(0).clearCanvas
        End If
        
        fullPDImageUnload = True
    End If

End Function

'Previously, we could unload images by just unloading their containing form.  This is no longer possible, so we must
' query unload images using this special function.
Public Function QueryUnloadPDImage(ByRef Cancel As Integer, ByRef UnloadMode As Integer, ByVal imageID As Long) As Boolean

    Debug.Print "(Image #" & imageID & " received a Query_Unload trigger)"
    
    'Failsafe to make sure the image was properly initialized
    If pdImages(imageID) Is Nothing Then Exit Function
    
    'If the user wants to be prompted about unsaved images, do it now
    If g_ConfirmClosingUnsaved And pdImages(imageID).IsActive And (Not pdImages(imageID).forInternalUseOnly) Then
    
        'Check the .HasBeenSaved property of the image associated with this form
        If Not pdImages(imageID).getSaveState Then
                        
            'If the user hasn't already told us to deal with all unsaved images in the same fashion, run some checks
            If Not g_DealWithAllUnsavedImages Then
            
                g_NumOfUnsavedImages = 0
                                
                'Loop through all images to count how many unsaved images there are in total.
                ' NOTE: we only need to do this if the entire program is being shut down or if the user has selected "close all";
                ' otherwise, this close action only affects the current image, so we shouldn't present a "repeat for all images" option
                If g_ProgramShuttingDown Or g_ClosingAllImages Then
                    Dim i As Long
                    For i = 1 To g_NumOfImagesLoaded
                        If pdImages(i).IsActive And (Not pdImages(i).forInternalUseOnly) And (Not pdImages(i).getSaveState) Then
                            g_NumOfUnsavedImages = g_NumOfUnsavedImages + 1
                        End If
                    Next i
                End If
            
                'Before displaying the "do you want to save this image?" dialog, bring the image in question to the foreground.
                If FormMain.Enabled Then activatePDImage imageID, "unsaved changes dialog required"
                
                'Show the "do you want to save this image?" dialog. On that form, the number of unsaved images will be
                ' displayed and the user will be given an option to apply their choice to all unsaved images.
                Dim confirmReturn As VbMsgBoxResult
                confirmReturn = confirmClose(imageID)
                        
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
                If FormMain.Enabled Then activatePDImage imageID, "image being saved during shutdown"
                
                'Attempt to save. Note that the user can still cancel at this point, and we want to honor their cancellation
                Dim saveSuccessful As Boolean
                saveSuccessful = MenuSave(CLng(imageID))
                
                'If something went wrong, or the user canceled the save dialog, stop the unload process
                Cancel = Not saveSuccessful
 
                'If we make it here and the save was successful, force an immediate unload
                If Cancel = False Then
                    UnloadPDImage Cancel, imageID
                
                '...but if the save was not successful, suspend all unload action
                Else
                    If g_ProgramShuttingDown Then g_ProgramShuttingDown = False
                    If g_ClosingAllImages Then g_ClosingAllImages = False
                    g_DealWithAllUnsavedImages = False
                End If
            
            'Do not save the image
            ElseIf confirmReturn = vbNo Then
                
                'I think this "Unload Me" statement may be causing some kind of infinite recursion - perhaps because it triggers this very
                ' QueryUnload statement? Not sure, but I may need to revisit it if the problems don't go away...
                'UPDATE 26 Aug 2014: after changing my subclassing code, the problem seems to have disappeared, but I'm leaving
                ' this comment here until I'm absolutely certain the problem has been resolved.
                UnloadPDImage Cancel, imageID
                'Set Me = Nothing
                
            End If
        
        End If
    
    End If

End Function

'Previously, we could unload images by just unloading their containing form.  This is no longer possible, so we must
' unload images using this special function.
Public Function UnloadPDImage(Cancel As Integer, ByVal imageID As Long)

    'Failsafe to make sure the image was properly initialized
    If pdImages(imageID) Is Nothing Then Exit Function
    
    If pdImages(imageID).loadedSuccessfully Then Message "Closing image..."
    
    'Decrease the open image count
    g_OpenImageCount = g_OpenImageCount - 1
        
    'Deactivate this DIB (note that this will take care of additional actions, like clearing the Undo/Redo cache
    ' for this image)
    pdImages(imageID).deactivateImage
    
    'If this was the last (or only) open image and the histogram is loaded, unload the histogram
    ' (If we don't do this, the histogram may attempt to update, and without an active image it will throw an error)
    'If g_OpenImageCount = 0 Then Unload FormHistogram
    
    'Remove this image from the thumbnail toolbar
    toolbar_ImageTabs.RemoveImage imageID
    
    'Before exiting, restore focus to the next child window in line.  (But only if this image was the active window!)
    If g_CurrentImage = CLng(imageID) Then
    
        If g_OpenImageCount > 0 Then
        
            Dim i As Long
            i = Val(imageID) + 1
            If i > UBound(pdImages) Then i = i - 2
            
            Dim directionAscending As Boolean
            directionAscending = True
            
            Do While i >= 0
            
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).IsActive Then
                        activatePDImage i, "previous image unloaded"
                        Exit Do
                    End If
                End If
                
                If directionAscending Then
                    i = i + 1
                    If i > UBound(pdImages) Then
                        directionAscending = False
                        i = imageID
                    End If
                Else
                    i = i - 1
                End If
            
            Loop
        
        End If
        
    End If
    
    'If this was the last unloaded image, we need to disable a number of menus and other items.
    'If g_OpenImageCount = 0 Then g_WindowManager.allImageWindowsUnloaded
    
    'Sync the interface to match the settings of whichever image is active (or disable a bunch of items if no images are active)
    syncInterfaceToCurrentImage
    
    Message "Finished."
    
End Function

'Previously, images could be activated by clicking on their window.  Now that all images are rendered to a single
' user control on the main form, we must activate them manually.
Public Sub activatePDImage(ByVal imageID As Long, Optional ByRef reasonForActivation As String = "")

    'If this form is already the active image, don't waste time re-activating it
    If g_CurrentImage <> imageID Then
    
        'Update the current form variable
        g_CurrentImage = imageID
    
        'Because activation is an expensive process (requiring viewport redraws and more), I track the calls that access it.  This is used
        ' to minimize repeat calls as much as possible.
        Debug.Print "(Image #" & g_CurrentImage & " was activated because " & reasonForActivation & ")"
        
        'Double-check which monitor we are appearing on (for color management reasons)
        FormMain.mainCanvas(0).checkParentMonitor True
        
        'Before displaying the form, redraw it, just in case something changed while it was deactivated (e.g. form resize)
        PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Form received focus"
        
        'Reflow any image-window-specific chrome (status bar, rulers, etc)
        FormMain.mainCanvas(0).fixChromeLayout
    
        'Use the window manager to bring the window to the foreground
        'g_WindowManager.notifyChildReceivedFocus Me
        
        'Notify the thumbnail bar that a new image has been selected
        toolbar_ImageTabs.notifyNewActiveImage imageID
        
        'Synchronize various interface elements to match values stored in this image.
        syncInterfaceToCurrentImage
        
    End If
    
End Sub
