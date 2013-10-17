Attribute VB_Name = "Image_Window_Handler"
'***************************************************************************
'Image Window Handler
'Copyright ©2002-2013 by Tanner Helland
'Created: 11/29/02
'Last updated: 12/October/13
'Last update: all functions have been rewritten to interact properly with the new window manager class.
'
'This module contains functions relating to the creation, sizing, and maintenance of the windows (forms) associated with
' each image loaded by the user.  Even in single-window mode, each loaded image receives its own form.  This form is
' used to display the image on-screen, and its size is used to determine a number of viewport characteristics (such as
' whether or not scroll bars are needed to move around the image).
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
Public Sub CreateNewImageForm(Optional ByVal forInternalUse As Boolean = False)

    'The viewport will automatically attempt to update whenever a form is resized.  We forcibly disable such updating by setting
    ' this value to FALSE prior to working with the viewport.  When we are finished, we must set it to TRUE for the viewport to work!
    g_AllowViewportRendering = False

    'Increase the number of images we're tracking
    g_NumOfImagesLoaded = g_NumOfImagesLoaded + 1
    ReDim Preserve pdImages(0 To g_NumOfImagesLoaded) As pdImage
    
    Set pdImages(g_NumOfImagesLoaded) = New pdImage

    'This is the actual, physical form object on which an image will reside
    Dim newImageForm As New FormImage
    
    'IMPORTANT: the form tag is the only way we can keep track of separate forms
    'DO NOT CHANGE THIS TAG VALUE!
    newImageForm.Tag = g_NumOfImagesLoaded
    
    'Remember this ID in the associated image class
    pdImages(g_NumOfImagesLoaded).IsActive = True
    pdImages(g_NumOfImagesLoaded).imageID = g_NumOfImagesLoaded
        
    'Set a default window size (in twips)
    newImageForm.Width = 4500
    
    newImageForm.Height = 1
    
    'Default image values
    Set pdImages(g_NumOfImagesLoaded).containingForm = newImageForm
    pdImages(g_NumOfImagesLoaded).UndoNum = 0
    pdImages(g_NumOfImagesLoaded).UndoMax = 0
    pdImages(g_NumOfImagesLoaded).UndoState = False
    pdImages(g_NumOfImagesLoaded).RedoState = False
    pdImages(g_NumOfImagesLoaded).CurrentZoomValue = ZOOM_100_PERCENT   'Default zoom is 100%
    
    'Hide the form off-screen while the loading takes place, but remember its location so we can restore it post-load.
    Dim mainClientRect As winRect
    g_WindowManager.getActualMainFormClientRect mainClientRect
    pdImages(g_NumOfImagesLoaded).WindowLeft = mainClientRect.x1
    pdImages(g_NumOfImagesLoaded).WindowTop = mainClientRect.y1
    newImageForm.Left = 0
    newImageForm.Top = g_cMonitors.DesktopHeight * Screen.TwipsPerPixelY
    
    newImageForm.Show vbModeless, FormMain
    newImageForm.Caption = g_Language.TranslateMessage("Loading image...")
    'If FormMain.Enabled Then newImageForm.SetFocus
    
    'Set this image as the current one
    g_CurrentImage = g_NumOfImagesLoaded
    
    'Track how many windows we currently have open
    g_OpenImageCount = g_OpenImageCount + 1
    
    'Run a separate subroutine (see bottom of this page) to enable/disable menus and such if this is the first image to be loaded
    UpdateMDIStatus
    
    'Re-enable viewport adjustments
    g_AllowViewportRendering = True
    
    'If this image wasn't loaded by the user (e.g. it's an internal PhotoDemon process), mark is as such
    pdImages(g_NumOfImagesLoaded).forInternalUseOnly = forInternalUse
    
End Sub

'Fit the active window tightly around the image, using its current zoom value.  It is generally assumed that the image has been set to a
' reasonable zoom value at this point (preferably by FitImageToViewport); otherwise, this function may result in a very large form.
Public Sub FitWindowToImage(Optional ByVal suppressRendering As Boolean = False, Optional ByVal isImageLoading As Boolean = False)
        
    If g_OpenImageCount = 0 Then Exit Sub
    
    'If image windows are docked, we don't need to perform this function, as the window manager will automatically handle all
    ' image window positioning.
    If Not g_WindowManager.getFloatState(IMAGE_WINDOW) Then Exit Sub
        
    'Make sure the window isn't minimized or maximized
    If pdImages(g_CurrentImage).containingForm.WindowState = 0 Then
    
        'Disable AutoScroll, because that messes with our calculations
        g_AllowViewportRendering = False
    
        'To minimize flickering, we will only apply width/height and top/left changes once.
        ' While calculations are being run, store all changes to variables.
        Dim curTop As Long, curLeft As Long
        Dim curWidth As Long, curHeight As Long
    
        'Because certain changes will trigger the appearance of scroll bars, which take up extra space in the viewport,
        ' we need to check for this and factor it into our window size calculation.
        Dim forceMaxWidth As Boolean, forceMaxHeight As Boolean
        forceMaxWidth = False
        forceMaxHeight = False
        
        'We need a copy of two rects handled by the window manager:
        ' 1) the main form's client area, which we will use to reposition the window if an image is being loaded for the first time.
        ' 2) the current image window's rect, which we need if this action is happening at some time other than image load.
        Dim mainClientRect As winRect, curWindowRect As winRect
        g_WindowManager.getActualMainFormClientRect mainClientRect
        g_WindowManager.getWindowRectByIndex pdImages(g_CurrentImage).indexInWindowManager, curWindowRect
        
        'As a convenience to the caller, we will never allow the form to be larger than...
        '1) the main form's client area, if this function is called when an image is first loaded
        '2) the desktop, if this function is called after an image is first loaded
        Dim maxRight As Long, maxBottom As Long
        If isImageLoading Then
            maxRight = mainClientRect.x2
            maxBottom = mainClientRect.y2
        Else
            maxRight = g_cMonitors.DesktopWidth
            maxBottom = g_cMonitors.DesktopHeight
        End If
        
        'Regardless of when this function is run, we do not allow it to cover the main window's menu bar.
        ' Determine a max left/top position accordingly.
        Dim maxLeft As Long, maxTop As Long
        maxLeft = mainClientRect.x1
        maxTop = mainClientRect.y1
        
        'Now let's get some dimensions for our calculations.  These values hold chrome (window border) size in both directions
        Dim wDif As Long, hDif As Long
        wDif = g_WindowManager.getHorizontalChromeSize(pdImages(g_CurrentImage).containingForm.hWnd)
        hDif = g_WindowManager.getVerticalChromeSize(pdImages(g_CurrentImage).containingForm.hWnd)
        
        'Start our calculations by setting the new width/height to equal the image's current size (while accounting for zoom)
        curWidth = wDif + (pdImages(g_CurrentImage).Width * g_Zoom.ZoomArray(toolbar_File.CmbZoom.ListIndex))
        curHeight = hDif + (pdImages(g_CurrentImage).Height * g_Zoom.ZoomArray(toolbar_File.CmbZoom.ListIndex))
        
        'We now handle the fit operation in two possible ways:
        ' 1) If the image is being loaded for the first time, constrain its location to the main form's client area.
        ' 2) If the image is NOT being loaded for the first time, constrain its location to the desktop, while
        '     retaining its current top-left position (if possible).
        If isImageLoading Then
            curLeft = mainClientRect.x1
            curTop = mainClientRect.y1
        Else
            curLeft = curWindowRect.x1
            curTop = curWindowRect.y1
        End If
                
        ' Check for the image being off-viewport, starting with the vertical measurement.
        If curTop + curHeight > maxBottom Then
        
            'If we can solve the problem by simply moving the form upward, while keeping it inside the viewport, do that.
            If curHeight < (maxBottom - maxTop) Then
                curTop = maxBottom - curHeight
        
            'If the window is taller than the viewport, and moving it doesn't solve the problem, we need to shrink the window accordingly.
            '  Move the image to the tallest acceptable location, then shrink the window vertically (which will force a scroll bar to appear).
            Else
                curTop = maxTop
                curHeight = maxBottom - maxTop
                forceMaxHeight = True
            End If
        
        End If
       
        'Next, check the horizontal measurement.
        If curLeft + curWidth > maxRight Then
       
            'If we can solve the problem by simply moving the form left, do that.
            If curWidth < (maxRight - maxLeft) Then
                curLeft = maxRight - curWidth
       
            'If the image is wider than the viewport, and too large to simply shift it left, we need to shrink the window horizontally.
            ' Move the image to the left-most acceptable location in the viewport, then shrink the window horizontally (which will force
            ' a scroll bar to appear).
            Else
                curLeft = maxLeft
                curWidth = maxRight - maxLeft
                forceMaxWidth = True
            End If
       
        End If
        
        'If the image does not fill the entire viewport, but one dimension is maxed out, add a little extra space for
        ' the scroll bar that must appear.
        If forceMaxHeight And (Not forceMaxWidth) Then
            
            curWidth = curWidth + pdImages(g_CurrentImage).containingForm.VScroll.Width
            
            'If this addition pushes the image off-screen, nudge it slightly left
            If curLeft + curWidth > maxRight Then curLeft = maxRight - curWidth
            
        End If
        
        If forceMaxWidth And (Not forceMaxHeight) Then
            curHeight = curHeight + pdImages(g_CurrentImage).containingForm.HScroll.Height
            
            'If this addition pushes the image off-screen, nudge it slightly up
            If curTop + curHeight > maxBottom Then curTop = maxBottom - curHeight
            
        End If
        
        'Apply the changes in whatever manner appropriate (again, this is handled differently if the image is newly loaded; in that case
        ' we want to delay positioning until all prep work is done, to prevent unsightly flickering and jitters)
        If isImageLoading Then
            pdImages(g_CurrentImage).WindowLeft = curLeft
            pdImages(g_CurrentImage).WindowTop = curTop
            pdImages(g_CurrentImage).containingForm.Width = Screen.TwipsPerPixelX * curWidth
            pdImages(g_CurrentImage).containingForm.Height = Screen.TwipsPerPixelY * curHeight
        Else
            pdImages(g_CurrentImage).containingForm.Move curLeft * Screen.TwipsPerPixelX, curTop * Screen.TwipsPerPixelY, curWidth * Screen.TwipsPerPixelX, curHeight * Screen.TwipsPerPixelY
        End If
        
        'Re-enable scrolling
        g_AllowViewportRendering = True
        
    End If
    
    'Because external functions may rely on this to redraw the viewport, force a redraw regardless of whether or not
    ' the window was actually fit to the image (unless suppressRendering is specified, obviously)
    If Not suppressRendering Then PrepareViewport pdImages(g_CurrentImage).containingForm, "FitWindowToImage"
    
End Sub

'Resize the window so that all four edges are within the current viewport.
Public Sub FitWindowToViewport(Optional ByVal suppressRendering As Boolean = False)
        
    If g_OpenImageCount = 0 Then Exit Sub
    
    Dim resizeNeeded As Boolean
    resizeNeeded = False
        
    'Make sure the window isn't minimized or maximized
    If pdImages(g_CurrentImage).containingForm.WindowState = 0 Then
    
        'Prevent automatic recalculation of the viewport scroll bars until we finish our calculations here
        g_AllowViewportRendering = False
        
        'We need a copy of two rects handled by the window manager:
        ' 1) the main form's client area, which we will use to reposition the window if an image is being loaded for the first time.
        ' 2) the current image window's rect, which we need if this action is happening at some time other than image load.
        Dim mainClientRect As winRect, curWindowRect As winRect
        g_WindowManager.getActualMainFormClientRect mainClientRect
        g_WindowManager.getWindowRectByIndex pdImages(g_CurrentImage).indexInWindowManager, curWindowRect
        
        'Start by determining if the image's canvas falls outside the viewport area.  Note that we will repeat this process
        ' twice: once for horizontal, and again for vertical.
        If curWindowRect.x2 > mainClientRect.x2 Then
            
            resizeNeeded = True
            
            'This variable determines the difference between the MDI client area's available width and the current child form's
            ' width, taking into account the .Left position and an arbitrary offset (currently 12 pixels)
            Dim newWidth As Long
            newWidth = mainClientRect.x2 - curWindowRect.x1
            pdImages(g_CurrentImage).containingForm.Width = Screen.TwipsPerPixelX * newWidth
            
        End If
        
        'Now repeat the process for the vertical measurement
        If curWindowRect.y2 > mainClientRect.y2 Then
        
            resizeNeeded = True
        
            Dim newHeight As Long
            newHeight = mainClientRect.y2 - curWindowRect.y1
            pdImages(g_CurrentImage).containingForm.Height = Screen.TwipsPerPixelY * newHeight
            
        End If
            
        'Re-enable scrolling
        g_AllowViewportRendering = True
        
    End If
    
    'Because external functions may rely on this to redraw the viewport, force a redraw regardless of whether or not
    ' the window was actually fit to the image (unless suppressRendering is specified, obviously)
    If (Not suppressRendering) And resizeNeeded Then PrepareViewport pdImages(g_CurrentImage).containingForm, "FitWindowToViewport"
    
End Sub

'Fit the current image onscreen at as large a size as possible (but never larger than 100% zoom)
Public Sub FitImageToViewport(Optional ByVal suppressRendering As Boolean = False)
    
    If g_OpenImageCount = 0 Then Exit Sub
    
    'Disable AutoScroll, because that messes with our calculations
    g_AllowViewportRendering = False
    
    'Note that all window dimension and position information comes from PD's window manager.  Because we modify window borders on-the-fly,
    ' VB's internal measurements are not accurate, so we must rely on the window manager's measurements.
    
    'Make sure the window isn't minimized
    If pdImages(g_CurrentImage).containingForm.WindowState = vbMinimized Then Exit Sub
    
    'In order to properly calculate auto-zoom, we need to know the largest possible area we have to work with. Ask the window manager
    ' for that value now (which is calculated based on the main form's area, minus toolbar sizes if docked).
    Dim maxWidth As Long, maxHeight As Long
    maxWidth = g_WindowManager.requestActualMainFormClientWidth
    maxHeight = g_WindowManager.requestActualMainFormClientHeight
    
    'If image windows are floating, we need to factor window chrome (borders) into the maximum available size
    If g_WindowManager.getFloatState(IMAGE_WINDOW) Then
        maxWidth = maxWidth - g_WindowManager.getHorizontalChromeSize(pdImages(g_CurrentImage).containingForm.hWnd)
        maxHeight = maxHeight - g_WindowManager.getVerticalChromeSize(pdImages(g_CurrentImage).containingForm.hWnd)
    End If
    
    'Use this to track the zoom value required to fit the image on-screen; we will start at 100%, then move downward until we find an ideal zoom.
    Dim zVal As Long
    zVal = ZOOM_100_PERCENT
    
    Dim i As Long
    
    'First, let's check to see if we need to adjust zoom because the width is too big
    If pdImages(g_CurrentImage).Width > maxWidth Then
        
        'The image is larger than the maximum available area.  Loop backwards through all possible zoom values until we find one that fits.
        For i = ZOOM_100_PERCENT To g_Zoom.ZoomCount Step 1
        
            If (pdImages(g_CurrentImage).Width * g_Zoom.ZoomArray(i)) < maxWidth Then
                zVal = i
                Exit For
            End If
        Next i
        
    End If
    
    'Repeat the above step, but for height.  Note that we start our "find best zoom" search from whatever zoom the horizontal search found.
    If (pdImages(g_CurrentImage).Height * g_Zoom.ZoomArray(zVal)) > maxHeight Then
    
        For i = zVal To g_Zoom.ZoomCount Step 1
            If (pdImages(g_CurrentImage).Height * g_Zoom.ZoomArray(i)) < maxHeight Then
                zVal = i
                Exit For
            End If
        Next i
        
    End If
    
    'Change the zoom combo box to reflect the new zoom value
    toolbar_File.CmbZoom.ListIndex = zVal
    pdImages(g_CurrentImage).CurrentZoomValue = zVal
    
    'Re-enable scrolling
    g_AllowViewportRendering = True
    
    'Now fix scrollbars and everything
    If Not suppressRendering Then PrepareViewport pdImages(g_CurrentImage).containingForm, "FitImageToViewport"
    
End Sub

'Fit the current image onscreen at as large a size as possible (including possibility of zoomed-in)
Public Sub FitOnScreen()
    
    If g_OpenImageCount = 0 Then Exit Sub
        
    'Disable AutoScroll, because that messes with our calculations
    g_AllowViewportRendering = False
    
    'Note that all window dimension and position information comes from PD's window manager.  Because we modify window borders on-the-fly,
    ' VB's internal measurements are not accurate, so we must rely on the window manager's measurements.
    
    'If the image is minimized, restore it
    If pdImages(g_CurrentImage).containingForm.WindowState = vbMinimized Then pdImages(g_CurrentImage).containingForm.WindowState = 0
    
    'In order to properly calculate auto-zoom, we need to know the largest possible area we have to work with. Ask the window manager
    ' for that value now (which is calculated based on the main form's area, minus toolbar sizes if docked).
    Dim maxWidth As Long, maxHeight As Long
    maxWidth = g_WindowManager.requestActualMainFormClientWidth
    maxHeight = g_WindowManager.requestActualMainFormClientHeight
    
    'If image windows are floating, we need to factor window chrome (borders) into the maximum available size
    If g_WindowManager.getFloatState(IMAGE_WINDOW) Then
        maxWidth = maxWidth - g_WindowManager.getHorizontalChromeSize(pdImages(g_CurrentImage).containingForm.hWnd)
        maxHeight = maxHeight - g_WindowManager.getVerticalChromeSize(pdImages(g_CurrentImage).containingForm.hWnd)
    End If
    
    'Use this to track zoom
    Dim zVal As Long
    zVal = 0
    
    Dim i As Long
    
    'Run a loop backwards through the possible zoom values, until we find one that fits the current image
    For i = 0 To g_Zoom.ZoomCount Step 1
        If (pdImages(g_CurrentImage).Width * g_Zoom.ZoomArray(i)) < maxWidth Then
            zVal = i
            Exit For
        End If
    Next i
    
    'Now do the same thing for the height, starting at whatever zoom value we previously found
    For i = zVal To g_Zoom.ZoomCount Step 1
        If (pdImages(g_CurrentImage).Height * g_Zoom.ZoomArray(i)) < maxHeight Then
            zVal = i
            Exit For
        End If
    Next i
    
    toolbar_File.CmbZoom.ListIndex = zVal
    pdImages(g_CurrentImage).CurrentZoomValue = zVal
    
    'Re-enable scrolling
    g_AllowViewportRendering = True
    
    'If the window is not maximized or minimized, fit the window to it
    If pdImages(g_CurrentImage).containingForm.WindowState = 0 Then FitWindowToImage True
    
    'Now fix scrollbars and everything
    PrepareViewport pdImages(g_CurrentImage).containingForm, "FitOnScreen"
    
End Sub

'When windows are created or destroyed, launch this routine to dis/en/able windows and toolbars, etc
Public Sub UpdateMDIStatus()

    'If two or more windows are open, enable the image tabstrip, and the Next/Previous image menu items
    If g_OpenImageCount >= 2 Then
        g_WindowManager.setWindowVisibility toolbar_ImageTabs.hWnd, True
        FormMain.MnuWindow(6).Enabled = True
        FormMain.MnuWindow(7).Enabled = True
    Else
        g_WindowManager.setWindowVisibility toolbar_ImageTabs.hWnd, False
        FormMain.MnuWindow(6).Enabled = False
        FormMain.MnuWindow(7).Enabled = False
    End If

    'If every window has been closed, disable all toolbar and menu options that are no longer applicable
    If g_OpenImageCount < 1 Then
        metaToggle tFilter, False
        metaToggle tSave, False
        metaToggle tSaveAs, False
        metaToggle tCopy, False
        metaToggle tUndo, False
        metaToggle tRedo, False
        metaToggle tImageOps, False
        metaToggle tFilter, False
        metaToggle tMacro, False
        metaToggle tRepeatLast, False
        metaToggle tSelection, False
        FormMain.MnuFile(7).Enabled = False
        FormMain.MnuFile(8).Enabled = False
        toolbar_File.cmdClose.Enabled = False
        FormMain.MnuFitWindowToImage.Enabled = False
        FormMain.MnuFitOnScreen.Enabled = False
        If toolbar_File.CmbZoom.Enabled And toolbar_File.Visible Then
            toolbar_File.CmbZoom.Enabled = False
            'FormMain.lblLeftToolBox(3).ForeColor = &H606060
            toolbar_File.CmbZoom.ListIndex = ZOOM_100_PERCENT   'Reset zoom to 100%
            toolbar_File.cmdZoomIn.Enabled = False
            toolbar_File.cmdZoomOut.Enabled = False
        End If
        
        toolbar_File.lblImgSize.ForeColor = &HD1B499
        toolbar_File.lblCoordinates.ForeColor = &HD1B499
        
        toolbar_File.lblImgSize.Caption = ""
        
        toolbar_File.lblCoordinates.Caption = ""
        
        Message "Please load an image.  (The large 'Open Image' button at the top-left should do the trick!)"
        
        'Finally, if dynamic icons are enabled, restore the main program icon and clear the icon cache
        If g_UserPreferences.GetPref_Boolean("Interface", "Dynamic Taskbar Icon", True) Then
            destroyAllIcons
            setNewTaskbarIcon origIcon32, FormMain.hWnd
            setNewAppIcon origIcon16
        End If
        
        'New addition: destroy all inactive pdImage objects.  This helps keep memory usage at a bare minimum.
        If g_NumOfImagesLoaded > 1 Then
        
            Dim i As Long
            
            'Loop through all pdImage objects and make sure they've been deactivated
            For i = 0 To g_NumOfImagesLoaded
                If (Not pdImages(i) Is Nothing) Then
                    pdImages(i).deactivateImage
                    Set pdImages(i) = Nothing
                End If
            Next i
        
            'Redim the pdImages array
            'Erase pdImages
        
            'Reset all window tracking variables
            g_NumOfImagesLoaded = 0
            g_CurrentImage = 0
            g_OpenImageCount = 0
                        
        End If
        
        'Erase any remaining viewport buffer
        eraseViewportBuffers
                
    'Otherwise, enable all of 'em
    Else
        metaToggle tFilter, True
        metaToggle tSave, True
        metaToggle tSaveAs, True
        metaToggle tCopy, True
        metaToggle tUndo, pdImages(g_CurrentImage).UndoState
        metaToggle tRedo, pdImages(g_CurrentImage).RedoState
        metaToggle tImageOps, True
        metaToggle tFilter, True
        metaToggle tMacro, True
        metaToggle tRepeatLast, pdImages(g_CurrentImage).RedoState
        FormMain.MnuFile(7).Enabled = True
        toolbar_File.cmdClose.Enabled = True
        FormMain.MnuFile(8).Enabled = True
        FormMain.MnuFitWindowToImage.Enabled = True
        FormMain.MnuFitOnScreen.Enabled = True
        toolbar_File.lblImgSize.ForeColor = &H544E43
        toolbar_File.lblCoordinates.ForeColor = &H544E43
        If toolbar_File.CmbZoom.Enabled = False Then
            toolbar_File.CmbZoom.Enabled = True
            'FormMain.lblLeftToolBox(3).ForeColor = &H544E43
            toolbar_File.cmdZoomIn.Enabled = True
            toolbar_File.cmdZoomOut.Enabled = True
        End If
    End If
    
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
