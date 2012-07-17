Attribute VB_Name = "MDIWindow"
'***************************************************************************
'MDI Window Handler
'�2000-2012 Tanner Helland
'Created: 11/29/02
'Last updated: 29/June/12
'Last update: added a SuppressUpdating component to two routines; this prevents a call to PrepareViewport.
'             Now, when an image is loaded, PrepareViewport only gets called once (instead of 3x, as before).
'             This speeds up image load time by a non-trivial amount.
'
'Interfaces with the main MDI active form; this module handles determining
'form size in relation to image size, etc.
'
'***************************************************************************

Option Explicit

'The image number we are currently at (always goes up, never down)
Public NumOfImagesLoaded As Long

'The current image we are working with (generally FormMain.ActiveForm.Tag)
Public CurrentImage As Long

'Number of existing windows
Public NumOfWindows As Long

'This array holds ALL IMPORTANT IMAGE INFORMATION for every loaded image.
'Undo functionality also exists only within these classes.
Public pdImages() As pdImage


'Create a new, blank MDI child
Public Sub CreateNewImageForm(Optional ByVal forInternalUse As Boolean = False)

    'Disable viewport adjustments
    FixScrolling = False

    'Increase the number of images we're tracking
    NumOfImagesLoaded = NumOfImagesLoaded + 1
    ReDim Preserve pdImages(0 To NumOfImagesLoaded) As pdImage
    
    Set pdImages(NumOfImagesLoaded) = New pdImage

    'This is the actual, physical form object on which an image will reside
    Dim frm As New FormImage
    
    'IMPORTANT: the form tag is the only way we can keep track of separate forms
    'DO NOT CHANGE THIS TAG VALUE!
    frm.Tag = NumOfImagesLoaded
    
    'Remember this ID in the associated image class
    pdImages(NumOfImagesLoaded).IsActive = True
    pdImages(NumOfImagesLoaded).ImageID = NumOfImagesLoaded
    
    'Default size (stupid twip measurements, unfortunately)
    frm.Width = 4500
    
    frm.Height = 1
    
    'Default image values
    Set pdImages(NumOfImagesLoaded).containingForm = frm
    pdImages(NumOfImagesLoaded).UndoNum = 0
    pdImages(NumOfImagesLoaded).UndoMax = 0
    pdImages(NumOfImagesLoaded).UndoState = False
    pdImages(NumOfImagesLoaded).RedoState = False
    pdImages(NumOfImagesLoaded).CurrentZoomValue = 15   'Default zoom is 100%
    
    'This is kind of cheap, but let's just set a random loading point between 0 and 99% :)
    Randomize Timer
    Dim RandPercent As Long
    RandPercent = Int(Rnd * 100)
    
    'Hide the form off-screen while the loading takes place, but remember its location so we can restore it post-load.
    pdImages(NumOfImagesLoaded).WindowLeft = frm.Left
    pdImages(NumOfImagesLoaded).WindowTop = frm.Top
    frm.Left = FormMain.ScaleWidth
    frm.Top = FormMain.ScaleHeight
    
    frm.Show
    frm.Caption = "Loading image (" & RandPercent & "%)..."
    frm.SetFocus
    
    'Set this image as the current one
    CurrentImage = NumOfImagesLoaded
    
    'Track how many windows we currently have open
    NumOfWindows = NumOfWindows + 1
    'Run a separate subroutine (see bottom of this page) to enable/disable
    'menus and stuff if no more windows are open
    UpdateMDIStatus
    
    'Re-enable viewport adjustments
    FixScrolling = True
    
    'If this image wasn't loaded by the user (e.g. it's an internal PhotoDemon process), mark is as such
    pdImages(NumOfImagesLoaded).forInternalUseOnly = forInternalUse
    
End Sub

'Fit the active window tightly around the image
Public Sub FitWindowToImage(Optional ByVal suppressRendering As Boolean = False)
    
    'Disable AutoScroll, because that messes with our calculations
    FixScrolling = False
    
    'Gotta change the scalemode to twips to match the MDI form
    FormMain.ActiveForm.ScaleMode = 1
    
    'Make sure the window isn't minimized or maximized
    FormMain.ActiveForm.WindowState = 0
    
    'Now let's get some dimensions for our calculations
    Dim tDif As Long, hDif As Long
    'This variable determines the difference between scalewidth and width...
    tDif = FormMain.ActiveForm.Width - FormMain.ActiveForm.ScaleWidth
    '...while this variable does the same thing for scaleheight and height
    hDif = FormMain.ActiveForm.Height - FormMain.ActiveForm.ScaleHeight
    
    'Now we set the form dimensions to match the image's
    FormMain.ActiveForm.Width = tDif + (FormMain.ActiveForm.BackBuffer.Width * Zoom.ZoomArray(FormMain.CmbZoom.ListIndex))
    FormMain.ActiveForm.Height = hDif + (FormMain.ActiveForm.BackBuffer.Height * Zoom.ZoomArray(FormMain.CmbZoom.ListIndex))
    
    'Set the scalemode back to a decent value
    FormMain.ActiveForm.ScaleMode = 3
    
    'Re-enable scrolling
    FixScrolling = True
    
    'Now fix scrollbars and everything
    If suppressRendering = False Then PrepareViewport FormMain.ActiveForm
    
End Sub

Public Sub FitImageToWindow(Optional ByVal suppressRendering As Boolean = False)
    
    'Disable AutoScroll, because that messes with our calculations
    FixScrolling = False
    
    'Gotta change the scalemode to twips to match the MDI form
    FormMain.ActiveForm.ScaleMode = 1
    
    'Make sure the window isn't minimized
    If FormMain.ActiveForm.WindowState = vbMinimized Then Exit Sub
    
    'Now let's get some dimensions for our calculations
    Dim tDif As Long, hDif As Long
    'This variable determines the difference between scalewidth and width...
    tDif = FormMain.ActiveForm.Width - FormMain.ActiveForm.ScaleWidth
    '...while this variable does the same thing for scaleheight and height
    hDif = FormMain.ActiveForm.Height - FormMain.ActiveForm.ScaleHeight
    
    'Use this to track zoom
    Dim zVal As Long
    zVal = 15
    
    'First, let's check to see if we need to adjust zoom because the width is too big
    If (FormMain.ActiveForm.BackBuffer.Width > FormMain.ScaleWidth - tDif) Then
        'If it is too big, run a loop backwards through the possible zoom values to see
        'if one will make it fit
        For x = 15 To 0 Step -1
            If (FormMain.ActiveForm.BackBuffer.Width * Zoom.ZoomArray(x)) < (FormMain.ScaleWidth - tDif) Then
                zVal = x
                Exit For
            End If
        Next x
        
    End If
    
    'Now we do the same thing for the height
    If (FormMain.ActiveForm.BackBuffer.Height > FormMain.ScaleHeight - hDif) Then
        'If the image's height is too big for the form, run a loop backwards through all
        ' possible zoom values to see if one will make it fit
        For x = zVal To 0 Step -1
            If (FormMain.ActiveForm.BackBuffer.Height * Zoom.ZoomArray(x)) < FormMain.ScaleHeight - hDif Then
                zVal = x
                Exit For
            End If
        Next x
        
    End If
    
    'Change the zoom combo box to reflect the new zoom value (or the default, if no changes were made)
    FormMain.CmbZoom.ListIndex = zVal
    pdImages(CurrentImage).CurrentZoomValue = zVal
    
    'Set the scalemode back to a decent value
    FormMain.ActiveForm.ScaleMode = 3
    
    'Re-enable scrolling
    FixScrolling = True
    
    'Now fix scrollbars and everything
    If suppressRendering = False Then PrepareViewport FormMain.ActiveForm
    
End Sub

'Fit the current image onscreen at as large a size as possible
Public Sub FitOnScreen()
    'Current plan: this needs to be an individual task.
    
    'Gotta change the scalemode to twips to match the MDI form
    FormMain.ActiveForm.ScaleMode = 1
    
    'Next, get the image data (so we have picwidthl and picheightl)
    'GetImageData
    
    'Make sure the window isn't minimized or maximized
    FormMain.ActiveForm.WindowState = 0
    
    'Disable AutoScroll, because that messes with our calculations
    FixScrolling = False
    
    'Now let's get some dimensions for our calculations
    Dim tDif As Long, hDif As Long
    'This variable determines the difference between scalewidth and width...
    tDif = FormMain.ActiveForm.Width - FormMain.ActiveForm.ScaleWidth
    '...while this variable does the same thing for scaleheight and height
    hDif = FormMain.ActiveForm.Height - FormMain.ActiveForm.ScaleHeight

    'Use this to track zoom
    Dim zVal As Long
    zVal = FormMain.CmbZoom.ListCount - 1
    
    'Run a loop backwards through the possible zoom values to see
    'if one will make it fit at the maximum possible size
    For x = FormMain.CmbZoom.ListCount - 1 To 0 Step -1
        If (FormMain.ActiveForm.BackBuffer.Width * Zoom.ZoomArray(x)) < FormMain.ScaleWidth - tDif Then
            zVal = x
            Exit For
        End If
    Next x
    
    'Now we do the same thing for the height
    For x = zVal To 0 Step -1
        If (FormMain.ActiveForm.BackBuffer.Height * Zoom.ZoomArray(x)) < FormMain.ScaleHeight - hDif Then
            zVal = x
            Exit For
        End If
    Next x
    FormMain.CmbZoom.ListIndex = zVal
    pdImages(CurrentImage).CurrentZoomValue = zVal
    
    'Set the scalemode back to pixels
    FormMain.ActiveForm.ScaleMode = 3
    'Re-enable scrolling
    FixScrolling = True
    FitWindowToImage
    'Now fix scrollbars and everything
    PrepareViewport FormMain.ActiveForm
    
End Sub

'When windows are created or destroyed, launch this routine to dis/en/able windows and toolbars, etc
Public Sub UpdateMDIStatus()
    'If every window has been closed, disable all toolbar and menu
    ' options that are no longer applicable
    If NumOfWindows < 1 Then
        tInit tFilter, False
        tInit tSave, False
        tInit tSaveAs, False
        tInit tCopy, False
        tInit tUndo, False
        tInit tRedo, False
        tInit tImageOps, False
        tInit tFilter, False
        tInit tHistogram, False
        tInit tMacro, False
        tInit tRepeatLast, False
        FormMain.MnuClose.Enabled = False
        FormMain.MnuFitWindowToImage.Enabled = False
        FormMain.MnuFitOnScreen.Enabled = False
        If FormMain.CmbZoom.Enabled = True Then
            FormMain.CmbZoom.Enabled = False
            FormMain.lblZoom.ForeColor = &HD1B499
            FormMain.CmbZoom.ListIndex = 15   'Reset zoom to 100%
        End If
        
        FormMain.lblImgSize.ForeColor = &HD1B499
        FormMain.lblCoordinates.ForeColor = &HD1B499
        
        FormMain.lblImgSize.Caption = ""
        
        FormMain.lblCoordinates.Caption = ""
        
        Message "Please load an image.  (The large 'Open Image' button at the top-left should do the trick!)"
        
    'Otherwise, enable all of 'em
    Else
        tInit tFilter, True
        tInit tSave, True
        tInit tSaveAs, True
        tInit tCopy, True
        tInit tUndo, pdImages(CurrentImage).UndoState
        tInit tRedo, pdImages(CurrentImage).RedoState
        tInit tImageOps, True
        tInit tFilter, True
        tInit tHistogram, True
        tInit tMacro, True
        tInit tRepeatLast, pdImages(CurrentImage).RedoState
        FormMain.MnuClose.Enabled = True
        FormMain.MnuFitWindowToImage.Enabled = True
        FormMain.MnuFitOnScreen.Enabled = True
        FormMain.lblImgSize.ForeColor = &H544E43
        FormMain.lblCoordinates.ForeColor = &H544E43
        If FormMain.CmbZoom.Enabled = False Then
            FormMain.CmbZoom.Enabled = True
            FormMain.lblZoom.ForeColor = &H544E43
        End If
    End If
End Sub
