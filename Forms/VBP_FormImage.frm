VERSION 5.00
Begin VB.Form FormImage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Image Window"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VBP_FormImage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   172
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.HScrollBar HScroll 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.VScrollBar VScroll 
      Height          =   3615
      LargeChange     =   10
      Left            =   6240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3960
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "FormImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Image Form (Child MDI form)
'Copyright ©2002-2013 by Tanner Helland
'Created: 11/29/02
'Last updated: 09/September/12
'Last update: manually unload this image's main layer when the attached form is unloaded (to conserve memory)
'
'Every time the user loads an image, one of these forms is spawned.  This form also interfaces with several
' specialized program components in the MDIWindow module.  Look there for more information.
'
'***************************************************************************

Option Explicit

'These are used to track use of the Ctrl, Alt, and Shift keys
Dim ShiftDown As Boolean, CtrlDown As Boolean, AltDown As Boolean

'Track mouse button use on this form
Dim lMouseDown As Boolean, rMouseDown As Boolean

'Track mouse movement on this form
Dim hasMouseMoved As Long

'Track initial mouse button locations
Dim initMouseX As Double, initMouseY As Double

'Used to prevent the obnoxious blinking effect of the main image scroll bars
Private Declare Function DestroyCaret Lib "user32" () As Long
    
'NOTE: _Activate and _GotFocus are confusing in VB6.  _Activate will be fired whenever a child form
' gains "focus."  _GotFocus will be pre-empted by controls on the form, so do not use it.

Private Sub Form_Activate()

    'Update the current form variable
    CurrentImage = Val(Me.Tag)
    
    'Display the size of this image in the status bar
    ' (NOTE: because this event will be fired when this form is first built, don't update the size values
    ' unless they actually exist.)
    If pdImages(CurrentImage).Width <> 0 Then DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height

    'If we are dynamically updating the taskbar icon to match the current image, we need to update those icons
    If g_UserPreferences.GetPreference_Boolean("General Preferences", "DynamicTaskbarIcon", True) And (MacroStatus <> MacroBATCH) Then
        If pdImages(CurrentImage).curFormIcon32 <> 0 Then
            setNewTaskbarIcon pdImages(CurrentImage).curFormIcon32
        Else
            setNewTaskbarIcon origIcon32
            setNewAppIcon origIcon16
        End If
        If pdImages(CurrentImage).curFormIcon16 <> 0 Then setNewAppIcon pdImages(CurrentImage).curFormIcon16
    End If

    'If this MDI child is maximized, double-check that it's been drawn correctly.
    ' (This is necessary because VB doesn't handle _Resize() properly when switching between maximized MDI child forms)
    If Me.WindowState = 2 Then
        DoEvents
        PrepareViewport Me, "Maximized MDI child redraw"
    End If
    
    'Determine whether Undo, Redo, Fade-last are available
    tInit tUndo, pdImages(CurrentImage).UndoState
    tInit tRedo, pdImages(CurrentImage).RedoState
    FormMain.MnuFadeLastEffect.Enabled = pdImages(CurrentImage).UndoState
    
    'Determine whether save is enabled
    tInit tSave, Not pdImages(CurrentImage).HasBeenSaved
    
    'Check the image's color depth, and check/uncheck the matching Image Mode setting
    If pdImages(CurrentImage).mainLayer.getLayerColorDepth() = 32 Then tInit tImgMode32bpp, True Else tInit tImgMode32bpp, False
    
    'Restore the zoom value for this particular image (again, only if the form has been initialized)
    If pdImages(CurrentImage).Width <> 0 Then FormMain.CmbZoom.ListIndex = pdImages(CurrentImage).CurrentZoomValue
    
    'If a selection is active on this image, update the text boxes to match
    If pdImages(CurrentImage).selectionActive Then
        tInit tSelection, True
        pdImages(CurrentImage).mainSelection.refreshTextBoxes
    Else
        tInit tSelection, False
    End If
    
    'Finally, if the histogram window is open, redraw it
    If (FormHistogram.Visible = True) And pdImages(Me.Tag).loadedSuccessfully Then
        FormHistogram.TallyHistogramValues
        FormHistogram.DrawHistogram
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    AltDown = (Shift And vbAltMask) > 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    AltDown = (Shift And vbAltMask) > 0
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Add support for scrolling with the mouse wheel
    If g_IsProgramCompiled Then Call WheelHook(Me.hWnd)
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

'Track which mouse buttons are pressed
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If the main form is disabled, exit
    If FormMain.Enabled = False Then Exit Sub
    
    'If the image has not yet been loaded, exit
    If pdImages(Me.Tag).loadedSuccessfully = False Then Exit Sub
    
    'These variables will hold the corresponding (x,y) coordinates on the image - NOT the viewport
    Dim imgX As Double, imgY As Double
    imgX = -1
    imgY = -1
    
    'Check mouse button use
    If Button = vbLeftButton Then
        
        'Check the location of the mouse to see if it's over the image
        If isMouseOverImage(x, y, Me) Then
        
            lMouseDown = True
            
            hasMouseMoved = 0
            
            'Remember this location
            initMouseX = x
            initMouseY = y
            
            'Display the image coordinates under the mouse pointer
            displayImageCoordinates x, y, Me, imgX, imgY
        
            'Check to see if a selection is already active.
            If pdImages(Me.Tag).selectionActive Then
            
                'Check the mouse coordinates of this click.
                Dim sCheck As Long
                sCheck = findNearestSelectionCoordinates(x, y, Me)
                
                'If that function did not return zero, notify the selection and exit
                If sCheck <> 0 Then
                
                    pdImages(Me.Tag).mainSelection.setTransformationType sCheck
                    pdImages(Me.Tag).mainSelection.setInitialTransformCoordinates imgX, imgY
                    
                    Exit Sub
                
                End If
            
            End If
                
            'Activate the selection and pass in the first two points
            pdImages(Me.Tag).selectionActive = True
            pdImages(Me.Tag).mainSelection.selLeft = 0
            pdImages(Me.Tag).mainSelection.selTop = 0
            pdImages(Me.Tag).mainSelection.selWidth = 0
            pdImages(Me.Tag).mainSelection.selHeight = 0
            pdImages(Me.Tag).mainSelection.setInitialCoordinates imgX, imgY
            pdImages(Me.Tag).mainSelection.refreshTextBoxes
            
            'Make the selection tools visible
            tInit tSelection, True

            'Render the new selection
            RenderViewport Me
            
        End If
        
    End If
    
    If Button = vbRightButton Then rMouseDown = True
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If the main form is disabled, exit
    If FormMain.Enabled = False Then Exit Sub
    
    'If the image has not yet been loaded, exit
    If pdImages(Me.Tag).loadedSuccessfully = False Then Exit Sub
    
    hasMouseMoved = hasMouseMoved + 1
    
    'These variables will hold the corresponding (x,y) coordinates on the image - NOT the viewport
    Dim imgX As Double, imgY As Double
    imgX = -1
    imgY = -1
    
    'Check the left mouse button
    If lMouseDown Then
    
        'First, check to see if a selection is active.  (In the future, we will be checking for other tools as well.)
        If pdImages(Me.Tag).selectionActive Then
            
            'Check the location of the mouse to see if it's over the image
            If isMouseOverImage(x, y, Me) Then
            
                'Display the image coordinates under the mouse pointer
                displayImageCoordinates x, y, Me, imgX, imgY
            
                'Pass new points to the active selection
                pdImages(Me.Tag).mainSelection.setAdditionalCoordinates imgX, imgY
            
            'If the mouse coordinates are NOT over the image, we need to find the closest points in the image and pass those instead
            Else
        
                imgX = x
                imgY = y
                findNearestImageCoordinates imgX, imgY, Me
                
                'Pass those points to the active selection
                pdImages(Me.Tag).mainSelection.setAdditionalCoordinates imgX, imgY
            
            End If
            
        End If
        
        'Force a redraw of the viewport
        If hasMouseMoved > 1 Then RenderViewport Me
    
    'This else means the LEFT mouse button is NOT down
    Else
    
        'Next, check to see if a selection is active.  If it is, we need to provide the user with visual cues about their
        ' ability to resize the selection.
        If pdImages(Me.Tag).selectionActive Then
        
            'This routine will return a best estimate for the location of the mouse.  The possible return values are:
            ' 0 - Cursor is not near a selection point
            ' 1 - NW corner
            ' 2 - NE corner
            ' 3 - SE corner
            ' 4 - SW corner
            ' 5 - N edge
            ' 6 - E edge
            ' 7 - S edge
            ' 8 - W edge
            ' 9 - interior of selection, not near a corner or edge
            Dim sCheck As Long
            sCheck = findNearestSelectionCoordinates(x, y, Me)
            
            'Based on that return value, assign a new mouse cursor to the form
            Select Case sCheck
                
                Case 0
                    setArrowCursor Me
                Case 1
                    setSizeNWSECursor Me
                Case 2
                    setSizeNESWCursor Me
                Case 3
                    setSizeNWSECursor Me
                Case 4
                    setSizeNESWCursor Me
                Case 5
                    setSizeNSCursor Me
                Case 6
                    setSizeWECursor Me
                Case 7
                    setSizeNSCursor Me
                Case 8
                    setSizeWECursor Me
                Case 9
                    setSizeAllCursor Me
                    
            End Select
        
            'Set the active selection's transformation type to match
            pdImages(Me.Tag).mainSelection.setTransformationType sCheck
        
        Else
        
            'Check the location of the mouse to see if it's over the image, and set the cursor accordingly.
            ' (NOTE: at present this has no effect, but once paint tools are implemented, it will be more important.)
            If isMouseOverImage(x, y, Me) Then
                setArrowCursor Me
            Else
                setArrowCursor Me
            End If
            
        End If
        
    End If
        
    'Display the image coordinates under the mouse pointer
    displayImageCoordinates x, y, Me
    
End Sub

'Track which mouse buttons are released
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If the image has not yet been loaded, exit
    If pdImages(Me.Tag).loadedSuccessfully = False Then Exit Sub
        
    'Check mouse buttons
    If Button = vbLeftButton Then
    
        lMouseDown = False
    
        'If a selection was being drawn, lock it into place
        If pdImages(Me.Tag).selectionActive Then
            
            'Check to see if this mouse location is the same as the initial mouse press.  If it is, and that particular
            ' point falls outside the selection, clear the selection from the image.
            If ((x = initMouseX) And (y = initMouseY) And (hasMouseMoved <= 1) And (findNearestSelectionCoordinates(x, y, Me) = 0)) Or ((pdImages(Me.Tag).mainSelection.selWidth <= 0) And (pdImages(Me.Tag).mainSelection.selHeight <= 0)) Then
                pdImages(Me.Tag).mainSelection.lockRelease
                pdImages(Me.Tag).selectionActive = False
                tInit tSelection, False
            Else
            
                'Lock the selection
                pdImages(Me.Tag).mainSelection.lockIn Me
                tInit tSelection, True
                
                'Message x & "," & initMouseX & "-" & y & "," & initMouseY & "-" & hasMouseMoved
            
            End If
            
            'Force a redraw of the screen
            RenderViewport Me
            
        Else
        
            'If the selection is not active, make sure it stays that way
            pdImages(Me.Tag).mainSelection.lockRelease
        
            
        End If
                        
    End If
    
    If Button = vbRightButton Then rMouseDown = False
    
    makeFormPretty Me
    
    'Reset the mouse movement tracker
    hasMouseMoved = 0
    
End Sub

'(This code is copied from FormMain's OLEDragOver event - please mirror any changes there)
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If FormMain.Enabled = False Then Exit Sub

    'Verify that the object being dragged is some sort of file or file list
    If Data.GetFormat(vbCFFiles) Then
        
        'Copy the filenames into an array
        Dim sFile() As String
        ReDim sFile(0 To Data.Files.Count) As String
        
        Dim oleFilename
        Dim tmpString As String
        
        Dim countFiles As Long
        countFiles = 0
        
        For Each oleFilename In Data.Files
            tmpString = CStr(oleFilename)
            If tmpString <> "" Then
                sFile(countFiles) = tmpString
                countFiles = countFiles + 1
            End If
        Next oleFilename
        
        'Because the OLE drop may include blank strings, verify the size of the array against countFiles
        ReDim Preserve sFile(0 To countFiles - 1) As String
        
        'Pass the list of filenames to PreLoadImage, which will load the images one-at-a-time
        PreLoadImage sFile
        
    End If
    
End Sub

'(This code is copied from FormMain's OLEDragOver event - please mirror any changes there)
Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If FormMain.Enabled = False Then Exit Sub

    'Check to make sure the type of OLE object is files
    If Data.GetFormat(vbCFFiles) Then
        'Inform the source (Explorer, in this case) that the files will be treated as "copied"
        Effect = vbDropEffectCopy And Effect
    Else
        'If it's not files, don't allow a drop
        Effect = vbDropEffectNone
    End If

End Sub

'In VB6, _QueryUnload fires before _Unload.  We check for unsaved images here.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'If the user wants to be prompted about unsaved images, do it now
    If g_ConfirmClosingUnsaved And pdImages(Me.Tag).IsActive And (Not pdImages(Me.Tag).forInternalUseOnly) Then
    
        'Check the .HasBeenSaved property of the image associated with this form
        If pdImages(Me.Tag).HasBeenSaved = False Then
                        
            'If the user hasn't already told us to deal with all unsaved images in the same fashion, run some checks
            If g_DealWithAllUnsavedImages = False Then
            
                g_NumOfUnsavedImages = 0
                                
                'Loop through all images to count how many unsaved images there are in total.
                ' NOTE: we only need to do this if the entire program is being shut down or if the user has selected "close all";
                ' otherwise, this close action only affects the current image, so we shouldn't present a "repeat for all images" option
                If g_ProgramShuttingDown Or g_ClosingAllImages Then
                    Dim i As Long
                    For i = 1 To NumOfImagesLoaded
                        If pdImages(i).IsActive And (Not pdImages(i).forInternalUseOnly) And (Not pdImages(i).HasBeenSaved) Then
                            g_NumOfUnsavedImages = g_NumOfUnsavedImages + 1
                        End If
                    Next i
                End If
            
                'Show the "do you want to save this image?" dialog.  On that form, the number of unsaved images will be
                ' displayed and the user will be given an option to apply their choice to all unsaved images.
                Dim confirmReturn As VbMsgBoxResult
                confirmReturn = confirmClose(Me.Tag)
                        
            Else
                confirmReturn = g_HowToDealWithAllUnsavedImages
            End If
        
            'There are now three possible courses of action:
            ' 1) The user canceled.  Quit and abandon all notion of closing.
            ' 2) The user asked us to save this image.  Pass control to MenuSave (which will in turn call SaveAs if necessary)
            ' 3) The user doesn't give a shit.  Exit without saving.
            
            'Cancel the close operation
            If confirmReturn = vbCancel Then
                
                Cancel = True
                If g_ProgramShuttingDown Then g_ProgramShuttingDown = False
                If g_ClosingAllImages Then g_ClosingAllImages = False
                g_DealWithAllUnsavedImages = False
                
            'Save the image
            ElseIf confirmReturn = vbYes Then
                
                'If the form being saved is enabled, bring that image to the foreground.  (If a "Save As" is required, this
                ' is the only way to show the user which image the Save As form is referencing.)
                If FormMain.Enabled Then
                    Me.SetFocus
                    DoEvents
                End If
                
                'Attempt to save.  Note that the user can still cancel at this point, and we want to honor their cancellation
                Dim saveSuccessful As Boolean
                saveSuccessful = MenuSave(CLng(Me.Tag))
                
                'If something went wrong, or the user canceled the save dialog, stop the unload process
                Cancel = Not saveSuccessful
 
                'If we make it here and the save was successful, force an immediate unload
                If Cancel = False Then
                    Unload Me
                
                '...but if the save was not successful, suspend all unload action
                Else
                    If g_ProgramShuttingDown Then g_ProgramShuttingDown = False
                    If g_ClosingAllImages Then g_ClosingAllImages = False
                    g_DealWithAllUnsavedImages = False
                End If
            
            'Do not save the image
            ElseIf confirmReturn = vbNo Then
                Unload Me
            
            End If
        
        End If
    
    End If
    
End Sub

Private Sub Form_Resize()
    
    'Redraw this form if certain criteria are met (image loaded, form visible, viewport adjustments allowed)
    If (pdImages(Me.Tag).Width > 0) And (pdImages(Me.Tag).Height > 0) And (Me.Visible = True) And (FormMain.WindowState <> vbMinimized) Then
        PrepareViewport Me, "Form_Resize(" & Me.ScaleWidth & "," & Me.ScaleHeight & ")"
    End If
    
    'The height of a newly created form is automatically set to 1.  This is normally changed when the image is
    ' resized to fit on screen, but if an image is loaded into a maximized window, the height value will remain
    ' at 1.  If the user ever un-maximized the window, it will leave a bare title bar behind, which looks
    ' terrible.  Thus, let's check for a height of 1, and if found resize the form to a larger (arbitrary) value.
    If (Me.WindowState = vbNormal) And (Me.ScaleHeight <= 1) Then
        Me.Height = 6000
        Me.Width = 8000
    End If
    
    Dim i As Long
    
    'If the window is being un-maximized, it's necessary to redraw every image buffer (to check for scroll bar enabling/disabling)
    If pdImages(Me.Tag).WindowState = vbMaximized And Me.WindowState = 0 Then
        
        'Run a loop through every child form to see if all windows are being un-maximized
        ' (This will only happen when the user presses the "unmaximize" window button)
        Dim allShrunk As Boolean
        allShrunk = True
        
        Dim tForm As Form
        For Each tForm In VB.Forms
            If tForm.Name = "FormImage" Then
                If tForm.WindowState = vbMaximized Then allShrunk = False
            End If
        Next
        
        'If the user has unmaximized all windows, we need to redraw them
        If allShrunk = True Then
        
            'Loop through every image, redrawing as we go
            For i = 1 To NumOfImagesLoaded
                If pdImages(i).IsActive Then
                    
                    'Remember this new window state and redraw the form containing this image
                    pdImages(i).WindowState = 0
                    PrepareViewport pdImages(i).containingForm, "Form_Resize(), user unmaximized MDI children"
                    
                    'While we're at it, make sure the images aren't still hidden off-form (which can happen if they were loaded while the window was maximized)
                    If pdImages(i).containingForm.Left >= FormMain.ScaleWidth Then pdImages(i).containingForm.Left = pdImages(i).WindowLeft
                    If pdImages(i).containingForm.Top >= FormMain.ScaleHeight Then pdImages(i).containingForm.Top = pdImages(i).WindowTop
    
                End If
            Next i
        End If
        
    End If
    
    'Remember this window state in the relevant pdImages object
    pdImages(Me.Tag).WindowState = Me.WindowState
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Message "Closing image..."
    NumOfWindows = NumOfWindows - 1
    
    'Release mouse wheel support
    If g_IsProgramCompiled Then Call WheelUnHook(Me.hWnd)
        
    Me.Visible = False
    
    'Deactivate this layer
    pdImages(Me.Tag).deactivateImage
    
    'Remove any undo files associated with this layer
    ClearUndo Me.Tag
    
    Message "Finished."

    'If this was the last (or only) open image and the histogram is loaded, unload the histogram
    ' (If we don't do this, the histogram may attempt to update, and without an active image it will throw an error)
    If NumOfWindows = 0 Then Unload FormHistogram
    
    UpdateMDIStatus
    
    ReleaseFormTheming Me
    
End Sub

Private Sub HScroll_Change()
    ScrollViewport Me
End Sub

Private Sub HScroll_GotFocus()
    DestroyCaret
End Sub

Private Sub HScroll_Scroll()
    ScrollViewport Me
End Sub

Private Sub VScroll_Change()
    ScrollViewport Me
End Sub

Private Sub VScroll_GotFocus()
    DestroyCaret
End Sub

Private Sub VScroll_Scroll()
    ScrollViewport Me
End Sub

'In VB6, a routine this like is required to support use of a mouse wheel.
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  
  On Error Resume Next
  
  'Vertical scrolling - only trigger it if the vertical scroll bar is actually visible
  If (VScroll.Visible = True) And (Not ShiftDown) And (Not CtrlDown) Then
  
    If rotation < 0 Then
        
        If VScroll.Value + VScroll.LargeChange > VScroll.Max Then
            VScroll.Value = VScroll.Max
        Else
            VScroll.Value = VScroll.Value + VScroll.LargeChange
        End If
        
        ScrollViewport Me
    
    ElseIf rotation > 0 Then
        
        If VScroll.Value - VScroll.LargeChange < VScroll.Min Then
            VScroll.Value = VScroll.Min
        Else
            VScroll.Value = VScroll.Value - VScroll.LargeChange
        End If
        
        ScrollViewport Me
        
    End If
  End If
  
  'Horizontal scrolling - only trigger if the horizontal scroll bar is visible AND a shift key has been pressed
  If (HScroll.Visible = True) And ShiftDown And (Not CtrlDown) Then
  
    If rotation < 0 Then
        
        If HScroll.Value + HScroll.LargeChange > HScroll.Max Then
            HScroll.Value = HScroll.Max
        Else
            HScroll.Value = HScroll.Value + HScroll.LargeChange
        End If
        
        ScrollViewport Me
    
    ElseIf rotation > 0 Then
        
        If HScroll.Value - HScroll.LargeChange < HScroll.Min Then
            HScroll.Value = HScroll.Min
        Else
            HScroll.Value = HScroll.Value - HScroll.LargeChange
        End If
        
        ScrollViewport Me
        
    End If
  End If
  
  'Zooming - only trigger when Ctrl has been pressed
  If CtrlDown And (Not ShiftDown) Then
  
    If rotation > 0 Then
        
        If FormMain.CmbZoom.ListIndex > 0 Then
            FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex - 1
            PrepareViewport Me, "Ctrl+Mousewheel"
        End If
    
    ElseIf rotation < 0 Then
        
        If FormMain.CmbZoom.ListIndex < (FormMain.CmbZoom.ListCount - 1) Then
            FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex + 1
            PrepareViewport Me, "Ctrl+Mousewheel"
        End If
        
    End If
  End If
    
End Sub

