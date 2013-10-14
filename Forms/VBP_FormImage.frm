VERSION 5.00
Begin VB.Form FormImage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Image Window"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   5430
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   362
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
'Last updated: 26/August/13
'Last update: Converted mouse tracking code to use bluMouseEvents
'
'Every time the user loads an image, one of these forms is spawned. This form also interfaces with several
' specialized program components in the MDIWindow module.
'
'As I start including more and more paint tools, this form is going to become a bit more complex. Stay tuned.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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
Dim m_initMouseX As Double, m_initMouseY As Double

'Used to prevent the obnoxious blinking effect of the main image scroll bars
Private Declare Function DestroyCaret Lib "user32" () As Long

'An outside class provides access to mousewheel events for scrolling the filter view
Private WithEvents cMouseEvents As bluMouseEvents
Attribute cMouseEvents.VB_VarHelpID = -1

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Public Sub ActivateWorkaround()

    'Update the current form variable
    g_CurrentImage = Val(Me.Tag)
    
    g_WindowManager.notifyChildReceivedFocus Me
    
    'If possible, set focus to this window
    'NOTE: this throws inexplicable RTEs, so leave it uncommented until debugged.
    'If Me.Visible Then Me.SetFocus
    
    'Display the size of this image in the status bar
    ' (NOTE: because this event will be fired when this form is first built, don't update the size values
    ' unless they actually exist.)
    If pdImages(g_CurrentImage).Width <> 0 Then DisplaySize pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height

    'If we are dynamically updating the taskbar icon to match the current image, we need to update those icons
    If g_UserPreferences.GetPref_Boolean("Interface", "Dynamic Taskbar Icon", True) And (MacroStatus <> MacroBATCH) Then
        If pdImages(g_CurrentImage).curFormIcon32 <> 0 Then
            setNewTaskbarIcon pdImages(g_CurrentImage).curFormIcon32, pdImages(g_CurrentImage).containingForm.hWnd
        Else
            setNewTaskbarIcon origIcon32, FormMain.hWnd
            setNewAppIcon origIcon16
        End If
        If pdImages(g_CurrentImage).curFormIcon16 <> 0 Then setNewAppIcon pdImages(g_CurrentImage).curFormIcon16
    End If

    'If this MDI child is maximized, double-check that it's been drawn correctly.
    ' (This is necessary because VB doesn't handle _Resize() properly when switching between maximized MDI child forms)
    'If Me.WindowState = 2 Then
        'DoEvents
        'PrepareViewport Me, "Maximized MDI child redraw"
    'End If
    
    'Determine whether Undo, Redo, Fade-last are available
    metaToggle tUndo, pdImages(g_CurrentImage).UndoState
    metaToggle tRedo, pdImages(g_CurrentImage).RedoState
    FormMain.MnuFadeLastEffect.Enabled = pdImages(g_CurrentImage).UndoState
    
    'Determine whether save is enabled
    metaToggle tSave, Not pdImages(g_CurrentImage).HasBeenSaved
    
    'Determine whether metadata is present, and dis/enable metadata menu items accordingly
    ' Note that if metadata loading uses the "on-demand" preference, we always set metadata browsing to TRUE so the user can
    ' browse at their leisure.
    If g_UserPreferences.GetPref_Boolean("Loading", "Automatically Load Metadata", True) Then
        
        'We are NOT using the on-demand model.  Determine whether metadata is present, and dis/enable metadata menu items accordingly.
        metaToggle tMetadata, pdImages(g_CurrentImage).imgMetadata.hasXMLMetadata
        metaToggle tGPSMetadata, pdImages(g_CurrentImage).imgMetadata.hasGPSMetadata()
        
    Else
    
        'We ARE using the on-demand model.  Always leave "browse metadata" set to TRUE, and same for GPS metadata, unless we
        ' have attempted to find GPS metadata (either by saving the image or the user manually loading metadata) but not found any.
        ' In that rare case, we can enable the GPS menu correctly.
        metaToggle tMetadata, True
        If pdImages(g_CurrentImage).imgMetadata.haveAttemptedToFindGPSData Then
            metaToggle tGPSMetadata, pdImages(g_CurrentImage).imgMetadata.hasGPSMetadata
        Else
            metaToggle tGPSMetadata, True
        End If
                
    End If
    
    'Check the image's color depth, and check/uncheck the matching Image Mode setting
    If pdImages(g_CurrentImage).mainLayer.getLayerColorDepth() = 32 Then metaToggle tImgMode32bpp, True Else metaToggle tImgMode32bpp, False
    
    'Restore the zoom value for this particular image (again, only if the form has been initialized)
    If pdImages(g_CurrentImage).Width <> 0 Then toolbar_File.CmbZoom.ListIndex = pdImages(g_CurrentImage).CurrentZoomValue
    
    'If a selection is active on this image, update the text boxes to match
    If pdImages(g_CurrentImage).selectionActive Then
        metaToggle tSelection, True
    Else
        metaToggle tSelection, False
    End If
    
    'Finally, if the histogram window is open, redraw it
    If FormHistogram.Visible And pdImages(Me.Tag).loadedSuccessfully Then
        FormHistogram.TallyHistogramValues
        FormHistogram.DrawHistogram
    End If
    
End Sub

'Mousekey back triggers the same thing as clicking Undo
Private Sub cMouseEvents_MouseBackButtonDown(ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
    If pdImages(Me.Tag).IsActive Then
        If pdImages(Me.Tag).UndoState Then Process "Undo", , , False
    End If
End Sub

'Mousekey forward triggers the same thing as clicking Redo
Private Sub cMouseEvents_MouseForwardButtonDown(ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
    If pdImages(Me.Tag).IsActive Then
        If pdImages(Me.Tag).RedoState Then Process "Redo", , , False
    End If
End Sub

Private Sub cMouseEvents_MouseHScroll(ByVal CharsScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)

    'Horizontal scrolling - only trigger if the horizontal scroll bar is visible AND a shift key has been pressed
    If HScroll.Visible And (Not CtrlDown) Then
  
        If CharsScrolled < 0 Then
            
            If HScroll.Value + HScroll.LargeChange > HScroll.Max Then
                HScroll.Value = HScroll.Max
            Else
                HScroll.Value = HScroll.Value + HScroll.LargeChange
            End If
            
            ScrollViewport Me
        
        ElseIf CharsScrolled > 0 Then
            
            If HScroll.Value - HScroll.LargeChange < HScroll.Min Then
                HScroll.Value = HScroll.Min
            Else
                HScroll.Value = HScroll.Value - HScroll.LargeChange
            End If
            
            ScrollViewport Me
            
        End If
        
    End If
  
End Sub

'When the mouse leaves the window, if no buttons are down, clear the coordinate display.
' (We must check for button states because the user is allowed to do things like drag selection nodes outside the image.)
Private Sub cMouseEvents_MouseOut()
    If (Not lMouseDown) And (Not rMouseDown) Then ClearImageCoordinatesDisplay
End Sub

Private Sub cMouseEvents_MouseVScroll(ByVal LinesScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
  
    'Vertical scrolling - only trigger it if the vertical scroll bar is actually visible
    If VScroll.Visible And (Not CtrlDown) Then
      
        If LinesScrolled < 0 Then
            
            If VScroll.Value + VScroll.LargeChange > VScroll.Max Then
                VScroll.Value = VScroll.Max
            Else
                VScroll.Value = VScroll.Value + VScroll.LargeChange
            End If
            
            ScrollViewport Me
        
        ElseIf LinesScrolled > 0 Then
            
            If VScroll.Value - VScroll.LargeChange < VScroll.Min Then
                VScroll.Value = VScroll.Min
            Else
                VScroll.Value = VScroll.Value - VScroll.LargeChange
            End If
            
            ScrollViewport Me
            
        End If
    
    End If
    
    'NOTE: horizontal scrolling is now handled in the separate _MouseHScroll event.  This is necessary to handle mice with
    '      a dedicated horizontal scroller.
    
    'Zooming - only trigger when Ctrl has been pressed
    If CtrlDown And (Not ShiftDown) Then
      
        If LinesScrolled > 0 Then
            
            If toolbar_File.CmbZoom.ListIndex > 0 Then
                toolbar_File.CmbZoom.ListIndex = toolbar_File.CmbZoom.ListIndex - 1
                'NOTE: a manual call to PrepareViewport is no longer required, as changing the combo box will automatically trigger a redraw
                'PrepareViewport Me, "Ctrl+Mousewheel"
            End If
        
        ElseIf LinesScrolled < 0 Then
            
            If toolbar_File.CmbZoom.ListIndex < (toolbar_File.CmbZoom.ListCount - 1) Then
                toolbar_File.CmbZoom.ListIndex = toolbar_File.CmbZoom.ListIndex + 1
                'PrepareViewport Me, "Ctrl+Mousewheel"
        End If
            
        End If
        
    End If
  
End Sub

'NOTE: _Activate and _GotFocus are confusing in VB6. _Activate will be fired whenever a child form
' gains "focus." _GotFocus will be pre-empted by controls on the form, so do not use it.

'Note also that _Activate has known problems - see http://support.microsoft.com/kb/190634
' This is why ActivateWorkaround exists. Some external functions call that if I know _Activate won't fire properly - see
' the Unload function in this block, for example.
Private Sub Form_Activate()
    ActivateWorkaround
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    
    'If a selection is active, notify it of any changes in the shift key (which is used to request 1:1 selections)
    If pdImages(Me.Tag).selectionActive Then pdImages(Me.Tag).mainSelection.requestSquare ShiftDown
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    
    'If a selection is active, notify it of any changes in the shift key (which is used to request 1:1 selections)
    If pdImages(Me.Tag).selectionActive Then pdImages(Me.Tag).mainSelection.requestSquare ShiftDown
    
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Enable mouse subclassing for events like mousewheel, forward/back keys, enter/leave
    Set cMouseEvents = New bluMouseEvents
    cMouseEvents.Attach Me.hWnd
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
End Sub

'Track which mouse buttons are pressed
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If the main form is disabled, exit
    If Not FormMain.Enabled Then Exit Sub
    
    'If the image has not yet been loaded, exit
    If Not pdImages(Me.Tag).loadedSuccessfully Then Exit Sub
    
    'These variables will hold the corresponding (x,y) coordinates on the IMAGE - not the VIEWPORT.
    ' (This is important if the user has zoomed into an image, and used scrollbars to look at a different part of it.)
    Dim imgX As Double, imgY As Double
    imgX = -1
    imgY = -1
    
    'Check mouse button use
    If Button = vbLeftButton Then
        
        lMouseDown = True
            
        hasMouseMoved = 0
            
        'Remember this location
        m_initMouseX = x
        m_initMouseY = y
            
        'Display the image coordinates under the mouse pointer
        displayImageCoordinates x, y, Me, imgX, imgY
        
        'Any further processing depends on which tool is currently active
        
        Select Case g_CurrentTool
        
            'Rectangular selection
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
            
                'Check to see if a selection is already active.  If it is, see if the user is allowed to transform it.
                If pdImages(Me.Tag).selectionActive Then
                
                    'Check the mouse coordinates of this click.
                    Dim sCheck As Long
                    sCheck = findNearestSelectionCoordinates(x, y, Me)
                    
                    'If that function did not return zero, notify the selection and exit
                    If (sCheck <> 0) And pdImages(Me.Tag).mainSelection.isTransformable Then
                    
                        'If the selection type matches the current selection tool, start transforming the selection.
                        If (pdImages(Me.Tag).mainSelection.getSelectionShape = g_CurrentTool) Then
                        
                            'Back up the current selection settings - those will be saved in a later step as part of the Undo/Redo chain
                            pdImages(Me.Tag).mainSelection.setBackupParamString
                            
                            'Initialize a selection transformation
                            pdImages(Me.Tag).mainSelection.setTransformationType sCheck
                            pdImages(Me.Tag).mainSelection.setInitialTransformCoordinates imgX, imgY
                            
                            Exit Sub
                            
                        'If the selection type does NOT match the current selection tool, select the proper tool, then start transforming
                        ' the selection.
                        Else
                        
                            toolbar_Selections.selectNewTool pdImages(Me.Tag).mainSelection.getSelectionShape
                            
                            'Back up the current selection settings - those will be saved in a later step as part of the Undo/Redo chain
                            pdImages(Me.Tag).mainSelection.setBackupParamString
                            
                            'Initialize a selection transformation
                            pdImages(Me.Tag).mainSelection.setTransformationType sCheck
                            pdImages(Me.Tag).mainSelection.setInitialTransformCoordinates imgX, imgY
                            
                            Exit Sub
                        
                        End If
                                        
                    'If it did return zero, erase any existing selection and start a new one
                    Else
                    
                        'Back up the current selection settings - those will be saved in a later step as part of the Undo/Redo chain
                        pdImages(Me.Tag).mainSelection.setBackupParamString
                    
                        initSelectionByPoint imgX, imgY
                    
                    End If
                
                Else
                    
                    initSelectionByPoint imgX, imgY
                    
                End If
            
        End Select
        
    End If
    
    If Button = vbRightButton Then rMouseDown = True
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    'If the main form is disabled, exit
    If Not FormMain.Enabled Then Exit Sub
    
    'If the image has not yet been loaded, exit
    If Not pdImages(Me.Tag).loadedSuccessfully Then Exit Sub
    
    hasMouseMoved = hasMouseMoved + 1
    
    'These variables will hold the corresponding (x,y) coordinates on the image - NOT the viewport
    Dim imgX As Double, imgY As Double
    imgX = -1
    imgY = -1
    
    'Check the left mouse button
    If lMouseDown Then
    
        Select Case g_CurrentTool
        
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
    
                'First, check to see if a selection is active. (In the future, we will be checking for other tools as well.)
                If pdImages(Me.Tag).selectionActive And pdImages(Me.Tag).mainSelection.isTransformable Then
                                        
                    'Display the image coordinates under the mouse pointer
                    displayImageCoordinates x, y, Me, imgX, imgY
                    
                    'If the SHIFT key is down, notify the selection engine that a square shape is requested
                    pdImages(Me.Tag).mainSelection.requestSquare ShiftDown
                    
                    'Pass new points to the active selection
                    pdImages(Me.Tag).mainSelection.setAdditionalCoordinates imgX, imgY
                    syncTextToCurrentSelection Me.Tag
                                        
                End If
                
                'Force a redraw of the viewport
                If hasMouseMoved > 1 Then RenderViewport Me
                
        End Select
    
    'This else means the LEFT mouse button is NOT down
    Else
    
        Select Case g_CurrentTool
        
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
            
                'Next, check to see if a selection is active. If it is, we need to provide the user with visual cues about their
                ' ability to resize the selection.
                If pdImages(Me.Tag).selectionActive And pdImages(Me.Tag).mainSelection.isTransformable Then
                
                    'This routine will return a best estimate for the location of the mouse.  We then pass its value
                    ' to a sub that will use it to select the most appropriate mouse cursor.
                    Dim sCheck As Long
                    sCheck = findNearestSelectionCoordinates(x, y, Me)
                    
                    'Based on that return value, assign a new mouse cursor to the form
                    setSelectionCursor sCheck
                    
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
        
            Case Else
        
                'Check the location of the mouse to see if it's over the image, and set the cursor accordingly.
                ' (NOTE: at present this has no effect, but once paint tools are implemented, it will be more important.)
                If isMouseOverImage(x, y, Me) Then
                    setArrowCursor Me
                Else
                    setArrowCursor Me
                End If
            
        End Select
        
    End If
        
    'Display the image coordinates under the mouse pointer (but only if this is the currently active image)
    If Me.Tag = g_CurrentImage Then displayImageCoordinates x, y, Me
    
End Sub

'Track which mouse buttons are released
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If the image has not yet been loaded, exit
    If Not pdImages(Me.Tag).loadedSuccessfully Then Exit Sub
        
    'Check mouse buttons
    If Button = vbLeftButton Then
    
        lMouseDown = False
    
        Select Case g_CurrentTool
        
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
            
                'If a selection was being drawn, lock it into place
                If pdImages(Me.Tag).selectionActive Then
                    
                    'Check to see if this mouse location is the same as the initial mouse press. If it is, and that particular
                    ' point falls outside the selection, clear the selection from the image.
                    If ((x = m_initMouseX) And (y = m_initMouseY) And (hasMouseMoved <= 1) And (findNearestSelectionCoordinates(x, y, Me) = 0)) Or ((pdImages(Me.Tag).mainSelection.selWidth <= 0) And (pdImages(Me.Tag).mainSelection.selHeight <= 0)) Then
                        Process "Remove selection", , pdImages(Me.Tag).mainSelection.getSelectionParamString, 2, g_CurrentTool
                    Else
                    
                        'Check to see if all selection coordinates are invalid.  If they are, forget about this selection.
                        If pdImages(Me.Tag).mainSelection.areAllCoordinatesInvalid Then
                            Process "Remove selection", , pdImages(Me.Tag).mainSelection.getSelectionParamString, 2, g_CurrentTool
                        Else
                        
                            'Depending on the type of transformation that may or may not have been applied, call the appropriate processor
                            ' function.  This has no practical purpose at present, except to give the user a pleasant name for this action.
                            Select Case pdImages(Me.Tag).mainSelection.getTransformationType
                            
                                'Creating a new selection
                                Case 0
                                    Process "Create selection", , pdImages(Me.Tag).mainSelection.getSelectionParamString, 2, g_CurrentTool
                                    
                                'Moving an existing selection
                                Case 9
                                    Process "Move selection", , pdImages(Me.Tag).mainSelection.getSelectionParamString, 2, g_CurrentTool
                                    
                                'Anything else is assumed to be resizing an existing selection
                                Case Else
                                    Process "Resize selection", , pdImages(Me.Tag).mainSelection.getSelectionParamString, 2, g_CurrentTool
                                    
                            End Select
                            
                        End If
                        
                    End If
                    
                    'Force a redraw of the screen
                    RenderViewport Me
                    
                Else
                    'If the selection is not active, make sure it stays that way
                    pdImages(Me.Tag).mainSelection.lockRelease
                End If
                
                'Synchronize the selection text box values with the final selection
                syncTextToCurrentSelection Me.Tag
                
            Case Else
                    
        End Select
                        
    End If
    
    If Button = vbRightButton Then rMouseDown = False
    
    makeFormPretty Me
    setArrowCursorToHwnd Me.hWnd
        
    'Reset the mouse movement tracker
    hasMouseMoved = 0
    
End Sub

'(This code is copied from FormMain's OLEDragOver event - please mirror any changes there)
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub

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
    If Not g_AllowDragAndDrop Then Exit Sub

    'Check to make sure the type of OLE object is files
    If Data.GetFormat(vbCFFiles) Then
        'Inform the source (Explorer, in this case) that the files will be treated as "copied"
        Effect = vbDropEffectCopy And Effect
    Else
        'If it's not files, don't allow a drop
        Effect = vbDropEffectNone
    End If

End Sub

'In VB6, _QueryUnload fires before _Unload. We check for unsaved images here.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'If the user wants to be prompted about unsaved images, do it now
    If g_ConfirmClosingUnsaved And pdImages(Me.Tag).IsActive And (Not pdImages(Me.Tag).forInternalUseOnly) Then
    
        'Check the .HasBeenSaved property of the image associated with this form
        If Not pdImages(Me.Tag).HasBeenSaved Then
                        
            'If the user hasn't already told us to deal with all unsaved images in the same fashion, run some checks
            If g_DealWithAllUnsavedImages = False Then
            
                g_NumOfUnsavedImages = 0
                                
                'Loop through all images to count how many unsaved images there are in total.
                ' NOTE: we only need to do this if the entire program is being shut down or if the user has selected "close all";
                ' otherwise, this close action only affects the current image, so we shouldn't present a "repeat for all images" option
                If g_ProgramShuttingDown Or g_ClosingAllImages Then
                    Dim i As Long
                    For i = 1 To g_NumOfImagesLoaded
                        If pdImages(i).IsActive And (Not pdImages(i).forInternalUseOnly) And (Not pdImages(i).HasBeenSaved) Then
                            g_NumOfUnsavedImages = g_NumOfUnsavedImages + 1
                        End If
                    Next i
                End If
            
                'Before displaying the "do you want to save this image?" dialog, bring the image in question to the foreground.
                If FormMain.Enabled Then Me.SetFocus
            
                'Show the "do you want to save this image?" dialog. On that form, the number of unsaved images will be
                ' displayed and the user will be given an option to apply their choice to all unsaved images.
                Dim confirmReturn As VbMsgBoxResult
                confirmReturn = confirmClose(Me.Tag)
                        
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
                If FormMain.Enabled Then Me.SetFocus
                
                'Attempt to save. Note that the user can still cancel at this point, and we want to honor their cancellation
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
                
                'I think this "Unload Me" statement may be causing some kind of infinite recursion - perhaps because it triggers this very
                ' QueryUnload statement? Not sure, but I may need to revisit it if the problems don't go away...
                'UPDATE 26 Aug 2013: after changing my subclassing code, the problem seems to have disappeared, but I'm leaving
                ' this comment here until I'm absolutely certain the problem has been resolved.
                Unload Me
                'Set Me = Nothing
                'Cancel = False
                'Me.Visible = False
            
            End If
        
        End If
    
    End If
    
End Sub

Private Sub Form_Resize()
    
    'Redraw this form if certain criteria are met (image loaded, form visible, viewport adjustments allowed)
    If (pdImages(Me.Tag).Width > 0) And (pdImages(Me.Tag).Height > 0) And Me.Visible And (FormMain.WindowState <> vbMinimized) And (g_WindowManager.getClientWidth(Me.hWnd) > 0) Then
        PrepareViewport Me, "Form_Resize(" & Me.ScaleWidth & "," & Me.ScaleHeight & ")"
    End If
    
    'The height of a newly created form is automatically set to 1. This is normally changed when the image is
    ' resized to fit on screen, but if an image is loaded into a maximized window, the height value will remain
    ' at 1. If the user ever un-maximized the window, it will leave a bare title bar behind, which looks
    ' terrible. Thus, let's check for a height of 1, and if found resize the form to a larger (arbitrary) value.
    'If (Me.WindowState = vbNormal) And (Me.ScaleHeight <= 1) Then
    '    Me.Height = 6000
    '    Me.Width = 8000
    'End If
    
    'Remember this window state in the relevant pdImages object
    pdImages(Me.Tag).WindowState = Me.WindowState
    
    Exit Sub
    
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
            For i = 1 To g_NumOfImagesLoaded
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
            
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Unload the mouse tracker
    Set cMouseEvents = Nothing
    
    Message "Closing image..."
    
    g_OpenImageCount = g_OpenImageCount - 1
            
    Me.Visible = False
    
    'Deactivate this layer
    pdImages(Me.Tag).deactivateImage
    
    'Remove any undo files associated with this layer
    ClearUndo Me.Tag
    
    'If this was the last (or only) open image and the histogram is loaded, unload the histogram
    ' (If we don't do this, the histogram may attempt to update, and without an active image it will throw an error)
    If g_OpenImageCount = 0 Then Unload FormHistogram
    
    UpdateMDIStatus
    
    ReleaseFormTheming Me
    
    'Notify the window manager that this hWnd will soon be dead - so stop subclassing it!
    g_WindowManager.unregisterForm Me
    
    'Before exiting, restore focus to the next child window in line.
    If g_OpenImageCount > 0 Then
    
        Dim i As Long
        For i = g_NumOfImagesLoaded To 0 Step -1
            If (Not pdImages(i) Is Nothing) Then
                If pdImages(i).IsActive Then
                    pdImages(i).containingForm.ActivateWorkaround
                    Exit For
                End If
            End If
        Next i
    
    'If this is the last window, we need to make the toolbars children of the main form again.  Otherwise, they will be mistakenly unloaded.
    Else
        g_WindowManager.allImageWindowsUnloaded
    End If
    
    Message "Finished."
            
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

'Selection tools utilize a variety of cursors.  To keep the main MouseMove sub clean, cursors are set separately
' by this routine.
Private Sub setSelectionCursor(ByVal transformID As Long)

    Select Case pdImages(Me.Tag).mainSelection.getSelectionShape()

        Case sRectangle, sCircle
        
            'For a rectangle or circle selection, the possible transform IDs are:
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
            Select Case transformID
        
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
            
        'For a line selection, the possible transform IDs are:
        ' 0 - Cursor is not near an endpoint
        ' 1 - Near x1/y1
        ' 2 - Near x2/y2
        Case sLine
        
            Select Case transformID
                Case 0
                    setArrowCursor Me
                Case 1
                    setSizeAllCursor Me
                Case 2
                    setSizeAllCursor Me
            End Select
        
    End Select

End Sub

'Selections can be initiated several different ways.  To cut down on duplicated code, all new selection instances for this form are referred
' to this function.  Initial X/Y values are required.
Private Sub initSelectionByPoint(ByVal x As Double, ByVal y As Double)

    'I don't have a good explanation, but without DoEvents here, creating a line selection for the first
    ' time may inexplicably fail.  While I try to track down the exact cause, I'll leave this here to
    ' maintain desired behavior...
    DoEvents
    
    'Activate the attached image's primary selection
    pdImages(Me.Tag).selectionActive = True
    pdImages(Me.Tag).mainSelection.lockRelease
    
    'Populate a variety of selection attributes using a single shorthand declaration.  A breakdown of these
    ' values and what they mean can be found in the corresponding pdSelection.initFromParamString function
    pdImages(Me.Tag).mainSelection.initFromParamString buildParams(g_CurrentTool, toolbar_Selections.cmbSelType(0).ListIndex, toolbar_Selections.cmbSelSmoothing(0).ListIndex, toolbar_Selections.sltSelectionFeathering.Value, toolbar_Selections.sltSelectionBorder.Value, toolbar_Selections.sltCornerRounding.Value, toolbar_Selections.sltSelectionLineWidth.Value, 0, 0, 0, 0, 0, 0, 0, 0)
    
    'Set the first two coordinates of this selection to this mouseclick's location
    pdImages(Me.Tag).mainSelection.setInitialCoordinates x, y
    syncTextToCurrentSelection Me.Tag
    pdImages(Me.Tag).mainSelection.requestNewMask
        
    'Make the selection tools visible
    metaToggle tSelection, True
                        
End Sub
