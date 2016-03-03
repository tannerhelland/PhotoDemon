VERSION 5.00
Begin VB.UserControl pdCanvas 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13290
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   KeyPreview      =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   886
   ToolboxBitmap   =   "pdCanvas.ctx":0000
   Begin PhotoDemon.pdStatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   7350
      Width           =   13290
      _ExtentX        =   23442
      _ExtentY        =   609
   End
   Begin PhotoDemon.pdCanvasView CanvasView 
      Height          =   4935
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   4575
      _ExtentX        =   8281
      _ExtentY        =   8916
   End
   Begin VB.PictureBox picProgressBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   886
      TabIndex        =   0
      Top             =   7095
      Visible         =   0   'False
      Width           =   13290
   End
   Begin PhotoDemon.pdButtonToolbox cmdCenter 
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      AutoToggle      =   -1  'True
      BackColor       =   -2147483626
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdScrollBar hScroll 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      OrientationHorizontal=   -1  'True
      VisualStyle     =   1
   End
   Begin PhotoDemon.pdScrollBar vScroll 
      Height          =   4935
      Left            =   5040
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   8705
      VisualStyle     =   1
   End
End
Attribute VB_Name = "pdCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Canvas User Control (previously a standalone form)
'Copyright 2002-2016 by Tanner Helland
'Created: 29/November/02
'Last updated: 03/March/16
'Last update: migrate status bar into its own dedicated control
'
'In 2013, PD's canvas was rebuilt as a dedicated user control, and instead of each image maintaining its own canvas inside
' separate, dedicated windows (which required a *ton* of code to keep in sync with the main PD window), a single canvas was
' integrated directly into the main window, and shared by all windows.
'
'Technically, the primary canvas is only the first entry in an array.  This was done deliberately in case I ever added support for
' multiple canvases being usable at once.  This has some neat possibilities - for example, having side-by-side canvases at
' different locations on an image - but there's a lot of messy UI considerations with something like this, especially if the two
' viewports can support different images simultaneously.  So I have postponed this work until some later date, with the caveat
' that implementing it will be a lot of work, and likely have unexpected interactions throughout the program.
'
'This canvas relies on pdInputMouse for all mouse interactions.  See the pdInputMouse class for details on why we do our own mouse
' management instead of using VB's intrinsic mouse functions.
'
'As much as possible, I've tried to keep paint tool operation within this canvas to a minimum.  Generally speaking, the only tool
' interactions the canvas should handle is reporting mouse events to external functions that actually handle paint tool processing
' and rendering.  To that end, try to adhere to the existing tool implementation format when adding new tool support.  (Selections
' are currently the exception to this rule, because they were implemented long before other tools and thus aren't as
' well-contained.  I hope to someday remedy this.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Enum PD_MOUSEEVENT
    pMouseDown = 0
    pMouseMove = 1
    pMouseUp = 2
End Enum

#If False Then
    Private Const pMouseDown = 0, pMouseMove = 1, pMouseUp = 2
#End If

Private Const SM_CXVSCROLL As Long = 2
Private Const SM_CYHSCROLL As Long = 3

'Mouse interactions are complicated in this form, so we sometimes need to cache button values and process them elsewhere
Private m_LMBDown As Boolean, m_RMBDown As Boolean

'Every time a canvas MouseMove event occurs, this number is incremented by one.  If mouse events are coming in fast and furious,
' we can delay renders between them to improve responsiveness.  (This number is reset to zero when the mouse is released.)
Private m_NumOfMouseMovements As Long

'If the mouse is currently over the canvas, this will be set to TRUE.
Private m_IsMouseOverCanvas As Boolean

'Track initial mouse button locations
Private m_InitMouseX As Double, m_InitMouseY As Double

'On the canvas's MouseDown event, this control will mark the relevant point of interest index for the active layer (if any).
' If a point of interest has not been selected, this value will be reset to -1.
Private m_CurPointOfInterest As Long

'As some POI interactions may cause the canvas to redraw, we also cache the *last* point of interest.  When this mismatches the
' current one, a UI-only viewport redraw is requested, and the last/current point values are synched.
Private m_LastPointOfInterest As Long

'To improve performance, we can ask the canvas to not refresh itself until we say so.
Private m_SuspendRedraws As Boolean

'Some tools support the ability to auto-activate a layer beneath the mouse.  If supported, during the MouseMove event,
' this value (m_LayerAutoActivateIndex) will be updated with the index of the layer that will be auto-activated if the
' user presses the mouse button.  This can be used to modify things like cursor behavior, to make sure the user receives
' accurate feedback on what a given action will affect.
Private m_LayerAutoActivateIndex As Long

'Selection tools need to know if a selection was active before mouse events start.  If it is, creation of an invalid new selection
' will add "Remove Selection" to the Undo/Redo chain; however, if no selection was active, the working selection will simply
' be erased.
Private m_SelectionActiveBeforeMouseEvents As Boolean

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDCANVAS_COLOR_LIST
    [_First] = 0
    PDC_Background = 0
    PDC_StatusBar = 1
    [_Last] = 1
    [_Count] = 2
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'External functions can call this to set the current network state (which in turn, draws a relevant icon to the status bar)
Public Sub SetNetworkState(ByVal newNetworkState As Boolean)
    StatusBar.SetNetworkState newNetworkState
End Sub

'Use these functions to forcibly prevent the canvas from redrawing itself.  REDRAWS WILL NOT HAPPEN AGAIN UNTIL YOU RESTORE ACCESS!
' (Also note that this function relays state changes to the underlying pdCanvasView object; as such, do not set m_SuspendRedraws
'  manually - only set it via this function, to ensure the canvas and underlying canvas view stay in sync.)
Public Function GetRedrawSuspension() As Boolean
    GetRedrawSuspension = m_SuspendRedraws Or CanvasView.GetRedrawSuspension()
End Function

Public Sub SetRedrawSuspension(ByVal newRedrawValue As Boolean)
    CanvasView.SetRedrawSuspension newRedrawValue
    m_SuspendRedraws = newRedrawValue
End Sub

'Need to wipe the canvas?  Call this function, but please be careful - it will literally erase the canvas's back buffer.
Public Sub ClearCanvas()
    
    CanvasView.ClearCanvas
    StatusBar.ClearCanvas
    
    If (g_OpenImageCount <= 0) Then
        SetScrollVisibility PD_HORIZONTAL, False
        SetScrollVisibility PD_VERTICAL, False
    
    'If any valid images are loaded, scroll bars are always made visible
    Else
        SetScrollVisibility PD_HORIZONTAL, True
        SetScrollVisibility PD_VERTICAL, True
    End If
    
    Me.AlignCanvasView
    
End Sub

'Get/Set scroll bar value
Public Function GetScrollValue(ByVal barType As PD_ORIENTATION) As Long
    If barType = PD_HORIZONTAL Then GetScrollValue = hScroll.Value Else GetScrollValue = vScroll.Value
End Function

Public Sub SetScrollValue(ByVal barType As PD_ORIENTATION, ByVal newValue As Long)
    
    Select Case barType
    
        Case PD_HORIZONTAL
            hScroll.Value = newValue
            
        Case PD_VERTICAL
            vScroll.Value = newValue
        
        Case PD_BOTH
            hScroll.Value = newValue
            vScroll.Value = newValue
        
    End Select
    
    'If automatic redraws are suspended, the scroll bars change events won't fire, so we must manually notify external UI elements
    If Me.GetRedrawSuspension Then RelayViewportChanges
    
End Sub

'Get/Set scroll max/min
Public Function GetScrollMax(ByVal barType As PD_ORIENTATION) As Long
    If barType = PD_HORIZONTAL Then GetScrollMax = hScroll.Max Else GetScrollMax = vScroll.Max
End Function

Public Function GetScrollMin(ByVal barType As PD_ORIENTATION) As Long
    If barType = PD_HORIZONTAL Then GetScrollMin = hScroll.Min Else GetScrollMin = vScroll.Min
End Function

Public Sub SetScrollMax(ByVal barType As PD_ORIENTATION, ByVal newMax As Long)
    If barType = PD_HORIZONTAL Then hScroll.Max = newMax Else vScroll.Max = newMax
End Sub

Public Sub SetScrollMin(ByVal barType As PD_ORIENTATION, ByVal newMin As Long)
    If barType = PD_HORIZONTAL Then hScroll.Min = newMin Else vScroll.Min = newMin
End Sub

'Set scroll bar LargeChange value
Public Sub SetScrollLargeChange(ByVal barType As PD_ORIENTATION, ByVal newLargeChange As Long)
    If barType = PD_HORIZONTAL Then hScroll.LargeChange = newLargeChange Else vScroll.LargeChange = newLargeChange
End Sub

'Set scrollbar visibility.  Note that visibility is only toggled as necessary, so this function is preferable to
' calling .Visible properties directly.
Public Sub SetScrollVisibility(ByVal barType As PD_ORIENTATION, ByVal newVisibility As Boolean)
    
    'If the scroll bar status wasn't actually changed, we can avoid a forced screen refresh
    Dim changesMade As Boolean
    changesMade = False
    
    Select Case barType
    
        Case PD_HORIZONTAL
            If newVisibility <> hScroll.Visible Then
                hScroll.Visible = newVisibility
                changesMade = True
            End If
        
        Case PD_VERTICAL
            If newVisibility <> vScroll.Visible Then
                vScroll.Visible = newVisibility
                changesMade = True
            End If
        
        Case PD_BOTH
            If (newVisibility <> hScroll.Visible) Or (newVisibility <> vScroll.Visible) Then
                hScroll.Visible = newVisibility
                vScroll.Visible = newVisibility
                changesMade = True
            End If
    
    End Select
    
    'The "center" button between the scroll bars has the same visibility as the scrollbars; it's only visible if both bars are visible
    cmdCenter.Visible = CBool(hScroll.Visible And vScroll.Visible)
    
    'When scroll bar visibility is changed, we must move the main canvas picture box to match
    If changesMade Then AlignCanvasView
    
End Sub

Public Sub DisplayImageSize(ByRef srcImage As pdImage, Optional ByVal clearSize As Boolean = False)
    StatusBar.DisplayImageSize srcImage, clearSize
End Sub

Public Sub DisplayCanvasMessage(ByRef cMessage As String)
    StatusBar.DisplayCanvasMessage cMessage
End Sub

Public Sub DisplayCanvasCoordinates(ByVal xCoord As Long, ByVal yCoord As Long, Optional ByVal clearCoords As Boolean = False)
    StatusBar.DisplayCanvasCoordinates xCoord, yCoord, clearCoords
End Sub

Public Sub RequestViewportRedraw(Optional ByVal refreshImmediately As Boolean = False)
    CanvasView.RequestRedraw refreshImmediately
End Sub

'Return the current width/height of the underlying canvas view
Public Function GetCanvasWidth() As Long
    GetCanvasWidth = CanvasView.GetCanvasWidth
End Function

Public Function GetCanvasHeight() As Long
    GetCanvasHeight = CanvasView.GetCanvasHeight
End Function

Public Function GetStatusBarHeight() As Long
    GetStatusBarHeight = StatusBar.GetHeight
End Function

Public Function GetProgBarReference() As PictureBox
    Set GetProgBarReference = picProgressBar
End Function

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

Public Property Get hWnd()
    hWnd = UserControl.hWnd
End Property

'Note that this control does *not* return its own DC.  Instead, it returns the DC of the underlying CanvasView object.
' This is by design.
Public Property Get hDC()
    hDC = CanvasView.hDC
End Property

'TODO: why is this exposed externally?
Public Sub EnableZoomIn(ByVal isEnabled As Boolean)
    'cmdZoomIn.Enabled = isEnabled
End Sub

Public Sub EnableZoomOut(ByVal isEnabled As Boolean)
    'cmdZoomOut.Enabled = isEnabled
End Sub

Public Sub EnableZoomFit(ByVal isEnabled As Boolean)
    'cmdZoomFit.Enabled = isEnabled
    'cmdZoomFit.Value = (cmbZoom.ListIndex = g_Zoom.getZoomFitAllIndex)
End Sub

Public Function GetZoomDropDownReference() As pdDropDown
    Set GetZoomDropDownReference = StatusBar.GetZoomDropDownReference
End Function

'Various input events are bubbled up from the underlying CanvasView control.  It provides no handling over paint and
' tool events, so we must reroute those events here.
Private Sub CanvasView_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    markEventHandled = False
    
    'Make sure canvas interactions are allowed (e.g. an image has been loaded, etc)
    If IsCanvasInteractionAllowed() Then
    
        Dim hOffset As Long, vOffset As Long
        Dim canvasUpdateRequired As Boolean

        'Any further processing depends on which tool is currently active
        Select Case g_CurrentTool
        
            'Drag-to-pan canvas
            Case NAV_DRAG
                
                canvasUpdateRequired = False
                
                'Suspend automatic redraws until all arrow keys have been processed
                Me.SetRedrawSuspension True
                
                'If scrollbars are visible, nudge the canvas in the direction of the arrows.
                If vScroll.Enabled Then
                    If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Then canvasUpdateRequired = True
                    If (vkCode = VK_UP) Then vScroll.Value = vScroll.Value - 1
                    If (vkCode = VK_DOWN) Then vScroll.Value = vScroll.Value + 1
                End If
                
                If hScroll.Enabled Then
                    If (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then canvasUpdateRequired = True
                    If (vkCode = VK_LEFT) Then hScroll.Value = hScroll.Value - 1
                    If (vkCode = VK_RIGHT) Then hScroll.Value = hScroll.Value + 1
                End If
                
                'Re-enable automatic redraws
                Me.SetRedrawSuspension False
                
                'Redraw the viewport if necessary
                If canvasUpdateRequired Then
                    markEventHandled = True
                    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
                End If
                    
            'Move stuff around
            Case NAV_MOVE
            
                'Handle arrow keys first
                If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Or (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then
            
                    'Calculate offset modifiers for the current layer
                    If (vkCode = VK_UP) Then vOffset = vOffset - 1
                    If (vkCode = VK_DOWN) Then vOffset = vOffset + 1
                    If (vkCode = VK_LEFT) Then hOffset = hOffset - 1
                    If (vkCode = VK_RIGHT) Then hOffset = hOffset + 1
                    
                    If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Or (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then canvasUpdateRequired = True
                    
                    'Apply the offsets
                    With pdImages(g_CurrentImage).getActiveLayer
                        .setLayerOffsetX .getLayerOffsetX + hOffset
                        .setLayerOffsetY .getLayerOffsetY + vOffset
                    End With
                    
                    'Redraw the viewport if necessary
                    If canvasUpdateRequired Then
                        markEventHandled = True
                        Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
                    End If
                    
                'Handle non-arrow keys next
                Else
                
                    'Delete key: delete the active layer (if allowed)
                    If (vkCode = VK_DELETE) And pdImages(g_CurrentImage).getNumOfLayers > 1 Then
                        markEventHandled = True
                        Process "Delete layer", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGE_VECTORSAFE
                    End If
                    
                    'Insert: raise Add New Layer dialog
                    If (vkCode = VK_INSERT) Then
                        markEventHandled = True
                        Process "Add new layer", True
                    End If
                
                    'Tab and Shift+Tab: move through layer stack
                    If (vkCode = VK_TAB) Then
                        
                        markEventHandled = True
                        
                        'Retrieve the active layer index
                        Dim curLayerIndex As Long
                        curLayerIndex = pdImages(g_CurrentImage).getActiveLayerIndex
                        
                        'Advance the layer index according to the Shift modifier
                        If (Shift And vbShiftMask) <> 0 Then
                            curLayerIndex = curLayerIndex + 1
                        Else
                            curLayerIndex = curLayerIndex - 1
                        End If
                        
                        If curLayerIndex < 0 Then curLayerIndex = pdImages(g_CurrentImage).getNumOfLayers - 1
                        If curLayerIndex > pdImages(g_CurrentImage).getNumOfLayers - 1 Then curLayerIndex = 0
                        
                        'Activate the new layer
                        pdImages(g_CurrentImage).setActiveLayerByIndex curLayerIndex
                        
                        'Redraw the viewport and interface to match
                        Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), Me
                        SyncInterfaceToCurrentImage
                        
                    End If
                
                    'Space bar: toggle active layer visibility
                    If (vkCode = VK_SPACE) Then
                        markEventHandled = True
                        pdImages(g_CurrentImage).getActiveLayer.setLayerVisibility (Not pdImages(g_CurrentImage).getActiveLayer.getLayerVisibility)
                        Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
                        SyncInterfaceToCurrentImage
                    End If
                
                End If
            
            'Selections
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
            
                'Handle arrow keys first
                If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Or (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then
            
                    'If a selection is active, nudge it using the arrow keys
                    If pdImages(g_CurrentImage).selectionActive And (pdImages(g_CurrentImage).mainSelection.getSelectionShape <> sRaster) Then
                        
                        markEventHandled = True
                        
                        'Disable automatic refresh requests
                        pdImages(g_CurrentImage).mainSelection.rejectRefreshRequests = True
                        
                        'Calculate offsets
                        If (vkCode = VK_UP) Then vOffset = vOffset - 1
                        If (vkCode = VK_DOWN) Then vOffset = vOffset + 1
                        If (vkCode = VK_LEFT) Then hOffset = hOffset - 1
                        If (vkCode = VK_RIGHT) Then hOffset = hOffset + 1
                        
                        'If offsets were generated, update the selection and redraw the screen
                        If (hOffset <> 0) Or (vOffset <> 0) Then
                        
                            'Some selection types can be modified by simply updating the selection text boxes.  Others cannot.
                            
                            'Non-textbox-compatible selections are handled here
                            If (g_CurrentTool = SELECT_POLYGON) Or (g_CurrentTool = SELECT_LASSO) Or (g_CurrentTool = SELECT_WAND) Then
                                pdImages(g_CurrentImage).mainSelection.rejectRefreshRequests = False
                                pdImages(g_CurrentImage).mainSelection.nudgeSelection hOffset, vOffset
                            
                            'Textbox-compatible selections are handled here
                            Else
                            
                                'Update the selection coordinate text boxes with the new offsets
                                toolpanel_Selections.tudSel(0).Value = toolpanel_Selections.tudSel(0).Value + hOffset
                                toolpanel_Selections.tudSel(1).Value = toolpanel_Selections.tudSel(1).Value + vOffset
                                
                                If g_CurrentTool = SELECT_LINE Then
                                    toolpanel_Selections.tudSel(2).Value = toolpanel_Selections.tudSel(2).Value + hOffset
                                    toolpanel_Selections.tudSel(3).Value = toolpanel_Selections.tudSel(3).Value + vOffset
                                End If
                                
                                'Update the screen
                                pdImages(g_CurrentImage).mainSelection.rejectRefreshRequests = False
                                pdImages(g_CurrentImage).mainSelection.updateViaTextBox
                            
                            End If
                            
                            Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                            
                        End If
                    
                    End If
                
                'Handle non-arrow keys next.  (Note: most non-arrow keys are not meant to work with key-repeating, so they
                ' are handled in the KeyUp event instead.)
                Else
                
                    
                                    
                End If
            
        End Select
        
    End If

End Sub

Private Sub CanvasView_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    markEventHandled = False

    'Make sure canvas interactions are allowed (e.g. an image has been loaded, etc)
    If IsCanvasInteractionAllowed() Then
        
        'Any further processing depends on which tool is currently active
        Select Case g_CurrentTool
        
            Case NAV_DRAG
            
            Case NAV_MOVE
            
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
                
                'Delete key: if a selection is active, erase the selected area
                If (vkCode = VK_DELETE) And pdImages(g_CurrentImage).selectionActive Then
                    markEventHandled = True
                    Process "Erase selected area", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_LAYER
                End If
                
                'Escape key: if a selection is active, clear it
                If (vkCode = VK_ESCAPE) And pdImages(g_CurrentImage).selectionActive Then
                    markEventHandled = True
                    Process "Remove selection", , , UNDO_SELECTION
                End If
                
                'Backspace key: for lasso and polygon selections, retreat back one or more coordinates, giving the user a chance to
                ' correct any potential mistakes.
                If ((g_CurrentTool = SELECT_LASSO) Or (g_CurrentTool = SELECT_POLYGON)) And (vkCode = VK_BACK) And pdImages(g_CurrentImage).selectionActive And (Not pdImages(g_CurrentImage).mainSelection.isLockedIn) Then
                    
                    markEventHandled = True
                    
                    'Polygons
                    If (g_CurrentTool = SELECT_POLYGON) Then
                    
                        'Do not allow point removal if the polygon has already been successfully closed.
                        If Not pdImages(g_CurrentImage).mainSelection.getPolygonClosedState Then pdImages(g_CurrentImage).mainSelection.removeLastPolygonPoint
                    
                    'Lassos
                    Else
                    
                        'Do not allow point removal if the lasso has already been successfully closed.
                        If Not pdImages(g_CurrentImage).mainSelection.getLassoClosedState Then
                    
                            'Ask the selection object to retreat its position
                            Dim newImageX As Double, newImageY As Double
                            pdImages(g_CurrentImage).mainSelection.retreatLassoPosition newImageX, newImageY
                            
                            'The returned coordinates will be in image coordinates.  Convert them to viewport coordinates.
                            Dim newCanvasX As Double, newCanvasY As Double
                            Drawing.ConvertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), newImageX, newImageY, newCanvasX, newCanvasY
                            
                            'Finally, convert the canvas coordinates to screen coordinates, and move the cursor accordingly
                            CanvasView.SetCursorToCanvasPosition newCanvasX, newCanvasY
                            
                        End If
                        
                    End If
                    
                    'Redraw the screen to reflect this new change.
                    Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                
                End If
            
        End Select
        
    End If
    
End Sub

Private Sub cmdCenter_Click()
    Image_Canvas_Handler.CenterOnScreen
End Sub

'At present, the only App Commands the canvas will handle are forward/back, which link to Undo/Redo
Private Sub CanvasView_AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If IsCanvasInteractionAllowed() Then
    
        Select Case cmdID
        
            'Back button
            Case AC_BROWSER_BACKWARD, AC_UNDO
            
                If pdImages(g_CurrentImage).IsActive Then
                    If pdImages(g_CurrentImage).undoManager.getUndoState Then Process "Undo", , , UNDO_NOTHING
                End If
            
            'Forward button
            Case AC_BROWSER_FORWARD, AC_REDO
            
                If pdImages(g_CurrentImage).IsActive Then
                    If pdImages(g_CurrentImage).undoManager.getRedoState Then Process "Redo", , , UNDO_NOTHING
                End If
        
        End Select

    End If

End Sub

'RELAY (partially)
Private Sub CanvasView_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Make sure interactions with this canvas are allowed
    If Not IsCanvasInteractionAllowed() Then Exit Sub
    
    'Because VB does not allow an user control to receive focus if it contains controls that can receive focus, the arrow buttons
    ' can behave unpredictably (for example, if the zoom box has focus, and the user clicks on the canvas, the canvas will not
    ' receive focus and arrow key presses will continue to interact with the zoom box instead of the viewport)
    ' (NOTE: this should be fixed as of 6.6, as a dedicated picture box is now used for rendering)
    
    'Note whether a selection is active when mouse interactions began
    m_SelectionActiveBeforeMouseEvents = (pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.isLockedIn)
    
    'These variables will hold the corresponding (x,y) coordinates on the IMAGE - not the VIEWPORT.
    ' (This is important if the user has zoomed into an image, and used scrollbars to look at a different part of it.)
    Dim imgX As Double, imgY As Double
    
    'Note that displayImageCoordinates returns a copy of the displayed coordinates via imgX/Y
    DisplayImageCoordinates x, y, pdImages(g_CurrentImage), Me, imgX, imgY
    
    'We also need a copy of the current mouse position relative to the active layer.  (This became necessary in PD 7.0, as layers
    ' may have non-destructive affine transforms active, which means we can't blindly switch between image and layer coordinate spaces!)
    Dim layerX As Single, layerY As Single
    Drawing.ConvertImageCoordsToLayerCoords pdImages(g_CurrentImage), pdImages(g_CurrentImage).getActiveLayer, imgX, imgY, layerX, layerY
    
    'Display a relevant cursor for the current action
    SetCanvasCursor pMouseUp, Button, x, y, imgX, imgY, layerX, layerY
    
    'Selection tools all use the same variable for tracking POIs
    Dim sCheck As Long
    
    'Check mouse button use
    If Button = vbLeftButton Then
        
        m_LMBDown = True
        m_NumOfMouseMovements = 0
            
        'Remember this location
        m_InitMouseX = x
        m_InitMouseY = y
        
        'Some functions may not operate on the current layer, but on the layer under the mouse
        Dim layerUnderMouse As Long
        
        'Ask the current layer if these coordinates correspond to a point of interest.  We don't always use this return value,
        ' but a number of functions could potentially ask for it, so we cache it at MouseDown time and hang onto it until
        ' the mouse is released.
        m_CurPointOfInterest = pdImages(g_CurrentImage).getActiveLayer.checkForPointOfInterest(layerX, layerY)
        
        'Any further processing depends on which tool is currently active
        Select Case g_CurrentTool
        
            'Drag-to-pan canvas
            Case NAV_DRAG
                setInitialCanvasScrollValues FormMain.mainCanvas(0)
                
            'Move stuff around
            Case NAV_MOVE
            
                'Prior to moving or transforming a layer, we need to check the state of the "auto-activate layer beneath mouse"
                ' option; if it is set, check (and possibly modify) the active layer based on the mouse position.
                If CBool(toolpanel_MoveSize.chkAutoActivateLayer) Then
                
                    layerUnderMouse = Layer_Handler.getLayerUnderMouse(imgX, imgY, True)
                    
                    'The "getLayerUnderMouse" function will return a layer index if the mouse is over a layer.  If the mouse is not
                    ' over a layer, it will return -1.
                    If layerUnderMouse > -1 Then
                    
                        'If the layer under the mouse is not already active, activate it now
                        If layerUnderMouse <> pdImages(g_CurrentImage).getActiveLayerIndex Then
                            Layer_Handler.setActiveLayerByIndex layerUnderMouse, False
                            Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), Me
                        End If
                    
                    End If
                
                End If
                
                'Initiate the layer transformation engine.  Note that nothing will happen until the user actually moves the mouse.
                Tool_Support.setInitialLayerToolValues pdImages(g_CurrentImage), pdImages(g_CurrentImage).getActiveLayer, imgX, imgY, pdImages(g_CurrentImage).getActiveLayer.checkForPointOfInterest(layerX, layerY)
        
            'Standard selections
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO
            
                'Check to see if a selection is already active.  If it is, see if the user is allowed to transform it.
                If pdImages(g_CurrentImage).selectionActive Then
                
                    'Check the mouse coordinates of this click.
                    sCheck = findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
                    
                    'If a point of interest was clicked, initiate a transform
                    If (sCheck <> -1) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape <> sPolygon) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape <> sRaster) Then
                        
                        'Initialize a selection transformation
                        pdImages(g_CurrentImage).mainSelection.setTransformationType sCheck
                        pdImages(g_CurrentImage).mainSelection.setInitialTransformCoordinates imgX, imgY
                                        
                    'If a point of interest was *not* clicked, erase any existing selection and start a new one
                    Else
                        
                        'Polygon selections require special handling, because they don't operate on the "mouse up = complete" assumption.
                        ' They are completed when the user re-clicks the first point.  Any clicks prior to that point are treated as
                        ' an instruction to add a new points.
                        If g_CurrentTool = SELECT_POLYGON Then
                            
                            'First, see if the selection is locked in.  If it is, treat this is a regular transformation.
                            If pdImages(g_CurrentImage).mainSelection.isLockedIn Then
                                pdImages(g_CurrentImage).mainSelection.setTransformationType sCheck
                                pdImages(g_CurrentImage).mainSelection.setInitialTransformCoordinates imgX, imgY
                            
                            'Selection is not locked in, meaning the user is still constructing it.
                            Else
                            
                                'If the user clicked on the initial polygon point, attempt to close the polygon
                                If (sCheck = 0) And (pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints > 2) Then
                                    pdImages(g_CurrentImage).mainSelection.setPolygonClosedState True
                                    pdImages(g_CurrentImage).mainSelection.setTransformationType 0
                                
                                'The user did not click the initial polygon point, meaning we should add this coordinate as a new polygon point.
                                Else
                                    
                                    'Remove transformation mode (if any)
                                    pdImages(g_CurrentImage).mainSelection.setTransformationType -1
                                    pdImages(g_CurrentImage).mainSelection.overrideTransformMode False
                                    
                                    'Add the new point
                                    If pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints = 0 Then
                                        Selection_Handler.initSelectionByPoint imgX, imgY
                                    Else
                                        
                                        If (sCheck = -1) Or (sCheck = pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints) Then
                                            pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                                            pdImages(g_CurrentImage).mainSelection.setTransformationType pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints - 1
                                        Else
                                            pdImages(g_CurrentImage).mainSelection.setTransformationType sCheck
                                        End If
                                        
                                    End If
                                    
                                    'Reinstate transformation mode, using the index of the new point as the transform ID
                                    pdImages(g_CurrentImage).mainSelection.setInitialTransformCoordinates imgX, imgY
                                    pdImages(g_CurrentImage).mainSelection.overrideTransformMode True
                                    
                                    'Redraw the screen
                                    Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), Me
                                    
                                End If
                            
                            End If
                            
                        Else
                            Selection_Handler.initSelectionByPoint imgX, imgY
                        End If
                        
                    End If
                
                'If a selection is not active, start a new one
                Else
                    
                    Selection_Handler.initSelectionByPoint imgX, imgY
                    
                    'Polygon selections require special handling, as usual.  After creating the initial point, we want to immediately initiate
                    ' transform mode, because dragging the mouse will simply move the newly created point.
                    If g_CurrentTool = SELECT_POLYGON Then
                        pdImages(g_CurrentImage).mainSelection.setTransformationType pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints - 1
                        pdImages(g_CurrentImage).mainSelection.overrideTransformMode True
                    End If
                    
                End If
            
            'Magic wand selections are easy.  They never transform - they only generate anew
            Case SELECT_WAND
                Selection_Handler.initSelectionByPoint imgX, imgY
                Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), Me
                
            'Text layer behavior varies depending on whether the current layer is a text layer or not
            Case VECTOR_TEXT, VECTOR_FANCYTEXT
                
                'One of two things can happen when the mouse is clicked in text mode:
                ' 1) The current layer is a text layer, and the user wants to edit it (move it around, resize, etc)
                ' 2) The user wants to add a new text layer, which they can do by clicking over a non-text layer portion of the image
                
                'Let's start by distinguishing between these two states.
                Dim userIsEditingCurrentTextLayer As Boolean
                
                'Check to see if the current layer is a text layer
                If pdImages(g_CurrentImage).getActiveLayer.isLayerText Then
                
                    'Did the user click on a POI for this layer?  If they did, the user is editing the current text layer.
                    If m_CurPointOfInterest >= 0 Then
                        userIsEditingCurrentTextLayer = True
                    Else
                        userIsEditingCurrentTextLayer = False
                    End If
                
                'The current active layer is not a text layer.
                Else
                    userIsEditingCurrentTextLayer = False
                End If
                
                'If the user is editing the current text layer, we can switch directly into layer transform mode
                If userIsEditingCurrentTextLayer Then
                    
                    'Initiate the layer transformation engine.  Note that nothing will happen until the user actually moves the mouse.
                    Tool_Support.setInitialLayerToolValues pdImages(g_CurrentImage), pdImages(g_CurrentImage).getActiveLayer, imgX, imgY, pdImages(g_CurrentImage).getActiveLayer.checkForPointOfInterest(layerX, layerY)
                    
                'The user is not editing a text layer.  Create a new text layer now.
                Else
                    
                    'Create a new text layer directly; note that we *do not* pass this command through the central processor, as we do not
                    ' want the delay associated with full Undo/Redo creation.
                    If g_CurrentTool = VECTOR_TEXT Then
                        Layer_Handler.addNewLayer pdImages(g_CurrentImage).getActiveLayerIndex, PDL_TEXT, 0, 0, 0, True, "", imgX, imgY, True
                    ElseIf g_CurrentTool = VECTOR_FANCYTEXT Then
                        Layer_Handler.addNewLayer pdImages(g_CurrentImage).getActiveLayerIndex, PDL_TYPOGRAPHY, 0, 0, 0, True, "", imgX, imgY, True
                    End If
                    
                    'Use a special initialization command that basically copies all existing text properties into the newly created layer.
                    Tool_Support.syncCurrentLayerToToolOptionsUI
                    
                    'Put the newly created layer into transform mode, with the bottom-right corner selected
                    Tool_Support.setInitialLayerToolValues pdImages(g_CurrentImage), pdImages(g_CurrentImage).getActiveLayer, imgX, imgY, 3
                                        
                    'Also, note that we have just created a new text layer.  The MouseUp event needs to know this, so it can initiate a full-image Undo/Redo event.
                    Tool_Support.setCustomToolState PD_TEXT_TOOL_CREATED_NEW_LAYER
                    
                    'Redraw the viewport immediately
                    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0), False, 3
                
                End If
                
            'In the future, other tools can be handled here
            Case Else
            
            
        End Select
    
    ElseIf Button = vbRightButton Then
    
        m_RMBDown = True
        
        'TODO: right-button functionality
    
    End If
    
End Sub

Private Sub CanvasView_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_IsMouseOverCanvas = True
End Sub

'When the mouse leaves the window, if no buttons are down, clear the coordinate display.
' (We must check for button states because the user is allowed to do things like drag selection nodes outside the image.)
' RELAY (partially)
Private Sub CanvasView_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_IsMouseOverCanvas = False
    If (Not m_LMBDown) And (Not m_RMBDown) Then ClearImageCoordinatesDisplay
End Sub

Private Sub CanvasView_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Make sure interactions with this canvas are allowed
    If Not IsCanvasInteractionAllowed() Then Exit Sub
    
    m_NumOfMouseMovements = m_NumOfMouseMovements + 1
    
    'These variables will hold the corresponding (x,y) coordinates on the image - NOT the viewport
    Dim imgX As Double, imgY As Double
    
    'Display the image coordinates under the mouse pointer
    DisplayImageCoordinates x, y, pdImages(g_CurrentImage), Me, imgX, imgY
    
    'We also need a copy of the current mouse position relative to the active layer.  (This became necessary in PD 7.0, as layers
    ' may have non-destructive affine transforms active, which means we can't reuse image coordinates as layer coordinates!)
    '
    'Note also that we refresh the layer transformation matrix if the mouse is not down
    Dim layerX As Single, layerY As Single
    Drawing.ConvertImageCoordsToLayerCoords pdImages(g_CurrentImage), pdImages(g_CurrentImage).getActiveLayer, imgX, imgY, layerX, layerY
        
    'Check the left mouse button
    If m_LMBDown Then
    
        Select Case g_CurrentTool
        
            'Drag-to-pan canvas
            Case NAV_DRAG
                panImageCanvas m_InitMouseX, m_InitMouseY, x, y, pdImages(g_CurrentImage), FormMain.mainCanvas(0)
            
            'Move stuff around
            Case NAV_MOVE
                Message "Shift key: preserve layer aspect ratio", "DONOTLOG"
                transformCurrentLayer imgX, imgY, pdImages(g_CurrentImage), pdImages(g_CurrentImage).getActiveLayer, FormMain.mainCanvas(0), (Shift And vbShiftMask)
        
            'Basic selection tools
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON
    
                'First, check to see if a selection is both active and transformable.
                If pdImages(g_CurrentImage).selectionActive And (pdImages(g_CurrentImage).mainSelection.getSelectionShape <> sRaster) Then
                    
                    'If the SHIFT key is down, notify the selection engine that a square shape is requested
                    pdImages(g_CurrentImage).mainSelection.requestSquare (Shift And vbShiftMask)
                    
                    'Pass new points to the active selection
                    pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                    syncTextToCurrentSelection g_CurrentImage
                                        
                End If
                
                'Force a redraw of the viewport
                If m_NumOfMouseMovements > 1 Then Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), Me
            
            'Lasso selections are handled specially, because mouse move events control the drawing of the lasso
            Case SELECT_LASSO
            
                'First, check to see if a selection is active
                If pdImages(g_CurrentImage).selectionActive Then
                    
                    'Pass new points to the active selection
                    pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                                        
                End If
                
                'To spare the debug logger from receiving too many events, forcibly prevent logging of this message
                ' while in debug mode.
                #If DEBUGMODE = 1 Then
                    Message "Release the mouse button to complete the lasso selection", "DONOTLOG"
                #Else
                    Message "Release the mouse button to complete the lasso selection"
                #End If
                
                'Force a redraw of the viewport
                If m_NumOfMouseMovements > 1 Then Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), Me
            
            'Wand selections are easier than other selection types, because they don't support any special transforms
            Case SELECT_WAND
                If pdImages(g_CurrentImage).selectionActive Then
                    pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                    Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), Me
                End If
                
            'Text layers are identical to the move tool
            Case VECTOR_TEXT, VECTOR_FANCYTEXT
                Message "Shift key: preserve layer aspect ratio"
                transformCurrentLayer imgX, imgY, pdImages(g_CurrentImage), pdImages(g_CurrentImage).getActiveLayer, FormMain.mainCanvas(0), (Shift And vbShiftMask)
            
        End Select
    
    'This else means the LEFT mouse button is NOT down
    Else
        
        'Display a relevant cursor for the current action
        SetCanvasCursor pMouseUp, Button, x, y, imgX, imgY, layerX, layerY
    
        Select Case g_CurrentTool
        
            'Drag-to-navigate
            Case NAV_DRAG
            
            'Move stuff around
            Case NAV_MOVE
            
                'If the user has the "auto-activate layer beneath mouse" option set, report the current layer name in the
                ' message bar; this is helpful for determining what layer will be affected by a given action.
                If CBool(toolpanel_MoveSize.chkAutoActivateLayer) Then
                
                    Dim layerUnderMouse As Long
                    layerUnderMouse = Layer_Handler.getLayerUnderMouse(imgX, imgY, True)
                    
                    'The "getLayerUnderMouse" function will return a layer index if the mouse is over a layer.  If the mouse is not
                    ' over a layer, it will return -1.
                    If layerUnderMouse > -1 Then
                        m_LayerAutoActivateIndex = layerUnderMouse
                        
                        'To spare the debug logger from receiving too many events, forcibly prevent logging of this message
                        ' while in debug mode.
                        #If DEBUGMODE = 1 Then
                            Message "Target layer: %1", pdImages(g_CurrentImage).getLayerByIndex(layerUnderMouse).getLayerName, "DONOTLOG"
                        #Else
                            Message "Target layer: %1", pdImages(g_CurrentImage).getLayerByIndex(layerUnderMouse).getLayerName
                        #End If
                    
                    'The mouse is not over a layer.  Default to the active layer, which allows the user to interact with the
                    ' layer even if it lies off-canvas.
                    Else
                        m_LayerAutoActivateIndex = pdImages(g_CurrentImage).getActiveLayerIndex
                    End If
                
                'Auto-activation is disabled.  Don't bother reporting the layer beneath the mouse to the user, as actions can
                ' only affect the active layer!
                Else
                    Message ""
                    m_LayerAutoActivateIndex = pdImages(g_CurrentImage).getActiveLayerIndex
                End If
                
            'Selection tools
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
            
            'Text tools
            Case VECTOR_TEXT, VECTOR_FANCYTEXT
            
            Case Else
            
        End Select
        
    End If
    
End Sub

Private Sub CanvasView_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    
    'Make sure interactions with this canvas are allowed
    If Not IsCanvasInteractionAllowed() Then Exit Sub
    
    'Display the image coordinates under the mouse pointer
    Dim imgX As Double, imgY As Double
    DisplayImageCoordinates x, y, pdImages(g_CurrentImage), Me, imgX, imgY
    
    'We also need a copy of the current mouse position relative to the active layer.  (This became necessary in PD 7.0, as layers
    ' may have non-destructive affine transforms active, which means we can't blindly switch between image and layer coordinate spaces!)
    Dim layerX As Single, layerY As Single
    Drawing.ConvertImageCoordsToLayerCoords pdImages(g_CurrentImage), pdImages(g_CurrentImage).getActiveLayer, imgX, imgY, layerX, layerY
    
    'Display a relevant cursor for the current action
    SetCanvasCursor pMouseUp, Button, x, y, imgX, imgY, layerX, layerY
    
    'Check mouse buttons
    If Button = vbLeftButton Then
    
        m_LMBDown = False
    
        Select Case g_CurrentTool
        
            'Click-to-drag navigation
            Case NAV_DRAG
                
            'Move stuff around
            Case NAV_MOVE
            
                'Pass a final transform request to the layer handler.  This will initiate Undo/Redo creation, among other things.
                If (m_NumOfMouseMovements > 0) Then transformCurrentLayer imgX, imgY, pdImages(g_CurrentImage), pdImages(g_CurrentImage).getActiveLayer, FormMain.mainCanvas(0), (Shift And vbShiftMask), True
                
                'Reset the generic tool mouse tracking function
                Tool_Support.terminateGenericToolTracking
                
            'Most selection tools
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_LASSO
            
                'If a selection was being drawn, lock it into place
                If pdImages(g_CurrentImage).selectionActive Then
                    
                    'Check to see if this mouse location is the same as the initial mouse press. If it is, and that particular
                    ' point falls outside the selection, clear the selection from the image.
                    If ((ClickEventAlsoFiring) And (findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage)) = -1)) Or ((pdImages(g_CurrentImage).mainSelection.selWidth <= 0) And (pdImages(g_CurrentImage).mainSelection.selHeight <= 0)) Then
                        
                        If (g_CurrentTool <> SELECT_WAND) Then
                            Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                        Else
                            Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                        End If
                    
                    'The mouse is being released after a significant move event, or on a point of interest to the current selection.
                    Else
                    
                        'If the selection is not raster-type, pass these final mouse coordinates to it
                        If (pdImages(g_CurrentImage).mainSelection.getSelectionShape <> sRaster) Then

                            pdImages(g_CurrentImage).mainSelection.requestSquare (Shift And vbShiftMask)
                            pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                            syncTextToCurrentSelection g_CurrentImage
                            
                        End If
                    
                        'Check to see if all selection coordinates are invalid (e.g. off-image).  If they are, forget about this selection.
                        If pdImages(g_CurrentImage).mainSelection.areAllCoordinatesInvalid Then
                            Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                        Else
                            
                            'Depending on the type of transformation that may or may not have been applied, call the appropriate processor function.
                            ' This is required to add the current selection event to the Undo/Redo chain.
                            Select Case g_CurrentTool
                            
                                Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
                                
                                    Select Case pdImages(g_CurrentImage).mainSelection.getTransformationType
                            
                                        'Creating a new selection
                                        Case -1
                                            Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                            
                                        'Moving an existing selection
                                        Case 8
                                            Process "Move selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                            
                                        'Anything else is assumed to be resizing an existing selection
                                        Case Else
                                            Process "Resize selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                            
                                    End Select
                                    
                                Case SELECT_LASSO
                                
                                    Select Case pdImages(g_CurrentImage).mainSelection.getTransformationType
                            
                                        'Creating a new selection
                                        Case -1
                                            Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                            
                                        'Moving an existing selection
                                        Case Else
                                            Process "Move selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                            
                                    End Select
                                    
                                Case SELECT_WAND
                                    Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                            
                            End Select
                            
                        End If
                        
                    End If
                    
                    'Force a redraw of the screen
                    Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), Me
                    
                Else
                    'If the selection is not active, make sure it stays that way
                    pdImages(g_CurrentImage).mainSelection.lockRelease
                End If
                
                'Synchronize the selection text box values with the final selection
                syncTextToCurrentSelection g_CurrentImage
                
            
            'As usual, polygon selections have some special considerations.
            Case SELECT_POLYGON
            
                'If a selection was being drawn, lock it into place
                If pdImages(g_CurrentImage).selectionActive Then
                
                    'Check to see if the selection is locked in.  If it is, we need to check for an "erase selection" click.
                    If pdImages(g_CurrentImage).mainSelection.getPolygonClosedState And ClickEventAlsoFiring And (findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage)) = -1) Then
                        Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                    
                    Else
                        
                        'If the polygon is already closed, we want to lock in the newly modified polygon
                        If pdImages(g_CurrentImage).mainSelection.getPolygonClosedState Then
                        
                            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
                            
                                Case pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints
                                    Process "Move selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                    
                                Case 0
                                    If ClickEventAlsoFiring Then
                                        Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                    Else
                                        Process "Resize selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                    End If
                                    
                                Case -1
                                
                                    'If the user has clicked off the selection, we want to remove it.
                                    If ClickEventAlsoFiring Then
                                        Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                                    
                                    'If they haven't clicked, this could simply indicate that they dragged a polygon point off the polygon
                                    ' and into some new region of the image.
                                    Else
                                        
                                        pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                                        Process "Resize selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                        
                                    End If
                                
                                Case Else
                                    Process "Resize selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                    
                            End Select
                            
                            'Check to see if all selection coordinates are invalid (e.g. off-image).  If they are, forget about this selection.
                            If pdImages(g_CurrentImage).mainSelection.isLockedIn And pdImages(g_CurrentImage).mainSelection.areAllCoordinatesInvalid Then
                                Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                            End If
                            
                        Else
                        
                            'Pass these final mouse coordinates to the selection engine
                            pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                            
                            'To spare the debug logger from receiving too many events, forcibly prevent logging of this message
                            ' while in debug mode.
                            #If DEBUGMODE = 1 Then
                                Message "Click on the first point to complete the polygon selection", "DONOTLOG"
                            #Else
                                Message "Click on the first point to complete the polygon selection"
                            #End If
                            
                        End If
                    
                    End If
                    
                    'Force a redraw of the screen
                    Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), Me
                
                Else
                    'If the selection is not active, make sure it stays that way
                    pdImages(g_CurrentImage).mainSelection.lockRelease
                End If
                
            'Magic wand selections are much easier than other selection types
            Case SELECT_WAND
                
                'If a selection was being drawn, lock it into place
                If pdImages(g_CurrentImage).selectionActive Then
                    
                    'Supply the final coords to the selection engine
                    pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                    
                    'Check to see if all selection coordinates are invalid (e.g. off-image).  If they are, forget about this selection.
                    If pdImages(g_CurrentImage).mainSelection.areAllCoordinatesInvalid Then
                        Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                    
                    'If the selection coordinates are valid, create it now.
                    Else
                        Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                    End If
                    
                    'Force a redraw of the screen
                    Viewport_Engine.Stage4_CompositeCanvas pdImages(g_CurrentImage), Me
                    
                Else
                    'If the selection is not active, make sure it stays that way
                    pdImages(g_CurrentImage).mainSelection.lockRelease
                End If
                
            'Text layers
            Case VECTOR_TEXT, VECTOR_FANCYTEXT
                
                'Pass a final transform request to the layer handler.  This will initiate Undo/Redo creation, among other things.
                
                '(Note that this function branches according to two states: whether this click is creating a new text layer (which requires a full
                ' image stack Undo/Redo), or whether we are simply modifying an existing layer.
                If Tool_Support.getCustomToolState = PD_TEXT_TOOL_CREATED_NEW_LAYER Then
                    
                    'Mark the current tool as busy to prevent any unwanted UI syncing
                    Tool_Support.setToolBusyState True
                    
                    'See if this was just a click (as it might be at creation time).
                    If ClickEventAlsoFiring Or (m_NumOfMouseMovements <= 2) Or (pdImages(g_CurrentImage).getActiveLayer.getLayerWidth < 4) Or (pdImages(g_CurrentImage).getActiveLayer.getLayerHeight < 4) Then
                        
                        'Update the layer's size.  At present, we simply make it fill the current viewport.
                        Dim curImageRectF As RECTF
                        pdImages(g_CurrentImage).imgViewport.getIntersectRectImage curImageRectF
                        
                        With pdImages(g_CurrentImage)
                            .getActiveLayer.setLayerOffsetX curImageRectF.Left
                            .getActiveLayer.setLayerOffsetY curImageRectF.Top
                            .getActiveLayer.setLayerWidth curImageRectF.Width
                            .getActiveLayer.setLayerHeight curImageRectF.Height
                        End With
                        
                        'If the current text box is empty, set some new text to orient the user
                        If g_CurrentTool = VECTOR_TEXT Then
                            
                            If Len(toolpanel_Text.txtTextTool.Text) = 0 Then
                                toolpanel_Text.txtTextTool.Text = g_Language.TranslateMessage("(enter text here)")
                            End If
                            
                        Else
                        
                            If Len(toolpanel_FancyText.txtTextTool.Text) = 0 Then
                                toolpanel_FancyText.txtTextTool.Text = g_Language.TranslateMessage("(enter text here)")
                            End If
                        
                        End If
                        
                        'Manually synchronize the new size values against their on-screen UI elements
                        Tool_Support.syncToolOptionsUIToCurrentLayer
                        
                        'Manually force a viewport redraw
                        Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                        
                    'If the user already specified a size, use their values to finalize the layer size
                    Else
                        transformCurrentLayer imgX, imgY, pdImages(g_CurrentImage), pdImages(g_CurrentImage).getActiveLayer, FormMain.mainCanvas(0), (Shift And vbShiftMask)
                    End If
                    
                    'Release the tool engine
                    Tool_Support.setToolBusyState False
                    
                    'Process the addition of the new layer; this will create proper Undo/Redo data for the entire image (required, as the layer order
                    ' has changed due to this new addition).
                    With pdImages(g_CurrentImage).getActiveLayer
                        Process "New text layer", , buildParams(.getLayerOffsetX, .getLayerOffsetY, .getLayerWidth, .getLayerHeight, .getVectorDataAsXML), UNDO_IMAGE_VECTORSAFE
                    End With
                    
                    'Manually synchronize menu, layer toolbox, and other UI settings against the newly created layer.
                    SyncInterfaceToCurrentImage
                    
                    'Finally, set focus to the text layer text entry box
                    If g_CurrentTool = VECTOR_TEXT Then
                        toolpanel_Text.txtTextTool.SetFocus
                        toolpanel_Text.txtTextTool.SelectAll
                    Else
                        toolpanel_FancyText.txtTextTool.SetFocus
                        toolpanel_FancyText.txtTextTool.SelectAll
                    End If
                    
                'The user is simply editing an existing layer.
                Else
                    
                    'As a convenience to the user, ignore clicks that don't actually change layer settings
                    If (m_NumOfMouseMovements > 0) Then transformCurrentLayer imgX, imgY, pdImages(g_CurrentImage), pdImages(g_CurrentImage).getActiveLayer, FormMain.mainCanvas(0), (Shift And vbShiftMask), True
                    
                End If
                
                'Reset the generic tool mouse tracking function
                Tool_Support.terminateGenericToolTracking
                
            Case Else
                    
        End Select
                        
    End If
    
    If Button = vbRightButton Then m_RMBDown = False
    
    'Reset any tracked point of interest value for this layer
    m_CurPointOfInterest = -1
        
    'Reset the mouse movement tracker
    m_NumOfMouseMovements = 0
    
End Sub

Public Sub CanvasView_MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    If Not IsCanvasInteractionAllowed() Then Exit Sub
    If hScroll.Visible Then hScroll.RelayMouseWheelEvent False, Button, Shift, x, y, scrollAmount
End Sub

'Vertical mousewheel scrolling.  Note that Shift+Wheel and Ctrl+Wheel modifiers do NOT raise this event; pdInputMouse automatically
' reroutes them to MouseWheelHorizontal and MouseWheelZoom, respectively.
Public Sub CanvasView_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    
    If Not IsCanvasInteractionAllowed() Then Exit Sub
    
    'PhotoDemon uses the standard photo editor convention of Ctrl+Wheel = zoom, Shift+Wheel = h_scroll, and Wheel = v_scroll.
    ' Some users (for reasons I don't understand??) expect plain mousewheel to zoom the image.  For these users, we now
    ' display a helpful message telling them to use the damn Ctrl modifier like everyone else.
    If vScroll.Visible Then
        vScroll.RelayMouseWheelEvent True, Button, Shift, x, y, scrollAmount
        
    'The user is using the mousewheel without Ctrl/Shift modifiers, even without a visible scrollbar.
    ' Display a message about how mousewheels are supposed to work.
    Else
        Message "Mouse Wheel = VERTICAL SCROLL,  Shift + Wheel = HORIZONTAL SCROLL,  Ctrl + Wheel = ZOOM"
    End If
    
    'NOTE: horizontal scrolling via Shift+Vertical Wheel is handled in the separate _MouseWheelHorizontal event.
    'NOTE: zooming via Ctrl+Vertical Wheel is handled in the separate _MouseWheelZoom event.
    
End Sub

Public Sub CanvasView_MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
    
    If Not IsCanvasInteractionAllowed() Then Exit Sub
    
    'Before doing anything else, cache the current mouse coordinates (in both Canvas and Image coordinate spaces)
    Dim imgX As Double, imgY As Double
    ConvertCanvasCoordsToImageCoords Me, pdImages(g_CurrentImage), x, y, imgX, imgY, True
    
    'Suspend automatic viewport redraws until we are done with our calculations
    g_AllowViewportRendering = False
    
    'Calculate a new zoom value
    If zoomAmount > 0 Then
        
        If FormMain.mainCanvas(0).GetZoomDropDownReference().Enabled And FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex > 0 Then
            FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex = g_Zoom.GetNearestZoomInIndex(FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex)
        End If
           
    ElseIf zoomAmount < 0 Then
        
        If FormMain.mainCanvas(0).GetZoomDropDownReference().Enabled And FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex <> g_Zoom.GetZoomCount Then
            FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex = g_Zoom.GetNearestZoomOutIndex(FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex)
        End If
        
    End If
    
    'Re-enable automatic viewport redraws
    g_AllowViewportRendering = True
    
    'Request a manual redraw from Viewport_Engine.Stage1_InitializeBuffer, while supplying our x/y coordinates so that it can preserve mouse position
    ' relative to the underlying image.
    Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_PreservePointPosition, x, y, imgX, imgY
    
    'Notify external UI elements of the change
    RelayViewportChanges

End Sub

Private Sub UserControl_Initialize()
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDCANVAS_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDCanvas", colorCount
    If Not g_IsProgramRunning Then UpdateColorList
    
    If g_IsProgramRunning Then
        
        'Allow the control to generate its own redraw requests
        Me.SetRedrawSuspension False
        
        'Set scroll bar size to match the current system default (which changes based on DPI, theming, and other factors)
        hScroll.Height = GetSystemMetrics(SM_CYHSCROLL)
        vScroll.Width = GetSystemMetrics(SM_CXVSCROLL)
        
        'Align the main picture box
        AlignCanvasView
        
        'Reset any POI trackers
        m_CurPointOfInterest = -1
        m_LastPointOfInterest = -1
        
    End If
    
End Sub

Private Sub HScroll_Scroll(ByVal eventIsCritical As Boolean)
    
    'Regardless of viewport state, cache the current scroll bar value inside the current image
    If Not pdImages(g_CurrentImage) Is Nothing Then
        pdImages(g_CurrentImage).imgViewport.setHScrollValue hScroll.Value
    End If
    
    If (Not Me.GetRedrawSuspension) Then
        
        'Request the scroll-specific viewport pipeline stage
        Viewport_Engine.Stage3_ExtractRelevantRegion pdImages(g_CurrentImage), Me
        
        'Notify any other relevant UI elements
        RelayViewportChanges
        
    End If
    
End Sub

Private Sub UserControl_Resize()

    'Align the canvas picture box to fill the available area
    AlignCanvasView
    
End Sub

Public Sub UpdateCanvasLayout()
    If g_OpenImageCount = 0 Then Me.ClearCanvas Else Me.AlignCanvasView
    StatusBar.ReflowStatusBar (g_OpenImageCount > 0)
End Sub

'RELAY: (not really, but all these measurements need to be migrated to ucSupport)
Public Sub AlignCanvasView()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = UserControl.ScaleWidth
    bHeight = UserControl.ScaleHeight
    
    'As of version 7.0, scroll bars are always visible.  This matches the behavior of paint-centric software like Krita,
    ' and makes it much easier to enable scrolling past the edge of an image (without resorting to stupid click-hold
    ' scroll behavior like GIMP).
    Dim hScrollTop As Long, hScrollLeft As Long, vScrollTop As Long, vScrollLeft As Long
    hScrollLeft = 0
    hScrollTop = bHeight - (Me.GetStatusBarHeight + hScroll.GetHeight)
    
    vScrollLeft = bWidth - vScroll.GetWidth
    vScrollTop = 0
    
    'With scroll bar positions calculated, calculate width/height values for the main canvas picture box
    Dim cvTop As Long, cvLeft As Long, cvWidth As Long, cvHeight As Long
    cvTop = 0
    cvLeft = 0
    cvWidth = vScrollLeft - cvLeft
    cvHeight = hScrollTop - cvTop
    
    'Move the CanvasView box into position first
    If (CanvasView.GetLeft <> cvLeft) Or (CanvasView.GetTop <> cvTop) Or (CanvasView.GetWidth <> cvWidth) Or (CanvasView.GetHeight <> cvHeight) Then
        If cvWidth > 0 And cvHeight > 0 Then CanvasView.SetPositionAndSize cvLeft, cvTop, cvWidth, cvHeight
    End If
    
    '...Followed by the scrollbars
    If (hScroll.Left <> hScrollLeft) Or (hScroll.Top <> hScrollTop) Or (hScroll.Width <> cvWidth) Then
        If cvWidth > 0 Then hScroll.SetPositionAndSize hScrollLeft, hScrollTop, cvWidth, hScroll.GetHeight
    End If
    
    If (vScroll.Left <> vScrollLeft) Or (vScroll.Top <> vScrollTop) Or (vScroll.Height <> cvHeight) Then
        If cvHeight > 0 Then vScroll.SetPositionAndSize vScrollLeft, vScrollTop, vScroll.GetWidth, cvHeight
    End If
    
    '...Followed by the "center" button (which sits between the scroll bars)
    If (cmdCenter.GetLeft <> vScrollLeft) Or (cmdCenter.GetTop <> hScrollTop) Then
        cmdCenter.SetLeft vScrollLeft
        cmdCenter.SetTop hScrollTop
    End If
    
End Sub

Private Sub UserControl_Show()

    If g_IsProgramRunning Then
        
        'Prep the command buttons
        cmdCenter.AssignImage "SB_ZOOM_CENTER"
        
        'XP users may not have Segoe UI available, which will cause the following lines to throw an error;
        ' it's not really a problem, as the labels will just keep their Tahoma font, but we must catch it anyway.
        On Error GoTo CanvasShowError
                
        'TODO: see if this is still valid, since we've moved theming much earlier in the load process
        'Now comes a bit of an odd case.  This control's _Show event happens relatively early in the load process due to it
        ' being present on FormMain.  Because of that, the global interface font value may not be loaded yet.  To avoid
        ' problems from this, we will just load Segoe UI by default, and if that fails (as it may on XP), the labels will
        ' retain their default Tahoma label.
        
        'Convert all labels to the current interface font
        If Len(g_InterfaceFont) = 0 Then g_InterfaceFont = "Segoe UI"
        
        'Request an update against the current theme
        ' TODO: do we really need to do this manually, or is PD's central theme function smart enough to do this on its own?
        Me.UpdateAgainstCurrentTheme
        
CanvasShowError:
        
    End If
    
    Exit Sub

End Sub

Private Sub VScroll_Scroll(ByVal eventIsCritical As Boolean)
        
    'Regardless of viewport state, cache the current scroll bar value inside the current image
    If Not pdImages(g_CurrentImage) Is Nothing Then
        pdImages(g_CurrentImage).imgViewport.setVScrollValue vScroll.Value
    End If
        
    If (Not Me.GetRedrawSuspension) Then
    
        'Request the scroll-specific viewport pipeline stage
        Viewport_Engine.Stage3_ExtractRelevantRegion pdImages(g_CurrentImage), Me
        
        'Notify any other relevant UI elements
        RelayViewportChanges
        
    End If
    
End Sub

Public Function PopulateSizeUnits()
    StatusBar.PopulateSizeUnits
End Function

'Whenever the mouse cursor needs to be reset, use this function to do so.  Also, when a new tool is created or a new tool feature
' is added, make sure to visit this sub and make any necessary cursor changes!
'
'A lot of extra values are passed to this function.  Individual tools can use those at their leisure to customize their cursor requests.
' RELAY: the actual cursor request needs to be passed to pdCanvasView, and we need to make sure its MouseEnter event also calls this.
Private Sub SetCanvasCursor(ByVal curMouseEvent As PD_MOUSEEVENT, ByVal Button As Integer, ByVal x As Single, ByVal y As Single, ByVal imgX As Double, ByVal imgY As Double, ByVal layerX As Double, ByVal layerY As Double)
    
    'Some cursor functions operate on a POI basis
    Dim curPOI As Long

    'Obviously, cursor setting is handled separately for each tool.
    Select Case g_CurrentTool
        
        Case NAV_DRAG
        
            'When click-dragging the image to scroll around it, the cursor depends on being over the image
            If IsMouseOverImage(x, y, pdImages(g_CurrentImage)) Then
                
                If Button <> 0 Then
                    CanvasView.RequestCursor_PNG "HANDCLOSED", 0, 0
                Else
                    CanvasView.RequestCursor_PNG "HANDOPEN", 0, 0
                End If
            
            'If the cursor is not over the image, change to an arrow cursor
            Else
                CanvasView.RequestCursor_System IDC_ARROW
            End If
        
        Case NAV_MOVE
            
            'When transforming layers, the cursor depends on the active POI
            curPOI = pdImages(g_CurrentImage).getActiveLayer.checkForPointOfInterest(layerX, layerY)
            
            Select Case curPOI
            
                'Mouse is not over the current layer
                Case -1
                    CanvasView.RequestCursor_System IDC_ARROW
                    
                'Mouse is over the top-left corner
                Case 0
                    CanvasView.RequestCursor_System IDC_SIZENWSE
                    
                'Mouse is over the top-right corner
                Case 1
                    CanvasView.RequestCursor_System IDC_SIZENESW
                    
                'Mouse is over the bottom-left corner
                Case 2
                    CanvasView.RequestCursor_System IDC_SIZENESW
                    
                'Mouse is over the bottom-right corner
                Case 3
                    CanvasView.RequestCursor_System IDC_SIZENWSE
                    
                'Mouse is over a rotation handle
                Case 4 To 7
                    CanvasView.RequestCursor_System IDC_SIZEALL
                    
                'Mouse is within the layer, but not over a specific node
                Case 8
                
                    'This case is unique because if the user has elected to ignore transparent pixels, they cannot move a layer
                    ' by dragging the mouse within a transparent region of the layer.  Thus, before changing the cursor,
                    ' check to see if the hovered layer index is the same as the current layer index; if it isn't, don't display
                    ' the Move cursor.  (Note that this works because the getLayerUnderMouse function, called during the MouseMove
                    ' event, automatically factors the transparency check into its calculation.  Thus we don't have to
                    ' re-evaluate the setting here.)
                    If m_LayerAutoActivateIndex = pdImages(g_CurrentImage).getActiveLayerIndex Then
                        CanvasView.RequestCursor_System IDC_SIZEALL
                    Else
                        CanvasView.RequestCursor_System IDC_ARROW
                    End If
                    
            End Select
            
            'The move tool is unique because it will request a redraw of the viewport when the POI changes, so that the current
            ' POI can be highlighted.
            If m_LastPointOfInterest <> curPOI Then
                m_LastPointOfInterest = curPOI
                Viewport_Engine.Stage5_FlipBufferAndDrawUI pdImages(g_CurrentImage), Me, curPOI
            End If
            
        Case SELECT_RECT, SELECT_CIRC
        
            'When transforming selections, the cursor image depends on its proximity to a point of interest.
            '
            'For a rectangle or circle selection, the possible transform IDs are:
            ' -1 - Cursor is not near a selection point
            ' 0 - NW corner
            ' 1 - NE corner
            ' 2 - SE corner
            ' 3 - SW corner
            ' 4 - N edge
            ' 5 - E edge
            ' 6 - S edge
            ' 7 - W edge
            ' 8 - interior of selection, not near a corner or edge
            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
            
                Case -1
                    CanvasView.RequestCursor_System IDC_ARROW
                Case 0
                    CanvasView.RequestCursor_System IDC_SIZENWSE
                Case 1
                    CanvasView.RequestCursor_System IDC_SIZENESW
                Case 2
                    CanvasView.RequestCursor_System IDC_SIZENWSE
                Case 3
                    CanvasView.RequestCursor_System IDC_SIZENESW
                Case 4
                    CanvasView.RequestCursor_System IDC_SIZENS
                Case 5
                    CanvasView.RequestCursor_System IDC_SIZEWE
                Case 6
                    CanvasView.RequestCursor_System IDC_SIZENS
                Case 7
                    CanvasView.RequestCursor_System IDC_SIZEWE
                Case 8
                    CanvasView.RequestCursor_System IDC_SIZEALL
            
            End Select
        
        Case SELECT_LINE
        
            'When transforming selections, the cursor image depends on its proximity to a point of interest.
            '
            'For a line selection, the possible transform IDs are:
            ' -1 - Cursor is not near an endpoint
            ' 0 - Near x1/y1
            ' 1 - Near x2/y2
            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
            
                Case -1
                    CanvasView.RequestCursor_System IDC_ARROW
                Case 0
                    CanvasView.RequestCursor_System IDC_SIZEALL
                Case 1
                    CanvasView.RequestCursor_System IDC_SIZEALL
            
            End Select
        
         Case SELECT_POLYGON
            
            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
            
                '-1: mouse is outside the lasso selection area
                Case -1
                    CanvasView.RequestCursor_System IDC_ARROW
                
                'numOfPolygonPoints: mouse is inside the polygon, but not over a polygon node
                Case pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints
                    If pdImages(g_CurrentImage).mainSelection.isLockedIn Then
                        CanvasView.RequestCursor_System IDC_SIZEALL
                    Else
                        CanvasView.RequestCursor_System IDC_ARROW
                    End If
                    
                'Everything else: mouse is over a polygon node
                Case Else
                    CanvasView.RequestCursor_System IDC_SIZEALL
                    
            End Select
        
        Case SELECT_LASSO
            
            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
            
                '-1: mouse is outside the lasso selection area
                Case -1
                    CanvasView.RequestCursor_System IDC_ARROW
                
                '0: mouse is inside the lasso selection area.  As a convenience to the user, we don't update the cursor
                '   if they're still in "drawing" mode - we only update it if the selection is complete.
                Case 0
                    If pdImages(g_CurrentImage).mainSelection.isLockedIn Then
                        CanvasView.RequestCursor_System IDC_SIZEALL
                    Else
                        CanvasView.RequestCursor_System IDC_ARROW
                    End If
                    
            End Select
            
        Case SELECT_WAND
        
            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
            
                '-1: mouse is outside the lasso selection area
                Case -1
                    CanvasView.RequestCursor_System IDC_ARROW
                
                '0: mouse is inside the lasso selection area.  As a convenience to the user, we don't update the cursor
                '   if they're still in "drawing" mode - we only update it if the selection is complete.
                Case Else
                    CanvasView.RequestCursor_System IDC_SIZEALL
                    
            End Select
        
        Case VECTOR_TEXT, VECTOR_FANCYTEXT

            'The text tool bears a lot of similarity to the Move / Size tool, although the resulting behavior is
            ' obviously quite different.
            
            'First, see if the active layer is a text layer.  If it is, we need to check for POIs.
            If pdImages(g_CurrentImage).getActiveLayer.isLayerText Then
                
                'When transforming layers, the cursor depends on the active POI
                curPOI = pdImages(g_CurrentImage).getActiveLayer.checkForPointOfInterest(layerX, layerY)
                
                Select Case curPOI
    
                    'Mouse is not over the current layer
                    Case -1
                        CanvasView.RequestCursor_System IDC_IBEAM
    
                    'Mouse is over the top-left corner
                    Case 0
                        CanvasView.RequestCursor_System IDC_SIZENWSE
                    
                    'Mouse is over the top-right corner
                    Case 1
                        CanvasView.RequestCursor_System IDC_SIZENESW
                    
                    'Mouse is over the bottom-left corner
                    Case 2
                        CanvasView.RequestCursor_System IDC_SIZENESW
                    
                    'Mouse is over the bottom-right corner
                    Case 3
                        CanvasView.RequestCursor_System IDC_SIZENWSE
                        
                    'Mouse is over a rotation handle
                    Case 4 To 7
                        CanvasView.RequestCursor_System IDC_SIZEALL
                    
                    'Mouse is within the layer, but not over a specific node
                    Case 8
                        CanvasView.RequestCursor_System IDC_SIZEALL
                    
                End Select
                
                'Similar to the move tool, texts tools will request a redraw of the viewport when the POI changes, so that the current
                ' POI can be highlighted.
                If m_LastPointOfInterest <> curPOI Then
                    m_LastPointOfInterest = curPOI
                    Viewport_Engine.Stage5_FlipBufferAndDrawUI pdImages(g_CurrentImage), Me, curPOI
                End If
                
            'If the current layer is *not* a text layer, clicking anywhere will create a new text layer
            Else
                CanvasView.RequestCursor_System IDC_IBEAM
            End If
        
        Case Else
            CanvasView.RequestCursor_System IDC_ARROW
                    
    End Select

End Sub

'Simple, unified way to see if canvas interaction is allowed.
Public Function IsCanvasInteractionAllowed() As Boolean
    IsCanvasInteractionAllowed = CanvasView.IsCanvasInteractionAllowed
End Function

'If the viewport experiences changes to scroll or zoom values, this function will be automatically called.  Any relays to external
' functions (functions that rely on viewport settings, obviously) should be handled here.
' TODO: migrate this function elsewhere, so things other than the canvas can easily utilize it.
Public Sub RelayViewportChanges()
    toolbar_Layers.NotifyViewportChange
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDC_Background, "Background", IDE_GRAY
    m_Colors.LoadThemeColor PDC_StatusBar, "StatusBar", IDE_GRAY
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating all button captions against the current language.
Public Sub UpdateAgainstCurrentTheme()
    
    'Suspend redraws until all theme updates are complete
    Me.SetRedrawSuspension True
    
    UpdateColorList
    CanvasView.UpdateAgainstCurrentTheme
    StatusBar.UpdateAgainstCurrentTheme
    
    'Reassign tooltips to any relevant controls.  (This also triggers a re-translation against language changes.)
    cmdCenter.AssignTooltip "Center the image inside the viewport"
    If Not (g_Themer Is Nothing) Then cmdCenter.BackColor = g_Themer.GetThemeColor(PDTC_BACKGROUND_COMMANDBAR)
    cmdCenter.UpdateAgainstCurrentTheme
    
    hScroll.UpdateAgainstCurrentTheme
    vScroll.UpdateAgainstCurrentTheme
    
    'Any controls that utilize a custom background color must now be updated to match *our* background color.
    Dim sbBackColor As Long
    sbBackColor = m_Colors.RetrieveColor(PDC_StatusBar, Me.Enabled)
    cmdCenter.BackColor = sbBackColor
    
    Me.UpdateCanvasLayout
    
    'Restore redraw capabilities
    Me.SetRedrawSuspension False
        
End Sub

