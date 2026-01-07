VERSION 5.00
Begin VB.Form toolbar_Layers 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Layers"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   ControlBox      =   0   'False
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
   HasDC           =   0   'False
   Icon            =   "Toolbar_Layers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer ctlContainer 
      Height          =   615
      Index           =   0
      Left            =   240
      Top             =   3840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   476
      Caption         =   "search"
      Draggable       =   -1  'True
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   476
      Caption         =   "colors"
      Draggable       =   -1  'True
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   476
      Caption         =   "overview"
      Draggable       =   -1  'True
   End
   Begin PhotoDemon.pdContainer ctlContainer 
      Height          =   615
      Index           =   1
      Left            =   240
      Top             =   4560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdContainer ctlContainer 
      Height          =   615
      Index           =   2
      Left            =   240
      Top             =   5280
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   270
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   476
      Caption         =   "layers"
      Draggable       =   -1  'True
   End
   Begin PhotoDemon.pdContainer ctlContainer 
      Height          =   615
      Index           =   3
      Left            =   240
      Top             =   6000
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
   End
End
Attribute VB_Name = "toolbar_Layers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Right-side ("Layers") Toolbar
'Copyright 2014-2026 by Tanner Helland
'Created: 25/March/14
'Last updated: 08/October/21
'Last update: fix the way the form is initialized on large screens (ensure saved panel positions are loaded correctly,
'             by *first* ensuring the form is its correct final size)
'
'For historical reasons, I call this the "layers" toolbar, but it actually encompasses everything that appears on
' the right-side toolbar.  Most of the code in this window is dedicated to supporting collapsible/resizable panels,
' so it's 90+% UX-related.
'
'For details on the individual panels, refer to the various layerpanel_* forms.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Because this form is resizable at run-time, we need to play some games with mouse capturing
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTLEFT As Long = 10

'The value of all controls on this form are saved and loaded to file by this class
' (Normally this is declared WithEvents, but this dialog doesn't require custom settings behavior.)
Private m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'Helper class to synchronize various subpanels with the picture boxes we use for positioning
Private m_WindowSync As pdWindowSync

'When we are responsible for this window resizing (because the user is resizing our window manually), we set this to TRUE.
' This variable is then checked before requesting additional redraws during our resize event.
Private m_WeAreResponsibleForResize As Boolean

'How close does the mouse have to be to the form border to allow resizing? Currently we use this constant,
' while accounting for DPI variance (e.g. this value represents (n) pixels *at 96 dpi*)
Private Const RESIZE_BORDER As Long = 6

'A dedicated mouse handler helps provide cursor handling
Private WithEvents m_MouseEvents As pdInputMouse
Attribute m_MouseEvents.VB_VarHelpID = -1

'Number of panels; set automatically at Form_Load
Private m_NumOfPanels As Long

'When the form is first loaded, we need to suspend all layout decisions until the form height has been
' correctly set.  This flag means DO NOT CALCULATE LAYOUT data yet.
Private m_FormStillLoading As Boolean

'When the user is in the midst of resizing a vertical panel, this will be set to a value >= 0
' (corresponding to the panel being resized).
Private m_PanelResizeActive As Long

'Similarly, during a panel resize, we track both the panel's initial size, and its "net change" amount.  We use these
' to resize the panel "on the fly"
Private m_NetResizeAmount As Long, m_PanelStartHeight As Long

Private Sub Form_Load()
        
    'Before doing anything else, we need to set this form to its final height - the client height of the
    ' main window.  (If we don't do this first, all subsequent layout initialization - like setting
    ' initial panel heights - may be wrong, because they're contingent on the size of the full panel.)
    If PDMain.IsProgramRunning() And (Not g_WindowManager Is Nothing) Then
        
        Dim parentClientHeight As Long
        parentClientHeight = g_WindowManager.GetClientHeight(FormMain.hWnd)
        
        'Set a flag to prevent additional layout calculations, then *immediately* apply the new height.
        m_FormStillLoading = True
        g_WindowManager.SetSizeByHWnd Me.hWnd, g_WindowManager.GetWindowWidth(Me.hWnd), parentClientHeight, True
        m_FormStillLoading = False
        
    End If
        
    'All layout decisions on this form are contingent on the number of panels, so set this first as subsequent code
    ' will likely rely on it.
    m_NumOfPanels = ttlPanel.Count
    m_PanelResizeActive = -1
    
    'Initialize panel height values.
    ' (Note that we do not calculate a hard-coded size for the final panel (layers).  It is autosized to fill whatever
    '  space remains after the panels above it are positioned.)
    Dim pnlDefaultHeight As Long, targetHeight As Long, pnlHeightSearch As Long
    pnlHeightSearch = Interface.FixDPI(24)
    pnlDefaultHeight = Interface.FixDPI(100)
    
    Dim i As Long
    For i = 0 To m_NumOfPanels - 1
    
        If UserPrefs.IsReady Then
            
            targetHeight = UserPrefs.GetPref_Long("Toolbox", "RightPanelSize" & CStr(i + 1), pnlDefaultHeight)
            If (i = 0) Then
                targetHeight = pnlHeightSearch
            Else
                
                'Scale by current DPI (which may have changed since last session - PD is portable!)
                Dim lastDPI As Single
                lastDPI = Interface.GetLastSessionDPI_Ratio()
                targetHeight = targetHeight * (Interface.GetSystemDPIRatio() / lastDPI)
                
            End If
            
        Else
            If (i = 0) Then
                targetHeight = pnlHeightSearch
            Else
                targetHeight = pnlDefaultHeight
            End If
        End If
        
        ctlContainer(i).SetHeight targetHeight
        
    Next i
    
    'Prep a mouse handler for the underlying form
    Set m_MouseEvents = New pdInputMouse
    m_MouseEvents.AddInputTracker Me.hWnd, , True
    
    'Prep a window synchronizer and add each subpanel to it
    Set m_WindowSync = New pdWindowSync
    
    'It can take quite some time to load these panels, so during debugging, it's helpful to track
    ' any unintentional changes to load time (which in turn harm PD's average startup time).
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    Load layerpanel_Search
    m_WindowSync.SynchronizeWindows ctlContainer(0).hWnd, layerpanel_Search.hWnd
    layerpanel_Search.Show
    PDDebug.LogTiming "right toolbox / search panel", VBHacks.GetTimerDifferenceNow(startTime)
    VBHacks.GetHighResTime startTime
    
    Load layerpanel_Navigator
    m_WindowSync.SynchronizeWindows ctlContainer(1).hWnd, layerpanel_Navigator.hWnd
    layerpanel_Navigator.Show
    PDDebug.LogTiming "right toolbox / navigator panel", VBHacks.GetTimerDifferenceNow(startTime)
    VBHacks.GetHighResTime startTime
    
    Load layerpanel_Colors
    m_WindowSync.SynchronizeWindows ctlContainer(2).hWnd, layerpanel_Colors.hWnd
    layerpanel_Colors.Show
    PDDebug.LogTiming "right toolbox / color panel", VBHacks.GetTimerDifferenceNow(startTime)
    VBHacks.GetHighResTime startTime
    
    Load layerpanel_Layers
    m_WindowSync.SynchronizeWindows ctlContainer(ctlContainer.UBound).hWnd, layerpanel_Layers.hWnd
    layerpanel_Layers.Show
    PDDebug.LogTiming "right toolbox / layers panel", VBHacks.GetTimerDifferenceNow(startTime)
    VBHacks.GetHighResTime startTime
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
        
    'Theme everything
    Me.UpdateAgainstCurrentTheme True
    
    PDDebug.LogTiming "right toolbox / everything else", VBHacks.GetTimerDifferenceNow(startTime)
    
    'Technically, we would now want to call ReflowInterface() to make sure everything is correctly aligned.
    ' However, UpdateAgainstCurrentTheme now calls that function automatically.
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    If Not (m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues
        m_lastUsedSettings.SetParentForm Nothing
    End If
    
    'Some settings are not stored inside the last-used settings file, but in the central PD settings file.
    ' (This is done so that a full "reset" of the core settings file appropriately resets the panel sizes, too.)
    If UserPrefs.IsReady Then
        Dim i As Long
        For i = 0 To m_NumOfPanels - 1
            UserPrefs.SetPref_Long "Toolbox", "RightPanelSize" & CStr(i + 1), ctlContainer(i).GetHeight
        Next i
    End If
    
End Sub

Private Sub Form_Resize()
    If (Not m_FormStillLoading) Then ReflowInterface
End Sub

'Toolbars can never be unloaded, EXCEPT when the whole program is going down.  Check for the program-wide closing flag prior
' to exiting; if it is not found, cancel the unload and simply hide this form.  (Note that the ToggleToolboxVisibility sub
' will also keep this toolbar's Window menu entry in sync with the form's current visibility.)
Private Sub Form_Unload(Cancel As Integer)
    
    If g_ProgramShuttingDown Then
        
        'Release this window from any program-wide handlers
        ReleaseFormTheming Me
        
        'Release our custom mouse handler
        Set m_MouseEvents = Nothing
        
        'Release the subpanel subclasser
        Set m_WindowSync = Nothing
        
        'Unload all child forms
        Unload layerpanel_Search
        Unload layerpanel_Navigator
        Unload layerpanel_Colors
        Unload layerpanel_Layers
        
    Else
        PDDebug.LogAction "WARNING!  toolbar_Layers was unloaded prematurely - why??"
        Cancel = True
    End If
    
End Sub

Private Sub m_MouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'Ignore all mouse events while the user is interacting with the canvas
    If FormMain.MainCanvas(0).IsMouseDown(pdLeftButton Or pdRightButton) Then Exit Sub
    
    'If the mouse is near the resizable edge of the toolbar (the left edge, currently), allow the user to resize
    ' the layer toolbox.
    Dim mouseInResizeTerritory As Boolean
    
    'Check the mouse position to see if it's in resize territory (along the left edge of the toolbox)
    mouseInResizeTerritory = (y > 0) And (y < Me.ScaleHeight) And (x < Interface.FixDPI(RESIZE_BORDER))
    
    'If the left mouse button is down, and the mouse is in resize territory, initiate an API resize event
    If mouseInResizeTerritory Then
        
        'Change the cursor to a resize cursor
        m_MouseEvents.SetCursor_System IDC_SIZEWE
        
        If (Button And vbLeftButton <> 0) Then
            
            m_WeAreResponsibleForResize = True
            ReleaseCapture
            VBHacks.SendMsgW Me.hWnd, WM_NCLBUTTONDOWN, HTLEFT, 0&
            
            'After the toolbox has been resized, we need to manually notify the toolbox manager, so it can
            ' notify any neighboring toolboxes (and/or the central canvas)
            Toolboxes.SetConstrainingSize PDT_RightToolbox, Me.ScaleWidth
            FormMain.UpdateMainLayout
            
            'A premature exit is required, because the end of this sub contains code to detect the release of the
            ' mouse after a drag event.  Because the event is not being initiated normally, we can't detect a standard
            ' MouseUp event, so instead, we mimic it by checking MouseMove and m_WeAreResponsibleForResize = TRUE.
            Exit Sub
            
        End If
    
    Else
        m_MouseEvents.SetCursor_System IDC_DEFAULT
    End If
    
    'Check for mouse release; we will only reach this point if the mouse is *not* in resize territory, which in turn
    ' means we can free the release code and resize the window now.  (On some OS/theme combinations, the canvas will
    ' live-resize as the mouse is moved.  On others, the canvas won't redraw until the mouse is released.)
    If m_WeAreResponsibleForResize Then
        m_WeAreResponsibleForResize = False
        m_MouseEvents.SetCursor_System IDC_DEFAULT
    End If
    
End Sub

Private Sub m_MouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    m_MouseEvents.SetCursor_System IDC_DEFAULT
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    
    'If a panel is opening, redraw any elements that have may been suppressed while the panel was invisible
    If newState Then NotifyLayerChange
    
    'Reflow the interface to account for the changed size
    ReflowInterface
    
End Sub

Private Sub ttlPanel_MouseDownCustom(Index As Integer, ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'Only panels after the first one can be resized (as the first panel sits at the top of the toolbox, and it must
    ' always remain aligned there).  Note also that dragging a titlebar resizes the panel *above* this one
    ' (hence the -1 on the line below).
    If (Index > 1) And (Not g_WindowManager Is Nothing) And ((Button And pdLeftButton) <> 0) Then
        m_PanelResizeActive = Index - 1
        m_PanelStartHeight = ctlContainer(m_PanelResizeActive).GetHeight
        m_NetResizeAmount = 0
    End If
    
End Sub

Private Sub ttlPanel_MouseDrag(Index As Integer, ByVal xChange As Long, ByVal yChange As Long)
    
    'The user is click-dragging a titlebar to resize its associated panel.  Calculate a new height and immediately
    ' reflow the interface to match.
    If (m_PanelResizeActive >= 0) Then
        m_NetResizeAmount = yChange
        ReflowInterface
    End If
    
End Sub

Private Sub ttlPanel_MouseUpCustom(Index As Integer, ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    
    'After a drag event, reset the resize tracker
    If (m_PanelResizeActive >= 0) And ((Button And pdLeftButton) <> 0) Then m_PanelResizeActive = -1
    
End Sub

'When one or more layers are modified (via painting, effects, whatever), PD's various interface control functions
' will notify this toolbar via this function.  The toolbar will then redraw individual panels as necessary.
'
'Note that a layerID of -1 means multiple/all layers have changed, while a value >= 0 tells you which layer changed,
' perhaps sparing the amount of redraw work required.
Public Sub NotifyLayerChange(Optional ByVal layerID As Long = -1)
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Ideally, we wouldn't redraw the layer box unless it's actually visible, but we need to ensure that the layer
    ' box's internal caches of things like layer thumbnails stays relevant to image state.  (Otherwise, if the panel
    ' is closed and then the user later opens it, it would be completely out of sync!)  As such, we always redraw
    ' the layer box, regardless of whether it's visible or not.
    layerpanel_Layers.ForceRedraw True, layerID
    
    If ttlPanel(1).Value Then layerpanel_Navigator.nvgMain.NotifyNewThumbNeeded
    'pdDebug.LogAction "toolbar_Layers.NotifyLayerChange finished in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub

'If the current viewport position and/or size changes, this toolbar will be notified.  At present, the only subpanel
' affected by viewport changes is the navigator panel.
Public Sub NotifyViewportChange()
    If ttlPanel(1).Value Then layerpanel_Navigator.nvgMain.NotifyNewViewportPosition
End Sub

Public Sub ResetInterface()

    'Reset all panels to their default heights
    Dim i As Long
    For i = 0 To m_NumOfPanels - 1
        If (i = 0) Then
            ctlContainer(i).SetHeight Interface.FixDPI(24)
        Else
            ctlContainer(i).SetHeight Interface.FixDPI(100)
        End If
        ttlPanel(i).Value = True
    Next i
    
    'Reflow the interface to match
    ReflowInterface
    
End Sub

'Whenever the layer toolbox is resized, we must reflow all objects to fill the available space.  Note that we do not do
' specialized handling for the vertical direction; vertically, the only change we handle is resizing the layer box itself
' to fill whatever vertical space is available.
Private Sub ReflowInterface()
    
    'If the form is invisible (due to minimize or something else), just exit now
    Dim formWidth As Long, formHeight As Long
    If (g_WindowManager Is Nothing) Then
        formWidth = Me.ScaleWidth
        formHeight = Me.ScaleHeight
    Else
        formWidth = g_WindowManager.GetClientWidth(Me.hWnd)
        formHeight = g_WindowManager.GetClientHeight(Me.hWnd)
    End If
    
    If (formWidth <= 0) Or (formHeight <= 0) Then Exit Sub
    
    'When the parent form is resized, resize the layer list (and other items) to properly fill the
    ' available horizontal and vertical space.
    
    'Next, we want to resize all subpanel picture boxes, so that their size reflects the new form size.  This is a
    ' bit complicated, as each form has a different base size, and the user can toggle panel visibility at any time.
    
    'Start by calculating initial x/y offsets
    Dim yOffset As Long, xOffset As Long, xWidth As Long
    xOffset = Interface.FixDPI(RESIZE_BORDER)
    yOffset = Interface.FixDPI(2)
    xWidth = formWidth - xOffset
    
    'Treat the following values as constants
    Dim MIN_PANEL_SIZE As Long, MAX_PANEL_SIZE As Long, MIN_LAYER_PANEL_SIZE As Long
    MIN_PANEL_SIZE = Interface.FixDPI(70)
    MAX_PANEL_SIZE = Interface.FixDPI(320)
    MIN_LAYER_PANEL_SIZE = Interface.FixDPI(173)
    
    Dim FIXED_PANEL_SIZE_SEARCH As Long
    FIXED_PANEL_SIZE_SEARCH = Interface.FixDPI(24)
    
    'We now calculate the toolbar's layout in two passes.  First, we calculate new layout rects for all objects on
    ' the form (including titlebars and containers).  Next, we validate all positions by ensuring that all visible
    ' containers have enough room to correctly display their contents.  (If they don't, we shift stuff around until
    ' validity is reached.)  Finally, we apply all the new positions and render the results.
    Dim ttlRects() As RectF, pnlRects() As RectF
    ReDim ttlRects(0 To m_NumOfPanels - 1) As RectF
    ReDim pnlRects(0 To m_NumOfPanels - 1) As RectF
    
    'First pass: calculate all rects using the user's current layout settings.
    Dim i As Long, tmpHeight As Long
    For i = 0 To m_NumOfPanels - 1
        
        'Move the titlebar of this panel into position
        With ttlRects(i)
            .Left = xOffset
            .Top = yOffset
            .Width = xWidth - xOffset + Interface.FixDPI(2)
            .Height = ttlPanel(i).GetHeight
        End With
        
        'Move the yOffset beneath the panel
        yOffset = yOffset + ttlRects(i).Height + Interface.FixDPI(1)
        
        'If the title bar state is TRUE, calculate a layout rect for its associated panel
        If ttlPanel(i).Value Then
            
            'Move this panel into position.  (The x-check is a failsafe check only, for weird circumstances
            ' when the form is created and its size is not yet properly set.)
            If (xWidth - xOffset > 0) Then
                
                With pnlRects(i)
                    .Left = Int(CSng(xOffset) * 1.5 + 0.5)
                    .Top = yOffset
                    .Width = xWidth - xOffset
                End With
                    
                'The bottom panel (the layer panel) is handled specially, as it auto-sizes to fill any remaining
                ' vertical space.
                If (i = m_NumOfPanels - 1) Then
                    tmpHeight = (formHeight - yOffset)
                    pnlRects(i).Height = tmpHeight
                    
                'The top panel (the search panel) is also handled specially, as it remains a fixed size no matter what.
                ElseIf (i = 0) Then
                
                    pnlRects(i).Height = FIXED_PANEL_SIZE_SEARCH
                
                'All other panels
                Else
                
                    'If a panel is in the midst of a user-initiated resize, let's calculate its height as the sum of
                    ' its original height, and however far the user has vertically dragged the mouse.
                    If (i = m_PanelResizeActive) Then
                        tmpHeight = m_PanelStartHeight + m_NetResizeAmount
                    
                    '...otherwise, try to preserve the container's existing height
                    Else
                        tmpHeight = ctlContainer(i).GetHeight
                    End If
                    
                    'Because the user has control over panel height, we need to perform some checks to ensure the target
                    ' panel's height is an acceptable value
                    If (tmpHeight < MIN_PANEL_SIZE) Then tmpHeight = MIN_PANEL_SIZE
                    If (tmpHeight > MAX_PANEL_SIZE) Then tmpHeight = MAX_PANEL_SIZE
                    pnlRects(i).Height = tmpHeight
                    
                End If
                
            Else
                With pnlRects(i)
                    .Left = 0
                    .Top = yOffset
                    .Width = 1
                    .Height = 1
                End With
            End If
            
            'Add this panel's height to the running offset calculation.
            yOffset = yOffset + pnlRects(i).Height
        
        Else
            pnlRects(i).Height = 1
        End If
        
        'Calculate the new top position of the next panel in line.
        yOffset = yOffset + Interface.FixDPI(2)
        
    Next i
    
    Dim spaceNeeded As Long, j As Long
    Dim initHeight As Long, heightChange As Long
    
    'With all positions calculated, we now need to ensure that there is a valid amount of space for all panels.
    ' At present, this mostly just means ensuring that the layer box (if open) has enough room to display correctly.
    If ttlPanel(m_NumOfPanels - 1).Value Then
    
        'Figure out how much space we need to "make available" for the layer panel
        spaceNeeded = MIN_LAYER_PANEL_SIZE - pnlRects(m_NumOfPanels - 1).Height
        
        If (spaceNeeded > 0) Then
        
            'Set the layers panel to the minimum allowable size
            pnlRects(m_NumOfPanels - 1).Height = MIN_LAYER_PANEL_SIZE
            
            'Starting at the bottom and moving up, remove space from other panels until we have enough space to
            ' properly fit the layer panel.
            For i = (m_NumOfPanels - 2) To 0 Step -1
                
                'If this panel is open, remove as much height from it as we physically can
                If ttlPanel(i).Value Then
                
                    initHeight = pnlRects(i).Height
                    pnlRects(i).Height = pnlRects(i).Height - spaceNeeded
                    If (pnlRects(i).Height < MIN_PANEL_SIZE) Then pnlRects(i).Height = MIN_PANEL_SIZE
                    
                    'If we were able to remove 1+ pixels from this panel (because it was larger than the minimum
                    ' allowed size), shift all subsequent panels upward to compensate.
                    heightChange = (initHeight - pnlRects(i).Height)
                    If (heightChange > 0) Then
                    
                        'Adjust the running "space still needed" value to account for however many pixels we
                        ' just removed.
                        spaceNeeded = spaceNeeded - heightChange
                        
                        'Adjust the top position of subsequent panels and titlebars to match this new panel size
                        For j = i To m_NumOfPanels - 2
                            ttlRects(j + 1).Top = pnlRects(j).Top + pnlRects(j).Height + Interface.FixDPI(2)
                            pnlRects(j + 1).Top = ttlRects(j + 1).Top + ttlRects(j + 1).Height + Interface.FixDPI(1)
                        Next j
                    
                    End If
                    
                    'If we've removed sufficient space for everything to fit, our work here is done!
                    If (spaceNeeded <= 0) Then Exit For
                    
                End If
                
            Next i
            
        End If
    
    'If the layer box is *not* open, we still need to ensure there is enough room for its titlebar, at least.
    Else
        
        'See if the titlebar's position is valid (e.g. it is fully visible).
        spaceNeeded = (ttlRects(m_NumOfPanels - 1).Top + ttlRects(m_NumOfPanels - 1).Height) - (formHeight - Interface.FixDPI(2))
        
        If (spaceNeeded > 0) Then
        
            'Starting at the bottom and moving up, remove space from other panels until we have enough space to
            ' properly fit the layer titlebar.
            For i = (m_NumOfPanels - 2) To 0 Step -1
                
                If ttlPanel(i).Value Then
                
                    'Remove as much height from this panel as we physically can
                    initHeight = pnlRects(i).Height
                    pnlRects(i).Height = pnlRects(i).Height - spaceNeeded
                    If (pnlRects(i).Height < MIN_PANEL_SIZE) Then pnlRects(i).Height = MIN_PANEL_SIZE
                    
                    'If we were able to remove 1+ pixels from this panel (because it was larger than the minimum
                    ' allowed size), shift all subsequent panels upward to compensate.
                    heightChange = (initHeight - pnlRects(i).Height)
                    If (heightChange > 0) Then
                    
                        'Adjust the running "space still needed" value to account for however many pixels we just removed.
                        spaceNeeded = spaceNeeded - heightChange
                        
                        'Adjust the top position of subsequent panels and titlebars to match this new panel size
                        For j = i To m_NumOfPanels - 2
                            ttlRects(j + 1).Top = pnlRects(j).Top + pnlRects(j).Height + Interface.FixDPI(2)
                            pnlRects(j + 1).Top = ttlRects(j + 1).Top + ttlRects(j + 1).Height + Interface.FixDPI(1)
                        Next j
                    
                    End If
                    
                    'If we've removed sufficient space for everything to fit, our work here is done!
                    If (spaceNeeded <= 0) Then Exit For
                    
                End If
                
            Next i
            
        End If
    
    End If
    
    'With all positions calculated, we can now move everything into position in one fell swoop
    For i = 0 To m_NumOfPanels - 1
        
        'Move the titlebar of this panel into position
        With ttlRects(i)
            ttlPanel(i).SetPositionAndSize .Left, .Top, .Width, .Height
        End With
        
        '...same for its attached panel
        With pnlRects(i)
            If ttlPanel(i).Value Then
                ctlContainer(i).SetPositionAndSize .Left, .Top, .Width, .Height
            
            'If this panel is hidden, don't bother settings its size
            Else
                ctlContainer(i).SetPosition .Left, .Top
            End If
        End With
        
        'If the title bar state is TRUE, open its corresponding panel.
        ctlContainer(i).Visible = ttlPanel(i).Value
        
    Next i
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal isFirstLoad As Boolean = False)
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me
    
    'Pass the theme update request to any active child forms.
    ' (Note that we don't have to do this on our initial load, because the panels will automatically
    ' theme themselves.)
    If ((Not layerpanel_Search Is Nothing) And (Not isFirstLoad)) Then layerpanel_Search.UpdateAgainstCurrentTheme
    If ((Not layerpanel_Navigator Is Nothing) And (Not isFirstLoad)) Then layerpanel_Navigator.UpdateAgainstCurrentTheme
    If ((Not layerpanel_Colors Is Nothing) And (Not isFirstLoad)) Then layerpanel_Colors.UpdateAgainstCurrentTheme
    If ((Not layerpanel_Layers Is Nothing) And (Not isFirstLoad)) Then layerpanel_Layers.UpdateAgainstCurrentTheme
    
    'Reflow the interface, to account for any language changes.  (This will also trigger a redraw of the layer list box.)
    ReflowInterface
    
End Sub

Public Sub SetFocusToSearchBox()
    If (Not ttlPanel(0).Value) Then ttlPanel(0).Value = True
    layerpanel_Search.SetFocusToSearchBox
End Sub
