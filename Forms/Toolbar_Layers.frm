VERSION 5.00
Begin VB.Form toolbar_Layers 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Layers"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   3735
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdContainer ctlContainer 
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1560
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
      Caption         =   "overview"
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
      Caption         =   "layers"
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
      Caption         =   "color selector"
   End
   Begin PhotoDemon.pdContainer ctlContainer 
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdContainer ctlContainer 
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
   End
   Begin VB.Line lnSeparatorLeft 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   440
   End
End
Attribute VB_Name = "toolbar_Layers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Right-side ("Layers") Toolbar
'Copyright 2014-2018 by Tanner Helland
'Created: 25/March/14
'Last updated: 19/February/18
'Last update: implement vertically resizable panels
'
'For historical reasons, I call this the "layers" toolbar, but it actually encompasses everything that appears on
' the right-side toolbar.  Most of the code in this window is dedicated to supporting collapsible panels, and all
' the messy UX handling that goes along with that.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'Helper class to synchronize various subpanels with the picture boxes we use for positioning
Private m_WindowSync As pdWindowSync

'When we are responsible for this window resizing (because the user is resizing our window manually), we set this to TRUE.
' This variable is then checked before requesting additional redraws during our resize event.
Private m_WeAreResponsibleForResize As Boolean

'How close does the mouse have to be to the form border to allow resizing? Currently we use this constant,
' while accounting for DPI variance (e.g. this value represents (n) pixels *at 96 dpi*)
Private Const RESIZE_BORDER As Long = 5

'A dedicated mouse handler helps provide cursor handling
Private WithEvents m_MouseEvents As pdInputMouse
Attribute m_MouseEvents.VB_VarHelpID = -1

'Panels within this right-side toolbox store a number of extra information bits.  These help us reflow the
' panel correctly at run-time.
Private Type PD_Panel

    'Initial height of each panel.  This is (currently) hard-coded for all panels except the layers panel;
    ' that panel is dynamically sized to fit any remaining vertical space in the toolbox.
    InitialHeight As Long
    
    'Current height of each panel.  This is typically identical to the initial height value, *except* during
    ' a resize operation.  This value should be used for all layout decisions.
    CurrentHeight As Long
    
    'To allow the user to dynamically resize panels, we need to track mouse events on the underlying form object.
    ' These rects define "interactive" areas for vertical resize operations.
    VerticalResizeRect As RectF
    
End Type

Private m_Panels() As PD_Panel

'Number of panels; set automatically at Form_Load
Private m_NumOfPanels As Long

'When the user is in the midst of resizing a vertical panel, this will be set to a value >= 0 (corresponding to the
' panel being resized), and the initial y value will be populated to some non-zero value.
Private m_PanelResizeActive As Long, m_PanelResizeInitY As Long

'Panel resizing requires mouse capturing.  (Otherwise, if the mouse leaves the underlying control - a likely scenario,
' as the user will "drag" the mouse over neighboring panels - other controls will steal mouse focus mid-resize.)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

'If the user has resized the right-side panels in a way that doesn't leave enough room for the layers panel to operate,
' this will be set to TRUE.  The user must reduce one or more other panels before the layer panel will be allowed to open.
Private m_LayerPanelMustStayHidden As Boolean

Private Sub Form_Load()
    
    'All layout decisions on this form are contingent on the number of panels, so set this first as subsequent code
    ' will likely rely on it.
    m_NumOfPanels = ttlPanel.Count
    ReDim m_Panels(0 To m_NumOfPanels - 1) As PD_Panel
    m_PanelResizeActive = -1
    
    'Some panel heights are hard-coded.  Calculate those now.
    ' (Note that we do not calculate a hard-coded size for the final panel (layers).  It is autosized to fill whatever
    '  space remains after other panels are positioned.)
    ' (Also, in a perfect world the user could resize each panel vertically.  I'm writing each sub-panel UI so that
    '  it technically supports this behavior, but there's no framework for that kind of resizing just yet.)
    Dim i As Long
    If (Not g_UserPreferences Is Nothing) Then
        For i = 0 To m_NumOfPanels - 1
            m_Panels(i).InitialHeight = g_UserPreferences.GetPref_Long("Toolbox", "RightPanelWidth-" & CStr(i + 1), Interface.FixDPI(100))
        Next i
    Else
        For i = 0 To m_NumOfPanels - 1
            m_Panels(i).InitialHeight = Interface.FixDPI(100)
        Next i
    End If
    
    'Synchronize all panel heights
    For i = 0 To m_NumOfPanels - 1
        m_Panels(i).CurrentHeight = m_Panels(i).InitialHeight
    Next i
    
    'Prep a mouse handler for the underlying form
    Set m_MouseEvents = New pdInputMouse
    m_MouseEvents.AddInputTracker Me.hWnd, , True
    
    'Prep a window synchronizer and add each subpanel to it
    Set m_WindowSync = New pdWindowSync
    
    'It can take quite some time to load these panels, so during debugging, it's helpful to track
    ' any unintentional changes to load time (which in turn harm PD's average startup time).
    #If DEBUGMODE = 1 Then
        Dim startTime As Currency
        VBHacks.GetHighResTime startTime
    #End If
    
    Load layerpanel_Navigator
    m_WindowSync.SynchronizeWindows ctlContainer(0).hWnd, layerpanel_Navigator.hWnd
    layerpanel_Navigator.Show
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogTiming "right toolbox / navigator panel", VBHacks.GetTimerDifferenceNow(startTime)
        VBHacks.GetHighResTime startTime
    #End If
    
    Load layerpanel_Colors
    m_WindowSync.SynchronizeWindows ctlContainer(1).hWnd, layerpanel_Colors.hWnd
    layerpanel_Colors.Show
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogTiming "right toolbox / color panel", VBHacks.GetTimerDifferenceNow(startTime)
        VBHacks.GetHighResTime startTime
    #End If
    
    Load layerpanel_Layers
    m_WindowSync.SynchronizeWindows ctlContainer(ctlContainer.UBound).hWnd, layerpanel_Layers.hWnd
    layerpanel_Layers.Show
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogTiming "right toolbox / layers panel", VBHacks.GetTimerDifferenceNow(startTime)
        VBHacks.GetHighResTime startTime
    #End If
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
        
    'Theme everything
    Me.UpdateAgainstCurrentTheme True
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogTiming "right toolbox / everything else", VBHacks.GetTimerDifferenceNow(startTime)
    #End If
    
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
    If (Not g_UserPreferences Is Nothing) Then
        Dim i As Long
        For i = 0 To m_NumOfPanels - 1
            g_UserPreferences.SetPref_Long "Toolbox", "RightPanelWidth-" & CStr(i + 1), m_Panels(i).CurrentHeight
        Next i
    End If
    
End Sub

Private Sub Form_Resize()
    ReflowInterface
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
        Unload layerpanel_Navigator
        Unload layerpanel_Colors
        Unload layerpanel_Layers
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  toolbar_Layers was unloaded prematurely - why??"
        #End If
        Cancel = True
    End If
    
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
    
    'Before doing anything complicated, left-align the separator line between the canvas area and the toolbox
    lnSeparatorLeft.x1 = 0
    lnSeparatorLeft.y1 = 0
    lnSeparatorLeft.x2 = 0
    lnSeparatorLeft.y2 = formHeight
    
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
    MIN_LAYER_PANEL_SIZE = Interface.FixDPI(166)
    
    Dim i As Long, tmpHeight As Long
    For i = 0 To m_NumOfPanels - 1
        
        'Move the titlebar of this panel into position
        ttlPanel(i).SetPositionAndSize xOffset, yOffset, xWidth - xOffset + Interface.FixDPI(2), ttlPanel(i).GetHeight
        
        'Move the yOffset beneath the panel
        yOffset = yOffset + ttlPanel(i).GetHeight + Interface.FixDPI(1)
        
        'If the title bar state is TRUE, open its corresponding panel.
        If ttlPanel(i).Value Then
            
            'Move this panel into position.
            If (xWidth - xOffset > 0) Then
                
                'All panels follow an identical pattern, *except* for the layers panel (which auto-fills any remaining
                ' vertical space - see below)
                If (i < (m_NumOfPanels - 1)) Then
                
                    'Because the user has control over panel height, we need to perform some checks to ensure the target
                    ' panel's height is an acceptable value
                    tmpHeight = m_Panels(i).CurrentHeight
                    If (tmpHeight < MIN_PANEL_SIZE) Then tmpHeight = MIN_PANEL_SIZE
                    If (tmpHeight > MAX_PANEL_SIZE) Then tmpHeight = MAX_PANEL_SIZE
                    ctlContainer(i).SetPositionAndSize Int(CSng(xOffset) * 1.5 + 0.5), yOffset, xWidth - xOffset, tmpHeight
                    
                'The layers panel is unique, because it shrinks to fit all available vertical space.
                Else
                    
                    'Calculate an "ideal" height; if this isn't available (because previous panels are too tall),
                    ' close the panel entirely.
                    tmpHeight = (formHeight - yOffset)
                    
                    'There's not enough room to operate the layer panel; force it to hide until the user frees up space
                    If (tmpHeight < MIN_LAYER_PANEL_SIZE) Then
                        m_LayerPanelMustStayHidden = True
                        ttlPanel(i).Value = False
                        ctlContainer(i).Visible = False
                    Else
                        m_LayerPanelMustStayHidden = False
                        ctlContainer(i).SetPositionAndSize Int(CSng(xOffset) * 1.5 + 0.5), yOffset, xWidth - xOffset, tmpHeight
                    End If
                    
                End If
                
            End If
            
            'Show the panel, and add its height to the running offset calculation
            ' (Also, it looks weird, but we need to re-check that the title bar is still set to TRUE here;
            '  previous steps may have deactivated it due to size and/or layout constraints.)
            If ttlPanel(i).Value Then ctlContainer(i).Visible = True
            yOffset = yOffset + ctlContainer(i).GetHeight
            
        'If the title bar state is FALSE, close its corresponding panel.
        Else
            ctlContainer(i).Visible = False
        End If
        
        'We now want to determine the offset *between* panels.  This step is important as it determines the interactive
        ' region between panels where the user can click-drag to resize individual panels.
        With m_Panels(i).VerticalResizeRect
            .Left = xOffset
            .Width = ctlContainer(i).GetWidth
            .Top = yOffset
            .Height = Interface.FixDPI(RESIZE_BORDER)
        End With
        
        'Calculate the new top position of the next panel in line.
        yOffset = yOffset + m_Panels(i).VerticalResizeRect.Height
        
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
    
    'The left separator line is colored according to the current shadow accent color
    If Not (g_Themer Is Nothing) Then
        lnSeparatorLeft.borderColor = g_Themer.GetGenericUIColor(UI_GrayDark)
    Else
        lnSeparatorLeft.borderColor = vbHighlight
    End If
    
    'Pass the theme update request to any active child forms.
    ' (Note that we don't have to do this on our initial load, because the panels will automatically
    ' theme themselves.)
    If ((Not layerpanel_Navigator Is Nothing) And (Not isFirstLoad)) Then layerpanel_Navigator.UpdateAgainstCurrentTheme
    If ((Not layerpanel_Colors Is Nothing) And (Not isFirstLoad)) Then layerpanel_Colors.UpdateAgainstCurrentTheme
    If ((Not layerpanel_Layers Is Nothing) And (Not isFirstLoad)) Then layerpanel_Layers.UpdateAgainstCurrentTheme
    
    'Reflow the interface, to account for any language changes.  (This will also trigger a redraw of the layer list box.)
    ReflowInterface
    
End Sub

Private Sub m_MouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

    'If we're not currently in the midst of a panel resize, see if we need to initiate one
    If (x > Interface.FixDPI(RESIZE_BORDER)) And (m_PanelResizeActive < 0) And ((Button And pdLeftButton) <> 0) Then
        m_PanelResizeActive = GetResizeRectUnderMouse(x, y)
        m_PanelResizeInitY = y
        
        'To ensure that other PD hWnds don't steal the mouse from us, lock it to this window until the drag is complete
        m_MouseEvents.SetCaptureOverride False
        SetCapture Me.hWnd
        
    End If
    
End Sub

Private Sub m_MouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'If the mouse is near the resizable edge of the toolbar (the left edge, currently), allow the user to resize
    ' the layer toolbox.
    Dim mouseInResizeTerritory As Boolean
    Dim hitCode As Long
    
    'Check the mouse position to see if it's in resize territory (along the left edge of the toolbox)
    If (y > 0) And (y < Me.ScaleHeight) And (x < Interface.FixDPI(RESIZE_BORDER)) Then
        mouseInResizeTerritory = True
        hitCode = HTLEFT
    Else
        mouseInResizeTerritory = False
    End If
    
    'If the left mouse button is down, and the mouse is in resize territory, initiate an API resize event
    If (mouseInResizeTerritory And (m_PanelResizeActive < 0)) Then
    
        'Change the cursor to a resize cursor
        m_MouseEvents.SetSystemCursor IDC_SIZEWE
        
        If (Button = vbLeftButton) Then
        
            m_WeAreResponsibleForResize = True
            ReleaseCapture
            SendMessage Me.hWnd, WM_NCLBUTTONDOWN, hitCode, ByVal 0&
            
            'After the toolbox has been resized, we need to manually notify the toolbox manager, so it can
            ' notify any neighboring toolboxes (and/or the central canvas)
            Toolboxes.SetConstrainingSize PDT_RightToolbox, Me.ScaleWidth
            FormMain.UpdateMainLayout
            
            'A premature exit is required, because the end of this sub contains code to detect the release of the
            ' mouse after a drag event.  Because the event is not being initiated normally, we can't detect a standard
            ' MouseUp event, so instead, we mimic it by checking MouseMove and m_WeAreResponsibleForResize = TRUE.
            Exit Sub
            
        End If
    
    End If
    
    'Next, see if the mouse is in the interactive area "between" panels.
    If (m_PanelResizeActive < 0) Then
        
        Dim tmpIndex As Long
        tmpIndex = GetResizeRectUnderMouse(x, y)
        
        'If the mouse is inside an interactive area, and the left mouse button *isn't* down,
        ' change the cursor to reflect that the user can resize via this position.
        If (tmpIndex >= 0) Then
            m_MouseEvents.SetSystemCursor IDC_SIZENS
        Else
            If (Not mouseInResizeTerritory) Then m_MouseEvents.SetSystemCursor IDC_DEFAULT
        End If
    
    'We are already in the midst of a resize.  Calculate a new height and immediately reflow the interface to match.
    Else
        m_Panels(m_PanelResizeActive).CurrentHeight = m_Panels(m_PanelResizeActive).InitialHeight + (y - m_PanelResizeInitY)
        ReflowInterface
    End If
    
    'Check for mouse release; we will only reach this point if the mouse is *not* in resize territory, which in turn
    ' means we can free the release code and resize the window now.  (On some OS/theme combinations, the canvas will
    ' live-resize as the mouse is moved.  On others, the canvas won't redraw until the mouse is released.)
    If m_WeAreResponsibleForResize Then
        m_WeAreResponsibleForResize = False
        m_MouseEvents.SetSystemCursor IDC_DEFAULT
    End If
    
End Sub

Private Sub m_MouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    
    'If the user just finished a vertical panel resize, update our stored height value for the resized panel
    If (m_PanelResizeActive >= 0) And ((Button And pdLeftButton) <> 0) Then
        
        m_Panels(m_PanelResizeActive).InitialHeight = m_Panels(m_PanelResizeActive).CurrentHeight
        m_PanelResizeActive = -1
        
        'We also need to reset the mouse handler to its default behavior
        ReleaseCapture
        m_MouseEvents.SetCaptureOverride True
        
    End If
    
    'See where the mouse is now and set the cursor accordingly
    Dim tmpIndex As Long
    tmpIndex = GetResizeRectUnderMouse(x, y)
    
    'If the mouse is inside an interactive area, and the left mouse button *isn't* down,
    ' change the cursor to reflect that the user can resize via this position.
    If (tmpIndex >= 0) Then
        m_MouseEvents.SetSystemCursor IDC_SIZENS
    Else
        m_MouseEvents.SetSystemCursor IDC_DEFAULT
    End If
    
End Sub

Private Function GetResizeRectUnderMouse(ByVal x As Single, ByVal y As Single) As Long
    GetResizeRectUnderMouse = -1
    Dim i As Long
    For i = 0 To m_NumOfPanels - 1
        If PDMath.IsPointInRectF(x, y, m_Panels(i).VerticalResizeRect) Then
            GetResizeRectUnderMouse = i
            Exit For
        End If
    Next i
End Function

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    
    'If a panel is opening, redraw any elements that have may been suppressed while the panel was invisible
    If newState Then NotifyLayerChange
    
    'Reflow the interface to account for the changed size
    ReflowInterface
    
End Sub

'When one or more layers are modified (via painting, effects, whatever), PD's various interface control functions
' will notify this toolbar via this function.  The toolbar will then redraw individual panels as necessary.
'
'Note that a layerID of -1 means multiple/all layers have changed, while a value >= 0 tells you which layer changed,
' perhaps sparing the amount of redraw work required.
Public Sub NotifyLayerChange(Optional ByVal layerID As Long = -1)
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    If ttlPanel(2).Value And (Not m_LayerPanelMustStayHidden) Then layerpanel_Layers.ForceRedraw True, layerID
    If ttlPanel(0).Value Then layerpanel_Navigator.nvgMain.NotifyNewThumbNeeded
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "toolbar_Layers.NotifyLayerChange finished in " & VBHacks.GetTimeDiffNowAsString(startTime)
    #End If
End Sub

'If the current viewport position and/or size changes, this toolbar will be notified.  At present, the only subpanel
' affected by viewport changes is the navigator panel.
Public Sub NotifyViewportChange()
    If ttlPanel(0).Value Then layerpanel_Navigator.nvgMain.NotifyNewViewportPosition
End Sub
