VERSION 5.00
Begin VB.Form toolbar_Layers 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Layers"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3735
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   2
      Left            =   240
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   5
      Top             =   3960
      Width           =   1815
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   1
      Left            =   360
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   476
      Caption         =   "overview"
      Value           =   0   'False
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   480
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   3
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
      TabIndex        =   4
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   476
      Caption         =   "color selector"
      Value           =   0   'False
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
'Copyright 2014-2015 by Tanner Helland
'Created: 25/March/14
'Last updated: 30/September/15
'Last update: implement collapsible panels
'
'For historical reasons, I call this the "layers" toolbar, but it actually encompasses everything that appears on
' the right-side toolbar.  Most of the code in this window is dedicated to supporting collapsible panels, and all
' the messy UI that goes along with that.
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

'How close does the mouse have to be to the form border to allow resizing? Currently we use 7 pixels, while accounting
' for DPI variance (e.g. 7 pixels at 96 dpi)
Private Const RESIZE_BORDER As Long = 7

'A dedicated mouse handler helps provide cursor handling
Private WithEvents m_MouseEvents As pdInputMouse
Attribute m_MouseEvents.VB_VarHelpID = -1

'Default heights of each panel.  These are (currently) hard-coded for all panels except the layers panel; it is dynamically
' sized to fit whatever remaining space we have in the panel as a whole.
Private m_defaultPanelHeight() As Long

'Number of panels; set automatically at Form_Load
Private m_numOfPanels As Long

Private Sub Form_Load()
    
    'All layout decisions on this form are contingent on the number of panels, so set this first as subsequent code
    ' will likely rely on it.
    m_numOfPanels = ttlPanel.Count
    ReDim m_defaultPanelHeight(0 To m_numOfPanels - 1) As Long
    
    'Some panel heights are hard-coded.  Calculate those now.
    ' (Note that we do not calculate a hard-coded size for the final panel (layers).  It is autosized to fill whatever
    '  space remains after other panels are positioned.)
    ' (Also, in a perfect world the user could resize each panel vertically.  I'm writing each sub-panel UI so that
    '  it technically supports this behavior, but there's no framework for that kind of resizing just yet.)
    m_defaultPanelHeight(0) = FixDPI(100)
    m_defaultPanelHeight(1) = FixDPI(100)
    
    'Prep a mouse handler
    Set m_MouseEvents = New pdInputMouse
    m_MouseEvents.addInputTracker Me.hWnd, True, True, , True, True
    
    'Prep a window synchronizer and add each subpanel to it
    Set m_WindowSync = New pdWindowSync
    
    Load layerpanel_Navigator
    m_WindowSync.SynchronizeWindows picContainer(0).hWnd, layerpanel_Navigator.hWnd
    layerpanel_Navigator.Show
    
    Load layerpanel_Colors
    m_WindowSync.SynchronizeWindows picContainer(1).hWnd, layerpanel_Colors.hWnd
    layerpanel_Colors.Show
    
    Load layerpanel_Layers
    m_WindowSync.SynchronizeWindows picContainer(picContainer.UBound).hWnd, layerpanel_Layers.hWnd
    layerpanel_Layers.Show
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.setParentForm Me
    m_lastUsedSettings.loadAllControlValues
    
    'Theme everything
    UpdateAgainstCurrentTheme
    
    'Reflow the interface to match its current size
    ReflowInterface
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    m_lastUsedSettings.saveAllControlValues
    
End Sub

Private Sub Form_Resize()
    ReflowInterface
End Sub

'Toolbars can never be unloaded, EXCEPT when the whole program is going down.  Check for the program-wide closing flag prior
' to exiting; if it is not found, cancel the unload and simply hide this form.  (Note that the toggleToolbarVisibility sub
' will also keep this toolbar's Window menu entry in sync with the form's current visibility.)
Private Sub Form_Unload(Cancel As Integer)
    
    If g_ProgramShuttingDown Then
    
        'Release the subpanel subclasser
        Set m_WindowSync = Nothing
        
        'Unload all child forms
        Unload layerpanel_Navigator
        Unload layerpanel_Colors
        Unload layerpanel_Layers
        
        'Release our custom mouse handler
        Set m_MouseEvents = Nothing
        
        'Release this window from any program-wide handlers
        ReleaseFormTheming Me
        g_WindowManager.UnregisterForm Me
        
    Else
        Cancel = True
        ToggleToolbarVisibility LAYER_TOOLBOX
    End If
    
End Sub

'Whenever the layer toolbox is resized, we must reflow all objects to fill the available space.  Note that we do not do
' specialized handling for the vertical direction; vertically, the only change we handle is resizing the layer box itself
' to fill whatever vertical space is available.
Private Sub ReflowInterface()
    
    'If the form is invisible (due to minimize or something else), just exit now
    If Not Me.Visible Then Exit Sub
    If (Me.ScaleHeight = 0) Or (Me.ScaleWidth = 0) Then Exit Sub
    
    'When the parent form is resized, resize the layer list (and other items) to properly fill the
    ' available horizontal and vertical space.
    
    'Before doing anything complicated, left-align the separator line between the canvas area and the toolbox
    lnSeparatorLeft.x1 = 0
    lnSeparatorLeft.y1 = 0
    lnSeparatorLeft.x2 = 0
    lnSeparatorLeft.y2 = Me.ScaleHeight
    
    'Next, we want to resize all subpanel picture boxes, so that their size reflects the new form size.  This is a
    ' bit complicated, as each form has a different base size, and the user can toggle panel visibility at any time.
    
    'Start by calculating initial x/y offsets
    Dim yOffset As Long, xOffset As Long, xWidth As Long
    xOffset = FixDPI(RESIZE_BORDER)
    yOffset = FixDPI(2)
    xWidth = Me.ScaleWidth - xOffset
    
    Dim i As Long
    For i = 0 To m_numOfPanels - 1
        
        'Move the titlebar of this panel into position
        ttlPanel(i).Move xOffset, yOffset, xWidth - xOffset + FixDPI(2)
        
        'Move the yOffset beneath the panel
        yOffset = yOffset + ttlPanel(i).Height + FixDPI(2)
        
        'If the title bar state is TRUE, open its corresponding panel.
        If ttlPanel(i).Value Then
            
            'Move the panel into position.  For all panels except the layers panel, height is hard-coded at design-time.
            If i < m_numOfPanels - 1 Then
                picContainer(i).Move xOffset * 2, yOffset, xWidth - xOffset, m_defaultPanelHeight(i)
                
            'The layers panel is unique, because it shrinks to fit all available space.
            Else
                picContainer(i).Move xOffset * 2, yOffset, xWidth - xOffset, Me.ScaleHeight - yOffset
            End If
            
            'Show the panel, and add its height to the running offset calculation
            picContainer(i).Visible = True
            yOffset = yOffset + picContainer(i).Height
            
        'If the title bar state is FALSE, close its corresponding panel.
        Else
            picContainer(i).Visible = False
        End If
        
        'Regardless of visibility, always add some padding to the running offset
        yOffset = yOffset + FixDPI(4)
        
    Next i
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) MakeFormPretty is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    MakeFormPretty Me
    
    'The left separator line is colored according to the current shadow accent color
    If Not (g_Themer Is Nothing) Then
        lnSeparatorLeft.BorderColor = g_Themer.getThemeColor(PDTC_GRAY_SHADOW)
    Else
        lnSeparatorLeft.BorderColor = vbHighlight
    End If
    
    'TODO: pass along the request to any active child forms.
    If Not (layerpanel_Navigator) Is Nothing Then layerpanel_Layers.UpdateAgainstCurrentTheme
    If Not (layerpanel_Colors) Is Nothing Then layerpanel_Layers.UpdateAgainstCurrentTheme
    If Not (layerpanel_Layers) Is Nothing Then layerpanel_Layers.UpdateAgainstCurrentTheme
    
    'Reflow the interface, to account for any language changes.  (This will also trigger a redraw of the layer list box.)
    ReflowInterface
    
End Sub

Private Sub m_MouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'If the mouse is near the resizable edge of the toolbar (the left edge, currently), allow the user to resize
    ' the layer toolbox.
    Dim mouseInResizeTerritory As Boolean
    Dim hitCode As Long
    
    'Check the mouse position to see if it's in resize territory (along the left edge of the toolbox)
    If (y > 0) And (y < Me.ScaleHeight) And (x < FixDPI(RESIZE_BORDER)) Then
        mouseInResizeTerritory = True
        hitCode = HTLEFT
    Else
        mouseInResizeTerritory = False
    End If
    
    'If the left mouse button is down, and the mouse is in resize territory, initiate an API resize event
    If mouseInResizeTerritory Then
    
        'Change the cursor to a resize cursor
        m_MouseEvents.setSystemCursor IDC_SIZEWE
        
        If (Button = vbLeftButton) Then
            m_WeAreResponsibleForResize = True
            ReleaseCapture
            SendMessage Me.hWnd, WM_NCLBUTTONDOWN, hitCode, ByVal 0&
            
            'A premature exit is required, because the end of this sub contains code to detect the release of the
            ' mouse after a drag event.  Because the event is not being initiated normally, we can't detect a standard
            ' MouseUp event, so instead, we mimic it by checking MouseMove and m_WeAreResponsibleForResize = TRUE.
            Exit Sub
            
        End If
        
    Else
        m_MouseEvents.setSystemCursor IDC_DEFAULT
    End If
    
    'Check for mouse release; we will only reach this point if the mouse is *not* in resize territory, which in turn
    ' means we can free the release code and resize the window now.  (On some OS/theme combinations, the canvas will
    ' live-resize as the mouse is moved.  On others, the canvas won't redraw until the mouse is released.)
    If m_WeAreResponsibleForResize Then
        
        m_WeAreResponsibleForResize = False
        m_MouseEvents.setSystemCursor IDC_DEFAULT
        
        'If theming is disabled, window performance is so poor that the window manager will automatically
        ' disable canvas updates until the mouse is released.  Request a full update now.
        If (Not g_IsThemingEnabled) Then g_WindowManager.NotifyToolboxResized Me.hWnd, True
        
    End If
    
End Sub

Private Sub m_MouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    m_MouseEvents.setSystemCursor IDC_DEFAULT
End Sub

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

    'Optimizing the layer listbox is TODO!
    If ttlPanel(2).Value Then layerpanel_Layers.forceRedraw True
    
    'Redraw the navigator to match
    If ttlPanel(0).Value Then layerpanel_Navigator.nvgMain.NotifyNewThumbNeeded

End Sub

'If the current viewport position and/or size changes, this toolbar will be notified.  At present, the only subpanel
' affected by viewport changes is the navigator panel.
Public Sub NotifyViewportChange()
    If ttlPanel(0).Value Then layerpanel_Navigator.nvgMain.NotifyNewViewportPosition
End Sub
