VERSION 5.00
Begin VB.UserControl pdCanvas 
   Appearance      =   0  'Flat
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
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   886
   ToolboxBitmap   =   "pdCanvas.ctx":0000
   Begin PhotoDemon.pdContainer pnlNoImages 
      Height          =   3375
      Left            =   6240
      Top             =   2160
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5953
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   495
         Index           =   0
         Left            =   120
         Top             =   2760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         Caption         =   "quick start"
         CustomDragDropEnabled=   -1  'True
         FontSize        =   18
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   495
         Index           =   1
         Left            =   1200
         Top             =   2760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Caption         =   "recent images"
         CustomDragDropEnabled=   -1  'True
         FontSize        =   18
      End
      Begin PhotoDemon.pdHyperlink hypRecentFiles 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Alignment       =   2
         Caption         =   "clear recent image list"
         RaiseClickEvent =   -1  'True
      End
      Begin PhotoDemon.pdCheckBox chkRecentFiles 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Caption         =   "show recent images, if any"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdButton cmdStart 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         Caption         =   "New image..."
         CustomDragDropEnabled=   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdButton cmdStart 
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         Caption         =   "Open image..."
         CustomDragDropEnabled=   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdButton cmdStart 
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         Caption         =   "Import from clipboard..."
         CustomDragDropEnabled=   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdButton cmdStart 
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         Caption         =   "Batch process..."
         CustomDragDropEnabled=   -1  'True
         FontSize        =   12
      End
      Begin PhotoDemon.pdButton cmdRecent 
         Height          =   495
         Index           =   0
         Left            =   960
         TabIndex        =   12
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         CustomDragDropEnabled=   -1  'True
      End
   End
   Begin PhotoDemon.pdProgressBar mainProgBar 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
   End
   Begin PhotoDemon.pdRuler vRuler 
      Height          =   4935
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   8705
      Orientation     =   1
   End
   Begin PhotoDemon.pdRuler hRuler 
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
   End
   Begin PhotoDemon.pdImageStrip ImageStrip 
      Height          =   990
      Left            =   6240
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   1746
      Alignment       =   0
   End
   Begin PhotoDemon.pdStatusBar StatusBar 
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   7350
      Width           =   13290
      _ExtentX        =   23442
      _ExtentY        =   609
   End
   Begin PhotoDemon.pdCanvasView CanvasView 
      Height          =   4935
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   4575
      _ExtentX        =   8281
      _ExtentY        =   8916
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
'Copyright 2002-2026 by Tanner Helland
'Created: 29/November/02
'Last updated: 28/December/24
'Last update: handle new user preference for mouse wheel zoom vs scroll
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As PointAPI) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

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

'Mouse interactions are complicated in this form, so we sometimes need to cache button values
' and process them elsewhere
Private m_LMBDown As Boolean, m_RMBDown As Boolean, m_MMBDown As Boolean

'Every time a canvas MouseMove event occurs, this number is incremented by one.  If mouse events are coming in fast and furious,
' we can delay renders between them to improve responsiveness.  (This number is reset to zero when the mouse is released.)
Private m_NumOfMouseMovements As Long

'If the mouse is currently over the canvas, this will be set to TRUE.
Private m_IsMouseOverCanvas As Boolean

'Track initial mouse button locations
Private m_InitMouseX As Double, m_InitMouseY As Double

'Last mouse x/y positions, in canvas coordinates
Private m_LastCanvasX As Double, m_LastCanvasY As Double

'Last mouse x/y positions, in image coordinates
Private m_LastImageX As Double, m_LastImageY As Double

'On the canvas's MouseDown event, this control will mark the relevant point of interest index for the active layer (if any).
' If a point of interest has not been selected, this value will be reset to poi_Undefined (-1).
Private m_CurPOI As PD_PointOfInterest

'As some POI interactions may cause the canvas to redraw, we also cache the *last* point of interest.  When this mismatches the
' current one, a UI-only viewport redraw is requested, and the last/current point values are synched.
Private m_LastPOI As PD_PointOfInterest

'To improve performance, we can ask the canvas to not refresh itself until we say so.
Private m_SuspendRedraws As Boolean

'Some tools support the ability to auto-activate a layer beneath the mouse.  If supported, during the MouseMove event,
' this value (m_LayerAutoActivateIndex) will be updated with the index of the layer that will be auto-activated if the
' user presses the mouse button.  This can be used to modify things like cursor behavior, to make sure the user receives
' accurate feedback on what a given action will affect.
Private m_LayerAutoActivateIndex As Long

'Selection tools need to know if a selection was active before mouse events start.
' If it was, creation of an invalid new selection will add "Remove Selection" to the
' Undo/Redo chain; however, if no selection was active, the working selection will
' simply be erased.
Private m_SelectionActiveBeforeMouseEvents As Boolean

'If the user attempts to initiate a paint operation on an invisible or locked layer,
' we will ignore subsequent mouse events until the mouse button is released.
Private m_IgnoreMouseActions As Boolean

'When we reflow the interface, we mark a special "resize" state to prevent recursive automatic reflow event notifications
Private m_InternalResize As Boolean

'Are various canvas elements currently visible?  Outside callers can access/modify these via dedicated Get/Set functions.
Private m_RulersVisible As Boolean, m_StatusBarVisible As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDCANVAS_COLOR_LIST
    [_First] = 0
    PDC_StatusBar = 0
    PDC_SpecialButtonBackground = 1
    [_Last] = 1
    [_Count] = 2
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'Popup menu for the image strip
Private WithEvents m_PopupMenu As pdPopupMenu
Attribute m_PopupMenu.VB_VarHelpID = -1

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_Canvas
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function

'Helper functions to ensure ideal UI behavior
Public Function IsScreenCoordInsideCanvasView(ByVal srcX As Long, ByVal srcY As Long) As Boolean

    'Get the canvas view's window
    Dim tmpRect As RECT
    If (Not g_WindowManager Is Nothing) Then
        g_WindowManager.GetWindowRect_API_Universal CanvasView.hWnd, VarPtr(tmpRect)
        IsScreenCoordInsideCanvasView = PDMath.IsPointInRect(srcX, srcY, tmpRect)
    Else
        IsScreenCoordInsideCanvasView = False
    End If
    
End Function

Public Function GetCanvasViewHWnd() As Long
    GetCanvasViewHWnd = CanvasView.hWnd
End Function

Public Sub ManuallyNotifyCanvasMouse(ByVal mouseX As Long, ByVal mouseY As Long)
    CanvasView.NotifyExternalMouseMove mouseX, mouseY
End Sub

'External functions can call this to set the current network state (which in turn, draws a relevant icon to the status bar)
Public Sub SetNetworkState(ByVal newNetworkState As Boolean)
    StatusBar.SetNetworkState newNetworkState
End Sub

'External functions can call this to set the current selection state
' (which updates the status bar with a little selection size notification).
Public Sub SetSelectionState(ByVal newSelectionState As Boolean)
    StatusBar.SetSelectionState newSelectionState
End Sub

'Use these functions to forcibly prevent the canvas from redrawing itself.
' REDRAWS WILL NOT HAPPEN AGAIN UNTIL YOU RESTORE ACCESS!
'
'(Also note that this function relays state changes to the underlying pdCanvasView object; as such, do not set
' m_SuspendRedraws manually - only set it via this function, to ensure the canvas and underlying canvas view
' remain in sync.)
Public Function GetRedrawSuspension() As Boolean
    GetRedrawSuspension = m_SuspendRedraws Or CanvasView.GetRedrawSuspension()
End Function

Public Sub SetRedrawSuspension(ByVal newRedrawValue As Boolean)
    CanvasView.SetRedrawSuspension newRedrawValue
    If m_RulersVisible Then
        hRuler.SetRedrawSuspension newRedrawValue
        vRuler.SetRedrawSuspension newRedrawValue
    End If
    m_SuspendRedraws = newRedrawValue
End Sub

'Need to wipe the canvas?  Call this function, but please be careful - it will literally erase the canvas's back buffer.
Public Sub ClearCanvas()
    
    CanvasView.ClearCanvas
    StatusBar.ClearCanvas
    
    'If any valid images are loaded, scroll bars are always made visible
    SetScrollVisibility pdo_Horizontal, PDImages.IsImageActive()
    SetScrollVisibility pdo_Vertical, PDImages.IsImageActive()
    
    'With appropriate elements shown/hidden, we can now align everything
    Me.AlignCanvasView
    
End Sub

'Get/Set scroll bar value
Public Function GetScrollValue(ByVal barType As PD_Orientation) As Long
    If (barType = pdo_Horizontal) Then GetScrollValue = hScroll.Value Else GetScrollValue = vScroll.Value
End Function

Public Sub SetScrollValue(ByVal barType As PD_Orientation, ByVal newValue As Long)
    
    If (barType = pdo_Horizontal) Then
        hScroll.Value = newValue
    ElseIf (barType = pdo_Vertical) Then
        vScroll.Value = newValue
    Else
        hScroll.Value = newValue
        vScroll.Value = newValue
    End If
    
    'If automatic redraws are suspended, the scroll bars change events won't fire, so we must manually notify external UI elements
    If Me.GetRedrawSuspension Then Viewport.NotifyEveryoneOfViewportChanges
    
End Sub

'Get/Set scroll max/min
Public Function GetScrollMax(ByVal barType As PD_Orientation) As Long
    If (barType = pdo_Horizontal) Then GetScrollMax = hScroll.Max Else GetScrollMax = vScroll.Max
End Function

Public Function GetScrollMin(ByVal barType As PD_Orientation) As Long
    If (barType = pdo_Horizontal) Then GetScrollMin = hScroll.Min Else GetScrollMin = vScroll.Min
End Function

Public Sub SetScrollMax(ByVal barType As PD_Orientation, ByVal newMax As Long)
    If (barType = pdo_Horizontal) Then hScroll.Max = newMax Else vScroll.Max = newMax
End Sub

Public Sub SetScrollMin(ByVal barType As PD_Orientation, ByVal newMin As Long)
    If (barType = pdo_Horizontal) Then hScroll.Min = newMin Else vScroll.Min = newMin
End Sub

'Set scroll bar LargeChange value
Public Sub SetScrollLargeChange(ByVal barType As PD_Orientation, ByVal newLargeChange As Long)
    If (barType = pdo_Horizontal) Then hScroll.LargeChange = newLargeChange Else vScroll.LargeChange = newLargeChange
End Sub

'Get/Set scrollbar visibility.  Note that visibility is only toggled as necessary, so this function is preferable to
' calling .Visible properties directly.
Public Function GetScrollVisibility(ByVal barType As PD_Orientation) As Boolean
    If (barType = pdo_Horizontal) Then GetScrollVisibility = hScroll.Visible Else GetScrollVisibility = vScroll.Visible
End Function

Public Sub SetScrollVisibility(ByVal barType As PD_Orientation, ByVal newVisibility As Boolean)
    
    'If the scroll bar status wasn't actually changed, we can avoid a forced screen refresh
    Dim changesMade As Boolean
    changesMade = False
    
    If (barType = pdo_Horizontal) Then
        If (newVisibility <> hScroll.Visible) Then
            hScroll.Visible = newVisibility
            changesMade = True
        End If
    
    ElseIf (barType = pdo_Vertical) Then
        If (newVisibility <> vScroll.Visible) Then
            vScroll.Visible = newVisibility
            changesMade = True
        End If
    
    Else
        If (newVisibility <> hScroll.Visible) Or (newVisibility <> vScroll.Visible) Then
            hScroll.Visible = newVisibility
            vScroll.Visible = newVisibility
            changesMade = True
        End If

    End If
    
    'When scroll bar visibility is changed, we must move the main canvas picture box to match
    If changesMade Then
    
        'The "center" button between the scroll bars has the same visibility as the scrollbars;
        ' it's only visible if *both* bars are visible.
        cmdCenter.Visible = (hScroll.Visible And vScroll.Visible)
        Me.AlignCanvasView
        
    End If
    
End Sub

Public Sub DisplayImageSize(ByRef srcImage As pdImage, Optional ByVal clearSize As Boolean = False)
    StatusBar.DisplayImageSize srcImage, clearSize
End Sub

Public Sub DisplayCanvasMessage(ByRef cMessage As String)
    StatusBar.DisplayCanvasMessage cMessage
End Sub

Public Sub DisplayCanvasCoordinates(ByVal xCoord As Double, ByVal yCoord As Double, Optional ByVal clearCoords As Boolean = False)
    StatusBar.DisplayCanvasCoordinates xCoord, yCoord, clearCoords
    If m_RulersVisible Then
        hRuler.NotifyMouseCoords m_LastCanvasX, m_LastCanvasY, xCoord, yCoord, clearCoords
        vRuler.NotifyMouseCoords m_LastCanvasX, m_LastCanvasY, xCoord, yCoord, clearCoords
    End If
End Sub

Public Sub RequestRulerUpdate()
    If m_RulersVisible Then
        hRuler.NotifyViewportChange
        vRuler.NotifyViewportChange
    End If
End Sub

Public Sub RequestViewportRedraw(Optional ByVal refreshImmediately As Boolean = False)
    CanvasView.RequestRedraw refreshImmediately
End Sub

'Tabstrip relays include the next five functions
Public Sub NotifyTabstripAddNewThumb(ByVal pdImageIndex As Long)
    ImageStrip.AddNewThumb pdImageIndex
End Sub

Public Sub NotifyTabstripNewActiveImage(ByVal pdImageIndex As Long)
    ImageStrip.NotifyNewActiveImage pdImageIndex
End Sub

Public Sub NotifyTabstripUpdatedImage(ByVal pdImageIndex As Long)
    ImageStrip.NotifyUpdatedImage pdImageIndex
End Sub

Public Sub NotifyTabstripRemoveThumb(ByVal pdImageIndex As Long, Optional ByVal refreshStrip As Boolean = True)
    ImageStrip.RemoveThumb pdImageIndex, refreshStrip
End Sub

Public Sub NotifyTabstripTotalRedrawRequired(Optional ByVal regenerateThumbsToo As Boolean = False)
    ImageStrip.RequestTotalRedraw regenerateThumbsToo
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

Public Function ProgBar_GetVisibility() As Boolean
    ProgBar_GetVisibility = mainProgBar.Visible
End Function

Public Sub ProgBar_SetVisibility(ByVal newVisibility As Boolean)
    mainProgBar.Visible = newVisibility
End Sub

Public Function ProgBar_GetMax() As Double
    ProgBar_GetMax = mainProgBar.Max
End Function

Public Sub ProgBar_SetMax(ByVal newMax As Double)
    mainProgBar.Max = newMax
End Sub

Public Function ProgBar_GetValue() As Double
    ProgBar_GetValue = mainProgBar.Value
End Function

Public Sub ProgBar_SetValue(ByVal newValue As Double)
    mainProgBar.Value = newValue
End Sub

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

'Note that this control does *not* return its own DC.  Instead, it returns the DC of the underlying CanvasView object.
' This is by design.
Public Property Get hDC() As Long
    hDC = CanvasView.hDC
End Property

'To support high-DPI settings properly, we expose specialized move+size functions
Public Function GetLeft() As Long
    GetLeft = ucSupport.GetControlLeft
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    ucSupport.RequestNewPosition newLeft, , True
End Sub

Public Function GetTop() As Long
    GetTop = ucSupport.GetControlTop
End Function

Public Sub SetTop(ByVal newTop As Long)
    ucSupport.RequestNewPosition , newTop, True
End Sub

Public Function GetWidth() As Long
    GetWidth = ucSupport.GetControlWidth
End Function

Public Sub SetWidth(ByVal newWidth As Long)
    ucSupport.RequestNewSize newWidth, , True
End Sub

Public Function GetHeight() As Long
    GetHeight = ucSupport.GetControlHeight
End Function

Public Sub SetHeight(ByVal newHeight As Long)
    ucSupport.RequestNewSize , newHeight, True
End Sub

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

'Most tools don't respond to double-clicks, but polygon selections follows Photoshop convention
' (also GIMP, Krita) and close the current polygon selection, if any, on a double-click event.
Private Sub CanvasView_DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'These variables will hold the corresponding (x,y) coordinates on the IMAGE - not the VIEWPORT.
    ' (This is important if the user has zoomed into an image, and used scrollbars to look at a different part of it.)
    Dim imgX As Double, imgY As Double
    DisplayImageCoordinates x, y, PDImages.GetActiveImage(), Me, imgX, imgY
    
    If (g_CurrentTool = SELECT_POLYGON) Then
        SelectionUI.NotifySelectionMouseDblClick Me, imgX, imgY
    ElseIf (g_CurrentTool = ND_CROP) Then
        Tools_Crop.NotifyDoubleClick Button, Shift, x, y
    End If
    
End Sub

Private Sub CanvasView_LostFocusAPI()
    m_LMBDown = False
    m_RMBDown = False
    m_MMBDown = False
End Sub

Private Sub chkRecentFiles_Click()
    If chkRecentFiles.Visible Then
        UserPrefs.SetPref_Boolean "Interface", "WelcomeScreenRecentFiles", chkRecentFiles.Value
        LayoutNoImages True
    End If
End Sub

Private Sub chkRecentFiles_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    If shiftTabWasPressed Then
        newTargetHwnd = cmdStart(cmdStart.UBound).hWnd
    Else
        If (chkRecentFiles.Value And cmdRecent(cmdRecent.lBound).Visible) Then
            newTargetHwnd = cmdRecent(cmdRecent.lBound).hWnd
        Else
            newTargetHwnd = cmdStart(cmdStart.lBound).hWnd
        End If
    End If
    
End Sub

Private Sub cmdRecent_Click(Index As Integer)
    If (Not g_RecentFiles Is Nothing) Then
        If (LenB(g_RecentFiles.GetFullPath(Index)) <> 0) Then Loading.LoadFileAsNewImage g_RecentFiles.GetFullPath(Index)
    End If
End Sub

Private Sub cmdRecent_CustomDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Loading.LoadFromDragDrop Data, Effect, Button, Shift, x, y
End Sub

Private Sub cmdRecent_CustomDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Loading.HelperForDragOver Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub cmdRecent_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    If shiftTabWasPressed Then
        
        If (Index > 0) Then
            newTargetHwnd = cmdRecent(Index - 1).hWnd
        Else
            newTargetHwnd = chkRecentFiles.hWnd
        End If
        
    Else
        If (Index < cmdRecent.UBound) Then
            If cmdRecent(Index + 1).Visible Then
                newTargetHwnd = cmdRecent(Index + 1).hWnd
            Else
                newTargetHwnd = hypRecentFiles.hWnd
            End If
        Else
            newTargetHwnd = hypRecentFiles.hWnd
        End If
    End If
    
End Sub

Private Sub cmdStart_Click(Index As Integer)

    'Some indices are hard-coded, others are contingent on current user settings (like recent files)
    If (Index = 0) Then
        Actions.LaunchAction_ByName "file_new"
    ElseIf (Index = 1) Then
        Actions.LaunchAction_ByName "file_open"
    ElseIf (Index = 2) Then
        Actions.LaunchAction_ByName "edit_pasteasimage"
    ElseIf (Index = 3) Then
        Actions.LaunchAction_ByName "file_batch_process"
    End If

End Sub

Private Sub cmdStart_CustomDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Loading.LoadFromDragDrop Data, Effect, Button, Shift
End Sub

Private Sub cmdStart_CustomDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Loading.HelperForDragOver Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub cmdStart_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    If shiftTabWasPressed Then
        If (Index > cmdStart.lBound) Then
            newTargetHwnd = cmdStart(Index - 1).hWnd
        Else
            If hypRecentFiles.Visible Then
                newTargetHwnd = hypRecentFiles.hWnd
            Else
                newTargetHwnd = chkRecentFiles.hWnd
            End If
        End If
    Else
        If (Index < cmdStart.UBound) Then
            newTargetHwnd = cmdStart(Index + 1).hWnd
        Else
            newTargetHwnd = chkRecentFiles.hWnd
        End If
    End If
    
End Sub

Private Sub hScroll_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Allow the ESC key to restore focus to the canvas itself
    If (Not markEventHandled) And (whichSysKey = pdnk_Escape) Then FormMain.MainCanvas(0).SetFocusToCanvasView
    
End Sub

Private Sub hypRecentFiles_Click()
    If (Not g_RecentFiles Is Nothing) Then g_RecentFiles.ClearList
    LayoutNoImages
End Sub

Private Sub hypRecentFiles_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        If (cmdRecent.UBound >= 0) Then
            
            'Find the highest-index visible recent file button
            Dim i As Long
            For i = cmdRecent.UBound To cmdRecent.lBound Step -1
                If cmdRecent(i).Visible Then
                    newTargetHwnd = cmdRecent(i).hWnd
                    Exit Sub
                End If
            Next i
            
            'If we failed to find a visible recent file button, set focus to the left column
            newTargetHwnd = chkRecentFiles.hWnd
            
        Else
            newTargetHwnd = chkRecentFiles.hWnd
        End If
    Else
        newTargetHwnd = cmdStart(cmdStart.lBound).hWnd
    End If
End Sub

Private Sub lblTitle_CustomDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Loading.LoadFromDragDrop Data, Effect, Button, Shift
End Sub

Private Sub lblTitle_CustomDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Loading.HelperForDragOver Data, Effect, Button, Shift, x, y, State
End Sub

Private Sub m_PopupMenu_MenuClicked(ByRef clickedMenuID As String, ByVal idxMenuTop As Long, ByVal idxMenuSub As Long)

    Select Case idxMenuTop
        
        'Save
        Case 0
            Actions.LaunchAction_ByName "file_save"
            
        'Save copy (lossless)
        Case 1
            Actions.LaunchAction_ByName "file_savecopy"
        
        'Save as
        Case 2
            Actions.LaunchAction_ByName "file_saveas"
        
        'Revert
        Case 3
            Actions.LaunchAction_ByName "file_revert"
        
        '(separator)
        Case 4
        
        'Show in file manager
        Case 5
            Actions.LaunchAction_ByName "image_showinexplorer"
            
        '(separator)
        Case 6
        
        'Close
        Case 7
            Actions.LaunchAction_ByName "file_close"
        
        'Close all but this
        Case 8
            
            Dim curImageID As Long
            curImageID = PDImages.GetActiveImage.imageID
            
            Dim listOfOpenImageIDs As pdStack
            If PDImages.GetListOfActiveImageIDs(listOfOpenImageIDs) Then
                
                Dim tmpImageID As Long
                Do While listOfOpenImageIDs.PopInt(tmpImageID)
                    If PDImages.IsImageActive(tmpImageID) Then
                        If (PDImages.GetImageByID(tmpImageID).imageID <> curImageID) Then CanvasManager.FullPDImageUnload tmpImageID
                    End If
                Loop
                
            End If
    
    End Select

End Sub

Private Sub pnlNoImages_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If (Not g_AllowDragAndDrop) Then Exit Sub
    
    'Use the external function (in the clipboard handler, as the code is roughly identical to
    ' clipboard pasting) to load the OLE source.
    g_Clipboard.LoadImageFromDragDrop Data, Effect, False
    
End Sub

Private Sub pnlNoImages_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    'PD supports a lot of potential drop sources these days.  These values are defined and addressed by the main
    ' clipboard handler, as Drag/Drop and clipboard actions share a ton of similar code.
    If g_Clipboard.IsObjectDragDroppable(Data) And g_AllowDragAndDrop Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
    
End Sub

'When the control receives focus, if the focus isn't received via mouse click, display a focus rect around the active button
Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

'When the control loses focus, erase any focus rects it may have active
Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

'Get/set zoom-related UI elements
Public Function IsZoomEnabled() As Boolean
    IsZoomEnabled = StatusBar.IsZoomEnabled
End Function

Public Sub SetZoomDropDownIndex(ByVal newIndex As Long)
    StatusBar.SetZoomDropDownIndex newIndex
End Sub

Public Function GetZoomDropDownIndex() As Long
    GetZoomDropDownIndex = StatusBar.GetZoomDropDownIndex
End Function

'Only use this function for initially populating the zoom drop-down
Public Function GetZoomDropDownReference() As pdDropDown
    Set GetZoomDropDownReference = StatusBar.GetZoomDropDownReference
End Function

'Various input events are bubbled up from the underlying CanvasView control.  It provides no handling over paint and
' tool events, so we must reroute those events here.

'At present, the only App Commands the canvas handles are forward/back, which link to Undo/Redo
Private Sub CanvasView_AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If Me.IsCanvasInteractionAllowed() Then
    
        Select Case cmdID
        
            'Back button: currently triggers Undo
            Case AC_BROWSER_BACKWARD, AC_UNDO
                If PDImages.GetActiveImage.UndoManager.GetUndoState Then Process "Undo", , , UNDO_Nothing
                
            'Forward button: currently triggers Redo
            Case AC_BROWSER_FORWARD, AC_REDO
                If PDImages.GetActiveImage.UndoManager.GetRedoState Then Process "Redo", , , UNDO_Nothing
                
        End Select

    End If

End Sub

Private Sub CanvasView_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)

    markEventHandled = False
    
    'Make sure canvas interactions are allowed (e.g. an image has been loaded, etc)
    If Me.IsCanvasInteractionAllowed() Then
        
        'If certain tools are currently active, and the user presses the ALT key (but *no other*
        ' key modifiers), we will silently switch to the color picker.
        Dim tmpIsSwitchableTool As Boolean
        tmpIsSwitchableTool = tmpIsSwitchableTool Or (g_CurrentTool = PAINT_PENCIL) Or (g_CurrentTool = PAINT_SOFTBRUSH) Or (g_CurrentTool = PAINT_ERASER) Or (g_CurrentTool = PAINT_FILL)
        If (tmpIsSwitchableTool And (Not Tools.GetToolAltState()) And (vkCode = VK_ALT) And (Not ucSupport.IsKeyDown(VK_CONTROL)) And (Not ucSupport.IsKeyDown(VK_SHIFT))) Then
            
            'Addendum 8.4: before activating, notify the color tool of the current mouse position;
            ' this lets it prep some internal values, so we don't get flicker when we active it
            If m_IsMouseOverCanvas Then Tools_ColorPicker.NotifyMouseXY False, m_LastImageX, m_LastImageY, FormMain.MainCanvas(0)
            
            Tools.SetToolAltState True
            toolbar_Toolbox.SelectNewTool COLOR_PICKER
            
            'Now, something weird: we need to eat this keypress so that Windows doesn't
            ' steal focus and give it to the menu bar.  When we do this, however, PD's central
            ' hotkey handler won't be able to update its tracker for Alt key state (which is
            ' important for handling rapid keypresses, hence why we can't use async key
            ' state detection) - so we must manually notify the hotkey tracker of the state change.
            markEventHandled = True
            FormMain.HotkeyManager.NotifyAltKeystateChange True
            
        End If
        
        'Any further processing depends on which tool is currently active
        Select Case g_CurrentTool
            
            'Pan/zoom
            Case NAV_DRAG
                Tools_Move.NotifyKeyDown_HandTool Shift, vkCode, markEventHandled
                
            'Move stuff around
            Case NAV_MOVE
                Tools_Move.NotifyKeyDown Shift, vkCode, markEventHandled
            
            'Crop tool handles some buttons
            Case ND_CROP
                Tools_Crop.NotifyKeyDown Shift, vkCode, markEventHandled
                
            'Selection tools use a universal handler
            Case SELECT_RECT, SELECT_CIRC, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
                SelectionUI.NotifySelectionKeyDown Me, Shift, vkCode, markEventHandled
                
            'Pencil and paint tools redraw cursors under certain conditions
            Case PAINT_PENCIL, PAINT_SOFTBRUSH, PAINT_ERASER, PAINT_CLONE, PAINT_GRADIENT
            
                'First, notify the correct module
                If (g_CurrentTool = PAINT_PENCIL) Then
                    Tools_Pencil.NotifyBrushXY m_LMBDown, Shift, m_LastImageX, m_LastImageY, 0&, Me
                ElseIf (g_CurrentTool = PAINT_SOFTBRUSH) Or (g_CurrentTool = PAINT_ERASER) Then
                    Tools_Paint.NotifyBrushXY m_LMBDown, Shift, m_LastImageX, m_LastImageY, 0&, Me
                ElseIf (g_CurrentTool = PAINT_CLONE) Then
                    Tools_Clone.NotifyBrushXY m_LMBDown, Shift, m_LastImageX, m_LastImageY, 0&, Me
                ElseIf (g_CurrentTool = PAINT_GRADIENT) Then
                    Tools_Gradient.NotifyToolXY m_LMBDown, Shift, m_LastImageX, m_LastImageY, 0&, Me
                End If
                
                'Then, update the cursor to reflect any changes
                SetCanvasCursor pMouseMove, 0&, m_LastCanvasX, m_LastCanvasY, m_LastImageX, m_LastImageY, m_LastImageX, m_LastImageY
                
        End Select
        
    End If

End Sub

Private Sub CanvasView_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    markEventHandled = False

    'Make sure canvas interactions are allowed (e.g. an image has been loaded, etc)
    If IsCanvasInteractionAllowed() Then
        
        'If the color picker is currently in-use, and it was activated using the ALT key,
        ' we need to restore the user's original tool.
        If (g_CurrentTool = COLOR_PICKER) And Tools.GetToolAltState() And (vkCode = VK_ALT) Then
            
            Tools.SetToolAltState False
            toolbar_Toolbox.SelectNewTool g_PreviousTool
            
            'See detailed notes on these next two lines in the matching _KeyDown statement
            markEventHandled = True
            FormMain.HotkeyManager.NotifyAltKeystateChange False
            
        End If
        
        'Any further processing depends on which tool is currently active
        Select Case g_CurrentTool
                
            'Selection tools use a universal handler
            Case SELECT_RECT, SELECT_CIRC, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
                SelectionUI.NotifySelectionKeyUp Me, Shift, vkCode, markEventHandled
                
            'Pencil and paint tools redraw cursors under certain conditions
            Case PAINT_PENCIL, PAINT_SOFTBRUSH, PAINT_ERASER, PAINT_CLONE, PAINT_GRADIENT
                
                'First, notify the correct module
                If (g_CurrentTool = PAINT_PENCIL) Then
                    Tools_Pencil.NotifyBrushXY m_LMBDown, Shift, m_LastImageX, m_LastImageY, 0&, Me
                ElseIf (g_CurrentTool = PAINT_SOFTBRUSH) Or (g_CurrentTool = PAINT_ERASER) Then
                    Tools_Paint.NotifyBrushXY m_LMBDown, Shift, m_LastImageX, m_LastImageY, 0&, Me
                ElseIf (g_CurrentTool = PAINT_CLONE) Then
                    Tools_Clone.NotifyBrushXY m_LMBDown, Shift, m_LastImageX, m_LastImageY, 0&, Me
                ElseIf (g_CurrentTool = PAINT_GRADIENT) Then
                    Tools_Gradient.NotifyToolXY m_LMBDown, Shift, m_LastImageX, m_LastImageY, 0&, Me
                End If
                
                'Finally, update the canvas cursor to reflect any changes
                SetCanvasCursor pMouseMove, 0&, m_LastCanvasX, m_LastCanvasY, m_LastImageX, m_LastImageY, m_LastImageX, m_LastImageY
                
        End Select
        
        'Perform a special check for arrow keys.  VB may attempt to use these to control on-form navigation,
        ' which we do not want.
        If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Or (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then markEventHandled = True
        
    End If
    
End Sub

Private Sub cmdCenter_Click(ByVal Shift As ShiftConstants)
    Actions.LaunchAction_ByName "view_center_on_screen", pdas_Menu
End Sub

Private Sub CanvasView_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
        
    m_LastCanvasX = x
    m_LastCanvasY = y
    
    'Make sure interactions with this canvas are allowed
    If (Not Me.IsCanvasInteractionAllowed()) Then Exit Sub
    
    'Note whether a selection is active when mouse interactions began
    m_SelectionActiveBeforeMouseEvents = (PDImages.GetActiveImage.IsSelectionActive And PDImages.GetActiveImage.MainSelection.IsLockedIn)
    
    'These variables will hold the corresponding (x,y) coordinates on the IMAGE - not the VIEWPORT.
    ' (This is important if the user has zoomed into an image, and used scrollbars to look at a different part of it.)
    Dim imgX As Double, imgY As Double
    
    'Note that displayImageCoordinates returns a copy of the displayed coordinates via imgX/Y
    DisplayImageCoordinates x, y, PDImages.GetActiveImage(), Me, imgX, imgY
    m_LastImageX = imgX
    m_LastImageY = imgY
    
    'We also need a copy of the current mouse position relative to the active layer.
    ' (This became necessary in PD 7.0, as layers may have non-destructive affine transforms active,
    ' which means we can't blindly switch between image and layer coordinate spaces!)
    Dim layerX As Single, layerY As Single
    Drawing.ConvertImageCoordsToLayerCoords_Full PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, imgX, imgY, layerX, layerY
    
    'Display a relevant cursor for the current action
    SetCanvasCursor pMouseDown, Button, x, y, imgX, imgY, layerX, layerY
            
    'Check mouse button use
    If (Button = vbLeftButton) Then
        
        m_LMBDown = True
        m_NumOfMouseMovements = 0
        
        'Remember this location
        m_InitMouseX = x
        m_InitMouseY = y
        
        'Ask the current layer if these coordinates correspond to a point of interest.  We don't always use this return value,
        ' but a number of functions could potentially ask for it, so we cache it at MouseDown time and hang onto it until
        ' the mouse is released.
        m_CurPOI = PDImages.GetActiveImage.GetActiveLayer.CheckForPointOfInterest(imgX, imgY)
        
        'Any further processing depends on which tool is currently active
        Select Case g_CurrentTool
        
            'Drag-to-pan canvas
            Case NAV_DRAG
                Tools.SetInitialCanvasScrollValues FormMain.MainCanvas(0)
            
            'Zoom in/out (and click-drag to set zoom area)
            Case NAV_ZOOM
                Tools_Zoom.NotifyMouseDown Button, Shift, Me, PDImages.GetActiveImage, x, y
            
            'Move stuff around
            Case NAV_MOVE
                Tools_Move.NotifyMouseDown Me, Shift, imgX, imgY
                
            'Color picker
            Case COLOR_PICKER
                Tools_ColorPicker.NotifyMouseXY m_LMBDown, imgX, imgY, Me
            
            'Measure tool
            Case ND_MEASURE
                Tools_Measure.NotifyMouseDown FormMain.MainCanvas(0), imgX, imgY
            
            'Crop tool
            Case ND_CROP
                Tools_Crop.NotifyMouseDown Button, Shift, FormMain.MainCanvas(0), PDImages.GetActiveImage, x, y
            
            'Selections
            Case SELECT_RECT, SELECT_CIRC, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
                SelectionUI.NotifySelectionMouseDown Me, imgX, imgY
                
            'Text layer behavior varies depending on whether the current layer is a text layer or not
            Case TEXT_BASIC, TEXT_ADVANCED
                Tools_Text.NotifyMouseDown Button, Shift, imgX, imgY
                
            'Note for all paint tools: mouse interactions are disallowed if the active layer
            ' is locked or invisible.
            Case PAINT_PENCIL
                If DoesLayerAllowPainting(True) Then Tools_Pencil.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
            
            Case PAINT_SOFTBRUSH, PAINT_ERASER
                If DoesLayerAllowPainting(True) Then Tools_Paint.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                
            Case PAINT_CLONE
                If DoesLayerAllowPainting(True) Then Tools_Clone.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
            
            Case PAINT_FILL
                If DoesLayerAllowPainting(True) Then Tools_Fill.NotifyMouseXY m_LMBDown, imgX, imgY, Me
                
            Case PAINT_GRADIENT
                If DoesLayerAllowPainting(True) Then Tools_Gradient.NotifyToolXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                
            'In the future, other tools can be handled here
            Case Else
            
        End Select
    
    'TODO: right-button functionality?
    ElseIf (Button = vbRightButton) Then
        m_RMBDown = True
    
    ElseIf (Button = pdMiddleButton) Then
        m_MMBDown = True
        
        'Activate HAND TOOL behavior
        m_InitMouseX = x
        m_InitMouseY = y
        Tools.SetInitialCanvasScrollValues FormMain.MainCanvas(0)
        
        'Immediately update the cursor to reflect this change
        SetCanvasCursor pMouseDown, Button, x, y, imgX, imgY, layerX, layerY
        
    End If
    
End Sub

Private Sub CanvasView_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_IsMouseOverCanvas = True
    m_LastCanvasX = x
    m_LastCanvasY = y
End Sub

'When the mouse leaves the window, if no buttons are down, clear the coordinate display.
' (We must check for button states because the user is allowed to do things like drag selection nodes outside the image,
'  or paint outside the image.)
Private Sub CanvasView_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    m_IsMouseOverCanvas = False
    
    Select Case g_CurrentTool
        Case PAINT_PENCIL, PAINT_SOFTBRUSH, PAINT_ERASER, PAINT_CLONE, PAINT_FILL, PAINT_GRADIENT, COLOR_PICKER
            Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), Me
        Case SELECT_RECT, SELECT_CIRC, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
            SelectionUI.NotifySelectionMouseLeave Me
    End Select
    
    'If the mouse is not being used, clear the image coordinate display entirely
    If (Not m_LMBDown) And (Not m_RMBDown) Then
        
        'MouseLeave events are sent by the OS if the canvas view is disabled (which happens
        ' whenever an operation is processing in PD, to prevent click-through issues).
        ' This can lead to unwanted flickering when an operation completes, so to avoid it,
        ' we double-check that the mouse *has* actually left the canvas area.
        Dim tmpPoint As PointAPI, needToClear As Boolean
        If (GetCursorPos(tmpPoint) <> 0) Then needToClear = Not FormMain.MainCanvas(0).IsScreenCoordInsideCanvasView(tmpPoint.x, tmpPoint.y)
        If needToClear Then Interface.ClearImageCoordinatesDisplay
        
    End If
    
End Sub

Private Sub CanvasView_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    m_IsMouseOverCanvas = True
    m_LastCanvasX = x
    m_LastCanvasY = y
    
    'Make sure interactions with this canvas are allowed
    If Not IsCanvasInteractionAllowed() Then Exit Sub
    
    m_NumOfMouseMovements = m_NumOfMouseMovements + 1
    m_LMBDown = ((Button And pdLeftButton) <> 0)
    
    'These variables will hold the corresponding (x,y) coordinates on the image - NOT the viewport
    Dim imgX As Double, imgY As Double
    
    'Display the image coordinates under the mouse pointer, and if rulers are active, mirror the coordinates there, also
    Interface.DisplayImageCoordinates x, y, PDImages.GetActiveImage(), Me, imgX, imgY
    m_LastImageX = imgX
    m_LastImageY = imgY
    
    'We also need a copy of the current mouse position relative to the active layer.  (This became necessary in PD 7.0, as layers
    ' may have non-destructive affine transforms active, which means we can't reuse image coordinates as layer coordinates!)
    Dim layerX As Single, layerY As Single
    Drawing.ConvertImageCoordsToLayerCoords_Full PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, imgX, imgY, layerX, layerY
        
    'Check the left mouse button
    If m_LMBDown Then
    
        Select Case g_CurrentTool
        
            'Drag-to-pan canvas
            Case NAV_DRAG
                Tools.PanImageCanvas m_InitMouseX, m_InitMouseY, x, y, PDImages.GetActiveImage(), FormMain.MainCanvas(0)
            
            Case NAV_ZOOM
                Tools_Zoom.NotifyMouseMove Button, Shift, Me, PDImages.GetActiveImage, x, y
            
            'Move stuff around
            Case NAV_MOVE
                Tools_Move.NotifyMouseMove m_LMBDown, Shift, imgX, imgY
            
            'Crop tool
            Case ND_CROP
                Tools_Crop.NotifyMouseMove Button, Shift, FormMain.MainCanvas(0), PDImages.GetActiveImage, x, y
                SetCanvasCursor pMouseMove, Button, x, y, imgX, imgY, layerX, layerY
                
            'Color picker
            Case COLOR_PICKER
                Tools_ColorPicker.NotifyMouseXY m_LMBDown, imgX, imgY, Me
                SetCanvasCursor pMouseMove, Button, x, y, imgX, imgY, layerX, layerY
            
            'Measure tool
            Case ND_MEASURE
                Tools_Measure.NotifyMouseMove m_LMBDown, Shift, imgX, imgY
                SetCanvasCursor pMouseMove, Button, x, y, imgX, imgY, layerX, layerY
            
            'Selection tools
            Case SELECT_RECT, SELECT_CIRC, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
                SelectionUI.NotifySelectionMouseMove Me, True, Shift, imgX, imgY, m_NumOfMouseMovements
                
            'Text layers are identical to the move tool
            Case TEXT_BASIC, TEXT_ADVANCED
                Message "Shift key: preserve layer aspect ratio"
                Tools.TransformCurrentLayer imgX, imgY, PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, FormMain.MainCanvas(0), (Shift And vbShiftMask)
            
            'Unlike other tools, the paintbrush engine controls when the main viewport gets redrawn.
            ' (Some tricks are used to improve performance, including coalescing render events if they occur
            '  quickly enough.)  As such, there is no viewport redraw request here.
            Case PAINT_PENCIL
                If (Not m_IgnoreMouseActions) Then Tools_Pencil.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                
            Case PAINT_SOFTBRUSH, PAINT_ERASER
                If (Not m_IgnoreMouseActions) Then Tools_Paint.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
            
            Case PAINT_CLONE
                If (Not m_IgnoreMouseActions) Then Tools_Clone.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                
            Case PAINT_FILL
                If (Not m_IgnoreMouseActions) Then
                    Tools_Fill.NotifyMouseXY True, imgX, imgY, Me
                    SetCanvasCursor pMouseMove, Button, x, y, imgX, imgY, layerX, layerY
                End If
                
            Case PAINT_GRADIENT
                If (Not m_IgnoreMouseActions) Then Tools_Gradient.NotifyToolXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
            
        End Select
    
    'This else means the LEFT mouse button is NOT down
    Else
        
        'Middle-mouse button is a special case; we activate HAND TOOL behavior when it is used
        If m_MMBDown Then
            Tools.PanImageCanvas m_InitMouseX, m_InitMouseY, x, y, PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
        'Middle mouse-button is not down; treat this as a normal move event
        Else
            
            Select Case g_CurrentTool
            
                'Drag-to-navigate
                Case NAV_DRAG
                
                'Zoom in/out
                Case NAV_ZOOM
                    Tools_Zoom.NotifyMouseMove Button, Shift, Me, PDImages.GetActiveImage, x, y
                
                'Move stuff around
                Case NAV_MOVE
                    m_LayerAutoActivateIndex = Tools_Move.NotifyMouseMove(m_LMBDown, Shift, imgX, imgY)
                    
                'Crop tool
                Case ND_CROP
                    Tools_Crop.NotifyMouseMove Button, Shift, FormMain.MainCanvas(0), PDImages.GetActiveImage, x, y
                
                'Color picker
                Case COLOR_PICKER
                    Tools_ColorPicker.NotifyMouseXY m_LMBDown, imgX, imgY, Me
                
                'Measure tool
                Case ND_MEASURE
                    Tools_Measure.NotifyMouseMove m_LMBDown, Shift, imgX, imgY
                
                'Selection tools
                Case SELECT_RECT, SELECT_CIRC, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
                    SelectionUI.NotifySelectionMouseMove Me, False, Shift, imgX, imgY, m_NumOfMouseMovements
                    
                'Text tools
                Case TEXT_BASIC, TEXT_ADVANCED
                    'Nothing at present; the viewport renderer handles this for us
                
                Case PAINT_PENCIL
                    Tools_Pencil.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                
                Case PAINT_SOFTBRUSH, PAINT_ERASER
                    Tools_Paint.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                    
                Case PAINT_CLONE
                    Tools_Clone.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                    
                Case PAINT_FILL
                    Tools_Fill.NotifyMouseXY False, imgX, imgY, Me
                    
                Case PAINT_GRADIENT
                    Tools_Gradient.NotifyToolXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                    
                Case Else
                
            End Select
            
        End If
        
        'Now that everything's been updated, render a cursor to match
        SetCanvasCursor pMouseMove, Button, x, y, imgX, imgY, layerX, layerY
        
        'Yield for timer events only.  (This allows active UI animations, if any, to proceed.)
        VBHacks.DoEventsTimersOnly
        
    End If
    
End Sub

Private Sub CanvasView_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
        
    m_LastCanvasX = x
    m_LastCanvasY = y
    
    'Make sure interactions with this canvas are allowed
    If Not Me.IsCanvasInteractionAllowed() Then Exit Sub
    
    'Display the image coordinates under the mouse pointer
    Dim imgX As Double, imgY As Double
    DisplayImageCoordinates x, y, PDImages.GetActiveImage(), Me, imgX, imgY
    
    'We also need a copy of the current mouse position relative to the active layer.  (This became necessary in PD 7.0, as layers
    ' may have non-destructive affine transforms active, which means we can't blindly switch between image and layer coordinate spaces!)
    Dim layerX As Single, layerY As Single
    Drawing.ConvertImageCoordsToLayerCoords_Full PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, imgX, imgY, layerX, layerY
    
    'Display a relevant cursor for the current action
    SetCanvasCursor pMouseUp, Button, x, y, imgX, imgY, layerX, layerY
    
    'Check mouse buttons
    If (Button = vbLeftButton) Then
    
        m_LMBDown = False
    
        Select Case g_CurrentTool
        
            'Click-to-drag navigation requires no special behavior here
            Case NAV_DRAG
            
            'Zoom is handled later in this function, because it needs to track both left/right buttons
            'Case NAV_ZOOM
            
            'Move stuff around
            Case NAV_MOVE
                Tools_Move.NotifyMouseUp Button, Shift, imgX, imgY, m_NumOfMouseMovements
                
            'Color picker
            Case COLOR_PICKER
                Tools_ColorPicker.NotifyMouseXY m_LMBDown, imgX, imgY, Me
                
            'Measure tool
            Case ND_MEASURE
                Tools_Measure.NotifyMouseUp Button, Shift, imgX, imgY, m_NumOfMouseMovements, clickEventAlsoFiring
            
            'Crop is handled later in this function, because it needs to track both left/right buttons
            'Case NAV_CROP
            
            'Selection tools have their own dedicated handler
            Case SELECT_RECT, SELECT_CIRC, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
                SelectionUI.NotifySelectionMouseUp Me, Shift, imgX, imgY, clickEventAlsoFiring, m_SelectionActiveBeforeMouseEvents
                
            'Text layers
            Case TEXT_BASIC, TEXT_ADVANCED
                Tools_Text.NotifyMouseUp Button, Shift, imgX, imgY, m_NumOfMouseMovements, clickEventAlsoFiring
                
            'Notify the brush engine of the final result, then permanently commit this round of brush work
            Case PAINT_PENCIL
                If (Not m_IgnoreMouseActions) Then
                    Tools_Pencil.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                    Tools_Pencil.CommitBrushResults
                End If
                
            Case PAINT_SOFTBRUSH, PAINT_ERASER
                If (Not m_IgnoreMouseActions) Then
                    Tools_Paint.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                    Tools_Paint.CommitBrushResults
                End If
            
            Case PAINT_CLONE
                If (Not m_IgnoreMouseActions) Then
                    Tools_Clone.NotifyBrushXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                    Tools_Clone.CommitBrushResults
                End If
                
            Case PAINT_FILL
                If (Not m_IgnoreMouseActions) Then Tools_Fill.NotifyMouseXY m_LMBDown, imgX, imgY, Me
            
            Case PAINT_GRADIENT
                If (Not m_IgnoreMouseActions) Then
                    Tools_Gradient.NotifyToolXY m_LMBDown, Shift, imgX, imgY, timeStamp, Me
                    Tools_Gradient.CommitGradientResults
                End If
                
            Case Else
                    
        End Select
                        
    End If
    
    'Some controls handle multiple button possibilities themselves
    If (g_CurrentTool = NAV_ZOOM) Then
        Tools_Zoom.NotifyMouseUp Button, Shift, Me, PDImages.GetActiveImage(), x, y, m_NumOfMouseMovements, clickEventAlsoFiring
    ElseIf (g_CurrentTool = ND_CROP) Then
        Tools_Crop.NotifyMouseUp Button, Shift, FormMain.MainCanvas(0), PDImages.GetActiveImage, x, y, m_NumOfMouseMovements, clickEventAlsoFiring
    End If
    
    If (Button = pdRightButton) Then m_RMBDown = False
    If (Button = pdMiddleButton) Then m_MMBDown = False
    
    'Reset any tracked point of interest value for this layer
    m_CurPOI = poi_Undefined
        
    'Reset the mouse movement tracker
    m_NumOfMouseMovements = 0
    
End Sub

Public Sub CanvasView_MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    If (Not IsCanvasInteractionAllowed()) Then Exit Sub
    If hScroll.Visible Then hScroll.RelayMouseWheelEvent False, Button, Shift, x, y, scrollAmount
End Sub

'Vertical mousewheel scrolling.  Note that Shift+Wheel and Ctrl+Wheel modifiers do NOT raise this event; pdInputMouse automatically
' reroutes them to MouseWheelHorizontal and MouseWheelZoom, respectively.
Public Sub CanvasView_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    
    If Not IsCanvasInteractionAllowed() Then Exit Sub
    
    'The user can toggle mousewheel scroll vs zoom behavior from the Tools > Options > Interface panel
    If UserPrefs.GetZoomWithWheel Then
        HandleCanvasZoom Button, Shift, x, y, scrollAmount
    Else
        HandleCanvasScroll Button, Shift, x, y, scrollAmount
    End If
    
    'NOTE: horizontal scrolling via Shift+Vertical Wheel is handled in the separate _MouseWheelHorizontal event.
    'NOTE: zooming via Ctrl+Vertical Wheel is handled in the separate _MouseWheelZoom event.
    
End Sub

Public Sub CanvasView_MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
    
    If Not IsCanvasInteractionAllowed() Then Exit Sub
    
    'The user can toggle mousewheel scroll vs zoom behavior from the Tools > Options > Interface panel
    If UserPrefs.GetZoomWithWheel Then
        HandleCanvasScroll Button, Shift, x, y, zoomAmount
    Else
        HandleCanvasZoom Button, Shift, x, y, zoomAmount
    End If
    
End Sub

Private Sub HandleCanvasScroll(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

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
    
End Sub

Private Sub HandleCanvasZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
    If (zoomAmount <> 0) Then Tools_Zoom.RelayCanvasZoom Me, PDImages.GetActiveImage(), x, y, (zoomAmount > 0)
End Sub

Private Sub ImageStrip_Click(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    If ((Button And pdRightButton) <> 0) Then
        
        ucSupport.RequestCursor IDC_DEFAULT
        
        'Raise the context menu
        BuildPopupMenu
        If (Not m_PopupMenu Is Nothing) Then m_PopupMenu.ShowMenu ImageStrip.hWnd, x, y
        ShowCursor 1
        
    End If
    
End Sub

Private Sub ImageStrip_ItemClosed(ByVal itemIndex As Long)
    CanvasManager.FullPDImageUnload itemIndex
End Sub

Private Sub ImageStrip_ItemSelected(ByVal itemIndex As Long)
    CanvasManager.ActivatePDImage itemIndex, "user clicked image thumbnail"
End Sub

'When the image strip's position changes, we may need to move it to an entirely new position.  This also necessitates
' a layout adjustment of all other controls on the canvas.
Private Sub ImageStrip_PositionChanged()
    If (Not m_InternalResize) Then Me.AlignCanvasView
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    AlignCanvasView
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, False, True
    ucSupport.RequestExtraFunctionality True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDCANVAS_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDCanvas", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
    If PDMain.IsProgramRunning() Then
        
        'Allow the control to generate its own redraw requests
        Me.SetRedrawSuspension False
        
        'Set scroll bar size to match the current system default (which changes based on DPI, theming, and other factors)
        hScroll.Height = GetSystemMetrics(SM_CYHSCROLL)
        vScroll.Width = GetSystemMetrics(SM_CXVSCROLL)
        
        'Align the main picture box
        AlignCanvasView
        
        'Reset any POI trackers
        m_CurPOI = poi_Undefined
        m_LastPOI = poi_Undefined
        
    End If
    
End Sub

Private Sub HScroll_Scroll(ByVal eventIsCritical As Boolean)
    
    'Regardless of viewport state, cache the current scroll bar value inside the current image
    If PDImages.IsImageActive() Then PDImages.GetActiveImage.ImgViewport.SetHScrollValue hScroll.Value
    
    'Request the scroll-specific viewport pipeline stage, and notify all relevant UI elements of the change
    If (Not Me.GetRedrawSuspension) Then
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), Me
        Viewport.NotifyEveryoneOfViewportChanges
    End If
    
End Sub

Public Sub UpdateCanvasLayout()
    If (PDImages.GetNumOpenImages() = 0) Then Me.ClearCanvas Else Me.AlignCanvasView
    StatusBar.ReflowStatusBar PDImages.IsImageActive()
End Sub

Private Function ShouldImageStripBeVisible() As Boolean
    
    ShouldImageStripBeVisible = False
    
    'User preference = "Always visible"
    If (ImageStrip.VisibilityMode = 0) Then
        ShouldImageStripBeVisible = PDImages.IsImageActive()
    
    'User preference = "Visible if 2+ images loaded"
    ElseIf (ImageStrip.VisibilityMode = 1) Then
        ShouldImageStripBeVisible = (PDImages.GetNumOpenImages() > 1)
        
    'User preference = "Never visible"
    End If
    
End Function

'Given the current user control rect, modify it to account for the image tabstrip's position, and also fill a new rect
' with the tabstrip's dimensions.
Private Sub FillTabstripRect(ByRef ucRect As RectF, ByRef dstRect As RectF)
    
    Dim conSize As Long
    conSize = ImageStrip.ConstrainingSize
    
    With dstRect

        Select Case ImageStrip.Alignment
        
            Case vbAlignTop
                .Left = ucRect.Left
                .Top = ucRect.Top
                .Width = ucRect.Width
                .Height = conSize
                ucRect.Top = ucRect.Top + conSize
                ucRect.Height = ucRect.Height - conSize
            
            Case vbAlignBottom
                .Left = ucRect.Left
                .Top = (ucRect.Top + ucRect.Height) - conSize
                .Width = ucRect.Width
                .Height = conSize
                ucRect.Height = ucRect.Height - conSize
            
            Case vbAlignLeft
                .Left = ucRect.Left
                .Top = ucRect.Top
                .Width = conSize
                .Height = ucRect.Height
                ucRect.Left = ucRect.Left + conSize
                ucRect.Width = ucRect.Width - conSize
            
            Case vbAlignRight
                .Left = (ucRect.Left + ucRect.Width) - conSize
                .Top = ucRect.Top
                .Width = conSize
                .Height = ucRect.Height
                ucRect.Width = ucRect.Width - conSize
        
        End Select
        
    End With
    
End Sub

'Given the current user control rect, modify it to account for the status bar's position, and also fill a new rect
' with the status bar's dimensions.
Private Sub FillStatusBarRect(ByRef ucRect As RectF, ByRef dstRect As RectF)
    
    If m_StatusBarVisible Then ucRect.Height = ucRect.Height - StatusBar.GetHeight
    
    dstRect.Top = ucRect.Top + ucRect.Height
    dstRect.Height = StatusBar.GetHeight
    dstRect.Left = ucRect.Left
    dstRect.Width = ucRect.Width
    
End Sub

Public Sub AlignCanvasView()
        
    'Don't align anything until the program is up and running
    If (Not UserPrefs.IsReady) Then Exit Sub
        
    'Prevent recursive redraws by putting the entire UC into "resize mode"; while in this mode, we ignore anything that
    ' attempts to auto-initiate a canvas realignment request.
    If m_InternalResize Then Exit Sub
    m_InternalResize = True
    
    'Measurements must come from ucSupport (to guarantee that they're DPI-aware)
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetControlWidth
    bHeight = ucSupport.GetControlHeight
    
    'Using the DPI-aware measurements, construct a rect that defines the entire available control area
    Dim ucRect As RectF
    ucRect.Left = 0!
    ucRect.Top = 0!
    ucRect.Width = bWidth
    ucRect.Height = bHeight
    
    'The image tabstrip, if visible, gets placement preference
    Dim tabstripVisible As Boolean, tabstripRect As RectF
    tabstripVisible = ShouldImageStripBeVisible()
    
    'If we are showing the tabstrip for the first time, we need to position it prior to displaying it
    If tabstripVisible Then
        FillTabstripRect ucRect, tabstripRect
    Else
        ImageStrip.Visible = tabstripVisible
    End If
    
    'With the tabstrip rect in place, we now need to calculate a status bar rect (if any)
    Dim statusBarRect As RectF
    FillStatusBarRect ucRect, statusBarRect
    
    'Next comes ruler rects.  Note that rulers may or may not be visible, depending on user settings.
    Dim hRulerRectF As RectF, vRulerRectF As RectF
    With hRulerRectF
        .Top = ucRect.Top
        .Left = ucRect.Left
        .Width = ucRect.Width
        .Height = hRuler.GetHeight()
    End With
    
    With vRulerRectF
        .Top = ucRect.Top + hRulerRectF.Height
        .Left = ucRect.Left
        .Width = vRuler.GetWidth()
        .Height = ucRect.Height - hRulerRectF.Height
    End With
    
    'Account for ruler space - or if an image is *not* loaded, leave a ruler-sized border around
    ' the canvas area (for aesthetic reasons)
    If m_RulersVisible Or (Not PDImages.IsImageActive) Then
        ucRect.Top = ucRect.Top + hRulerRectF.Height
        ucRect.Height = ucRect.Height - hRulerRectF.Height
        ucRect.Left = ucRect.Left + vRulerRectF.Width
        ucRect.Width = ucRect.Width - vRulerRectF.Width
    End If
    
    'As of version 7.0, scroll bars are always visible.  This matches the behavior of paint-centric software like Krita,
    ' and makes it much easier to enable scrolling past the edge of an image (without resorting to stupid click-hold
    ' scroll behavior like GIMP).
    Dim hScrollTop As Long, hScrollLeft As Long, vScrollTop As Long, vScrollLeft As Long
    hScrollLeft = ucRect.Left
    hScrollTop = (ucRect.Top + ucRect.Height) - hScroll.GetHeight
    If hScroll.Visible Then ucRect.Height = ucRect.Height - hScroll.GetHeight
    
    vScrollTop = ucRect.Top
    vScrollLeft = (ucRect.Left + ucRect.Width) - vScroll.GetWidth
    If vScroll.Visible Then ucRect.Width = ucRect.Width - vScroll.GetWidth
    
    'With scroll bar positions calculated, calculate width/height values for the main canvas picture box
    Dim cvTop As Long, cvLeft As Long, cvWidth As Long, cvHeight As Long
    cvTop = ucRect.Top
    cvLeft = ucRect.Left
    cvWidth = ucRect.Width
    cvHeight = ucRect.Height
    
    'Move the CanvasView box into position first
    If (CanvasView.GetLeft <> cvLeft) Or (CanvasView.GetTop <> cvTop) Or (CanvasView.GetWidth <> cvWidth) Or (CanvasView.GetHeight <> cvHeight) Then
        If ((cvWidth > 0) And (cvHeight > 0)) Then CanvasView.SetPositionAndSize cvLeft, cvTop, cvWidth, cvHeight
    End If
    
    'The "no images loaded" panel always gets moved into the *same* position, but with a border
    ' around the outside.
    pnlNoImages.SetPositionAndSize cvLeft, cvTop, (cvWidth - cvLeft), (cvHeight - cvTop)
    
    '...Followed by the scrollbars
    If (hScroll.GetLeft <> hScrollLeft) Or (hScroll.GetTop <> hScrollTop) Or (hScroll.GetWidth <> cvWidth) Then
        If (cvWidth > 0) Then hScroll.SetPositionAndSize hScrollLeft, hScrollTop, cvWidth, hScroll.GetHeight
    End If
    
    If (vScroll.GetLeft <> vScrollLeft) Or (vScroll.GetTop <> vScrollTop) Or (vScroll.GetHeight <> cvHeight) Then
        If (cvHeight > 0) Then vScroll.SetPositionAndSize vScrollLeft, vScrollTop, vScroll.GetWidth, cvHeight
    End If
    
    '...Followed by the "center" button (which sits between the scroll bars)
    If (cmdCenter.GetLeft <> vScrollLeft) Or (cmdCenter.GetTop <> hScrollTop) Then
        cmdCenter.SetLeft vScrollLeft
        cmdCenter.SetTop hScrollTop
    End If
    
    '...Followed by rulers
    With hRulerRectF
        If m_RulersVisible And ((hRuler.GetLeft <> .Left) Or (hRuler.GetTop <> .Top) Or (hRuler.GetWidth <> .Width)) Then hRuler.SetPositionAndSize .Left, .Top, .Width, .Height
        hRuler.Visible = m_RulersVisible And PDImages.IsImageActive()
        hRuler.NotifyViewportChange
    End With
    
    With vRulerRectF
        If m_RulersVisible And ((vRuler.GetLeft <> .Left) Or (vRuler.GetTop <> .Top) Or (vRuler.GetHeight <> .Height)) Then vRuler.SetPositionAndSize .Left, .Top, .Width, .Height
        vRuler.Visible = m_RulersVisible And PDImages.IsImageActive()
        vRuler.NotifyViewportChange
    End With
    
    '...Followed by the status bar
    With statusBarRect
        StatusBar.SetPositionAndSize .Left, .Top, .Width, .Height
        StatusBar.Visible = m_StatusBarVisible
    End With
    
    '...and the progress bar placeholder.  (Note that it doesn't need a special rect - we always just position it
    ' above the status bar.)
    With statusBarRect
        mainProgBar.SetPositionAndSize .Left, .Top - mainProgBar.GetHeight, .Width, mainProgBar.GetHeight
    End With
    
    '...And finally, the image tabstrip (as relevant)
    With tabstripRect
        ImageStrip.SetPositionAndSize .Left, .Top, .Width, .Height
    End With
    
    If tabstripVisible And (Not ImageStrip.Visible) Then ImageStrip.Visible = True
    
    'If one or more images is loaded, ensure proper visibility between the primary canvas and
    ' PD's "quick start" panel.
    If PDImages.IsImageActive() Then
        CanvasView.Visible = True
        pnlNoImages.Visible = False
    Else
        
        'Retrieve relevant properties before refreshing, and hide the relevant UI elements
        ' so they don't trigger UI reflows
        chkRecentFiles.Visible = False
        chkRecentFiles.Value = UserPrefs.GetPref_Boolean("Interface", "WelcomeScreenRecentFiles", False)
        chkRecentFiles.Visible = True
        
        LayoutNoImages
        pnlNoImages.Visible = True
        CanvasView.Visible = False
        
    End If
    
    m_InternalResize = False
    
End Sub

Private Sub LayoutNoImages(Optional ByVal srcIsCheckBox As Boolean = False)
    
    Dim i As Long
    
    'Determine a good button width; this varies according to canvas area
    Dim curPanelWidth As Long, curPanelHeight As Long
    curPanelWidth = pnlNoImages.GetWidth()
    curPanelHeight = pnlNoImages.GetHeight()
    
    Dim btnWidth As Long, btnHeight As Long
    btnWidth = Interface.FixDPI(250)
    If (btnWidth > (curPanelWidth * 0.9)) Then btnWidth = curPanelWidth * 0.9
    btnHeight = Interface.FixDPI(50)
    
    Dim xPadding As Long, xOffset As Long
    xPadding = Interface.FixDPI(20)
    
    'If space is available (and recent files exist), make some layout adjustments to account
    ' for *two* columns of buttons - one for the standard new/open buttons, and another column
    ' for recent files.
    Dim showRecentFiles As Boolean, maxNumRecentFiles As Long
    If (Not g_RecentFiles Is Nothing) Then maxNumRecentFiles = g_RecentFiles.GetNumOfItems Else maxNumRecentFiles = 0
    showRecentFiles = (maxNumRecentFiles > 0) And ((btnWidth * 2 + xPadding) <= curPanelWidth)
    
    'Obviously, the value of the "show recent files" checkbox is also taken into account!
    If showRecentFiles Then showRecentFiles = chkRecentFiles.Value
    
    'If we *can* show recent files, we want to align the regular shortcuts differently.
    If showRecentFiles Then
        xOffset = (curPanelWidth - (btnWidth * 2 + xPadding)) \ 2
    Else
        xOffset = (curPanelWidth - btnWidth) \ 2
    End If
    
    'Before we can center our button collection vertically, we need to figure out its total height.
    ' As part of optimizing the redraw process, we're going to pre-calculate all item rects,
    ' then position them all at once.
    
    'Start with the left-side controls (which may be the only visible controls, if we're not
    ' displaying recent items).
    Dim totalHeight As Long, yPadding As Long, yOffset As Long
    yPadding = Interface.FixDPI(10)
    
    Dim objCount As Long
    objCount = cmdStart.UBound + 2
    
    Dim allRects() As RectL_WH
    ReDim allRects(0 To objCount) As RectL_WH
    
    totalHeight = 0
    
    For i = 0 To objCount
        allRects(i).Width = btnWidth
        allRects(i).Left = xOffset
        If (i = 0) Then
            allRects(i).Height = lblTitle(0).GetHeight
        ElseIf (i <= cmdStart.UBound + 1) Then
            allRects(i).Height = btnHeight
        Else
            allRects(i).Height = chkRecentFiles.GetHeight
        End If
        totalHeight = totalHeight + allRects(i).Height
        If (i < UBound(allRects)) Then totalHeight = totalHeight + yPadding
    Next i
    
    'Determine starting height
    yOffset = (curPanelHeight - totalHeight) \ 2
    
    'Position everything
    For i = LBound(allRects) To UBound(allRects)
        allRects(i).Top = yOffset
        If (i = 0) Then
            lblTitle(0).SetPositionAndSize allRects(i).Left, allRects(i).Top, allRects(i).Width, allRects(i).Height
        ElseIf (i <= cmdStart.UBound + 1) Then
            cmdStart(i - 1).SetPositionAndSize allRects(i).Left, allRects(i).Top, allRects(i).Width, allRects(i).Height
        Else
            chkRecentFiles.SetPositionAndSize allRects(i).Left, allRects(i).Top, allRects(i).Width, allRects(i).Height
        End If
        yOffset = yOffset + allRects(i).Height + yPadding
    Next i
    
    'If we *are* showing recent files, populate and position them next!
    If showRecentFiles Then
        
        xOffset = xOffset + btnWidth + xPadding
        
        'Time to repeat most the above steps, but for the recent files list
        objCount = maxNumRecentFiles + 2
        ReDim allRects(0 To objCount - 1) As RectL_WH
        
        totalHeight = 0
        
        For i = 0 To objCount - 1
            
            'If we're too tall for the current container, halt processing
            If ((totalHeight + btnHeight) > curPanelHeight) Then
                totalHeight = totalHeight - yPadding
                objCount = i
                If (objCount > 2) Then
                    allRects(objCount - 1).Height = hypRecentFiles.GetHeight
                    Exit For
                End If
            End If
            
            allRects(i).Width = btnWidth
            allRects(i).Left = xOffset
            If (i = 0) Then
                allRects(i).Height = lblTitle(1).GetHeight
            ElseIf (i < UBound(allRects)) Then
                allRects(i).Height = btnHeight
            Else
                allRects(i).Height = hypRecentFiles.GetHeight
            End If
            
            totalHeight = totalHeight + allRects(i).Height
            If (i < UBound(allRects)) Then
                totalHeight = totalHeight + yPadding
            
            'This branch ensures that buttons in the recent images column align with
            ' the standard buttons on the left (silly, I know, but it looks better!)
            Else
                totalHeight = totalHeight + (chkRecentFiles.GetHeight - hypRecentFiles.GetHeight)
            End If
            
        Next i
        
        'Ensure the number of available cmdRecent instances matches the number of recent file
        ' buttons we're gonna display
        Dim tooMany As Boolean
        tooMany = (cmdRecent.UBound > objCount - 2)
        
        If tooMany Then
            For i = cmdRecent.UBound To objCount - 2 Step -1
                Unload cmdRecent(i)
            Next i
        End If
        
        Dim tooFew As Boolean
        tooFew = cmdRecent.UBound < objCount - 2
        
        If tooFew Then
            For i = cmdRecent.UBound + 1 To objCount - 2
                Load cmdRecent(i)
            Next i
        End If
        
        'Determine starting height
        yOffset = (curPanelHeight - totalHeight) \ 2
        
        'Initialize a temporary object to hold thumbnails of recent files
        Dim btnImageSize As Long
        If OS.IsVistaOrLater Then
            btnImageSize = Interface.FixDPI(32)
        Else
            btnImageSize = Interface.FixDPI(16)
        End If
        
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateBlank btnImageSize, btnImageSize, 32, 0, 0
        tmpDIB.SetInitialAlphaPremultiplicationState True
        
        'Position everything
        For i = 0 To objCount - 1
            allRects(i).Top = yOffset
            If (i = 0) Then
                lblTitle(1).SetPositionAndSize allRects(i).Left, allRects(i).Top, allRects(i).Width, allRects(i).Height
                lblTitle(1).UpdateAgainstCurrentTheme
                lblTitle(1).Visible = True
            ElseIf (i < objCount - 1) Then
                cmdRecent(i - 1).SetPositionAndSize allRects(i).Left, allRects(i).Top, allRects(i).Width, allRects(i).Height
                If (Not g_RecentFiles Is Nothing) Then
                    
                    'Thumbnails may not exist; always check before accessing
                    If (Not g_RecentFiles.GetMRUThumbnail(i - 1) Is Nothing) Then
                        tmpDIB.ResetDIB 0
                        GDI_Plus.GDIPlus_StretchBlt tmpDIB, 0, 0, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, g_RecentFiles.GetMRUThumbnail(i - 1), 0, 0, g_RecentFiles.GetMRUThumbnail(i - 1).GetDIBWidth, g_RecentFiles.GetMRUThumbnail(i - 1).GetDIBHeight, interpolationType:=UserPrefs.GetThumbnailInterpolationPref(), dstCopyIsOkay:=True
                        cmdRecent(i - 1).AssignImage vbNullString, tmpDIB, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight
                    End If
                    
                    cmdRecent(i - 1).Caption = g_RecentFiles.GetMenuCaption(i - 1)
                    cmdRecent(i - 1).UpdateAgainstCurrentTheme
                    cmdRecent(i - 1).Visible = True
                    
                End If
            Else
                hypRecentFiles.SetPositionAndSize allRects(i).Left, allRects(i).Top, allRects(i).Width, allRects(i).Height
                hypRecentFiles.UpdateAgainstCurrentTheme
                hypRecentFiles.Visible = True
            End If
            
            yOffset = yOffset + allRects(i).Height + yPadding
            
        Next i
        
    '(We're not showing recent files.  Unload everything we can, then hide the rest)
    Else
        
        'Unload all buttons but the first one
        If (cmdRecent.UBound > 0) Then
            For i = 1 To cmdRecent.UBound
                Unload cmdRecent(i)
            Next i
        End If
        
        'Hide any remaining recent-files UI elements
        cmdRecent(0).Visible = False
        lblTitle(1).Visible = False
        hypRecentFiles.Visible = False
        
    End If
    
    'If this event *wasn't* initiated by the "show recent files" checkbox, try to set
    ' focus to a useful UI element.
    If (Not srcIsCheckBox) And (Not g_WindowManager Is Nothing) Then
        
        'Set keyboard focus to the first recent file
        If showRecentFiles And cmdRecent(0).Visible Then
            g_WindowManager.SetFocusAPI cmdRecent(0).hWnd
    
        'Set focus to the "open files" button
        Else
            g_WindowManager.SetFocusAPI cmdStart(1).hWnd
        End If
        
    End If
    
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Loading.LoadFromDragDrop Data, Effect, Button, Shift
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Loading.HelperForDragOver Data, Effect, Button, Shift, x, y, State
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

Private Sub vScroll_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Allow the ESC key to restore focus to the canvas itself
    If (Not markEventHandled) And (whichSysKey = pdnk_Escape) Then FormMain.MainCanvas(0).SetFocusToCanvasView
    
End Sub

Private Sub VScroll_Scroll(ByVal eventIsCritical As Boolean)
        
    'Regardless of viewport state, cache the current scroll bar value inside the current image
    If PDImages.IsImageActive() Then PDImages.GetActiveImage.ImgViewport.SetVScrollValue vScroll.Value
        
    If (Not Me.GetRedrawSuspension) Then
    
        'Request the scroll-specific viewport pipeline stage
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), Me
        
        'Notify any other relevant UI elements
        Viewport.NotifyEveryoneOfViewportChanges
        
    End If
    
End Sub

Public Sub PopulateSizeUnits()
    StatusBar.PopulateSizeUnits
End Sub

'To improve performance, the canvas is notified when it should read/write user preferences in a given session.
' This allows it to cache relevant values, then manage them independent of the preferences engine for the
' duration of a session.
Public Sub ReadUserPreferences()
    
    ImageStrip.ReadUserPreferences
    
    m_RulersVisible = UserPrefs.GetPref_Boolean("Toolbox", "RulersVisible", True)
    Menus.SetMenuChecked "view_rulers", m_RulersVisible
    
    m_StatusBarVisible = UserPrefs.GetPref_Boolean("Toolbox", "StatusBarVisible", True)
    Menus.SetMenuChecked "view_statusbar", m_StatusBarVisible
    
End Sub

Public Sub WriteUserPreferences()
    ImageStrip.WriteUserPreferences
    UserPrefs.SetPref_Boolean "Toolbox", "RulersVisible", m_RulersVisible
    UserPrefs.SetPref_Boolean "Toolbox", "StatusBarVisible", m_StatusBarVisible
End Sub

'Get/set ruler settings
Public Function GetRulerVisibility() As Boolean
    GetRulerVisibility = m_RulersVisible
End Function

Public Sub SetRulerVisibility(ByVal newState As Boolean)
    
    Dim changesMade As Boolean
    
    If (newState <> hRuler.Visible) Then
        m_RulersVisible = newState
        changesMade = True
    End If
    
    'When ruler visibility is changed, we must reflow the canvas area to match
    If changesMade Then Me.AlignCanvasView
    
End Sub

'Get/set status bar settings
Public Function GetStatusBarVisibility() As Boolean
    GetStatusBarVisibility = m_StatusBarVisible
End Function

Public Sub SetStatusBarVisibility(ByVal newState As Boolean)

    Dim changesMade As Boolean
    
    If (newState <> StatusBar.Visible) Then
        m_StatusBarVisible = newState
        changesMade = True
    End If
    
    If changesMade Then Me.AlignCanvasView
    
End Sub

'Various drawing tools support high-rate mouse input.  Change that behavior here.
Public Sub SetMouseInput_HighRes(ByVal newState As Boolean)
    CanvasView.SetMouseInput_HighRes newState
End Sub

Public Sub SetMouseInput_AutoDrop(ByVal newState As Boolean)
    CanvasView.SetMouseInput_AutoDrop newState
End Sub

Public Function GetNumMouseEventsPending() As Long
    GetNumMouseEventsPending = CanvasView.GetNumMouseEventsPending()
End Function

Public Function GetNextMouseMovePoint(ByVal ptrToDstMMP As Long) As Boolean
    GetNextMouseMovePoint = CanvasView.GetNextMouseMovePoint(ptrToDstMMP)
End Function

'Retrieve last-known mouse position; only valid if the mouse is actually over the canvas
Public Function GetLastMouseX() As Long
    GetLastMouseX = CanvasView.GetLastMouseX()
End Function

Public Function GetLastMouseY() As Long
    GetLastMouseY = CanvasView.GetLastMouseY()
End Function

'Paint tool mouse operations need to query this function before allowing paint interactions
Private Function DoesLayerAllowPainting(Optional ByVal displayUIFeedback As Boolean = False) As Boolean
    
    'This function only matters if the left mouse button is currently *down*.
    If m_LMBDown Then
    
        DoesLayerAllowPainting = PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility()
        If (Not DoesLayerAllowPainting) And displayUIFeedback Then
            Message "Action canceled: target layer isn't visible."
            UserControls.PostPDMessage WM_PD_FLASH_ACTIVE_LAYER, 500&, 4000&
        End If
        m_IgnoreMouseActions = (Not DoesLayerAllowPainting)
    
    Else
        DoesLayerAllowPainting = True
        m_IgnoreMouseActions = False
    End If
    
End Function

'Set the active canvas cursor via this function.
'
'Note that a number of extra values must be passed to this function.  Individual tools can use these parameters
' to customize cursor requests (e.g. the Move/Size tool will display different cursors depending on what's present
' at the given mouse location).
'
'Note that this function can also be used to hide the cursor completely, if you are using custom "cursor-ish"
' rendering for e.g. a paintbrush outline.
Private Sub SetCanvasCursor(ByVal curMouseEvent As PD_MOUSEEVENT, ByVal Button As Integer, ByVal x As Single, ByVal y As Single, ByVal imgX As Double, ByVal imgY As Double, ByVal layerX As Double, ByVal layerY As Double)
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'Some cursor functions operate on a POI basis
    Dim curPOI As PD_PointOfInterest
    
    'Prepare a default viewport parameter object; some cursor objects are rendered directly onto the
    ' primary canvas, so we may need to perform a viewport refresh
    Dim tmpViewportParams As PD_ViewportParams
    tmpViewportParams = Viewport.GetDefaultParamObject()
    
    'Handle some special cases first
    
    'If the MIDDLE MOUSE BUTTON is clicked, we silently activate the hand tool
    If m_MMBDown Then
        CanvasView.RequestCursor_Resource "cursor_handclosed", 0, 0
        Exit Sub
    End If
    
    'Obviously, cursor setting is handled separately for each tool.
    Select Case g_CurrentTool
        
        Case NAV_DRAG
        
            'When click-dragging the image to scroll around it, the cursor depends on being over the image
            If IsMouseOverImage(x, y, PDImages.GetActiveImage()) Then
                
                If (curMouseEvent = pMouseUp) Or (Button = 0) Then
                    CanvasView.RequestCursor_Resource "cursor_handopen", 0, 0
                Else
                    CanvasView.RequestCursor_Resource "cursor_handclosed", 0, 0
                End If
            
            'If the cursor is not over the image, change to an arrow cursor
            Else
                CanvasView.RequestCursor_System IDC_ARROW
            End If
        
        Case NAV_ZOOM
            If IsMouseOverImage(x, y, PDImages.GetActiveImage()) Then
                CanvasView.RequestCursor_Resource "cursor_zoom", 5, 6
            Else
                CanvasView.RequestCursor_System IDC_ARROW
            End If
            
        Case NAV_MOVE
            
            'When transforming layers, the cursor depends on the active POI
            curPOI = PDImages.GetActiveImage.GetActiveLayer.CheckForPointOfInterest(imgX, imgY, Tools_Move.GetDrawLayerRotateNodes())
            
            'If a POI has not been selected, but a selection is active, see if the mouse is
            ' in the selected area.
            Dim moveViaSelectionActive As Boolean
            If (curPOI = poi_Undefined) Then
                If PDImages.GetActiveImage.IsSelectionActive Then
                    moveViaSelectionActive = PDImages.GetActiveImage.MainSelection.IsPointSelected(imgX, imgY)
                    If moveViaSelectionActive Then curPOI = poi_Interior
                End If
            End If
            
            Select Case curPOI
            
                'Mouse is not over the current layer
                Case poi_Undefined
                    CanvasView.RequestCursor_System IDC_ARROW
                    
                'Mouse is over the top-left corner
                Case poi_CornerNW
                    CanvasView.RequestCursor_System IDC_SIZENWSE
                    
                'Mouse is over the top-right corner
                Case poi_CornerNE
                    CanvasView.RequestCursor_System IDC_SIZENESW
                    
                'Mouse is over the bottom-left corner
                Case poi_CornerSW
                    CanvasView.RequestCursor_System IDC_SIZENESW
                    
                'Mouse is over the bottom-right corner
                Case poi_CornerSE
                    CanvasView.RequestCursor_System IDC_SIZENWSE
                    
                'Mouse is over a rotation handle
                Case poi_EdgeE, poi_EdgeS, poi_EdgeW, poi_EdgeN
                    CanvasView.RequestCursor_System IDC_SIZEALL
                    'CanvasView.RequestCursor_Resource "cursor_rotate", 7, 7
                    
                'Mouse is within the layer, but not over a specific node
                Case poi_Interior
                
                    'This case is unique because if the user has elected to ignore transparent pixels, they cannot move a layer
                    ' by dragging the mouse within a transparent region of the layer.  Thus, before changing the cursor,
                    ' check to see if the hovered layer index is the same as the current layer index; if it isn't, don't display
                    ' the Move cursor.  (Note that this works because the getLayerUnderMouse function, called during the MouseMove
                    ' event, automatically factors the transparency check into its calculation.  Thus we don't have to
                    ' re-evaluate the setting here.)
                    If (m_LayerAutoActivateIndex = PDImages.GetActiveImage.GetActiveLayerIndex) Or moveViaSelectionActive Then
                        CanvasView.RequestCursor_System IDC_SIZEALL
                    Else
                        CanvasView.RequestCursor_System IDC_ARROW
                    End If
                    
            End Select
            
            'The move tool is unique because it will request a redraw of the viewport when the POI changes, so that the current
            ' POI can be highlighted.
            If (m_LastPOI <> curPOI) Then
                m_LastPOI = curPOI
                tmpViewportParams.curPOI = curPOI
                Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), Me, VarPtr(tmpViewportParams)
            End If
        
        'Crop tool handles cursor changes locally
        Case ND_CROP
            Tools_Crop.ReadyForCursor CanvasView
        
        'The color-picker custom-draws its own outline.
        Case COLOR_PICKER
            CanvasView.RequestCursor_System IDC_ICON
            Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), Me
        
        'The measurement tool uses a combination of cursors and on-canvas UI to do its thing
        Case ND_MEASURE
            If Tools_Measure.SpecialCursorWanted() Then CanvasView.RequestCursor_System IDC_SIZEALL Else CanvasView.RequestCursor_System IDC_ARROW
            Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), Me
        
        Case SELECT_RECT, SELECT_CIRC
        
            'When transforming selections, the cursor image depends on its proximity to a point of interest.
            Select Case IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage())
            
                Case poi_Undefined
                    CanvasView.RequestCursor_System IDC_ARROW
                Case poi_CornerNW
                    CanvasView.RequestCursor_System IDC_SIZENWSE
                Case poi_CornerNE
                    CanvasView.RequestCursor_System IDC_SIZENESW
                Case poi_CornerSE
                    CanvasView.RequestCursor_System IDC_SIZENWSE
                Case poi_CornerSW
                    CanvasView.RequestCursor_System IDC_SIZENESW
                Case poi_EdgeN
                    CanvasView.RequestCursor_System IDC_SIZENS
                Case poi_EdgeE
                    CanvasView.RequestCursor_System IDC_SIZEWE
                Case poi_EdgeS
                    CanvasView.RequestCursor_System IDC_SIZENS
                Case poi_EdgeW
                    CanvasView.RequestCursor_System IDC_SIZEWE
                Case poi_Interior
                    CanvasView.RequestCursor_System IDC_SIZEALL
            
            End Select
        
         Case SELECT_POLYGON
            
            Select Case IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage())
            
                Case poi_Undefined
                    CanvasView.RequestCursor_System IDC_ARROW
                
                'numOfPolygonPoints: mouse is inside the polygon, but not over a polygon node
                Case poi_Interior
                    If PDImages.GetActiveImage.MainSelection.IsLockedIn Then
                        CanvasView.RequestCursor_System IDC_SIZEALL
                    Else
                        CanvasView.RequestCursor_System IDC_ARROW
                    End If
                    
                'Everything else: mouse is over a polygon node
                Case Else
                    CanvasView.RequestCursor_System IDC_SIZEALL
                    
            End Select
        
        Case SELECT_LASSO
            
            Select Case IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage())
            
                Case poi_Undefined
                    CanvasView.RequestCursor_System IDC_ARROW
                
                'poi_Interior: mouse is inside the lasso selection area.  As a convenience to the user, we don't update the cursor
                '   if they're still in "drawing" mode - we only update it if the selection is complete.
                Case poi_Interior
                    If PDImages.GetActiveImage.MainSelection.IsLockedIn Then
                        CanvasView.RequestCursor_System IDC_SIZEALL
                    Else
                        CanvasView.RequestCursor_System IDC_ARROW
                    End If
                    
            End Select
            
        Case SELECT_WAND
        
            Select Case IsCoordSelectionPOI(imgX, imgY, PDImages.GetActiveImage())
            
                Case poi_Undefined
                    CanvasView.RequestCursor_System IDC_ARROW
                
                '0: mouse is inside the lasso selection area.  As a convenience to the user, we don't update the cursor
                '   if they're still in "drawing" mode - we only update it if the selection is complete.
                Case Else
                    CanvasView.RequestCursor_System IDC_SIZEALL
                    
            End Select
        
        Case TEXT_BASIC, TEXT_ADVANCED

            'The text tool bears a lot of similarity to the Move / Size tool, although the resulting behavior is
            ' obviously quite different.
            
            'First, see if the active layer is a text layer.  If it is, we need to check for POIs.
            If PDImages.GetActiveImage.GetActiveLayer.IsLayerText Then
                
                'When transforming layers, the cursor depends on the active POI
                curPOI = PDImages.GetActiveImage.GetActiveLayer.CheckForPointOfInterest(imgX, imgY)
                
                Select Case curPOI
    
                    'Mouse is not over the current layer
                    Case poi_Undefined
                        CanvasView.RequestCursor_System IDC_IBEAM
    
                    'Mouse is over the top-left corner
                    Case poi_CornerNW
                        CanvasView.RequestCursor_System IDC_SIZENWSE
                    
                    'Mouse is over the top-right corner
                    Case poi_CornerNE
                        CanvasView.RequestCursor_System IDC_SIZENESW
                    
                    'Mouse is over the bottom-left corner
                    Case poi_CornerSW
                        CanvasView.RequestCursor_System IDC_SIZENESW
                    
                    'Mouse is over the bottom-right corner
                    Case poi_CornerSE
                        CanvasView.RequestCursor_System IDC_SIZENWSE
                        
                    'Mouse is over a rotation handle
                    Case poi_EdgeE, poi_EdgeS, poi_EdgeW, poi_EdgeN
                        CanvasView.RequestCursor_System IDC_SIZEALL
                    
                    'Mouse is within the layer, but not over a specific node
                    Case poi_Interior
                        CanvasView.RequestCursor_System IDC_SIZEALL
                    
                End Select
                
                'Similar to the move tool, texts tools will request a redraw of the viewport when the POI changes, so that the current
                ' POI can be highlighted.
                If (m_LastPOI <> curPOI) Then
                    m_LastPOI = curPOI
                    tmpViewportParams.curPOI = curPOI
                    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), Me, VarPtr(tmpViewportParams)
                End If
                
            'If the current layer is *not* a text layer, clicking anywhere will create a new text layer
            Else
                CanvasView.RequestCursor_System IDC_IBEAM
            End If
        
        'Paint tools are a little weird, because we custom-draw the current brush outline - but *only*
        ' if no mouse button is down.  (If a button *is* down, the paint operation will automatically
        ' request a viewport refresh.)
        Case PAINT_PENCIL, PAINT_SOFTBRUSH, PAINT_ERASER, PAINT_CLONE
            CanvasView.RequestCursor_System IDC_ICON
            If (Button = 0) Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), Me
            
        'The fill tool needs to manually render a custom "paint bucket" icon regardless of mouse button state
        Case PAINT_FILL
            CanvasView.RequestCursor_System IDC_ICON
            Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), Me
            
        Case PAINT_GRADIENT
            CanvasView.RequestCursor_System IDC_ICON
            Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), Me
            
        Case Else
            CanvasView.RequestCursor_System IDC_ARROW
            
    End Select

End Sub

'Is the mouse currently over the canvas?
Public Function IsMouseOverCanvas() As Boolean
    IsMouseOverCanvas = m_IsMouseOverCanvas
End Function

'Is the user interacting with the canvas right now?
Public Function IsMouseDown(ByVal whichButton As PDMouseButtonConstants) As Boolean
    IsMouseDown = CanvasView.IsMouseDown(whichButton)
End Function

'Simple, unified way to see if canvas interaction is allowed.
Public Function IsCanvasInteractionAllowed() As Boolean
    IsCanvasInteractionAllowed = CanvasView.IsCanvasInteractionAllowed
End Function

'If the viewport experiences changes to scroll or zoom values, this function will be automatically called.
' Any relays to external functions (functions that rely on viewport settings, obviously) should be handled here.
' NOTE: external callers should call Viewport.NotifyEveryoneOfViewportChanges, which will automatically call
' this sub as well as a bunch of other notifiers throughout the project.
Public Sub NotifyViewportChanges()
    If m_RulersVisible Then
        hRuler.NotifyViewportChange
        vRuler.NotifyViewportChange
    End If
End Sub

'When the status bar changes its current measurement unit (e.g. pixels, cm, inches), it needs to notify the canvas
' so that other UI elements - like rulers - can synchronize.
Public Sub NotifyRulerUnitChange(ByVal newUnit As PD_MeasurementUnit)
    hRuler.NotifyUnitChange newUnit
    vRuler.NotifyUnitChange newUnit
End Sub

Public Function GetRulerUnit() As PD_MeasurementUnit
    GetRulerUnit = hRuler.GetCurrentUnit()
End Function

Public Sub NotifyImageStripVisibilityMode(ByVal newMode As Long)
    ImageStrip.VisibilityMode = newMode
End Sub

Public Sub NotifyImageStripAlignment(ByVal newAlignment As AlignConstants)
    ImageStrip.Alignment = newAlignment
End Sub

Public Sub SetCursorToCanvasPosition(ByVal newCanvasX As Double, ByVal newCanvasY As Double)
    CanvasView.SetCursorToCanvasPosition newCanvasX, newCanvasY
End Sub

Public Sub SetFocusToCanvasView()
    If CanvasView.Visible And (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI CanvasView.hWnd
End Sub

Private Sub BuildPopupMenu()
    
    Set m_PopupMenu = New pdPopupMenu
    
    With m_PopupMenu
        
        .AddMenuItem Menus.GetCaptionFromName("file_save"), menuIsEnabled:=Menus.IsMenuEnabled("file_save")
        .AddMenuItem Menus.GetCaptionFromName("file_savecopy"), menuIsEnabled:=Menus.IsMenuEnabled("file_savecopy")
        .AddMenuItem Menus.GetCaptionFromName("file_saveas"), menuIsEnabled:=Menus.IsMenuEnabled("file_saveas")
        .AddMenuItem Menus.GetCaptionFromName("file_revert"), menuIsEnabled:=Menus.IsMenuEnabled("file_revert")
        .AddMenuItem "-"
        
        'Open in Explorer only works if the image is currently on-disk
        .AddMenuItem g_Language.TranslateMessage("Show in file manager..."), menuIsEnabled:=(LenB(PDImages.GetActiveImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)) <> 0)
        
        .AddMenuItem "-"
        .AddMenuItem Menus.GetCaptionFromName("file_close")
        .AddMenuItem g_Language.TranslateMessage("Close all except this"), menuIsEnabled:=(PDImages.GetNumOpenImages() > 1)
        
    End With
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDC_StatusBar, "StatusBar", IDE_GRAY
        .LoadThemeColor PDC_SpecialButtonBackground, "SpecialButtonBackground", IDE_GRAY
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating all button captions against the current language.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0, Optional ByVal forceRefresh As Boolean = False)
    
    If (ucSupport.ThemeUpdateRequired Or forceRefresh) Then
        
        'Debug.Print "(the primary canvas is retheming itself - watch for excessive invocations!)"
        
        'Suspend redraws until all theme updates are complete
        Me.SetRedrawSuspension True
        
        UpdateColorList
        ucSupport.SetCustomBackcolor UserPrefs.GetCanvasColor()
        UserControl.BackColor = UserPrefs.GetCanvasColor()
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd, False
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        
        Dim btnImageSize As Long
        btnImageSize = Interface.FixDPI(26)
        pnlNoImages.UpdateAgainstCurrentTheme
        cmdStart(0).AssignImage "file_new", imgWidth:=btnImageSize, imgHeight:=btnImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
        cmdStart(1).AssignImage "file_open", imgWidth:=btnImageSize, imgHeight:=btnImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
        cmdStart(2).AssignImage "edit_paste", imgWidth:=btnImageSize, imgHeight:=btnImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
        
        'Use a slightly larger icon size for the batch icon, as it has fine details that look
        ' muddy at the size of the previous icons
        btnImageSize = Interface.FixDPI(28)
        cmdStart(3).AssignImage "file_batch", imgWidth:=btnImageSize, imgHeight:=btnImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
        
        Dim i As Long
        For i = cmdStart.lBound To cmdStart.UBound
            cmdStart(i).UpdateAgainstCurrentTheme
        Next i
        
        For i = lblTitle.lBound To lblTitle.UBound
            lblTitle(i).UpdateAgainstCurrentTheme
        Next i
        
        chkRecentFiles.UpdateAgainstCurrentTheme
        hypRecentFiles.UpdateAgainstCurrentTheme
        
        CanvasView.UpdateAgainstCurrentTheme
        StatusBar.UpdateAgainstCurrentTheme
        ImageStrip.UpdateAgainstCurrentTheme
        mainProgBar.UpdateAgainstCurrentTheme
        hRuler.UpdateAgainstCurrentTheme
        vRuler.UpdateAgainstCurrentTheme
        
        'Reassign tooltips to any relevant controls.  (This also triggers a re-translation against language changes.)
        btnImageSize = Interface.FixDPI(15)
        cmdCenter.AssignImage "zoom_center", Nothing, btnImageSize, btnImageSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
        cmdCenter.AssignTooltip "Center image in viewport"
        cmdCenter.BackColor = m_Colors.RetrieveColor(PDC_SpecialButtonBackground, Me.Enabled)
        cmdCenter.UpdateAgainstCurrentTheme
        
        hScroll.UpdateAgainstCurrentTheme
        vScroll.UpdateAgainstCurrentTheme
        
        'Any controls that utilize a custom background color must now be updated to match *our* background color.
        Dim sbBackColor As Long
        sbBackColor = m_Colors.RetrieveColor(PDC_StatusBar, Me.Enabled)
        
        Me.UpdateCanvasLayout
        
        'Restore redraw capabilities
        Me.SetRedrawSuspension False
    
    End If
    
End Sub
