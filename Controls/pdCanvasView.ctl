VERSION 5.00
Begin VB.UserControl pdCanvasView 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdCanvasView.ctx":0000
End
Attribute VB_Name = "pdCanvasView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "CanvasView" User Control (e.g. the primary component of pdCanvas)
'Copyright 2002-2026 by Tanner Helland
'Created: 29/November/02
'Last updated: 10/September/20
'Last update: change the way canvas background color is handled; it is now directly exposed via Tools > Options
'
'In 2013, PD's canvas was rebuilt as a dedicated user control, and instead of each image maintaining its own canvas inside
' separate, dedicated windows (which required a *ton* of code to keep in sync with the main PD window), a single canvas was
' integrated directly into the main window, and shared by all windows.
'
'In 2016, we refined this further, embedding this unique pdCanvasView object inside the larger pdCanvas object.  pdCanvas has
' manage a lot of things - scroll bars, buttons, a full status bar - and by migrating the main viewer portion into its own
' usercontrol, we gained a lot of flexibility over performance and code separation.  This is especially important as part of
' improving on-canvas tool responsiveness.
'
'To really understand how this control operates, you'll need to examine pdCanvas, as it ultimately deals with the many mouse
' and key events we potentially raise.
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

'This control bubbles every input event up to its parent control (pdCanvas).  Please do not add any tool-specific handling
' to this control instance; that level of decision-making should happen inside pdCanvas itself.
Public Event MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseHover(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
Public Event MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
Public Event MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
Public Event MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
Public Event ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)
Public Event KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)
Public Event AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

'To improve performance, external functions can request that we ignore any refresh requests.
' IMPORTANT NOTE: IT IS IMPERATIVE THAT YOU SET THIS VALUE CORRECTLY.  If you set it and forget to release it,
' the canvas will be locked in an inactive state.
Private m_SuspendRedraws As Boolean

'Because PD is single-threaded, we sometimes need to notify this control of mouse events it may not have caught
' (because something else was occurring in the background).  When manually notified, a temporary flag is set
' until the next time an "honest" WM_MOUSEMOVE message arrives.
Private m_ManualMouseMode As Boolean

'The last x/y coordinates recorded by standard canvas mouse events.  Importantly, these are *not* processed
' at high-DPI; they exist purely as support functions for things like hotkeys that modify their behavior
' to incorporate last-known mouse coordinates as part of an externally triggered function.

'Also, the value of these is "indeterminate" if the mouse is not currently over the canvas - always check
' that first!
Private m_LastMouseX As Long, m_LastMouseY As Long

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_CanvasView
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'This is one of the few controls in PD that manages a persistent DC.  This is a deliberate decision, so we can maintain expensive
' per-DC settings like color management.
Public Property Get hDC() As Long
    hDC = ucSupport.GetBackBufferDC
End Property

Public Function GetCanvasWidth() As Long
    GetCanvasWidth = ucSupport.GetControlWidth
End Function

Public Function GetCanvasHeight() As Long
    GetCanvasHeight = ucSupport.GetControlHeight
End Function

'To support high-DPI settings properly, we expose some specialized move+size functions
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

Public Sub NotifyExternalMouseMove(ByVal srcX As Long, ByVal srcY As Long)
    m_ManualMouseMode = True
    m_LastMouseX = srcX
    m_LastMouseY = srcY
    RaiseEvent MouseMoveCustom(0&, 0&, srcX, srcY, 0&)
End Sub

'Erase the current canvas.  If no images are loaded (which is really the only time this function should be called,
' we'll render the generic "please load an image" icon onto the background.
Public Sub ClearCanvas()

    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, UserPrefs.GetCanvasColor())
    If (bufferDC = 0) Then Exit Sub
    
    ucSupport.RequestRepaint
    
End Sub

'Retrieve last-known mouse position; only valid if the mouse is actually over the canvas
Public Function GetLastMouseX() As Long
    GetLastMouseX = m_LastMouseX
End Function

Public Function GetLastMouseY() As Long
    GetLastMouseY = m_LastMouseY
End Function

'Is the mouse over the canvas view right now?
Public Function IsMouseOverCanvasView() As Boolean
    IsMouseOverCanvasView = ucSupport.IsMouseInside() Or m_ManualMouseMode
End Function

'Is the user interacting with the canvas right now?
Public Function IsMouseDown(ByVal whichButton As PDMouseButtonConstants) As Boolean
    IsMouseDown = ucSupport.IsMouseButtonDown(whichButton)
End Function

'Certain criteria prevent the user from interacting with this canvas object (e.g. no images being loaded).  External functions
' *must* check this value before attempting to render to the canvas.
Public Function IsCanvasInteractionAllowed() As Boolean
    
    'By default, canvas interaction is allowed
    IsCanvasInteractionAllowed = True
    
    'Now, check a bunch of states that might indicate canvas interactions should not be allowed
    
    'If the main form is disabled, exit
    If (Not FormMain.Enabled) Then IsCanvasInteractionAllowed = False
        
    'If user input has been forcibly disabled by some other part of the program, exit
    If g_DisableUserInput And (Not m_ManualMouseMode) Then IsCanvasInteractionAllowed = False
    
    'If no images have been loaded, exit
    If (Not PDImages.IsImageActive()) Then IsCanvasInteractionAllowed = False
    
    'If our own internal redraw suspension flag is set, exit
    If m_SuspendRedraws Then IsCanvasInteractionAllowed = False
    
    'If any of the previous checks succeeded, exit immediately
    If (Not IsCanvasInteractionAllowed) Then Exit Function
    
    'If there is no active images or valid layers, canvas interactions are also disallowed.  (This is primarily a failsafe check.)
    If (Not PDImages.GetActiveImage.IsActive) Then IsCanvasInteractionAllowed = False
    If (PDImages.GetActiveImage.GetNumOfLayers = 0) Then IsCanvasInteractionAllowed = False
    
    'If the central processor is active, exit - but *only* if our internal notification flags have not been triggered.
    ' (If those flags are active, it means an external caller has notified us of something it wants rendered.)
    If m_ManualMouseMode Then
        IsCanvasInteractionAllowed = True
    Else
        If Processor.IsProgramBusy Then IsCanvasInteractionAllowed = False
    End If
    
End Function

Public Sub RequestCursor_System(Optional ByVal standardCursorType As SystemCursorConstant = IDC_DEFAULT)
    ucSupport.RequestCursor standardCursorType
End Sub

Public Sub RequestCursor_Resource(ByVal pngResourceName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0)
    ucSupport.RequestCursor_Resource pngResourceName, cursorHotspotX, cursorHotspotY
End Sub

'External functions can request an immediate redraw.  Please don't abuse this - it should really only be used
' when a UI element needs to be updated independent of PD's normal refresh cycles.
Public Sub RequestRedraw(Optional ByVal repaintImmediately As Boolean = False)
    RedrawBackBuffer repaintImmediately
End Sub

'For some tool actions, it may be helpful to move the cursor for the user.  Call this function to forcibly set a cursor position.
' (Note that this function will automatically handle the translation to screen coordinates, so passed coordinates should be
'  relative to the canvas itself - e.g. to the control's window.)
Public Sub SetCursorToCanvasPosition(ByVal canvasX As Double, ByVal canvasY As Double)
    ucSupport.RequestMousePosition canvasX, canvasY
End Sub

'Use these functions to forcibly prevent the canvas from redrawing itself.  REDRAWS WILL NOT HAPPEN AGAIN UNTIL YOU RESTORE ACCESS!
Public Function GetRedrawSuspension() As Boolean
    GetRedrawSuspension = m_SuspendRedraws
End Function

Public Sub SetRedrawSuspension(ByVal newRedrawValue As Boolean)
    m_SuspendRedraws = newRedrawValue
End Sub

'Manually request standard-rate or high-rate mouse tracking.  (Drawing tools support high-rate tracking.)
Public Sub SetMouseInput_HighRes(ByVal newState As Boolean)
    ucSupport.RequestHighResMouseInput newState
End Sub

Public Sub SetMouseInput_AutoDrop(ByVal newState As Boolean)
    ucSupport.RequestAutoDropMouseMessages newState
End Sub

Public Function GetNumMouseEventsPending() As Long
    GetNumMouseEventsPending = ucSupport.GetNumMouseEventsPending()
End Function

Public Function GetNextMouseMovePoint(ByVal ptrToDstMMP As Long) As Boolean
    Dim tmpMMP As MOUSEMOVEPOINT
    GetNextMouseMovePoint = ucSupport.GetNextMouseMovePoint(tmpMMP)
    CopyMemoryStrict ptrToDstMMP, VarPtr(tmpMMP), LenB(tmpMMP)
End Function

Private Sub ucSupport_AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent AppCommand(cmdID, Shift, x, y)
End Sub

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent ClickCustom(Button, Shift, x, y)
End Sub

Private Sub ucSupport_DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent DoubleClickCustom(Button, Shift, x, y)
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    RaiseEvent KeyDownCustom(Shift, vkCode, markEventHandled)
End Sub

Private Sub ucSupport_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    RaiseEvent KeyUpCustom(Shift, vkCode, markEventHandled)
End Sub

Private Sub ucSupport_LostFocusAPI()
    m_ManualMouseMode = False
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    RaiseEvent MouseDownCustom(Button, Shift, x, y, timeStamp)
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'If no images have been loaded, reset the cursor
    If (PDImages.GetNumOpenImages() = 0) Then ucSupport.RequestCursor IDC_DEFAULT
    RaiseEvent MouseEnter(Button, Shift, x, y)
    
End Sub

Private Sub ucSupport_MouseHover(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent MouseHover(Button, Shift, x, y)
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_ManualMouseMode = False
    RaiseEvent MouseLeave(Button, Shift, x, y)
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    m_ManualMouseMode = False
    m_LastMouseX = x
    m_LastMouseY = y
    RaiseEvent MouseMoveCustom(Button, Shift, x, y, timeStamp)
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    RaiseEvent MouseUpCustom(Button, Shift, x, y, clickEventAlsoFiring, timeStamp)
End Sub

Private Sub ucSupport_MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RaiseEvent MouseWheelHorizontal(Button, Shift, x, y, scrollAmount)
End Sub

Private Sub ucSupport_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RaiseEvent MouseWheelVertical(Button, Shift, x, y, scrollAmount)
End Sub

Private Sub ucSupport_MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
    RaiseEvent MouseWheelZoom(Button, Shift, x, y, zoomAmount)
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    
    'If no images are loaded, repaint ourselves automatically
    If (PDImages.GetNumOpenImages() = 0) And PDMain.IsProgramRunning() Then
        Me.ClearCanvas
    Else
    
        'If we don't manually suspend repainting, we'll get *crazy* flickering because ViewportEngine
        ' is slower than the automatic repaints requested by our parent pdUCSupport instance.  As such,
        ' we manually disable repaints until the viewport buffer is ready.
        ucSupport.SuspendAutoRepaintBehavior True
        Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        ucSupport.SuspendAutoRepaintBehavior False
        
        'Because we suspended auto-repaints, we must manually request a final paint-to-screen
        ucSupport.RequestRepaint True
        
    End If
    
End Sub

Private Sub UserControl_Initialize()

    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True, True, True
    ucSupport.SpecifyRequiredKeys VK_SHIFT, VK_ALT, VK_CONTROL, VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN, VK_DELETE, VK_INSERT, VK_TAB, VK_SPACE, VK_ESCAPE, VK_BACK
    ucSupport.RequestHighPerformanceRendering True
    
End Sub

'(This code is copied from FormMain's OLEDragDrop event - please mirror any changes there, or even better, stop being lazy
' and write a universal drag/drop handler!)
Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Loading.LoadFromDragDrop Data, Effect, Button, Shift, x, y
End Sub

'(This code is copied from FormMain's OLEDragOver event - please mirror any changes there, or even better, stop being lazy
' and write a universal drag/drop handler!)
Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Loading.HelperForDragOver Data, Effect, Button, Shift, x, y, State
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer(Optional ByVal refreshImmediately As Boolean = False)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(False)
    If (bufferDC = 0) Then Exit Sub
    
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight

    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint refreshImmediately
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UserControl.BackColor = UserPrefs.GetCanvasColor()
        If (PDImages.GetNumOpenImages() = 0) Then Me.ClearCanvas
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub
