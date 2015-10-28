VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdUCSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Comprehensive User Control Support class
'Copyright 2015-2015 by Tanner Helland
'Created: 25/October/15  (but built from many parts existing earlier)
'Last updated: 28/October/15
'Last update: continue work on initial build
'
'Now that PD provides so many of its own user controls (UCs), we have to repeat a lot of UC management code.
' Painting, custom input handling, subclassing custom window messages - the list goes on for awhile.
'
'To cut down on the amount of repeat work I have to do for each unique control, I'm trying to wrap as many generic
' UC capabilities as I can into this single class.  It then raises a boatload of events for individual UCs to
' handle as they see fit.
'
'In the interim, some user controls may still implement many of these features manually.  Migrating them over is
' a WIP, and one I hope to have done by the 7.0 release.  My prototype UC for testing has been the old ColorSelector
' control, because it requires a lot of weird features (like captions) and the code needed an overhaul anyway.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This class raises many events.  Some *must* be responded to.  Others are optional.
' (These are poorly documented at present - sorry!)
Public Event CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef bHandled As Boolean)
Public Event GotFocusAPI()
Public Event LostFocusAPI()
Public Event WindowMove(ByVal newLeft As Long, ByVal newTop As Long)
Public Event WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
Public Event RepaintRequired(ByVal updateLayoutToo As Boolean)
Public Event VisibilityChange(ByVal newVisibility As Boolean)

Public Event MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseHover(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
Public Event MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
Public Event MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
Public Event ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
Public Event DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

'API technique for drawing a focus rectangle; used only for designer mode (see the GetBackBufferDC method for details)
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT) As Long

'This class subclasses a *lot* of things.  Some are handled by this class itself; others use dedicated single-purpose
' support classes.
Private Const WM_SHOWWINDOW As Long = &H18

Private m_Subclass As cSelfSubHookCallback
Private WithEvents m_MouseEvents As pdInputMouse
Attribute m_MouseEvents.VB_VarHelpID = -1
Private WithEvents m_FocusDetector As pdFocusDetector
Attribute m_FocusDetector.VB_VarHelpID = -1
Private WithEvents m_Painter As pdWindowPainter
Attribute m_Painter.VB_VarHelpID = -1
Private WithEvents m_MoveSize As pdWindowSize
Attribute m_MoveSize.VB_VarHelpID = -1
Private m_ToolTipManager As pdToolTip
Private m_Caption As pdCaption

'Persistent copy of the associated user control's hWnd and DC.  (If tooltips are activated, we also cache the parent hWnd.)
Private m_hWnd As Long, m_hDC As Long, m_hWndParent As Long

'Size of the underlying usercontrol.  These values are retrieved via API, because VB has DPI-awareness issues.
Private m_Width As Long, m_Height As Long

'If the caller has registered one (or more) custom messages, this will be set to TRUE.  When the parent control is
' unloaded, we need to de-register the hWnd with PD's central message marshaler.
Private m_CustomInternalMessagesActive As Boolean

'Persistent back buffer, which we manage for the UC.  This provides a bunch of nice benefits, like automatic size syncing,
' color management when flipping to the screen, and perhaps someday, use of windowless controls, where appropriate.
Private m_BackBuffer As pdDIB

'Current back color of the underlying control, if any.  -1 if no back color has been specified.
Private m_BackColor As Long

'If a caption is active, this is the rect where it has been automatically drawn.  (The caller probably needs this in order to know
' where to position the rest of the control's contents.)
Private m_CaptionRect As RECT

'Many controls find it useful to know if the mouse is inside the control boundaries.  Rather than force them all to track
' this manually, we track it for them.  (Obviously, this value is only valid if they've specifically requested mouse tracking.)
Private m_MouseInsideUC As Boolean

'Visibility is also tracked via the API
Private m_Visibility As Boolean

'This class tries very hard to minimize redraw requests to the parent UC.  As a result, certain sets of conditions can lead to a
' control being shown, without any redraw requests being raised.  To prevent this, we use a failsafe check when the control is
' shown for the first time.  If no redraw requests have been raised, we raise one then, to ensure the control is shown correctly.
Private m_RedrawRequestCount As Long


'***************************************************************************
' SETUP AND INITIALIZATION FUNCTIONS
'***************************************************************************

'The FIRST THING a user control must do is call this sub.  This prepares the support class, applies basic subclassing,
' and initializes a bunch of other support modules.
Public Sub RegisterControl(ByVal hWnd_Control As Long, ByVal hDC_Control As Long)
    
    m_hWnd = hWnd_Control
    m_hDC = hDC_Control
    
    If (m_hWnd <> 0) Then
    
        'Subclass the user control immediately
        If g_IsProgramRunning Then m_Subclass.ssc_Subclass m_hWnd, , 1, Me, True, True
    
        'Some functionality is supplied "by default".  The caller does not have to request this behavior; they get it for free.
        Set m_ToolTipManager = New pdToolTip            'Themed, Unicode-friendly, multiline tooltips
        
        Set m_Painter = New pdWindowPainter             'Flicker-free double-buffered window painting
        If g_IsProgramRunning Then m_Painter.startPainter m_hWnd
        
        Set m_FocusDetector = New pdFocusDetector       'Reliable, API-based Got/LostFocus detection
        If g_IsProgramRunning Then m_FocusDetector.startFocusTracking m_hWnd
        
        Set m_MoveSize = New pdWindowSize               'Reliable, API-based, DPI-aware move/size functionality
        m_MoveSize.AttachToHWnd m_hWnd, g_IsProgramRunning
        
        CacheWindowDimensions
        
        'Initialize the backbuffer and sync it to the current control size
        SyncBackBufferSize True
    
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdUCSupport.RegisterControl was called with hWnd = 0."
        #End If
    End If
    
End Sub

'After RegisterControl has been called (above), the caller can request additional support via this function.
' (Because not all user controls require things like captions or custom input events, this functionality is restricted to
'  an "as-required" basis.)
Public Sub RequestExtraFunctionality(Optional ByVal customMouseEvents As Boolean = False, Optional ByVal customKeyEvents As Boolean = False)
    
    If m_hWnd <> 0 Then
    
        'Initialize special subclassers, as requested
        If customMouseEvents Then
            Set m_MouseEvents = New pdInputMouse
            If g_IsProgramRunning Then m_MouseEvents.addInputTracker m_hWnd, True, True, False, True, True
        End If
        
        If customKeyEvents Then
            'TODO:
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  You can't request custom functionality prior to calling pdUCSupport.RegisterControl()!"
        #End If
    End If
    
End Sub

'Want to allow some automatic text captioning on your control?  No problem; request it here, and the support class takes care of
' the rest.  (Obviously you also need to supply the caption text, font, and other things, but this function lets the class know
' that caption support is at least a *possibility*.)
Public Sub RequestCaptionSupport(Optional ByVal supportWordWrap As Boolean = False)
    m_Caption.setWordWrapSupport supportWordWrap
End Sub

'Get/Set functions for various caption-specific features.  Note that you can remove caption support by calling SetCaptionText and
' supplying a blank string.
Public Function GetCaptionText() As String
    GetCaptionText = m_Caption.getCaptionEn()
End Function

Public Sub SetCaptionText(Optional ByVal newCaption As String = "")
    
    If m_Caption Is Nothing Then Set m_Caption = New pdCaption
    
    'Pass the string to the pdCaption instance; it handles the actual mess of sizing and rendering text
    If m_Caption.setCaption(newCaption) Then
        
        'Caption changes may require us to resize the control to fit.
        SyncBackBufferSize True
        
        'Request an immediate redraw (particularly important inside the IDE, so the caller can see the new text)
        AskForARepaint True
        
    End If
    
End Sub

Public Function GetCaptionFontSize() As Single
    GetCaptionFontSize = m_Caption.getFontSize
End Function

Public Sub SetCaptionFontSize(ByVal newSize As Single)
    
    'Changes to fontsize require redraws
    If m_Caption.setFontSize(newSize) Then
        SyncBackBufferSize True
        AskForARepaint True
    End If
    
End Sub

Public Function GetCaptionBottom() As Long
    GetCaptionBottom = m_CaptionRect.Bottom
End Function

Public Function IsCaptionActive() As Boolean
    If m_Caption Is Nothing Then
        IsCaptionActive = False
    Else
        IsCaptionActive = m_Caption.IsCaptionActive()
    End If
End Function

'Because this support class manages a back buffer for its parent control, it sometimes needs to raise redraw requests.  To prevent
' an excessive amount of back-and-forth requests, the caller needs to explicitly specify a custom backcolor if it wants one.
' NOTE: call this function without supplying a newColor value to reset to the default program-wide window backcolor.
' ANOTHER NOTE: this function won't raise a redraw request, by design.
Public Sub SetCustomBackColor(Optional ByVal newColor As Long = -1&)
    m_BackColor = newColor
End Sub

'Want to receive custom window messages?  No problem.  Register them here, then watch for the CustomMessage event.
' Note, however, that if the message is specific to PD, the caller needs to let us know, as we will need to manually
' handle message marshaling.
Public Sub SubclassCustomMessage(ByVal wMsg As Long, Optional ByVal msgIsInternalToPD As Boolean = False)
    
    m_Subclass.ssc_AddMsg m_hWnd, MSG_BEFORE, wMsg
    
    'If the message is internal to PD, register the hWnd+message pair with PD's central marshaler
    If msgIsInternalToPD Then
        UserControl_Support.AddMessageRecipient m_hWnd, wMsg
        m_CustomInternalMessagesActive = True
    End If
    
End Sub

'***************************************************************************
' (end setup and initialization functions)
'***************************************************************************

'***************************************************************************
' INTERNAL FUNCTIONS
'***************************************************************************

Private Sub AskForARepaint(ByVal askForLayoutUpdateToo As Boolean)
    m_RedrawRequestCount = m_RedrawRequestCount + 1
    RaiseEvent RepaintRequired(askForLayoutUpdateToo)
End Sub

'Cache the current dimensions of the underlying UC.  Returns TRUE if dimensions were retrieved successfully.
Private Function CacheWindowDimensions() As Boolean
    
    If Not (m_MoveSize Is Nothing) Then
        m_Width = m_MoveSize.GetWidth
        m_Height = m_MoveSize.GetHeight
        CacheWindowDimensions = True
    Else
        CacheWindowDimensions = False
    End If
    
End Function

'The naming of this function is potentially confusing, but it restores the back buffer to the currently specified back color.
' It does not resize the DIB; it just fill it with a single uniform color.  Optionally, a specific erase rect can be supplied.
' (If coordinates are *not* supplied, the entire buffer will be wiped.)
Private Function ResetBackBuffer(Optional ByVal newBackColor As Long = -1&, Optional ByVal rpLeft As Single = 0#, Optional ByVal rpTop As Single = 0#, Optional ByVal rpWidth As Single = 0#, Optional ByVal rpHeight As Single = 0#)
    
    'If the caller specifies a background color, we want to use it.  If they don't, grab the default window background from
    ' PD's central themer.
    If newBackColor = -1 Then
        If g_Themer Is Nothing Then
            newBackColor = RGB(255, 255, 255)
        Else
            If m_BackColor = -1 Then
                newBackColor = g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
            Else
                newBackColor = m_BackColor
            End If
        End If
    End If
    
    'There are a couple different ways to apply the repaint.  If the caller specifies a fill rect, we obey it, but if they don't,
    ' we want to wipe the entire buffer.  As a convenience to the caller, we also apply the backcolor inside the IDE.
    If g_IsProgramRunning Then
        If (rpLeft = 0#) And (rpTop = 0#) And (rpWidth = 0#) And (rpHeight = 0#) Then
            GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0#, 0#, m_Width, m_Height, newBackColor
        Else
            GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, rpLeft, rpTop, rpWidth, rpHeight, newBackColor
        End If
    Else
        If (rpLeft = 0#) And (rpTop = 0#) And (rpWidth = 0#) And (rpHeight = 0#) Then
            Drawing.fillRectToDC m_BackBuffer.getDIBDC, -1, -1, m_Width + 2, m_Height + 2, newBackColor
        Else
            Drawing.fillRectToDC m_BackBuffer.getDIBDC, rpLeft - 1, rpTop - 1, rpWidth + 2, rpHeight + 2, newBackColor
        End If
    End If
            
End Function

'Sync the back buffer to the underlying user control size.  You can call this at any time, because the buffer will only be recreated
' as necessary.
' NOTE: If a custom backcolor has been specified, it won't be applied until the caller requests a copy of the DC.
Private Sub SyncBackBufferSize(Optional ByVal alsoResetDIB As Boolean = False)
    
    'At compile-time, this event will likely be fired; because the back buffer doesn't exist, ignore the request
    If m_BackBuffer Is Nothing Then Exit Sub
    
    'We only want to recreate the buffer when absolutely necessary, so start with a size check.
    Dim bufferWasResized As Boolean
    bufferWasResized = False
    
    If (m_BackBuffer.getDIBWidth <> m_Width) Or (m_BackBuffer.getDIBHeight <> m_Height) Then
        m_BackBuffer.createBlank m_Width, m_Height, 24
        bufferWasResized = True
    End If
    
    'Wipe the DIB as necessary.  Two states require this; caller-requested wipes, and changes to the caption.
    If alsoResetDIB Or bufferWasResized Then
        
        'Erase the back buffer
        ResetBackBuffer
        
        'If a caption is active, paint it now.
        If m_Caption.IsCaptionActive Then
            
            'Notify the caption renderer of our width.  It will auto-fit its font to match.
            m_Caption.setControlSize m_Width, m_Height
            
            'Retrieve the caption's height
            With m_CaptionRect
                .Top = 0
                .Left = 0
                .Right = m_Width
                .Bottom = m_Caption.getCaptionHeight() + FixDPI(6)
            End With
            
            'Paint the caption
            m_Caption.drawCaption m_BackBuffer.getDIBDC, 1, 1
            
        End If
        
        'Inside the designer, I really don't care about painting accurate controls.  It's a waste of time, energy, and code
        ' to cover this case, so simply trace the control outline to provide a bit of orientation, then exit.
        If Not g_IsProgramRunning Then
        
            Dim tmpRect As RECT
            With tmpRect
                .Left = 0
                .Top = 0
                .Right = m_BackBuffer.getDIBWidth
                .Bottom = m_BackBuffer.getDIBHeight
            End With
            
            DrawFocusRect m_BackBuffer.getDIBDC, tmpRect
            
        End If
    
    End If
    
End Sub

'***************************************************************************
' (end internal functions)
'***************************************************************************

'***************************************************************************
' RELAY FUNCTIONS
'***************************************************************************

Private Sub m_FocusDetector_GotFocusReliable()
    RaiseEvent GotFocusAPI
End Sub

Private Sub m_FocusDetector_LostFocusReliable()
    RaiseEvent LostFocusAPI
End Sub

Private Sub m_MouseEvents_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent ClickCustom(Button, Shift, x, y)
End Sub

Private Sub m_MouseEvents_DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent DoubleClickCustom(Button, Shift, x, y)
End Sub

Private Sub m_MouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Ensure that a focus event has been raised, if it wasn't already
    If Not m_FocusDetector.HasFocus Then m_FocusDetector.setFocusManually
    
    RaiseEvent MouseDownCustom(Button, Shift, x, y)
    
End Sub

Private Sub m_MouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideUC = True
    RaiseEvent MouseEnter(Button, Shift, x, y)
End Sub

Private Sub m_MouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideUC = False
    m_MouseEvents.setSystemCursor IDC_DEFAULT
    RaiseEvent MouseLeave(Button, Shift, x, y)
End Sub

Private Sub m_MouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    If Not m_MouseInsideUC Then m_MouseInsideUC = True
    RaiseEvent MouseMoveCustom(Button, Shift, x, y)
End Sub

Private Sub m_MouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    RaiseEvent MouseUpCustom(Button, Shift, x, y, ClickEventAlsoFiring)
End Sub

Private Sub m_MouseEvents_MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RaiseEvent MouseWheelHorizontal(Button, Shift, x, y, scrollAmount)
End Sub

Private Sub m_MouseEvents_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RaiseEvent MouseWheelVertical(Button, Shift, x, y, scrollAmount)
End Sub

Private Sub m_MouseEvents_MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)
    RaiseEvent MouseWheelZoom(Button, Shift, x, y, zoomAmount)
End Sub

Private Sub m_MoveSize_WindowMove(ByVal newLeft As Long, ByVal newTop As Long)
    RaiseEvent WindowMove(newLeft, newTop)
End Sub

Private Sub m_MoveSize_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    
    If (m_Width <> newWidth) Or (m_Height <> newHeight) Then
        
        'Cache the new dimensions
        m_Width = newWidth
        m_Height = newHeight
        
        'Resize the backbuffer to match
        SyncBackBufferSize True
        
        'Let the caller know about the change
        RaiseEvent WindowResize(newWidth, newHeight)
        
        'Request a repaint to match.  Normally we would notify the caller that a layout update is also required,
        ' but since we just raised a resize event (which would trigger the same thing), it's redundant to do it here.
        AskForARepaint False
        
    End If
    
End Sub

'Rather than bother the caller with paint requests, we handle them internally, by flipping the back buffer to the underlying
' control's DC.
Private Sub m_Painter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)
    BitBlt m_hDC, winLeft, winTop, winWidth, winHeight, m_BackBuffer.getDIBDC, winLeft, winTop, vbSrcCopy
End Sub

'***************************************************************************
' (end relay functions)
'***************************************************************************

'***************************************************************************
' INTERACTIVE FUNCTIONS
'***************************************************************************

'Visibility, as tracked by the API.  Note that this will be set immediately before the change is actually applied, so the UC
' can check it to know how the control visibility is *about* to change.
Public Function AmIVisible() As Boolean
    AmIVisible = m_Visibility
End Function

'Want a new tooltip?  No problem; call this function to perform an instant update.
' (IMPORTANT NOTE: per Windows rules, a parentHwnd is required.  We can't request the parent hWnd at creation time,
'                  because the control isn't sited yet, so it must be supplied here.)
' (IMPORTANT NOTE: the tooltip class handles translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal parentHwnd As Long, ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    If m_hWndParent = 0 Then m_hWndParent = parentHwnd
    m_ToolTipManager.setTooltip m_hWnd, m_hWndParent, newTooltip, newTooltipTitle, newTooltipIcon
End Sub

'If the control wants to know if it has focus, it can call this function.  API techniques are used, so even API windows
' will be handled correctly.
Public Function DoIHaveFocus() As Boolean
    DoIHaveFocus = m_FocusDetector.HasFocus
End Function

'Retrieve a copy of the back buffer's DC.  Optionally, you can request that the control wipe the background first.
Public Function GetBackBufferDC(Optional ByVal repaintBackground As Boolean = False, Optional ByVal newBackColor As Long = -1&) As Long
    
    'As a failsafe, make sure the back buffer size is up-to-date
    SyncBackBufferSize
    
    'Make sure the program is running, as we need to pull data from PD's central themer
    If g_IsProgramRunning Then
        
        'If the caller wants the background repainted, apply it now
        If repaintBackground Then
            
            'If a caption is active, we manually wipe only the non-caption area
            If m_Caption.IsCaptionActive Then
                ResetBackBuffer newBackColor, -1#, m_CaptionRect.Bottom, m_Width + 2, m_Height + 2
            Else
                ResetBackBuffer newBackColor
            End If
            
        End If
        
    End If
    
    'With the background successfully filled, allow the caller to paint to the DC at their leisure.
    GetBackBufferDC = m_BackBuffer.getDIBDC
    
End Function

Public Function GetBackBufferWidth() As Long
    GetBackBufferWidth = m_Width
End Function

Public Function GetBackBufferHeight() As Long
    GetBackBufferHeight = m_Height
End Function

Public Function IsMouseInside() As Boolean
    IsMouseInside = m_MouseInsideUC
End Function

Public Sub RequestCursor(Optional ByVal standardCursorType As SystemCursorConstant = IDC_DEFAULT)
    m_MouseEvents.setSystemCursor standardCursorType
End Sub

'Want a paint event fired?  Call this function.  By default, the painter will add a WM_PAINT message to the stack, and wait
' for the window to paint itself naturally.  If this is unacceptable, you can set raiseImmediateDrawEvent to TRUE, which
' forces the paint event to fire immediately.  (Because window invalidation does not occur, clipping may not be handled
' precisely, but that should be okay because the caller knows why a redraw is needed.)
Public Sub RequestRepaint(Optional ByVal raiseImmediateDrawEvent As Boolean = False)
    
    'When running, rely on the painter to paint the control for us
    If g_IsProgramRunning Then
        m_Painter.RequestRepaint raiseImmediateDrawEvent
    Else
        'Inside the IDE, subclassing is disabled, so we must paint the control manually
        If Not m_BackBuffer Is Nothing Then
            CacheWindowDimensions
            SyncBackBufferSize True
            BitBlt m_hDC, 0, 0, m_Width, m_Height, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy
        End If
    End If
    
End Sub

'Whenever the caller user control receives an UpdateAgainstTheme/Language request, it needs to relay that to this support
' class as well.  We'll update things like the tooltip and text caption (if any) to match the new settings.
Public Sub UpdateAgainstThemeAndLanguage()
    
    'Tooltips must be refreshed (in case the active language has changed; tooltips manage translations internally)
    m_ToolTipManager.UpdateAgainstCurrentTheme
    
    'The caption manager will also refresh itself
    If m_Caption.IsCaptionActive Then m_Caption.UpdateAgainstCurrentTheme
    
    'After a theme change, the caller needs to repaint everything from scratch.
    SyncBackBufferSize True
    AskForARepaint True
    
End Sub

'***************************************************************************
' (end interactive functions)
'***************************************************************************

Private Sub Class_Initialize()

    Set m_Subclass = New cSelfSubHookCallback
    m_BackColor = -1
    m_RedrawRequestCount = 0
    
    'Some classes are required for minimal drawing inside the IDE.  As such, we must prep them now, instead of waiting for a
    ' formal UC registration event.
    Set m_BackBuffer = New pdDIB
    Set m_Caption = New pdCaption
    
End Sub

Private Sub Class_Terminate()
    
    'Release all subclassing now.  Note that the order in which we do this is important - don't change it!
    
    'Release custom functionality first, in reverse-order
    If Not (m_Caption Is Nothing) Then Set m_Caption = Nothing
    '(TODO: keyboard input handler goes here)
    If Not (m_MouseEvents Is Nothing) Then Set m_MouseEvents = Nothing
    
    'Release default functionality next, also in reverse-order
    Set m_MoveSize = Nothing
    Set m_FocusDetector = Nothing
    Set m_Painter = Nothing
    Set m_ToolTipManager = Nothing
    
    'As a final step, release this class's internal subclasser
    m_Subclass.ssc_Terminate
    Set m_Subclass = Nothing
    
    'If the user control was registered for internal PD messages, de-register it with PD's central marshaler.
    If m_CustomInternalMessagesActive Then UserControl_Support.RemoveMessageRecipient m_hWnd
    
End Sub

'All messages subclassed by m_Subclass are handled here.
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
    Select Case uMsg
        
        'TODO: wrap WM_WINDOWPOSCHANGING as well, because it covers some visibility change cases that WM_SHOWWINDOW does not
        Case WM_SHOWWINDOW
            
            m_Visibility = CBool(wParam <> 0)
            
            'If we haven't asked the caller for a repaint yet, ask for one now.  (This is the last possible time we can request
            ' one before the control is actually shown.)
            If m_Visibility And (m_RedrawRequestCount = 0) Then AskForARepaint False
            
            'Finally, notify the parent of the pending visibility change
            RaiseEvent VisibilityChange(m_Visibility)
        
        'All other messages are custom requests from the window.  Use them to raise custom message events.
        Case Else
            RaiseEvent CustomMessage(uMsg, wParam, lParam, bHandled)
    
    End Select
    
    bHandled = False

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub
