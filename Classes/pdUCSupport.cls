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
'Copyright 2015-2026 by Tanner Helland
'Created: 25/October/15  (but built from many parts existing earlier)
'Last updated: 24/November/20
'Last update: improve aggressiveness of freeing resources when able
'
'Now that PD provides so many of its own user controls (UCs), we have to repeat a lot of UC management code.
' Painting, custom input handling, subclassing custom window messages - the list goes on for awhile.
'
'To cut down on the amount of repeat work I have to do for each unique control, I've wrapped many generic
' UC capabilities into this class.  It then raises a boatload of events for individual UCs to handle as they
' see fit.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This class raises many events.  Some *must* be responded to.  Others are optional.
Public Event CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef bHandled As Boolean, ByRef lReturn As Long)
Public Event GotFocusAPI()
Public Event LostFocusAPI()
Public Event RepaintRequired(ByVal updateLayoutToo As Boolean)
Public Event VisibilityChange(ByVal newVisibility As Boolean)
Public Event WindowMove(ByVal newLeft As Long, ByVal newTop As Long)
Public Event WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
Public Event WindowResizeFinished(ByVal newWidth As Long, ByVal newHeight As Long)

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
Public Event KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, ByRef markEventHandled As Boolean)
Public Event SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, ByRef newTargetHwnd As Long)

Public Event AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT) As Long

'This class subclasses a *lot* of things.  Some are handled by this class itself; others use dedicated single-purpose
' support classes.
Private Type tagWINDOWPOS
    hWnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    wFlags As Long
End Type
Private Const SWP_HIDEWINDOW As Long = &H80
Private Const SWP_SHOWWINDOW As Long = &H40

Private Declare Function CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptrDst As Long, ByVal ptrSrc As Long, ByVal numOfBytes As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal srcHWnd As Long, ByRef dstRectL As RectL) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Const WM_SHOWWINDOW As Long = &H18&
Private Const WM_WINDOWPOSCHANGING As Long = &H46&

'This class performs its own subclassing, as do a bunch of the other classes (pretty much everything declared WithEvents)
Implements ISubclass

Private WithEvents m_MouseEvents As pdInputMouse
Attribute m_MouseEvents.VB_VarHelpID = -1
Private WithEvents m_KeyEvents As pdInputKeyboard
Attribute m_KeyEvents.VB_VarHelpID = -1
Private WithEvents m_FocusDetector As pdFocusDetector
Attribute m_FocusDetector.VB_VarHelpID = -1
Private WithEvents m_WindowPainter As pdWindowPainter
Attribute m_WindowPainter.VB_VarHelpID = -1
Private WithEvents m_MoveSize As pdWindowSize
Attribute m_MoveSize.VB_VarHelpID = -1
Private m_Caption As pdCaption

'Persistent copy of the associated user control's hWnd; this is arguably the most important piece of information we store,
' as all features are tied to it.
Private m_hWnd As Long

'Size of the underlying usercontrol.  These values are retrieved via API, because VB has DPI-awareness issues.
Private m_Width As Long, m_Height As Long

'If the attached control is capable of receiving focus (e.g. not labels, panels, etc), this will be set to TRUE.
' This is populated by the initial RegisterControl call, and cannot be changed once set.
Private m_ControlCanGetFocus As Boolean

'If the attached control is capable of receiving focus (see m_ControlCanGetFocus, above), we will activate a keyboard
' detector automatically.  Note that some controls also want to receive keypress data (e.g. a spinner wants +/- keys)
' so they can perform custom handling.  When such custom handling is active, we set an additional flag telling us to
' relay key events directly to the child control, instead of handling them ourselves.
Private m_ControlWantsKeyData As Boolean

'If the caller has registered one (or more) custom messages, this will be set to TRUE.  When the parent control is
' unloaded, we need to de-register the hWnd with PD's central message marshaler.
Private m_CustomInternalMessagesActive As Boolean
Private m_CustomMessages As pdDictionary

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
' Same goes for mouse buttons.
Private m_MouseInsideUC As Boolean, m_MouseButtonState As Long
Private m_LastMouseX As Single, m_LastMouseY As Single

'Controls can register themselves as "non-interactive".  This is useful for things like pdSlider, which is really
' just a glorified container for a pdSliderStandalone and pdSpinner combination.  "Non-interactive" controls have their
' cursor automatically managed by this class.
Private m_ControlIsNotInteractive As Boolean

'Visibility is also tracked via the API
Private m_Visibility As Boolean

'This class tries very hard to minimize redraw requests to the parent UC.  As a result, certain sets of conditions can lead to a
' control being shown, without any redraw requests being raised.  To prevent this, we use a failsafe check when the control is
' shown for the first time.  If no redraw requests have been raised, we raise one then, to ensure the control is shown correctly.
Private m_RedrawRequestCount As Long

'If the callers asks us for a resize, chances are that they *don't* want a matching resize event fired, because they're already
' in the middle of some drawing logic.  We use this to note such an event, so we can suppress external resize events.
Private m_ResizeEventsSuspended As Boolean

'If the caller wants the caption rendered to a non-standard position, these values will store their requested position
Private m_CustomCaptionRectActive As Boolean, m_CustomCaptionRect As RECT

'By design, this class can handle all caption rendering for the child.  Some controls, however, need to paint the caption in
' weird ways (e.g. in the middle of other paint operations).  If they require custom painting, this will be set to TRUE.
Private m_CustomCaptionPaintingActive As Boolean

'Some controls want to take advantage of all the nice things this class offers, without managing a persistent back buffer.
' (For example, every pixel of pdCanvas is covered by child controls, so it's wasteful to manage a full backbuffer that
' will never be painted to!)
'This value is set to FALSE by default.
Private m_NoBackBufferMode As Boolean

'This support class also manages tooltips.  Note that translations are handled automatically; the original English tooltips
' are cached to make runtime translations possible.  Tooltips are considered active if the original, English caption string
' has a non-zero length.  Note also that by default, tooltips are raised on _MouseHover events (whose delay is controlled
' by Windows); however, in the AssignTooltip function, callers can request "immediate" tooltip behavior, in which case tips
' will be raised immediately upon mouse entrance.
Private m_TTX As Single, m_TTY As Single
Private m_TTCaptionEn As String, m_TTTitleEn As String
Private m_TTCaption As String, m_TTTitle As String
Private m_TTImmediate As Boolean

'Theme ID at the last UpdateAgainstCurrentTheme request.  If this ID has not changed, we don't need to redraw anything.
' (Note that this ID describes both the current *theme* and the current *language*; a change to either initiates a redraw.)
Private m_LastThemeID As String

'Some controls (e.g. pdCanvasView) may require a lot of time to repaint their contents after something like
' a Form_Resize event.  If they know a lengthy delay is in place, they can ask us to suspend repainting until
' they're ready.
Private m_SuspendAutoRepaint As Boolean

'Similarly, by default, this class is very aggressive about freeing rendering resources whenever
' it can.  However, some specialized control instances (like animation controls) are very
' perf-sensitive, and they can ask us to suspend our most aggressive resource-minimizing behavior.
Private m_HighPerfRequested As Boolean

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum UCSUPPORT_COLOR_LIST
    [_First] = 0
    UCS_Background = 0
    UCS_Caption = 1
    [_Last] = 1
    [_Count] = 2
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'Brush and surface objects are used by pd2D to paint the back buffer
Private m_Surface As pd2DSurface, m_Brush As pd2DBrush


'***************************************************************************
' SETUP AND INITIALIZATION FUNCTIONS
'***************************************************************************

'The FIRST THING a user control must do is call this sub.  This prepares the support class, applies basic subclassing,
' and initializes a bunch of other support modules.
Friend Sub RegisterControl(ByVal hWnd_Control As Long, ByVal controlCanReceiveFocus As Boolean, Optional ByVal noBackBufferRequired As Boolean = False)
    
    m_hWnd = hWnd_Control
    
    If (m_hWnd <> 0) Then
        
        'Subclass the user control immediately
        If PDMain.IsProgramRunning() Then
            UserControls.IncrementPDControlCount
            VBHacks.StartSubclassing m_hWnd, Me
        End If
        
        'Synchronize initial window visiblity
        m_Visibility = (IsWindowVisible(m_hWnd) <> 0)
        
        'Mark whether the control can receive focus; if it can, we'll immediately start tracking Tab keypresses
        m_ControlCanGetFocus = controlCanReceiveFocus
        If m_ControlCanGetFocus Then
            
            Set m_KeyEvents = New pdInputKeyboard
            If PDMain.IsProgramRunning() Then
                m_KeyEvents.CreateKeyboardTracker "pdUCSupport", m_hWnd
                m_KeyEvents.RequestMoreKeys vbKeyReturn, vbKeyTab, vbKeyEscape
            End If
            
            'We also want to subclass some special internal PD messages related to tab-key nav;
            ' these allow controls to specify custom tab and shift+tab targets (for pages with
            ' non-standard control layouts, like toolboxes)
            Me.SubclassCustomMessage WM_PD_TAB_KEY_TARGET, True
            Me.SubclassCustomMessage WM_PD_SHIFT_TAB_KEY_TARGET, True
            
        End If
        
        'Some functionality is supplied "by default".  The caller does not have to request this behavior; they get it for free.
        
        'Visual themes integration
        Set m_Colors = New pdThemeColors
        Dim colorCount As UCSUPPORT_COLOR_LIST: colorCount = [_Count]
        m_Colors.InitializeColorList "GenericPDControl", colorCount
        If (Not PDMain.IsProgramRunning()) Then UpdateColorList
        
        'Flicker-free double-buffered window painting.  For IDE support of windowless controls, we have no choice but to subclass.
        ' (Note that in older versions, we only requested erase events if noBackBufferRequired = TRUE.  That has since changed,
        '  and we now request erase events for all types of user controls.)
        Set m_WindowPainter = New pdWindowPainter
        If PDMain.IsProgramRunning() Then m_WindowPainter.StartPainter m_hWnd, True
        
        'Reliable, API-based Got/LostFocus detection
        Set m_FocusDetector = New pdFocusDetector
        If PDMain.IsProgramRunning() Then m_FocusDetector.StartFocusTracking m_hWnd
        
        'Reliable, API-based, DPI-aware move/size functionality
        Set m_MoveSize = New pdWindowSize
        m_MoveSize.AttachToHWnd m_hWnd, PDMain.IsProgramRunning()
        CacheWindowDimensions
        
        'The vast majority of PD's custom controls are double-buffered to prevent flickering, but some controls do not require this.
        ' For those controls, we avoid creating a backbuffer, which can spare quite a bit of memory.
        m_NoBackBufferMode = noBackBufferRequired
        
        'Initialize the backbuffer and sync it to the current control size
        If (Not m_NoBackBufferMode) Then SyncBackBufferSize True
        
    Else
        PDDebug.LogAction "WARNING!  pdUCSupport.RegisterControl was called with hWnd = 0."
    End If
    
End Sub

'After RegisterControl has been called (above), the caller can request additional support via this function.
' (Because not all user controls require things like captions or custom input events, this functionality is restricted to
'  an "as-required" basis.)
Friend Sub RequestExtraFunctionality(Optional ByVal customMouseEvents As Boolean = False, Optional ByVal customKeyEvents As Boolean = False, Optional ByVal appCommandEvents As Boolean = False, Optional ByVal controlIsInteractive As Boolean = True)
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    If (m_hWnd <> 0) Then
    
        'Initialize special subclassers, as requested
        If customMouseEvents Then
            Set m_MouseEvents = New pdInputMouse
            If PDMain.IsProgramRunning() Then
                m_MouseEvents.AddInputTracker m_hWnd, appCommandEvents
                m_ControlIsNotInteractive = (Not controlIsInteractive)
                If m_ControlIsNotInteractive Then m_MouseEvents.SetCursor_System IDC_ARROW, False
            End If
        End If
        
        If customKeyEvents Then
            
            m_ControlWantsKeyData = True
            
            'The keyboard tracker should already exist, because we now create it by default for all interactive controls
            ' (so we can grab dialog navigation events).
            If (m_KeyEvents Is Nothing) Then
                
                If m_ControlWantsKeyData Then PDDebug.LogAction "WARNING!  Why is a control requesting keyboard data, when the control itself is marked as [can't receive focus]??"
                Set m_KeyEvents = New pdInputKeyboard
                If PDMain.IsProgramRunning() Then m_KeyEvents.CreateKeyboardTracker "pdUCSupport", m_hWnd
            
            End If
            
        End If
        
    Else
        PDDebug.LogAction "WARNING!  You can't request custom functionality prior to calling pdUCSupport.RegisterControl()!"
    End If
    
End Sub

Friend Sub SpecifyRequiredKeys(ParamArray listOfKeys() As Variant)
    
    If (Not m_KeyEvents Is Nothing) And (m_hWnd <> 0) Then
        
        If PDMain.IsProgramRunning() Then
            
            'Iterate through all requested keys, and add them one-by-one to the tracker
            If UBound(listOfKeys) >= LBound(listOfKeys) Then
            
                Dim i As Variant
                For Each i In listOfKeys
                    m_KeyEvents.RequestMoreKeys CLng(i)
                Next i
                
            End If
    
        End If
        
    Else
        PDDebug.LogAction "WARNING!  You can't request keypresses without first calling pdUCSupport.RequestExtraFunctionality()!"
    End If
    
End Sub

'Want to allow some automatic text captioning on your control?  No problem; request it here, and the support class takes care of
' the rest.  (Obviously you also need to supply the caption text, font, and other things, but this function lets the class know
' that caption support is at least a *possibility*.)
Friend Sub RequestCaptionSupport(Optional ByVal supportWordWrap As Boolean = False)
    m_Caption.SetWordWrapSupport supportWordWrap
    m_Caption.SetControlSize m_Width, m_Height
End Sub

Friend Sub SetCaptionWordWrap(ByVal newWordWrapState As Boolean)
    m_Caption.SetWordWrapSupport newWordWrapState
End Sub

'Get/Set functions for various caption-specific features.  Note that you can remove caption support by calling SetCaptionText and
' supplying a blank string.
Friend Function GetCaptionText() As String
    GetCaptionText = m_Caption.GetCaptionEn()
End Function

Friend Function GetCaptionTextTranslated() As String
    GetCaptionTextTranslated = m_Caption.GetCaptionTranslated()
End Function

Friend Sub SetCaptionText(Optional ByRef newCaption As String = vbNullString)
    
    'Pass the string to the pdCaption instance; it handles the actual mess of sizing and rendering text
    If m_Caption.SetCaption(newCaption) Then
        
        'Caption changes may require us to resize the control to fit.
        SyncBackBufferSize True
        
        'Request an immediate redraw (particularly important inside the IDE, so the caller can see the new text)
        AskForARepaint True
        
    End If
    
End Sub

Friend Function GetCaptionFontSize() As Single
    GetCaptionFontSize = m_Caption.GetFontSize
End Function

Friend Sub SetCaptionFontSize(ByVal newSize As Single)
    
    'Changes to font size require redraws
    If m_Caption.SetFontSize(newSize) Then
        SyncBackBufferSize True
        AskForARepaint True
    End If
    
End Sub

Friend Function GetCaptionFontBold() As Boolean
    GetCaptionFontBold = m_Caption.GetFontBold
End Function

Friend Sub SetCaptionFontBold(ByVal newState As Boolean)
    
    'Changes to font style require redraws
    If m_Caption.SetFontBold(newState) Then
        SyncBackBufferSize True
        AskForARepaint True
    End If
    
End Sub

Friend Function GetCaptionFontItalic() As Boolean
    GetCaptionFontItalic = m_Caption.GetFontItalic
End Function

Friend Sub SetCaptionFontItalic(ByVal newState As Boolean)
    
    'Changes to font style require redraws
    If m_Caption.SetFontItalic(newState) Then
        SyncBackBufferSize True
        AskForARepaint True
    End If
    
End Sub

Friend Function GetCaptionFontUnderline() As Boolean
    GetCaptionFontUnderline = m_Caption.GetFontUnderline
End Function

Friend Sub SetCaptionFontUnderline(ByVal newState As Boolean, Optional ByVal suspendRepaint As Boolean = False)
    
    'Changes to font style require redraws
    If m_Caption.SetFontUnderline(newState) Then
        SyncBackBufferSize True
        If (Not suspendRepaint) Then AskForARepaint True
    End If
    
End Sub

Friend Function GetCaptionAlignment() As AlignmentConstants
    GetCaptionAlignment = m_Caption.GetAlignment
End Function

Friend Sub SetCaptionAlignment(ByVal newAlignment As AlignmentConstants)
    
    'Changes to alignment require redraws
    If m_Caption.SetAlignment(newAlignment) Then
        SyncBackBufferSize True
        AskForARepaint True
    End If
    
End Sub

'The caption will be positioned automatically, or the caller can set a custom rect here.
' To reset to default positioning, call with all zeroes.
Friend Sub SetCaptionCustomPosition(Optional ByVal newLeft As Long = 0, Optional ByVal newTop As Long = 0, Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0)
    
    If (newLeft = 0) And (newTop = 0) And (newWidth = 0) And (newHeight = 0) Then
        m_CustomCaptionRectActive = False
    Else
    
        m_CustomCaptionRectActive = True
        
        With m_CustomCaptionRect
            .Left = newLeft
            .Top = newTop
            .Right = newLeft + newWidth
            .Bottom = newTop + newHeight
        End With
        
    End If
    
    'Immediately notify the caption support class of this boundary rect
    If m_CustomCaptionRectActive Then
        m_Caption.SetControlSize m_CustomCaptionRect.Right - m_CustomCaptionRect.Left, m_CustomCaptionRect.Bottom - m_CustomCaptionRect.Top
    Else
        m_Caption.SetControlSize m_Width, m_Height
    End If

End Sub

'By default, this class manages caption painting for you.  If you want to do it manually, pass this function FALSE.
Friend Sub SetCaptionAutomaticPainting(ByVal newState As Boolean)
    m_CustomCaptionPaintingActive = Not newState
End Sub

'If automatic caption painting is turned off, the caller can use this function paint the caption to any location of their choosing.
Friend Sub PaintCaptionManually(Optional ByVal newX As Long = -1, Optional ByVal newY As Long = -1, Optional ByVal customColor As Long = -1, Optional ByVal useEllipses As Boolean = False)
    If m_NoBackBufferMode Then
        PDDebug.LogAction "WARNING!  You cannot paint outside a WM_PAINT cycle if double-buffering was deactivated!"
    Else
        If (newX = -1) Then newX = m_CustomCaptionRect.Left
        If (newY = -1) Then newY = m_CustomCaptionRect.Top
        m_Caption.DrawCaption_Clipped m_BackBuffer.GetDIBDC, newX, newY, m_BackBuffer.GetDIBWidth - newX, m_BackBuffer.GetDIBHeight - newY, customColor, useEllipses
        'NOTE: normally, we free the back buffer DC after painting with it, but we can't do that here as controls typically
        '      call this function inside RedrawBackBuffer
    End If
End Sub

'If automatic caption painting is turned off, the caller can use this function paint the caption to any rect of their choosing.
Friend Sub PaintCaptionManually_Clipped(ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, Optional ByVal customColor As Long = -1, Optional ByVal useEllipses As Boolean = False, Optional ByVal useOriginalFontSize As Boolean = False, Optional ByVal centerVertically As Boolean = False)
    If m_NoBackBufferMode Then
        PDDebug.LogAction "WARNING!  You cannot paint outside a WM_PAINT cycle if double-buffering was deactivated!"
    Else
        m_Caption.DrawCaption_Clipped m_BackBuffer.GetDIBDC, dstX, dstY, dstWidth, dstHeight, customColor, useEllipses, useOriginalFontSize, centerVertically
        'NOTE: normally, we free the back buffer DC after painting with it, but we can't do that here as controls typically
        '      call this function inside RedrawBackBuffer
    End If
End Sub

Friend Sub SetCaptionCustomColor(ByVal newCaptionColor As Long)
    m_Caption.SetCaptionColor newCaptionColor
End Sub

Friend Function GetCaptionBottom() As Long
    GetCaptionBottom = m_CaptionRect.Bottom
End Function

Friend Function GetCaptionHeight(Optional ByVal useAutoFit As Boolean = True) As Long
    If useAutoFit Then
        GetCaptionHeight = m_Caption.GetCaptionHeight
    Else
        GetCaptionHeight = m_Caption.GetCaptionHeight_NoFit
    End If
End Function

Friend Function GetCaptionWidth(Optional ByVal useAutoFit As Boolean = True) As Long
    If useAutoFit Then
        GetCaptionWidth = m_Caption.GetCaptionWidth
    Else
        GetCaptionWidth = m_Caption.GetCaptionWidth_NoFit
    End If
End Function

Friend Function IsCaptionActive() As Boolean
    If m_Caption Is Nothing Then
        IsCaptionActive = False
    Else
        IsCaptionActive = m_Caption.IsCaptionActive()
    End If
End Function

'Manually suspend repaint events.  IMPORTANTLY, you must *re-enable* repaints after calling this function,
' or no painting will occur.
Friend Sub SuspendAutoRepaintBehavior(ByVal newSuspendState As Boolean)
    m_SuspendAutoRepaint = newSuspendState
End Sub

'Ask us to *not* be as aggressive about freeing resources after rendering.  (This should be used sparingly,
' e.g. animation controls may need it because they need ms-accurate rendering, but few other controls
' require this, so they should err on the side of minimizing resource consumption.)
Friend Sub RequestHighPerformanceRendering(ByVal newState As Boolean)
    m_HighPerfRequested = newState
End Sub

'Want to turn on/off high-resolution mouse input?  Request changes here.  (By default, high-resolution
' mouse input is *not* enabled for most controls.)
Friend Sub RequestHighResMouseInput(ByVal newState As Boolean)
    If (Not m_MouseEvents Is Nothing) Then m_MouseEvents.SetHighResolutionTrackingMode newState
End Sub

'Want to turn off auto-dropping of delayed mouse messages?  Request changes here.  (By default, auto-drop
' is always enabled.)
Friend Sub RequestAutoDropMouseMessages(ByVal newState As Boolean)
    If (Not m_MouseEvents Is Nothing) Then m_MouseEvents.SetAutoDropOfDelayedEvents newState
End Sub

'INSIDE A KEY EVENT (the only place this is valid), you can call this function to retrieve
' the unprocessed key code and scan code from the keyboardproc.
Friend Sub GetLastKeyAndScanCode(ByRef dstKeyCode As Long, ByRef dstScanCode As Long)
    If (Not m_KeyEvents Is Nothing) Then m_KeyEvents.GetLastKeyAndScanCode dstKeyCode, dstScanCode
End Sub

'Return the number of "pending" mouse move events.  Only relevant if high-resolution input tracking has been manually activated.
Friend Function GetNumMouseEventsPending() As Long
    If (Not m_MouseEvents Is Nothing) Then GetNumMouseEventsPending = m_MouseEvents.GetNumMouseMovePointsRemaining()
End Function

Friend Function GetNextMouseMovePoint(ByRef dstMMP As MOUSEMOVEPOINT) As Boolean
    If (Not m_MouseEvents Is Nothing) Then GetNextMouseMovePoint = m_MouseEvents.GetNextMouseMovePoint(dstMMP)
End Function

'Because this support class manages a back buffer for its parent control, it sometimes needs to raise redraw requests.  To prevent
' an excessive amount of back-and-forth requests, the caller needs to explicitly specify a custom backcolor if it wants one.
' NOTE: call this function without supplying a newColor value to reset to the default program-wide window backcolor.
' ANOTHER NOTE: this function won't raise a redraw request, by design.
Friend Sub SetCustomBackcolor(Optional ByVal newBackColor As Long = -1&)
    m_BackColor = newBackColor
End Sub

'Want to receive custom window messages?  No problem.  Register them here, then watch for
' the CustomMessage event.  Note, however, that if the message is specific to PD, the caller
' needs to let us know, as we will need to manually handle message marshaling.
Friend Sub SubclassCustomMessage(ByVal wMsg As Long, Optional ByVal msgIsInternalToPD As Boolean = False)
    
    If PDMain.IsProgramRunning() Then
        
        'To ensure messages are routed correctly, we store all registered message IDs locally
        If (m_CustomMessages Is Nothing) Then Set m_CustomMessages = New pdDictionary
        m_CustomMessages.AddEntry wMsg, msgIsInternalToPD
        
        'If the message is internal to PD, register the hWnd+message pair with PD's central marshaler
        If msgIsInternalToPD Then
            UserControls.AddMessageRecipient m_hWnd, wMsg
            m_CustomInternalMessagesActive = True
        End If
        
    End If
    
End Sub

Friend Function IsResizeActive() As Boolean
    IsResizeActive = m_MoveSize.IsResizeActive()
End Function

Friend Sub NotifyMouseDragResize_Start()
    m_MouseEvents.SetCaptureOverride True
End Sub

Friend Sub NotifyMouseDragResize_End()
    m_MouseEvents.SetCaptureOverride False
    m_MouseEvents.ResetTracking
    Me.RequestCursor IDC_DEFAULT
End Sub

'***************************************************************************
' (end setup and initialization functions)
'***************************************************************************

'***************************************************************************
' INTERNAL FUNCTIONS
'***************************************************************************

'Sometimes, this control detects changes that require a repaint (like a resize event).  To avoid unnecessary paint requests, we try to
' suppress all requests until the control is actually visible, hence the use of this wrapper instead of directly raising the event.
Private Sub AskForARepaint(ByVal askForLayoutUpdateToo As Boolean, Optional ByVal forceEventRegardless As Boolean = False)
    
    'If the control is visible, raise the event immediately.  Otherwise, suppress it, and update the suppressed draw count.
    If (m_Visibility Or forceEventRegardless) Then
        RaiseEvent RepaintRequired(askForLayoutUpdateToo)
        m_RedrawRequestCount = 0
    Else
        m_RedrawRequestCount = m_RedrawRequestCount + 1
    End If
    
End Sub

'Cache the current dimensions of the underlying UC.  Returns TRUE if dimensions were retrieved successfully.
Private Function CacheWindowDimensions() As Boolean
    
    If ((Not m_MoveSize Is Nothing) And PDMain.IsProgramRunning()) Then
        m_Width = m_MoveSize.GetWidth
        m_Height = m_MoveSize.GetHeight
        CacheWindowDimensions = True
    Else
        CacheWindowDimensions = False
    End If
    
End Function

'The naming of this function is potentially confusing, but it restores the back buffer to the
' currently specified back color.  It does *not* resize the DIB; it just fills it with a single
' uniform color.  Optionally, a specific erase rect can be supplied.  (If coordinates are *not*
' supplied, the entire buffer will be wiped.)
Private Sub ResetBackBuffer(Optional ByVal newBackColor As Long = -1&, Optional ByVal rpLeft As Single = 0!, Optional ByVal rpTop As Single = 0!, Optional ByVal rpWidth As Single = 0!, Optional ByVal rpHeight As Single = 0!)
    
    'If the caller specifies a background color, we want to use it.  If they don't, grab the default window background from
    ' PD's central themer.
    If (newBackColor = -1) Then
        If (m_BackColor = -1) Then
            newBackColor = m_Colors.RetrieveColor(UCS_Background)
        Else
            newBackColor = m_BackColor
        End If
    End If
    
    'There are a couple different ways to apply the repaint.  If the caller specifies a fill rect, we obey it, but if they don't,
    ' we want to wipe the entire buffer.  As a convenience to the caller, we also apply the backcolor inside the IDE.
    If (Not m_NoBackBufferMode) Then
    
        If PDMain.IsProgramRunning() And (m_BackBuffer.GetDIBDC <> 0) Then
            
            'Use GDI+ to apply the fill
            If (m_Surface Is Nothing) Then
                Set m_Surface = New pd2DSurface
                m_Surface.SetSurfaceAntialiasing P2_AA_None
                m_Surface.SetSurfacePixelOffset P2_PO_Normal
            End If
            m_Surface.WrapSurfaceAroundDC m_BackBuffer.GetDIBDC
            
            If (m_Brush Is Nothing) Then Set m_Brush = New pd2DBrush
            m_Brush.SetBrushColor newBackColor
            
            If (rpLeft = 0!) And (rpTop = 0!) And (rpWidth = 0!) And (rpHeight = 0!) Then
                PD2D.FillRectangleI m_Surface, m_Brush, 0, 0, m_Width, m_Height
            Else
                PD2D.FillRectangleF m_Surface, m_Brush, rpLeft, rpTop, rpWidth, rpHeight
            End If
            
            m_Surface.ReleaseSurface
            
            'Normally, brushes self-release, but because this class outlives almost everything else
            ' in PD, an instance of it may still exist after GDI+ is shutdown.  To prevent crashes,
            ' we need to release this brush manually when we're done with it.
            m_Brush.ReleaseBrush
            
        Else
            If (rpLeft = 0!) And (rpTop = 0!) And (rpWidth = 0!) And (rpHeight = 0!) Then
                GDI.FillRectToDC m_BackBuffer.GetDIBDC, -1, -1, m_Width + 2, m_Height + 2, newBackColor
            Else
                GDI.FillRectToDC m_BackBuffer.GetDIBDC, rpLeft - 1, rpTop - 1, rpWidth + 2, rpHeight + 2, newBackColor
            End If
        End If
        
        m_BackBuffer.FreeFromDC
        
    End If
            
End Sub

'Sync the back buffer to the underlying user control size.  You can call this at any time, because the buffer will only be recreated
' as necessary.
' NOTE: If a custom backcolor has been specified, it won't be applied until the caller requests a copy of the DC.
Private Sub SyncBackBufferSize(Optional ByVal alsoResetDIB As Boolean = False)
    
    'At compile-time, this event will likely be fired; if the back buffer doesn't exist, ignore this request entirely
    If (m_BackBuffer Is Nothing) Or m_NoBackBufferMode Then Exit Sub
    
    'We only want to recreate the buffer when absolutely necessary, so start with a size check.
    Dim bufferWasResized As Boolean
    bufferWasResized = False
    
    If (m_BackBuffer.GetDIBWidth <> m_Width) Or (m_BackBuffer.GetDIBHeight <> m_Height) Then
        m_BackBuffer.CreateBlank m_Width, m_Height, 32, , 255
        bufferWasResized = True
    End If
    
    'Wipe the DIB as necessary.  Two states require this; caller-requested wipes, and changes to the caption.
    If alsoResetDIB Or bufferWasResized Then
        
        'Erase the back buffer
        ResetBackBuffer
        
        'If a caption is active, paint it now.
        If m_Caption.IsCaptionActive Then
            
            If m_CustomCaptionRectActive Then
                
                m_Caption.SetControlSize m_CustomCaptionRect.Right - m_CustomCaptionRect.Left, m_CustomCaptionRect.Bottom - m_CustomCaptionRect.Top
                
                With m_CaptionRect
                    .Top = m_CustomCaptionRect.Top
                    .Left = m_CustomCaptionRect.Left
                    .Right = m_CustomCaptionRect.Right
                    .Bottom = m_CustomCaptionRect.Bottom
                End With
                
            Else
            
                'Notify the caption renderer of our width.  It will auto-fit its font to match.
                m_Caption.SetControlSize m_Width, m_Height
                
                'Make a copy of the caption's position
                With m_CaptionRect
                    .Top = 1
                    .Left = 1
                    .Right = m_Width
                    .Bottom = m_Caption.GetCaptionHeight() + Interface.FixDPI(6)
                End With
                
            End If
            
            'Paint the caption, as required
            If (Not m_CustomCaptionPaintingActive) Then
                m_Caption.DrawCaption_Clipped m_BackBuffer.GetDIBDC, m_CaptionRect.Left, m_CaptionRect.Top, m_CaptionRect.Right - m_CaptionRect.Left, m_CaptionRect.Bottom - m_CaptionRect.Top
            End If
            
        End If
        
        'Inside the designer, I really don't care about painting accurate controls.  It's a waste of time, energy, and code
        ' to cover this case, so simply trace the control outline to provide a bit of orientation, then exit.
        If (Not PDMain.IsProgramRunning()) Then
        
            Dim tmpRect As RECT
            With tmpRect
                .Left = 0
                .Top = 0
                .Right = m_BackBuffer.GetDIBWidth
                .Bottom = m_BackBuffer.GetDIBHeight
            End With
            
            DrawFocusRect m_BackBuffer.GetDIBDC, tmpRect
            
        End If
        
        'Keep GDI object count low
        m_BackBuffer.FreeFromDC
    
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
    
    m_MouseButtonState = 0
    If m_MouseInsideUC Then TriggerMouseLeave
    RaiseEvent LostFocusAPI
    
    'After losing focus, we typically don't expect redraws - so suspend this DIB for now
    m_BackBuffer.SuspendDIB
    
End Sub

Private Sub m_KeyEvents_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    'If the control has specifically requested key data, pass all keypresses down the chain
    If m_ControlWantsKeyData Then
        RaiseEvent KeyDownCustom(Shift, vkCode, markEventHandled)
    Else
        markEventHandled = False
    End If
    
    'Handle navigation keys separately (unless the keypress was already eaten by the attached control)
    If (Not markEventHandled) And m_ControlCanGetFocus Then
        If (vkCode = pdnk_Enter) Or (vkCode = pdnk_Escape) Or (vkCode = pdnk_Space) Or (vkCode = pdnk_Tab) Then RaiseEvent KeyDownSystem(Shift, vkCode, markEventHandled)
    End If
    
End Sub

Private Sub m_KeyEvents_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    'If the control has specifically requested key data, pass all keypresses down the chain
    If m_ControlWantsKeyData Then RaiseEvent KeyUpCustom(Shift, vkCode, markEventHandled)
    
End Sub

Private Sub m_MouseEvents_AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent AppCommand(cmdID, Shift, x, y)
End Sub

Private Sub m_MouseEvents_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent ClickCustom(Button, Shift, x, y)
End Sub

Private Sub m_MouseEvents_DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    RaiseEvent DoubleClickCustom(Button, Shift, x, y)
End Sub

Private Sub m_MouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'Track button state persistently, so the control can easily query it outside of events
    m_MouseButtonState = m_MouseButtonState Or Button
    
    'On left-clicks, Ensure that a focus event has been raised, if it wasn't already
    If (Not m_FocusDetector.HasFocus) And ((Button And pdLeftButton) = pdLeftButton) Then m_FocusDetector.SetFocusManually
    RaiseEvent MouseDownCustom(Button, Shift, x, y, timeStamp)
    
End Sub

Private Sub m_MouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    TriggerMouseEnter Button, Shift, x, y
End Sub

Private Sub TriggerMouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    m_MouseInsideUC = True
    HideTooltip True
    
    m_LastMouseX = x
    m_LastMouseY = y
    
    'Non-interactive controls get automatic cursor management from us
    If m_ControlIsNotInteractive Then m_MouseEvents.SetCursor_System IDC_ARROW
    
    'If the caller wants immediate tooltips, raise them now
    If m_TTImmediate Then RaiseTooltip x, y
    
    'Notify our parent
    RaiseEvent MouseEnter(Button, Shift, x, y)
    
End Sub

'On hover, request a tooltip (if one exists)
Private Sub m_MouseEvents_MouseHover(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    If (Not m_TTImmediate) Then RaiseTooltip x, y
End Sub

Private Sub m_MouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    TriggerMouseLeave Button, Shift, x, y
End Sub

Private Sub TriggerMouseLeave(Optional ByVal Button As PDMouseButtonConstants = pdLeftButton, Optional ByVal Shift As ShiftConstants = 0&, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0)
    m_MouseInsideUC = False
    m_MouseEvents.SetCursor_System IDC_DEFAULT
    HideTooltip False
    RaiseEvent MouseLeave(Button, Shift, x, y)
End Sub

Private Sub m_MouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    m_LastMouseX = x
    m_LastMouseY = y
    
    If (Not m_MouseInsideUC) Then
        TriggerMouseEnter Button, Shift, x, y
    Else
        RaiseEvent MouseMoveCustom(Button, Shift, x, y, timeStamp)
    End If
    
End Sub

Private Sub m_MouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    m_MouseButtonState = m_MouseButtonState And (Not Button)
    RaiseEvent MouseUpCustom(Button, Shift, x, y, clickEventAlsoFiring, timeStamp)
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
    
    'Let the caller know about the change, but note that this behavior can be deliberately suspended.
    ' (We typically suspend events when the caller requests a resize event - they know what's happening,
    '  so they don't need to be double- or triple-notified.)
    If (Not m_ResizeEventsSuspended) Then RaiseEvent WindowMove(newLeft, newTop)
    
End Sub

Private Sub m_MoveSize_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    
    If (m_Width <> newWidth) Or (m_Height <> newHeight) Then
        
        'Cache the new dimensions
        m_Width = newWidth
        m_Height = newHeight
        
        'Resize the backbuffer to match
        SyncBackBufferSize True
        
        'Let the caller know about the change
        If (Not m_ResizeEventsSuspended) Then
            
            'Some controls may respond to the window resize event in special ways; most do not.  This is primarily
            ' provided as an "FYI" event, and it *should not* be used to initiate redraws (as we're going to raise
            ' an event for that specific purpose momentarily).
            RaiseEvent WindowResize(newWidth, newHeight)
        
            'Request a repaint, and notify the caller that a layout update is also required.
            AskForARepaint True
            
        End If
        
        'If the mouse is *not* over the control, reset the last-known mouse (x, y).  These values are used to calculate
        ' tooltip positioning, and for keyboard-centric controls (like edit boxes), we want the tooltip to default to
        ' displaying immediately below the control.
        If (Not m_MouseInsideUC) Then
            m_LastMouseX = newWidth \ 2
            m_LastMouseY = newHeight
        End If
        
    End If
    
End Sub

'Convenience wrapper; some windows prefer to handle intensive resize events *after* the resize
' has completed (rather than interactively)
Private Sub m_MoveSize_WindowResizeFinal(ByVal newWidth As Long, ByVal newHeight As Long)
    RaiseEvent WindowResizeFinished(newWidth, newHeight)
End Sub

'Erase events must be responded to, regardless of paint mode.  (Previously, they were only handled if m_NoBackBufferMode = TRUE,
' but because the events may be raised post-dialog-shown, we must handle them regardless to avoid flicker.)
Private Sub m_WindowPainter_EraseBkgnd()
    
    If m_SuspendAutoRepaint Then Exit Sub
    
    Dim targetDC As Long
    targetDC = m_WindowPainter.GetPaintStructDC
    
    'If we are currently using "no backbuffer" mode, just paint our backcolor to the DC
    If m_NoBackBufferMode Then
    
        Dim finalBackColor As Long
        If (m_BackColor = -1) Then
            finalBackColor = m_Colors.RetrieveColor(UCS_Background)
        Else
            finalBackColor = m_BackColor
        End If
        
        GDI.FillRectToDC targetDC, 0, 0, Me.GetControlWidth + 1, Me.GetControlHeight + 1, finalBackColor
    
    'If backbuffering is active (as it is with most PD controls), copy the relevant portion of the buffer into place.
    Else
        GDI.BitBltWrapper targetDC, 0, 0, Me.GetControlWidth + 1, Me.GetControlHeight + 1, m_BackBuffer.GetDIBDC, 0, 0, vbSrcCopy
        If (Not m_HighPerfRequested) Then m_BackBuffer.SuspendDIB
    End If
    
End Sub

'Rather than bother the caller with paint requests, we handle them internally, by flipping the back buffer to the underlying
' control's DC.
Private Sub m_WindowPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)
    
    If m_SuspendAutoRepaint Then Exit Sub
    
    Dim targetDC As Long
    targetDC = m_WindowPainter.GetPaintStructDC
    
    'If we are currently using "no backbuffer" mode, just paint our backcolor to the DC
    If m_NoBackBufferMode Then
    
        Dim finalBackColor As Long
        If (m_BackColor = -1) Then
            finalBackColor = m_Colors.RetrieveColor(UCS_Background)
        Else
            finalBackColor = m_BackColor
        End If
        
        GDI.FillRectToDC targetDC, winLeft - 1, winTop - 1, winWidth + 2, winHeight + 2, finalBackColor
    
    'If backbuffering is active (as it is with most PD controls), copy the relevant portion of the buffer into place.
    Else
        GDI.BitBltWrapper targetDC, winLeft, winTop, winWidth, winHeight, m_BackBuffer.GetDIBDC, winLeft, winTop, vbSrcCopy
        If (Not m_HighPerfRequested) Then m_BackBuffer.SuspendDIB
    End If
    
End Sub

'***************************************************************************
' (end relay functions)
'***************************************************************************

'***************************************************************************
' INTERACTIVE FUNCTIONS
'***************************************************************************

'Visibility, as tracked by the API.  Note that this will be set immediately before the change is actually applied, so the UC
' can check it to know how the control visibility is *about* to change.
Friend Function AmIVisible() As Boolean
    AmIVisible = m_Visibility
End Function

Friend Function AmIEnabled() As Boolean
    AmIEnabled = (IsWindowEnabled(m_hWnd) = 1)
End Function

'Want a new tooltip?  No problem; call this function to perform an instant update.
' (IMPORTANT NOTE: per Windows rules, a parentHwnd is required.  We can't request the parent hWnd at creation time,
'                  because the control isn't sited yet, so it must be supplied here.)
' (IMPORTANT NOTE: the tooltip class handles translations automatically.  Always pass the original English text!)
Friend Sub AssignTooltip(ByVal parentHwnd As Long, ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False, Optional ByVal raiseImmediatelyRegardlessOfMouse As Boolean = False)
    
    If PDMain.IsProgramRunning() Then
    
        'Cache the original English versions of this tooltip
        m_TTCaptionEn = newTooltip
        m_TTTitleEn = newTooltipTitle
        
        'If translations are active, apply a translation now
        If g_Language.TranslationActive() Then
            m_TTCaption = g_Language.TranslateMessage(m_TTCaptionEn)
            m_TTTitle = g_Language.TranslateMessage(m_TTTitleEn)
        Else
            m_TTCaption = m_TTCaptionEn
            m_TTTitle = m_TTTitleEn
        End If
        
        m_TTImmediate = raiseTipsImmediately
        
        'If the tooltip is empty, release any existing tooltips
        If (LenB(m_TTCaption) = 0) Then
            HideTooltip True
        Else
            'If a tooltip is already active, or if the caller wants immediate tooltip display and the mouse
            ' is inside this control, update the existing tooltip immediately.
            If UserControls.IsTooltipActive(m_hWnd) Or (m_TTImmediate And m_MouseInsideUC) Or (m_TTImmediate And raiseImmediatelyRegardlessOfMouse) Then RaiseTooltip m_LastMouseX, m_LastMouseY
        End If
        
    End If
    
End Sub

Private Sub RaiseTooltip(Optional ByVal mouseHoverX As Single = -1!, Optional ByVal mouseHoverY As Single = -1!)
    
    If ((LenB(m_TTCaption) <> 0) And Me.AmIVisible And Me.AmIEnabled) Then
        
        'Cache the hovered (x, y) values; we'll re-use these if we need to refresh the tooltip
        If (mouseHoverX <> -1) Then
            m_TTX = mouseHoverX
            m_TTY = mouseHoverY
        End If
        
        'Note that we also pass the current control rect; this is used for optimal tooltip positioning.
        ' (This rect must also be in *screen coordinates*, FYI.)
        Dim curControlRect As RectL
        GetWindowRect m_hWnd, curControlRect
        
        'Request a tooltip, while converting the (x, y) coordinates to ratios; this simplifies positioning
        UserControls.ShowUCTooltip m_hWnd, curControlRect, m_TTX, m_TTY, m_TTCaption, m_TTTitle
        
    End If
    
End Sub

Private Sub HideTooltip(Optional ByVal forciblyRelease As Boolean = True)
    If forciblyRelease Then
        UserControls.HideUCTooltip
    Else
        If (LenB(m_TTCaption) <> 0) Then UserControls.HideUCTooltip
    End If
End Sub

'If the control wants to know if it has focus, it can call this function.  API techniques are used, so even API windows
' will be handled correctly.
Friend Function DoIHaveFocus() As Boolean
    DoIHaveFocus = m_FocusDetector.HasFocus
End Function

'Relay mouse capture requests to our internal mouse handler
Friend Sub SetCaptureOverride(ByVal newState As Boolean)
    If (Not m_MouseEvents Is Nothing) Then m_MouseEvents.SetCaptureOverride newState
End Sub

'Retrieve a copy of the back buffer's DC.  Optionally, you can request that the control wipe the background first.
Friend Function GetBackBufferDC(Optional ByVal repaintBackground As Boolean = False, Optional ByVal newBackColor As Long = -1&) As Long
    
    If m_NoBackBufferMode Then
        PDDebug.LogAction "WARNING!  You cannot request a back buffer DC when backbuffering has been disabled!"
        GetBackBufferDC = 0
        Exit Function
    End If
    
    'As a failsafe, make sure the back buffer size is up-to-date
    SyncBackBufferSize
    
    'Make sure the program is running, as we need to pull data from PD's central themer
    If PDMain.IsProgramRunning() Then
        
        'If the caller wants the background repainted, apply it now
        If repaintBackground Then
            
            'If a caption is active, we manually wipe only the non-caption area
            If m_Caption.IsCaptionActive And (Not m_CustomCaptionPaintingActive) Then
                ResetBackBuffer newBackColor, -1#, m_CaptionRect.Bottom, m_Width + 2, m_Height + 2
            Else
                ResetBackBuffer newBackColor
            End If
            
        End If
        
    Else
        If repaintBackground And (newBackColor <> -1) Then ResetBackBuffer newBackColor
    End If
    
    'With the background successfully filled, allow the caller to paint to the DC at their leisure.
    GetBackBufferDC = m_BackBuffer.GetDIBDC
    
End Function

Friend Function GetBackBufferWidth() As Long
    GetBackBufferWidth = m_Width
End Function

Friend Function GetBackBufferHeight() As Long
    GetBackBufferHeight = m_Height
End Function

Friend Sub GetControlRect(ByRef dstRect As RectL)
    With dstRect
        .Left = m_MoveSize.GetLeft
        .Top = m_MoveSize.GetTop
        .Right = .Left + m_MoveSize.GetWidth
        .Bottom = .Top + m_MoveSize.GetHeight
    End With
End Sub

Friend Function GetControlLeft() As Long
    GetControlLeft = m_MoveSize.GetLeft
End Function

Friend Function GetControlTop() As Long
    GetControlTop = m_MoveSize.GetTop
End Function

Friend Function GetControlWidth() As Long
    If PDMain.IsProgramRunning() Then GetControlWidth = m_MoveSize.GetWidth Else GetControlWidth = m_Width
End Function

Friend Function GetControlHeight() As Long
    If PDMain.IsProgramRunning() Then GetControlHeight = m_MoveSize.GetHeight Else GetControlHeight = m_Height
End Function

Friend Function IsKeyDown(ByVal vkCode As Long) As Boolean
    IsKeyDown = OS.IsVirtualKeyDown(vkCode)
End Function

Friend Function IsMouseInside() As Boolean
    IsMouseInside = m_MouseInsideUC
End Function

Friend Function IsMouseButtonDown(ByVal Button As PDMouseButtonConstants) As Boolean
    IsMouseButtonDown = ((Button And m_MouseButtonState) <> 0)
End Function

Friend Sub RequestCursor(Optional ByVal standardCursorType As SystemCursorConstant = IDC_DEFAULT)
    If (Not m_MouseEvents Is Nothing) Then m_MouseEvents.SetCursor_System standardCursorType, True
End Sub

Friend Sub RequestCursor_Resource(ByVal pngResourceName As String, Optional ByVal cursorHotspotX As Long = 0, Optional ByVal cursorHotspotY As Long = 0)
    If (Not m_MouseEvents Is Nothing) Then m_MouseEvents.SetCursor_Resource pngResourceName, cursorHotspotX, cursorHotspotY
End Sub

Friend Sub RequestMousePosition(ByVal mouseX As Double, ByVal mouseY As Double)
    If (Not m_MouseEvents Is Nothing) Then m_MouseEvents.MoveCursorToNewPosition mouseX, mouseY
End Sub

'Want a new control size?  No problem.  Request a resize here.  Missing width/height parameters will be set to match the
' current control width/height.
Friend Sub RequestNewSize(Optional ByVal newWidth As Long = 0&, Optional ByVal newHeight As Long = 0&, Optional ByVal alsoNotifyMeViaEvent As Boolean = False)
    
    'Fill in missing parameters, if any
    If (newWidth = 0) Then newWidth = m_MoveSize.GetWidth
    If (newHeight = 0) Then newHeight = m_MoveSize.GetHeight
    
    m_ResizeEventsSuspended = Not alsoNotifyMeViaEvent
    If (newWidth <> m_MoveSize.GetWidth) Or (newHeight <> m_MoveSize.GetHeight) Then
        m_MoveSize.SetSize newWidth, newHeight, alsoNotifyMeViaEvent
    End If
    m_ResizeEventsSuspended = False
    
End Sub

'Want a new control position?  No problem.  Request a move here.  Missing left/top parameters will be set to match the
' current control width/height.
Friend Sub RequestNewPosition(Optional ByVal newLeft As Long = -100, Optional ByVal newTop As Long = -100, Optional ByVal alsoNotifyMeViaEvent As Boolean = False)
    
    'Fill in missing parameters
    If (newLeft = -100) Then newLeft = m_MoveSize.GetLeft
    If (newTop = -100) Then newTop = m_MoveSize.GetTop
    
    m_ResizeEventsSuspended = Not alsoNotifyMeViaEvent
    If (newLeft <> m_MoveSize.GetLeft) Or (newTop <> m_MoveSize.GetTop) Then m_MoveSize.SetPosition newLeft, newTop
    m_ResizeEventsSuspended = False
    
End Sub

Friend Sub RequestFullMove(Optional ByVal newLeft As Long = -100, Optional ByVal newTop As Long = -100, Optional ByVal newWidth As Long = 0, Optional ByVal newHeight As Long = 0, Optional ByVal alsoNotifyMeViaEvent As Boolean = False)
    
    'Fill in missing parameters
    If (newLeft = -100) Then newLeft = m_MoveSize.GetLeft
    If (newTop = -100) Then newTop = m_MoveSize.GetTop
    If (newWidth = 0) Then newWidth = m_MoveSize.GetWidth
    If (newHeight = 0) Then newHeight = m_MoveSize.GetHeight
    
    m_ResizeEventsSuspended = Not alsoNotifyMeViaEvent
    m_MoveSize.SetSizeAndPosition newLeft, newTop, newWidth, newHeight
    m_ResizeEventsSuspended = False
    
End Sub

'Some controls may want to color-manage certain regions of their back buffer DIB.  By design, this user control support
' class does not expose the DIB directly (which allows for certain caching mechanisms that are otherwise very complicated).
' As such, if our owner wants something color-managed, they need to notify us, and we will handle it for them.
'
'To avoid the unwanted management of UI elements, a target rect needs to be explicitly specified.  *Only* the contents of
' this rect will be color-managed.  Everything else will not.  Please ensure that no colored UI elements are included
' in the managed rect, as the will look out-of-place next to the rest of the UI (which is explicitly *not* color-managed
' for performance reasons).
'
'Also, note that requests are not persistent, by design.  The caller must re-request management any time the underlying
' contents change (as this class has no way of being auto-notified).
Friend Sub RequestBufferColorManagement(ByVal ptrToSrcRectF As Long)
    Dim tmpRectF As RectF
    If (ptrToSrcRectF <> 0) Then
        CopyMemoryStrict VarPtr(tmpRectF), ptrToSrcRectF, 16&
    Else
        With tmpRectF
            .Left = 0!
            .Top = 0!
            .Width = Me.GetBackBufferWidth()
            .Height = Me.GetBackBufferHeight()
        End With
    End If
    ColorManagement.ApplyDisplayColorManagement_RectF m_BackBuffer, tmpRectF, , False
End Sub

'Want a paint event fired?  Call this function.  By default, the painter will add a WM_PAINT message to the stack, and wait
' for the window to paint itself naturally.  If this is unacceptable, you can set raiseImmediateDrawEvent to TRUE, which
' forces the paint event to fire immediately.  (Because window invalidation does not occur, clipping may not be handled
' precisely, but that should be okay because the caller knows why a redraw is needed.)
Friend Sub RequestRepaint(Optional ByVal raiseImmediateDrawEvent As Boolean = False)
    
    'When running, rely on the painter to paint the control for us
    If PDMain.IsProgramRunning() Then
        
        'In the event that a paint event is not called right away (because the control is invisible,
        ' or some other reason), free the back buffer DIB from its DC to free up a bit of resources.
        ' If a paint event *is* called right away, this has no ill effects, because the DC is
        ' auto-generated as needed.
        m_BackBuffer.FreeFromDC
        m_WindowPainter.RequestRepaint raiseImmediateDrawEvent
        
    Else
    
        'Inside the IDE, subclassing is disabled, so we must paint the control manually
        If (Not m_BackBuffer Is Nothing) And (Not m_NoBackBufferMode) Then
            
            CacheWindowDimensions
            If (Not m_CustomCaptionPaintingActive) Then SyncBackBufferSize True
            
            Dim targetDC As Long
            targetDC = m_WindowPainter.GetPaintStructDC
            GDI.BitBltWrapper targetDC, 0, 0, m_Width, m_Height, m_BackBuffer.GetDIBDC, 0, 0, vbSrcCopy
            
        End If
    End If
    
End Sub

'To support minimal rendering inside the IDE, this class must be notified of IDE resize events (as we don't want to subclass
' inside the designer)
Friend Sub NotifyIDEResize(ByVal newWidth As Long, ByVal newHeight As Long)
    m_Width = newWidth \ 15
    m_Height = newHeight \ 15
    RaiseEvent RepaintRequired(True)
    Me.RequestRepaint True
End Sub

'To support IDE painting of windowless controls, we have to jump through some unpleasant hoops.  This sub needs to be called from
' inside the parent user control's _Paint event, and the destination DC *must be passed* as it only exists for the lifetime of
' the Paint event (and we don't want to subclass inside the designer).
Friend Sub RequestIDERepaint(ByVal targetDC As Long)

    If (Not PDMain.IsProgramRunning()) And (Not m_NoBackBufferMode) Then
        If (Not m_BackBuffer Is Nothing) Then
            CacheWindowDimensions
            If m_CustomCaptionPaintingActive Then
                AskForARepaint False
            Else
                SyncBackBufferSize True
            End If
            GDI.BitBltWrapper targetDC, 0, 0, m_Width, m_Height, m_BackBuffer.GetDIBDC, 0, 0, vbSrcCopy
        End If
    End If

End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
        
    'Color list retrieval is pretty darn easy - just load each color one at a time, and leave the rest to the color class.
    ' It will build an internal hash table of the colors we request, which makes rendering much faster.
    m_Colors.LoadThemeColor UCS_Background, "Background", IDE_WHITE
    m_Colors.LoadThemeColor UCS_Caption, "Caption", IDE_GRAY
    
End Sub

'Child controls can query us to see the ID of the last theme ID we updated against.  If this ID hasn't changed, we don't
' need to redraw anything to match.  (Note that the ID contains data about both the UI theme and the current language;
' changing either should initiate a full redraw.)
Friend Function ThemeUpdateRequired() As Boolean
    If (LenB(m_LastThemeID) = 0) Then
        ThemeUpdateRequired = True
    Else
        ThemeUpdateRequired = Strings.StringsNotEqual(m_LastThemeID, Interface.GetCurrentInterfaceID, False)
    End If
End Function

'Whenever the caller user control receives an UpdateAgainstTheme/Language request, it needs to relay that to this support
' class as well.  We'll update things like the tooltip and text caption (if any) to match the new settings.
Friend Sub UpdateAgainstThemeAndLanguage(Optional ByRef ctlName As String = vbNullString)
    
    If PDMain.IsProgramRunning() Then
        
        If Me.ThemeUpdateRequired() Then
        
            'Tooltips must be refreshed (in case the active language has changed)
            If g_Language.TranslationActive() Then
                m_TTCaption = g_Language.TranslateMessage(m_TTCaptionEn)
                m_TTTitle = g_Language.TranslateMessage(m_TTTitleEn)
            Else
                m_TTCaption = m_TTCaptionEn
                m_TTTitle = m_TTTitleEn
            End If
            
            'This control requires some colors from the central themer; update its color cache now
            UpdateColorList
            
            'The caption manager will also refresh itself
            If m_Caption.IsCaptionActive Then m_Caption.UpdateAgainstCurrentTheme ctlName
            
            'After a theme change, the caller needs to repaint everything from scratch.
            SyncBackBufferSize True
            AskForARepaint True
            
            'Make a note of the current theme+language ID.  If this function gets called again but this ID hasn't changed,
            ' we can skip additional redraws.
            m_LastThemeID = Interface.GetCurrentInterfaceID()
            
        End If
        
    End If
    
End Sub

'***************************************************************************
' (end interactive functions)
'***************************************************************************

Private Sub Class_Initialize()

    m_BackColor = -1
    m_RedrawRequestCount = 0
    m_Visibility = False
    m_NoBackBufferMode = False
    
    'Some classes are required for minimal drawing inside the IDE.  As such, we must prep them now, instead of waiting for a
    ' formal UC registration event.
    Set m_BackBuffer = New pdDIB
    Set m_Caption = New pdCaption
    
End Sub

Private Sub Class_Terminate()
    
    'When PD was using old subclassing techniques, the order in which we released our subclassers was crucial to preventing leaks
    ' and/or crashes.  Now that we use comctl32 subclassing, order is basically irrelevant - but I've left the old order-specific
    ' code in place as I find it instructive.
    
    'Release all internal classes that may perform their own subclassing.
    StopAllSubclassing
    
    'Release our subclasser.
    If (m_hWnd <> 0) Then
        VBHacks.StopSubclassing m_hWnd, Me
        m_hWnd = 0
    End If
    
    UserControls.DecrementPDControlCount
    
End Sub

'Release all internal classes where subclassing may be active.  Note that order may be important here - do not change it
' without careful study of the order that the classes are created!
Private Sub StopAllSubclassing()
    
    'Release any non-subclassing classes first
    If (Not m_Caption Is Nothing) Then Set m_Caption = Nothing
    
    'If our parent control registered to receive internal PD messages, de-register it with PD's central marshaler.
    If m_CustomInternalMessagesActive Then
        UserControls.RemoveMessageRecipient m_hWnd
        m_CustomInternalMessagesActive = False
    End If
    
    'Release our custom subclassed message collection (which just stores message IDs, not hWnds or anything crucial to
    ' stopping the subclass process).
    Set m_CustomMessages = Nothing
    
    'Release custom functionality first, in reverse-order
    If (Not m_KeyEvents Is Nothing) Then
        m_KeyEvents.Shutdown
        Set m_KeyEvents = Nothing
    End If
    
    If (Not m_MouseEvents Is Nothing) Then
        m_MouseEvents.Shutdown
        Set m_MouseEvents = Nothing
    End If

    'Release default functionality next, also in reverse-order
    If (Not m_MoveSize Is Nothing) Then Set m_MoveSize = Nothing
    If (Not m_FocusDetector Is Nothing) Then Set m_FocusDetector = Nothing
    If (Not m_WindowPainter Is Nothing) Then Set m_WindowPainter = Nothing
    
End Sub

Private Sub HandleWM_HideChild()
    If (Not m_NoBackBufferMode) Then m_BackBuffer.SuspendDIB
End Sub

Private Function HandleWM_WindowPosChanging(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim initVisibility As Boolean
    initVisibility = m_Visibility
                
    Dim tmpWindowPos As tagWINDOWPOS
    CopyMemoryStrict VarPtr(tmpWindowPos), lParam, LenB(tmpWindowPos)
    
    If (tmpWindowPos.wFlags And SWP_HIDEWINDOW) <> 0 Then m_Visibility = False
    If (tmpWindowPos.wFlags And SWP_SHOWWINDOW) <> 0 Then
        
        m_Visibility = True
        
        'If we haven't asked the caller for a repaint yet, ask for one now.  (This is the last possible time we can request
        ' one before the control is actually shown.)
        If (m_RedrawRequestCount > 0) Then AskForARepaint True, True
        
    End If
    
    'Finally, notify the parent of the pending visibility change (if any), and suspend our
    ' back buffer if the window has been made invisible
    If (initVisibility <> m_Visibility) Then RaiseEvent VisibilityChange(m_Visibility)
    If (Not m_Visibility) Then m_BackBuffer.SuspendDIB
    
End Function

Private Function HandleWM_ShowWindow(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim initVisibility As Boolean
    initVisibility = m_Visibility
    
    m_Visibility = (wParam <> 0)
    
    'If we haven't asked the caller for a repaint yet, ask for one now.  (This is the last possible time we can request
    ' one before the control is actually shown.)
    If m_Visibility And (Not initVisibility) And (m_RedrawRequestCount > 0) Then AskForARepaint True, True
    
    'Finally, notify the parent of the pending visibility change (if any), and suspend our
    ' back buffer if the window has been made invisible
    If (initVisibility <> m_Visibility) Then RaiseEvent VisibilityChange(m_Visibility)
    If (Not m_Visibility) Then m_BackBuffer.SuspendDIB
    
End Function

'Note that most (all?) of the messages subclassed by this class are "notification"-style only.  Default wndprocs are still called
' after we gather the info we need.
Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    
    Dim eatMessage As Boolean: eatMessage = False
    
    'WM_WINDOWPOSCHANGING covers some visibility change cases that WM_SHOWWINDOW does not;
    ' see http://blogs.msdn.com/b/oldnewthing/archive/2008/01/15/7113860.aspx for details
    If (uiMsg = WM_WINDOWPOSCHANGING) Then
        HandleWM_WindowPosChanging hWnd, uiMsg, wParam, lParam
    
    ElseIf (uiMsg = WM_SHOWWINDOW) Then
        HandleWM_ShowWindow hWnd, uiMsg, wParam, lParam
        
    ElseIf (uiMsg = WM_NCDESTROY) Then
        
        'Stop all 3rd-party subclassing
        StopAllSubclassing
        
        'Stop internal subclassing and free our hWnd (which prevents subclassing from re-starting)
        If (m_hWnd <> 0) Then
            VBHacks.StopSubclassing m_hWnd, Me
            m_hWnd = 0
        End If
    
    'Child windows don't normally receive WM_SHOWWINDOW notifications.  In PD, we manually
    ' forward these messages so that we can free up various UI-related bits to keep memory
    ' as low as humanly possible.
    ElseIf (uiMsg = WM_PD_HIDECHILD) Then
        HandleWM_HideChild
        
    Else
        
        'See if this message is a message that our parent control wants notification for.
        ' If it is, relay it accordingly.
        If (Not m_CustomMessages Is Nothing) Then
            If m_CustomMessages.DoesKeyExist(uiMsg) Then
                
                'Look for some special internal message IDs
                
                'If tab or shift+tab was pressed, give the target control an opportunity to override
                ' the default focus target.
                Dim newHWnd As Long: newHWnd = 0
                If (uiMsg = WM_PD_TAB_KEY_TARGET) And (wParam = m_hWnd) Then
                    RaiseEvent SetCustomTabTarget(False, newHWnd)
                    eatMessage = True
                    PutMem4 lParam, newHWnd
                
                ElseIf (uiMsg = WM_PD_SHIFT_TAB_KEY_TARGET) And (wParam = m_hWnd) Then
                    RaiseEvent SetCustomTabTarget(True, newHWnd)
                    eatMessage = True
                    PutMem4 lParam, newHWnd
                    
                'Any other messages are up to the target control to deal with on their own.
                Else
                    RaiseEvent CustomMessage(uiMsg, wParam, lParam, eatMessage, ISubclass_WindowMsg)
                End If
                
            End If
        End If
        
    End If
    
    'At present, none of our subclassed messages get "eaten".  Always call the default handler after we're done.
    If (Not eatMessage) Then ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
End Function
