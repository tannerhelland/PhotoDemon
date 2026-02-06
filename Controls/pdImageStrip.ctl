VERSION 5.00
Begin VB.UserControl pdImageStrip 
   Alignable       =   -1  'True
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
   ToolboxBitmap   =   "pdImageStrip.ctx":0000
End
Attribute VB_Name = "pdImageStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Strip control (e.g. a scrollable line of image thumbnails)
'Copyright 2013-2026 by Tanner Helland
'Created: 15/October/13
'Last updated: 05/January/17
'Last update: delay resource loading until absolutely required
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this image strip control, specifically:
'
' 1) At present, this control is only used by pdCanvas when multiple images are loaded.  Changes may be required
'     to make it work as a general-purpose image strip.
' 2) High-DPI settings are handled automatically.
' 3) A hand cursor is automatically applied, and clicks are returned via the Click event.
' 4) Coloration is automatically handled by PD's internal theming engine.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This control takes a slightly different approach to events.  It raises a standard Click() event, as expected, but as part
' of this event it provides a full collection of mouse information, too.  This is to facilitate RMB popup menu support.
Public Event Click(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

'When the control's position is changed in any way (size, alignment, visibility), this event is raised.  It is up to
' the caller to handle the event and perform any relevant positioning adjustments.
Public Event PositionChanged()

'In addition, this class also raises events for when a new item is selected, and another when a given item is closed.
' These are much simpler than trying to reverse-engineer item indices from the generic Click() event.
Public Event ItemSelected(ByVal itemIndex As Long)
Public Event ItemClosed(ByVal itemIndex As Long)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'This window is resizable at run-time
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTLEFT As Long = 10
Private Const HTTOP As Long = 12
Private Const HTRIGHT As Long = 11
Private Const HTBOTTOM As Long = 15

'A collection of all currently active thumbnails; this is dynamically resized as thumbnails are added/removed.
Private Type ImageThumbEntry
    thumbDIB As pdDIB
    indexInPDImages As Long
End Type

Private m_Thumbs() As ImageThumbEntry
Private m_NumOfThumbs As Long

'Because the user can resize the thumbnail bar, we must track thumbnail width/height dynamically
Private m_ThumbWidth As Long, m_ThumbHeight As Long

'We don't want thumbnails to fill the full size of their individual blocks, so we apply a border of this many pixels
' to each side of the thumbnail
Private Const THUMB_BORDER_PADDING As Long = 5

'The currently selected and currently hovered thumbnail
Private m_CurrentThumb As Long, m_CurrentThumbHover As Long

'As a convenience to the user, we provide a small notification when an image has unsaved changes
Private m_ModifiedIcon As pdDIB

'In Feb '15, Raj added the very nice capability to close an image by hovering its tab, then clicking the X that magically appears.
' A few DIBs are required for this: normal gray, red highlight when hovered, and an underlying shadow (to help it stand out against
' dark thumbnails).
Private m_CloseIconRed As pdDIB, m_CloseIconGray As pdDIB, m_CloseIconShadow As pdDIB

'We also need a few tracking variables, for example if the user closes a tab that is *not* currently the active one
Private m_CloseTriggeredOnThumbnail As Long
Private m_CloseIconHovered As Long

'Thumbnails can be right-clicked to see a context menu
Private m_RightClickedThumbnail As Long

'If the user loads tons of images, the tabstrip may overflow the available area.  We now allow them to drag-scroll the list.
' In order to allow that, we must track a few extra things, like initial mouse x/y.
Private m_MouseDown As Boolean, m_ScrollingOccured As Boolean
Private m_InitX As Long, m_InitY As Long, m_InitOffset As Long
Private m_ListScrollable As Boolean
Private m_MouseDistanceTraveled As Long

'Instead of using an actual scrollbar, the image strip is currently scrollable by click+drag behavior.
Private m_ScrollValue As Long, m_ScrollMax As Long

'Some tabstrip update functions are allowed to process without actually redrawing the tabstrip.  (This is helpful
' for batching together multiple updates in a row.)  If one or more changes have been made to tabstrip contents,
' but the strip has *not* been redrawn, this will be set to TRUE.
Private m_RedrawNeeded As Boolean

'Horizontal or vertical layout; obviously, all our rendering and mouse detection code changes depending on the orientation
' of the tabstrip.
Private m_VerticalLayout As Boolean

'When we are responsible for this window resizing (because the user is resizing our window manually), we set this to TRUE.
' This variable is then checked before requesting additional redraws during our resize event.
Private weAreResponsibleForResize As Boolean

'Most importantly for scrolling, this value is set to TRUE on cMouseEvents_MouseDownCustom, *if* the mouse is clicked near the resizable edge of the
' toolbar (which varies according to its alignment, obviously).
Private m_MouseInResizeTerritory As Boolean

'Current image strip alignment and visibility mode.  (Visibility mode controls when the tabstrip is visible -
' always, on multiple loaded images, or never.)
Private m_Alignment As AlignConstants, m_VisibilityMode As Long

'Minimum and maximum allowable size.  Note that the actual dimension that correlates to this measurement
' (width or height) changes depending on orientation.  Also, the maximum size may be further limited by
' available viewport space.
Private Const MIN_STRIP_SIZE As Long = 40
Private Const MAX_STRIP_SIZE As Long = 300

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by the control, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDIMAGESTRIP_COLOR_LIST
    [_First] = 0
    PDIS_Background = 0
    PDIS_SelectedFill = 1
    PDIS_SelectedBorder = 2
    PDIS_UnselectedFill = 3
    PDIS_UnselectedBorder = 4
    PDIS_Separator = 5
    [_Last] = 5
    [_Count] = 6
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_ImageStrip
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

Public Property Get Alignment() As AlignConstants
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal newAlignment As AlignConstants)
    
    'When switching between horizontal and vertical layouts, we need to adjust our size to match.
    Dim prevVerticalLayout As Boolean, prevConstrainingSize As Long
    prevVerticalLayout = m_VerticalLayout
    prevConstrainingSize = Me.ConstrainingSize
    
    'From the new alignment setting, determine whether we are in horizontal or vertical mode
    If newAlignment = vbAlignNone Then newAlignment = vbAlignTop
    m_VerticalLayout = ((newAlignment = vbAlignLeft) Or (newAlignment = vbAlignRight))
    
    'If we've just switched between horizontal and vertical modes, resize the control to reflect our current height
    If (prevVerticalLayout <> m_VerticalLayout) Then
        ucSupport.RequestNewSize prevConstrainingSize, prevConstrainingSize, True
    End If
    
    If ((m_Alignment <> newAlignment) Or (prevVerticalLayout <> m_VerticalLayout)) And PDMain.IsProgramRunning() Then
        m_Alignment = newAlignment
        UpdateControlLayout
        UpdateAgainstTabstripPreferences
        PropertyChanged "Alignment"
    End If
    
End Property

'Constraining size is the size of the image strip in the non-fitting direction.  This size is adjustable by the user.
Public Property Get ConstrainingSize() As Long
    If m_VerticalLayout Then ConstrainingSize = ucSupport.GetControlWidth Else ConstrainingSize = ucSupport.GetControlHeight
End Property

Public Property Get VisibilityMode() As Long
    VisibilityMode = m_VisibilityMode
End Property

Public Property Let VisibilityMode(ByVal newMode As Long)
    If m_VisibilityMode <> newMode Then
        m_VisibilityMode = newMode
        UpdateAgainstTabstripPreferences
    End If
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
    If PDMain.IsProgramRunning() Then RedrawBackBuffer
End Property

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
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

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    Dim lbClick As Boolean, rbClick As Boolean
    lbClick = ((Button And pdLeftButton) <> 0)
    rbClick = ((Button And pdRightButton) <> 0)
    
    'LMB clicks can select a new thumb, or close a thumb (if over the corner-aligned "close image" icon)
    If (lbClick Or rbClick) Then
        
        'As a failsafe, the initial MouseDown event will mark whether a close icon is being clicked; this is to prevent
        ' the (admittedly weird) fringe case of click-dragging the list using a "close icon" region.
        If (m_CloseTriggeredOnThumbnail <> -1) And lbClick Then
            
            If GetThumbWithCloseIconAtPosition(x, y) = m_CloseTriggeredOnThumbnail Then
                RaiseEvent ItemClosed(m_Thumbs(m_CloseTriggeredOnThumbnail).indexInPDImages)
            End If

            'Reset the close identifier
            m_CloseTriggeredOnThumbnail = -1
            
        Else
            
            Dim potentialNewThumb As Long
            potentialNewThumb = GetThumbAtPosition(x, y)
            
            'Notify the program that a new image has been selected; it will then bring that image to the foreground,
            ' which will automatically trigger a toolbar redraw.
            If (potentialNewThumb >= 0) And (Not m_ScrollingOccured) Then
                m_CurrentThumb = potentialNewThumb
                RaiseEvent ItemSelected(m_Thumbs(m_CurrentThumb).indexInPDImages)
            End If
            
        End If
            
    End If
    
    'Also raise a generic "click" event, which our parent can deal with however they want
    RedrawBackBuffer
    ucSupport.RequestCursor IDC_HAND
    RaiseEvent Click(Button, Shift, x, y)
    
End Sub

Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    If (wMsg = WM_PD_COLOR_MANAGEMENT_CHANGE) Then Me.RequestTotalRedraw True
End Sub

Private Sub ucSupport_GotFocusAPI()
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_LostFocusAPI()
    RedrawBackBuffer
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'On left-button presses, make a note of the initial mouse position
    If (Button And pdLeftButton) <> 0 Then
    
        m_MouseDown = True
        m_InitX = x
        m_InitY = y
        m_MouseDistanceTraveled = 0
        m_InitOffset = m_ScrollValue
        
        'Detect close icon click, and store the clicked thumbnail
        m_CloseTriggeredOnThumbnail = GetThumbWithCloseIconAtPosition(x, y)
        
        'We must also detect if the mouse is over the edge of the form that allows live-resizing.  (This varies by tabstrip orientation, obviously.)
        m_MouseInResizeTerritory = IsMouseOverResizeBorder(x, y)
        
    ElseIf Button = vbRightButton Then
        m_RightClickedThumbnail = GetThumbAtPosition(x, y)
    End If
    
    'Reset the "resize in progress" tracker
    weAreResponsibleForResize = False
    
    'Reset the "scrolling occured" tracker
    m_ScrollingOccured = False
    
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If m_CurrentThumbHover <> -1 Then
        m_CurrentThumbHover = -1
        RedrawBackBuffer
    End If
    
    ucSupport.RequestCursor IDC_DEFAULT
    
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'Ignore all mouse events while the user is interacting with the canvas
    If FormMain.MainCanvas(0).IsMouseDown(pdLeftButton Or pdRightButton) Then Exit Sub
    
    'We require a few mouse movements to fire before doing anything; otherwise this function will fire constantly.
    m_MouseDistanceTraveled = m_MouseDistanceTraveled + 1
    
    'We handle several different _MouseMove scenarios, in this order:
    ' 1) If the mouse is near the resizable edge of the form, and the left button is depressed, activate live resizing.
    ' 2) If a button is depressed, activate tabstrip scrolling (if the list is long enough)
    ' 3) If no buttons are depressed, hover the image at the current position (if any)
        
    'Check mouse button state; if it's down, check for resize or scrolling of the image list
    If m_MouseDown Then
        
        If m_MouseInResizeTerritory Then
                
            If ((Button And pdLeftButton) <> 0) Then
            
                'Figure out which resize message to send to Windows; this varies by tabstrip orientation (obviously)
                Dim hitCode As Long
    
                Select Case Me.Alignment
                
                    Case vbAlignLeft
                        hitCode = HTRIGHT
                    
                    Case vbAlignTop
                        hitCode = HTBOTTOM
                    
                    Case vbAlignRight
                        hitCode = HTLEFT
                    
                    Case vbAlignBottom
                        hitCode = HTTOP
                
                End Select
                
                'Initiate resizing, and set a form-level marker so that other functions know we're responsible for any resize-related events
                weAreResponsibleForResize = True
                ucSupport.NotifyMouseDragResize_Start
                VBHacks.SendMsgW Me.hWnd, WM_NCLBUTTONDOWN, hitCode, 0&
                
                'After the drag operation is complete, the code will resume right here
                ucSupport.NotifyMouseDragResize_End
                m_MouseDown = False
                RaiseEvent PositionChanged
                
            End If
        
        'The mouse is not in resize territory.  This means the user is click-dragging to scroll a long list.
        Else
            
            'If the list is scrollable (due to tons of images being loaded), calculate a new offset now
            If (m_ListScrollable And (m_MouseDistanceTraveled > 5) And (Not weAreResponsibleForResize)) Then
            
                m_ScrollingOccured = True
            
                Dim mouseOffset As Long
                If m_VerticalLayout Then mouseOffset = (m_InitY - y) Else mouseOffset = (m_InitX - x)
                
                'Change the invisible scroll bar to match the new offset
                Dim newScrollValue As Long
                newScrollValue = m_InitOffset + mouseOffset
                
                If (newScrollValue < 0) Then
                    m_ScrollValue = 0
                ElseIf (newScrollValue > m_ScrollMax) Then
                    m_ScrollValue = m_ScrollMax
                Else
                    m_ScrollValue = newScrollValue
                End If
                
            End If
        
        End If
    
    'The left mouse button is not down.  Hover the image beneath the cursor (if any)
    Else
    
        'We want to highlight a close icon, if it's being hovered
        m_CloseIconHovered = GetThumbWithCloseIconAtPosition(x, y)
        
        Dim oldThumbHover As Long
        oldThumbHover = m_CurrentThumbHover
        
        'Retrieve the thumbnail at this position, and change the mouse pointer accordingly
        m_CurrentThumbHover = GetThumbAtPosition(x, y)
                
        'To prevent flickering, only update the tooltip when absolutely necessary
        If (m_CurrentThumbHover <> oldThumbHover) Then
        
            'If the cursor is over a thumbnail, update the tooltip to display that image's filename
            If (m_CurrentThumbHover <> -1) Then
                
                Dim strKeyLocation As String, strKeyName As String
                strKeyLocation = "CurrentLocationOnDisk"
                strKeyName = "OriginalFileName"
                
                If (LenB(PDImages.GetImageByID(m_Thumbs(m_CurrentThumbHover).indexInPDImages).ImgStorage.GetEntry_String(strKeyLocation)) <> 0) Then
                    Me.AssignTooltip PDImages.GetImageByID(m_Thumbs(m_CurrentThumbHover).indexInPDImages).ImgStorage.GetEntry_String(strKeyLocation), PDImages.GetImageByID(m_Thumbs(m_CurrentThumbHover).indexInPDImages).ImgStorage.GetEntry_String(strKeyName)
                Else
                    Me.AssignTooltip "Once this image has been saved to disk, its filename will appear here.", "This image does not have a filename."
                End If
            
            'The cursor is not over a thumbnail; let the user know they can hover if they want more information.
            Else
                Me.AssignTooltip "Hover an image thumbnail to see its name and current file location."
            End If
            
        End If
        
    End If
    
    'Set a mouse pointer according to the handling above
    If IsMouseOverResizeBorder(x, y) Then
        If m_VerticalLayout Then ucSupport.RequestCursor IDC_SIZEWE Else ucSupport.RequestCursor IDC_SIZENS
    
    'Display a hand cursor if over an image; default cursor otherwise
    Else
        If m_CurrentThumbHover = -1 Then ucSupport.RequestCursor IDC_DEFAULT Else ucSupport.RequestCursor IDC_HAND
    End If
    
    RedrawBackBuffer
    
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    
    If m_MouseDown Then
        m_MouseDown = False
        m_InitX = 0
        m_InitY = 0
        m_MouseDistanceTraveled = 0
        weAreResponsibleForResize = False
    End If
    
End Sub

Private Sub ucSupport_MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    If m_ListScrollable Then ScrollStripByWheel scrollAmount, x, y
End Sub

Private Sub ucSupport_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    If m_ListScrollable Then ScrollStripByWheel -1 * scrollAmount, x, y
End Sub

Private Sub ScrollStripByWheel(ByVal scrollAmount As Single, ByVal x As Long, ByVal y As Long)
    
    Dim scrollPixels As Long
    scrollPixels = FixDPI(16)
    
    If (scrollAmount > 0) Then
        m_ScrollValue = m_ScrollValue + scrollPixels
        If (m_ScrollValue > m_ScrollMax) Then m_ScrollValue = m_ScrollMax
    ElseIf (scrollAmount < 0) Then
        m_ScrollValue = m_ScrollValue - scrollPixels
        If (m_ScrollValue < 0) Then m_ScrollValue = 0
    End If
    
    m_CurrentThumbHover = GetThumbAtPosition(x, y)
    RedrawBackBuffer
        
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
        
    'Enforce minimum and maximum size restrictions
    Dim relevantSize As Long
    If m_VerticalLayout Then
        relevantSize = ucSupport.GetControlWidth
    Else
        relevantSize = ucSupport.GetControlHeight
    End If
    
    If relevantSize < FixDPI(MIN_STRIP_SIZE) Then
        If m_VerticalLayout Then
            ucSupport.RequestNewSize FixDPI(MIN_STRIP_SIZE), ucSupport.GetControlHeight, False
        Else
            ucSupport.RequestNewSize ucSupport.GetControlWidth, FixDPI(MIN_STRIP_SIZE), False
        End If
    ElseIf relevantSize > FixDPI(MAX_STRIP_SIZE) Then
        If m_VerticalLayout Then
            ucSupport.RequestNewSize FixDPI(MAX_STRIP_SIZE), ucSupport.GetControlHeight, False
        Else
            ucSupport.RequestNewSize ucSupport.GetControlWidth, FixDPI(MAX_STRIP_SIZE), False
        End If
    End If
    
    'Normally we would want to update the control's layout here, but ucSupport will automatically raise
    ' a redraw request of its own (making this one redundant!)
    'UpdateControlLayout
    
End Sub

'New images are currently added by their ID value; at some point it might be nice to modify this to
' accept any arbitrary image, but for the primary canvas, this method avoids a lot of unnecessary "glue" code.
Public Sub AddNewThumb(ByVal pdImageIndex As Long)

    'Request a thumbnail from the relevant pdImage object
    If m_VerticalLayout Then
        PDImages.GetImageByID(pdImageIndex).RequestThumbnail m_Thumbs(m_NumOfThumbs).thumbDIB, m_ThumbHeight - (FixDPI(THUMB_BORDER_PADDING) * 2)
    Else
        PDImages.GetImageByID(pdImageIndex).RequestThumbnail m_Thumbs(m_NumOfThumbs).thumbDIB, m_ThumbWidth - (FixDPI(THUMB_BORDER_PADDING) * 2)
    End If
    
    'Make a note of this thumbnail's index in the main pdImages array
    m_Thumbs(m_NumOfThumbs).indexInPDImages = pdImageIndex
    m_CurrentThumb = m_NumOfThumbs
    
    'Prepare the array to receive another entry in the future, then redraw the strip
    m_NumOfThumbs = m_NumOfThumbs + 1
    If (m_NumOfThumbs > UBound(m_Thumbs)) Then ReDim Preserve m_Thumbs(0 To UBound(m_Thumbs) * 2 + 1) As ImageThumbEntry
    
    'This is a little clumsy, but the redraw function also calculates new scroll bar maximum values.  As such,
    ' we need to perform one redraw to calculate a new max scroll value, then a *second* redraw if the image still
    ' lies off-screen.
    RedrawBackBuffer
    If FitThumbnailOnscreen(m_CurrentThumb) Then RedrawBackBuffer
    
End Sub

'Call this function to forcibly adjust the scrollbar so that the currently active thumbnail is moved on-screen.
' Note that it *does not actually perform a redraw* - instead, it will return TRUE if the scroll value changed.
' It is up to the caller to check that value and request a redraw accordingly.
Friend Function FitThumbnailOnscreen(ByVal thumbIndex As Long) As Boolean

    Dim isThumbnailOnscreen As Boolean

    'Because this control does not dynamically track thumb position, we must first figure out where this
    ' image thumbnail is currently positioned.  (Note that its position changes according to alignment, obviously)
    Dim hPosition As Long, vPosition As Long
    
    'Use the tabstrip's size to determine if this thumbnail lies off-screen
    If m_VerticalLayout Then
    
        vPosition = (thumbIndex * m_ThumbHeight) - m_ScrollValue
        
        If (vPosition < 0) Then
            isThumbnailOnscreen = False
        ElseIf ((vPosition + m_ThumbHeight - 1) > ucSupport.GetControlHeight) Then
            isThumbnailOnscreen = False
        Else
            isThumbnailOnscreen = True
        End If
        
    Else
        
        hPosition = (thumbIndex * m_ThumbWidth) - m_ScrollValue
        
        If (hPosition < 0) Then
            isThumbnailOnscreen = False
        ElseIf ((hPosition + m_ThumbWidth - 1) > ucSupport.GetControlWidth) Then
            isThumbnailOnscreen = False
        Else
            isThumbnailOnscreen = True
        End If
        
    End If
    
    'If the thumbnail is not onscreen, make it so!
    If (Not isThumbnailOnscreen) Then
        
        Dim newScrollValue As Long
        
        If m_VerticalLayout Then
            If (vPosition < 0) Then
                newScrollValue = thumbIndex * m_ThumbHeight
            Else
                newScrollValue = ((thumbIndex + 1) * m_ThumbHeight) - ucSupport.GetControlHeight
            End If
        Else
            If (hPosition < 0) Then
                newScrollValue = thumbIndex * m_ThumbWidth
            Else
                newScrollValue = ((thumbIndex + 1) * m_ThumbWidth) - ucSupport.GetControlWidth
            End If
        End If
        
        If (newScrollValue > m_ScrollMax) Then newScrollValue = m_ScrollMax
        m_ScrollValue = newScrollValue
    
    End If
    
    FitThumbnailOnscreen = (Not isThumbnailOnscreen)
            
End Function

'Given an (x, y) mouse pair, return the thumbnail index at that location.  If the cursor is not over a thumbnail, -1 is returned.
Private Function GetThumbAtPosition(ByVal x As Long, ByVal y As Long) As Long
    If m_VerticalLayout Then
        GetThumbAtPosition = (y + m_ScrollValue) \ m_ThumbHeight
        If GetThumbAtPosition > (m_NumOfThumbs - 1) Then GetThumbAtPosition = -1
    Else
        GetThumbAtPosition = (x + m_ScrollValue) \ m_ThumbWidth
        If GetThumbAtPosition > (m_NumOfThumbs - 1) Then GetThumbAtPosition = -1
    End If
End Function

'Images in PD are referenced by an unchangeable "ID" value.  That value *may not* correlate
' with an image's index in our current thumbnail collection.  Use this function to correlate the two.
Private Function GetThumbIndexFromPDIndex(ByVal pdImageIndex As Long) As Long
    
    GetThumbIndexFromPDIndex = -1
    
    Dim i As Long
    For i = 0 To m_NumOfThumbs - 1
        If (m_Thumbs(i).indexInPDImages = pdImageIndex) Then
            GetThumbIndexFromPDIndex = i
            Exit For
        End If
    Next i
    
End Function

'Given an (x, y) coordinate pair, determine whether that position lies within the clickable "close image" icon region.
' RETURNS: relevant thumbnail index if the position correlates to the "close icon" region, -1 otherwise
Private Function GetThumbWithCloseIconAtPosition(ByVal x As Long, ByVal y As Long) As Long
    
    GetThumbWithCloseIconAtPosition = -1
    
    Dim thumbIndex As Long
    thumbIndex = GetThumbAtPosition(x, y)
    
    If (thumbIndex <> -1) Then
        
        'Start by determing the boundary region of the underlying thumbnail
        Dim thumbnailStartOffsetX As Long, thumbnailStartOffsetY As Long
        If m_VerticalLayout Then
            thumbnailStartOffsetY = m_ThumbHeight * thumbIndex - m_ScrollValue
        Else
            thumbnailStartOffsetX = m_ThumbWidth * thumbIndex - m_ScrollValue
        End If
        
        'From this, determine where the "close icon" would appear on the thumbnail
        Dim closeButtonStartOffsetX As Long, closeButtonStartOffsetY As Long
        If (m_CloseIconGray Is Nothing) Then GetCloseImageResources
        closeButtonStartOffsetX = thumbnailStartOffsetX + (m_ThumbWidth - (FixDPI(THUMB_BORDER_PADDING) + m_CloseIconGray.GetDIBWidth + FixDPI(2)))
        closeButtonStartOffsetY = thumbnailStartOffsetY + FixDPI(THUMB_BORDER_PADDING) + FixDPI(2)
        
        Dim clickBoundaryX As Long, clickBoundaryY As Long
        clickBoundaryX = x - closeButtonStartOffsetX
        clickBoundaryY = y - closeButtonStartOffsetY
        
        If ((clickBoundaryX >= 0) And (clickBoundaryY >= 0)) Then
            If ((clickBoundaryX < m_CloseIconGray.GetDIBWidth) And (clickBoundaryY < m_CloseIconGray.GetDIBHeight)) Then
                GetThumbWithCloseIconAtPosition = thumbIndex
            End If
        End If
        
    End If
    
End Function

'Given an x/y mouse coordinate, return TRUE if the coordinate falls over the form resize area.  Tabstrip alignment is automatically handled.
Private Function IsMouseOverResizeBorder(ByVal mouseX As Single, ByVal mouseY As Single) As Boolean

    'How close does the mouse have to be to the form border to allow resizing?  We currently use 5 pixels, while accounting
    ' for DPI variance (e.g. 5 pixels at 96 dpi)
    Dim resizeBorderAllowance As Long
    resizeBorderAllowance = FixDPI(5)
    
    Select Case Me.Alignment
    
        Case vbAlignLeft
            If (mouseY > 0) And (mouseY < ucSupport.GetControlHeight) And (mouseX > ucSupport.GetControlWidth - resizeBorderAllowance) Then IsMouseOverResizeBorder = True
            
        Case vbAlignTop
            If (mouseX > 0) And (mouseX < ucSupport.GetControlWidth) And (mouseY > ucSupport.GetControlHeight - resizeBorderAllowance) Then IsMouseOverResizeBorder = True
            
        Case vbAlignRight
            If (mouseY > 0) And (mouseY < ucSupport.GetControlHeight) And (mouseX < resizeBorderAllowance) Then IsMouseOverResizeBorder = True
            
        Case vbAlignBottom
            If (mouseX > 0) And (mouseX < ucSupport.GetControlWidth) And (mouseY < resizeBorderAllowance) Then IsMouseOverResizeBorder = True
            
    End Select

End Function

'Sometimes an external component will have reason to change the active image.  If it notifies us, we'll adjust our layout
' to bring that image on-screen (among other redraw necessities).
Public Sub NotifyNewActiveImage(ByVal pdImageIndex As Long)
    
    Dim newThumbIndex As Long
    newThumbIndex = GetThumbIndexFromPDIndex(pdImageIndex)
    
    If (newThumbIndex <> m_CurrentThumb) Or m_RedrawNeeded Then
        m_RedrawNeeded = False
        m_CurrentThumb = newThumbIndex
        FitThumbnailOnscreen pdImageIndex
        RedrawBackBuffer
    End If
        
End Sub

'If some external action changes one of our image(s), the caller must notify us so that we can request an updated thumbnail
Public Sub NotifyUpdatedImage(ByVal pdImageIndex As Long)
    
    'Since we'll be interacting with the passed pdImages object, perform a quick failsafe check to make sure we don't crash
    Dim okayToUpdate As Boolean
    okayToUpdate = PDImages.IsImageActive(pdImageIndex)
    
    If okayToUpdate Then
    
        Dim thumbIndex As Long
        thumbIndex = GetThumbIndexFromPDIndex(pdImageIndex)
        
        If (thumbIndex >= 0) Then
        
            If m_VerticalLayout Then
                PDImages.GetImageByID(pdImageIndex).RequestThumbnail m_Thumbs(thumbIndex).thumbDIB, m_ThumbHeight - (FixDPI(THUMB_BORDER_PADDING) * 2)
            Else
                PDImages.GetImageByID(pdImageIndex).RequestThumbnail m_Thumbs(thumbIndex).thumbDIB, m_ThumbWidth - (FixDPI(THUMB_BORDER_PADDING) * 2)
            End If
            
            RedrawBackBuffer
        
        End If
        
    End If
        
End Sub

'When removing a thumbnail image, we should probably redraw the strip to match.  However, at shutdown time PD will instruct
' the bar to ignore redraws, so we can shut down the program more efficiently.
Public Sub RemoveThumb(ByVal pdImageIndex As Long, Optional ByVal refreshStrip As Boolean = True)

    'Find the matching thumbnail in our collection
    Dim thumbIndex As Long
    thumbIndex = GetThumbIndexFromPDIndex(pdImageIndex)
    
    If (thumbIndex <> -1) Then
    
        'Immediately free any resources associated with the removed image
        Set m_Thumbs(thumbIndex).thumbDIB = Nothing
        
        'Shift all subsequent entries downward
        If thumbIndex < (m_NumOfThumbs - 1) Then
            
            Dim i As Long
            For i = thumbIndex To m_NumOfThumbs - 2
                Set m_Thumbs(i).thumbDIB = m_Thumbs(i + 1).thumbDIB
                m_Thumbs(i).indexInPDImages = m_Thumbs(i + 1).indexInPDImages
            Next i
            
            Set m_Thumbs(m_NumOfThumbs - 1).thumbDIB = Nothing
            
        End If
        
        'Update the number of active thumbnails.  (This must remain synchronized, so we can calculate proper metrics for
        ' large lists that need to be scrollable.)
        m_NumOfThumbs = m_NumOfThumbs - 1
        If (m_NumOfThumbs < 0) Then
            m_NumOfThumbs = 0
            m_CurrentThumb = 0
        End If
        
        If (thumbIndex <= m_CurrentThumb) Then m_CurrentThumb = m_CurrentThumb - 1
        
    End If
    
    'Rendering metrics are regenerated by the back buffer renderer, so it takes care of everything beyond this point
    m_RedrawNeeded = (Not refreshStrip)
    If refreshStrip Then RedrawBackBuffer

End Sub

'INITIALIZE control
Private Sub UserControl_Initialize()
    
    'Reset the thumbnail array
    m_NumOfThumbs = 0
    ReDim m_Thumbs(0 To 3) As ImageThumbEntry
        
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True
    ucSupport.SubclassCustomMessage WM_PD_COLOR_MANAGEMENT_CHANGE, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDIMAGESTRIP_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDImageStrip", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
    ' Track the last thumbnail whose close icon has been clicked.
    ' -1 means no close icon has been clicked yet
    m_CloseTriggeredOnThumbnail = -1
    
    ' Track the last right-clicked thumbnail.
    m_RightClickedThumbnail = -1
        
    'If the tabstrip ever becomes long enough to scroll, this will be set to TRUE
    m_ListScrollable = False
    
End Sub

Private Sub UserControl_InitProperties()
    Alignment = vbAlignTop
    Enabled = True
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Not g_AllowDragAndDrop) Then Exit Sub
    g_Clipboard.LoadImageFromDragDrop Data, Effect, False
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    If g_Clipboard.IsObjectDragDroppable(Data) And g_AllowDragAndDrop Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Alignment = .ReadProperty("Alignment", vbAlignTop)
        Enabled = .ReadProperty("Enabled", True)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Alignment", m_Alignment, vbAlignTop
        .WriteProperty "Enabled", Me.Enabled, True
    End With
End Sub

Public Sub ReadUserPreferences()

    'Constraining size is settable by the user
    Dim cSize As Long
    cSize = UserPrefs.GetPref_Long("Core", "Image Tabstrip Size", Me.ConstrainingSize)
    
    If m_VerticalLayout Then
        If (ucSupport.GetControlWidth <> cSize) Then Me.SetWidth cSize
    Else
        If (ucSupport.GetControlHeight <> cSize) Then Me.SetHeight cSize
    End If
    
End Sub

Public Sub WriteUserPreferences()
    UserPrefs.SetPref_Long "Core", "Image Tabstrip Size", Me.ConstrainingSize
End Sub

Private Sub GetChangedImageResources()

    'Retrieve the unsaved image notification icon from the resource file, and stroke an outline around it
    ' to make it more visible.
    
    'Start by retrieving the original image at a much larger size than we actually need
    Dim unsavedNoteSizeTmp As Long:    unsavedNoteSizeTmp = Interface.FixDPI(64)
    Dim tmpDIB As pdDIB
    LoadResourceToDIB "generic_asterisk", tmpDIB, unsavedNoteSizeTmp, unsavedNoteSizeTmp, 2
    
    'Create an outline pen and stroke the image outline
    Dim cPen As pd2DPen
    Drawing2D.QuickCreateSolidPen cPen, 2.8, 0&, 80#, P2_LJ_Round, P2_LC_Round
    DIBs.OutlineDIB tmpDIB, cPen
    
    'Shrink the outlined DIB down to the size we actually need.  This results in a higher-quality outline
    ' since we're basically supersampling it.
    If (m_ModifiedIcon Is Nothing) Then Set m_ModifiedIcon = New pdDIB
    Dim unsavedNoteSizeFinal As Long
    unsavedNoteSizeFinal = Interface.FixDPI(16)
    m_ModifiedIcon.CreateBlank unsavedNoteSizeFinal, unsavedNoteSizeFinal, 32, 0, 0
    GDI_Plus.GDIPlus_StretchBlt m_ModifiedIcon, 1, 1, unsavedNoteSizeFinal - 2, unsavedNoteSizeFinal - 2, tmpDIB, 0, 0, unsavedNoteSizeTmp, unsavedNoteSizeTmp, , , , , , True
    Set tmpDIB = Nothing
    
End Sub

Private Sub GetCloseImageResources()

    'Retrieve all PNGs necessary to render the "close by hovering" X that appears
    Dim xCloseSize As Long, xClosePadding As Long
    xCloseSize = Interface.FixDPI(16): xClosePadding = Interface.FixDPI(0)
    
    If (m_CloseIconRed Is Nothing) Then Set m_CloseIconRed = New pdDIB
    LoadResourceToDIB "file_close", m_CloseIconRed, xCloseSize, xCloseSize, xClosePadding, g_Themer.GetGenericUIColor(UI_ErrorRed)
    
    If (m_CloseIconGray Is Nothing) Then Set m_CloseIconGray = New pdDIB
    LoadResourceToDIB "file_close", m_CloseIconGray, xCloseSize, xCloseSize, xClosePadding, g_Themer.GetGenericUIColor(UI_GrayNeutral)
    
    'Generate a drop-shadow for the X.  (We can use the same one for both red and gray, obviously.)
    If (m_CloseIconShadow Is Nothing) Then Set m_CloseIconShadow = New pdDIB
    Filters_Layers.CreateShadowDIB m_CloseIconGray, m_CloseIconShadow
    
    'Pad and blur the drop-shadow
    Dim tmpLUT() As Byte
    Dim cFilter As pdFilterLUT
    Set cFilter = New pdFilterLUT
    cFilter.FillLUT_Invert tmpLUT
    PadDIB m_CloseIconShadow, FixDPI(THUMB_BORDER_PADDING)
    QuickBlurDIB m_CloseIconShadow, Interface.FixDPI(2), False
    m_CloseIconShadow.SetAlphaPremultiplication False
    cFilter.ApplyLUTToAllColorChannels m_CloseIconShadow, tmpLUT
    m_CloseIconShadow.SetAlphaPremultiplication True
    
End Sub

Public Sub RequestTotalRedraw(Optional ByVal regenerateThumbsToo As Boolean = False)
    UpdateControlLayout regenerateThumbsToo
End Sub

'Because thumbnail sizes are closely tied to the size of the image strip, we generally need to recalculate a bunch
' of internal rendering metrics whenever the control size changes.
Private Sub UpdateControlLayout(Optional ByVal thumbsMustBeUpdated As Boolean = False)
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetControlWidth
    bHeight = ucSupport.GetControlHeight
    
    'Detect alignment changes (if any)
    If PDMain.IsProgramRunning() Then
        
        'If the control's size has changed in the dimension that determines thumb size, we need to recreate all image thumbnails
        Dim oldThumbWidth As Long, oldThumbHeight As Long
        oldThumbWidth = m_ThumbWidth: oldThumbHeight = m_ThumbHeight
        
        'Calculate new thumbnail sizes
        If m_VerticalLayout Then
            m_ThumbWidth = bWidth - 2
            m_ThumbHeight = m_ThumbWidth
        Else
            m_ThumbHeight = bHeight - 2
            m_ThumbWidth = m_ThumbHeight
        End If
        
        'Determine thumb refreshing by comparing old and new thumb values (but only in the relevant direction!)
        If ((Not m_VerticalLayout) And (m_ThumbHeight <> oldThumbHeight)) Then thumbsMustBeUpdated = True
        If (m_VerticalLayout And (m_ThumbWidth <> oldThumbWidth)) Then thumbsMustBeUpdated = True
        
        If thumbsMustBeUpdated Then
        
            Dim i As Long
            For i = 0 To m_NumOfThumbs - 1
                PDImages.GetImageByID(m_Thumbs(i).indexInPDImages).RequestThumbnail m_Thumbs(i).thumbDIB, m_ThumbWidth - (FixDPI(THUMB_BORDER_PADDING) * 2)
            Next i
            
        End If
        
    End If
    
    'Redraw the toolbar
    RedrawBackBuffer
    
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDIS_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
    
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    If PDMain.IsProgramRunning() And (m_NumOfThumbs > 0) And ucSupport.AmIVisible Then
        
        'Horizontal/vertical layout changes the constraining dimension (e.g. the dimension used to detect if the number
        ' of image tabs currently visible is long enough that it needs to be scrollable).
        Dim constrainingDimension As Long, constrainingMax As Long
        If m_VerticalLayout Then
            constrainingDimension = m_ThumbHeight
            constrainingMax = bHeight
        Else
            constrainingDimension = m_ThumbWidth
            constrainingMax = bWidth
        End If
        
        'Determine if the scrollbar needs to be accounted for or not
        Dim maxThumbSize As Long
        maxThumbSize = constrainingDimension * m_NumOfThumbs - 1
        
        If (maxThumbSize < constrainingMax) Then
            
            'Scrolling is unnecessary, as the entire thumbnail list fits on-screen.  Reset the scroll offset (if any).
            m_ScrollValue = 0
            m_ListScrollable = False
            
        Else
            
            'The current thumbnail list is too large to fit on-screen at once.  Note the maximum scroll value, and restrict the
            ' current offset (if any) to match.
            m_ListScrollable = True
            m_ScrollMax = maxThumbSize - constrainingMax
            If (m_ScrollValue > m_ScrollMax) Then m_ScrollValue = m_ScrollMax
            
        End If
        
        'pd2D is used for rendering
        Dim cSurface As pd2DSurface, cPen As pd2DPen, cBrush As pd2DBrush
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, False
        Drawing2D.QuickCreateSolidBrush cBrush
        Drawing2D.QuickCreateSolidPen cPen, 3!
        
        Dim isEnabled As Boolean, isHovered As Boolean, isSelected As Boolean
        isEnabled = Me.Enabled
        
        'Render each thumbnail block in turn
        Dim thumbRect As RectF
        thumbRect.Width = m_ThumbWidth
        thumbRect.Height = m_ThumbHeight
        
        Dim i As Long, tabVisible As Boolean
        For i = 0 To m_NumOfThumbs - 1
            
            tabVisible = False
            
            'Fill in the rest of this thumbnail's rect
            If m_VerticalLayout Then
                thumbRect.Top = (i * m_ThumbHeight) - m_ScrollValue
                If (Me.Alignment = vbAlignLeft) Then thumbRect.Left = 0 Else thumbRect.Left = 2
                tabVisible = ((thumbRect.Top + thumbRect.Height) >= 0) And (thumbRect.Top <= bHeight)
            Else
                thumbRect.Left = (i * m_ThumbWidth) - m_ScrollValue
                If (Me.Alignment = vbAlignTop) Then thumbRect.Top = 0 Else thumbRect.Top = 2
                tabVisible = ((thumbRect.Left + thumbRect.Width) >= 0) And (thumbRect.Left <= bWidth)
            End If
            
            If tabVisible Then
                
                isSelected = (i = m_CurrentThumb)
                isHovered = (i = m_CurrentThumbHover)
                
                If isSelected Then
                    cBrush.SetBrushColor m_Colors.RetrieveColor(PDIS_SelectedFill, isEnabled, , isHovered)
                    cPen.SetPenColor m_Colors.RetrieveColor(PDIS_SelectedBorder, isEnabled, , isHovered)
                Else
                    cBrush.SetBrushColor m_Colors.RetrieveColor(PDIS_UnselectedFill, isEnabled, , isHovered)
                    cPen.SetPenColor m_Colors.RetrieveColor(PDIS_UnselectedBorder, isEnabled, , isHovered)
                End If
                
                RenderThumbTab i, thumbRect, cSurface, cBrush, cPen
                
            End If
            
        Next i
        
        'Finally, draw a colored line between the tabstrip and the canvas (to create a little more
        ' visual separation)
        Set cPen = Nothing
        Drawing2D.QuickCreateSolidPen cPen, 2!, m_Colors.RetrieveColor(PDIS_Separator, Me.Enabled)
        
        Select Case Me.Alignment
        
            Case vbAlignLeft
                PD2D.DrawLineI cSurface, cPen, bWidth - 1, 0, bWidth - 1, bHeight
                
            Case vbAlignTop
                PD2D.DrawLineI cSurface, cPen, 0, bHeight - 1, bWidth, bHeight - 1
                
            Case vbAlignRight
                PD2D.DrawLineI cSurface, cPen, 1, 0, 1, bHeight
                
            Case vbAlignBottom
                PD2D.DrawLineI cSurface, cPen, 0, 1, bWidth, 1
                
        End Select
        
        Set cPen = Nothing: Set cBrush = Nothing: Set cSurface = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint True
    
End Sub

'Render a given thumbnail onto the background form at the specified offset
Private Sub RenderThumbTab(ByVal thumbIndex As Long, ByRef thumbRectF As RectF, ByRef dstSurface As pd2DSurface, ByRef fillBrush As pd2DBrush, ByRef outlinePen As pd2DPen)
    
    'Fill the thumbnail's background
    PD2D.FillRectangleF_FromRectF dstSurface, fillBrush, thumbRectF
    
    '...then paint a border around it (if it's selected)
    With thumbRectF
        PD2D.DrawRectangleF dstSurface, outlinePen, .Left + 1!, .Top + 1!, .Width - 2!, .Height - 2!
    End With
    
    '...then paint the thumbnail image itself...
    Dim offsetX As Long, offsetY As Long
    offsetX = thumbRectF.Left
    offsetY = thumbRectF.Top
    
    m_Thumbs(thumbIndex).thumbDIB.AlphaBlendToDC dstSurface.GetSurfaceDC(), 255, offsetX + FixDPI(THUMB_BORDER_PADDING), offsetY + FixDPI(THUMB_BORDER_PADDING)
    m_Thumbs(thumbIndex).thumbDIB.FreeFromDC
    
    '...then paint an asterisk in the bottom-left if the parent image has unsaved changes...
    If PDImages.IsImageActive(m_Thumbs(thumbIndex).indexInPDImages) Then
        If (Not PDImages.GetImageByID(m_Thumbs(thumbIndex).indexInPDImages).GetSaveState(pdSE_AnySave)) Then
            If (m_ModifiedIcon Is Nothing) Then GetChangedImageResources
            m_ModifiedIcon.AlphaBlendToDC dstSurface.GetSurfaceDC(), 230, offsetX + FixDPI(THUMB_BORDER_PADDING) + FixDPI(2), offsetY + m_ThumbHeight - FixDPI(THUMB_BORDER_PADDING) - m_ModifiedIcon.GetDIBHeight - FixDPI(2)
            m_ModifiedIcon.FreeFromDC
        End If
    End If
    
    '...and finally, if this thumb is being hovered, we paint a "close" icon in the top-right corner.
    If (thumbIndex = m_CurrentThumbHover) Then
        
        If (m_CloseIconShadow Is Nothing) Then GetCloseImageResources
        m_CloseIconShadow.AlphaBlendToDC dstSurface.GetSurfaceDC(), 230, offsetX + (m_ThumbWidth - (FixDPI(THUMB_BORDER_PADDING) * 2 + m_CloseIconRed.GetDIBWidth + FixDPI(2))), offsetY + FixDPI(2)
        m_CloseIconShadow.FreeFromDC
        
        If (thumbIndex = m_CloseIconHovered) Then
            m_CloseIconRed.AlphaBlendToDC dstSurface.GetSurfaceDC(), 230, offsetX + (m_ThumbWidth - (FixDPI(THUMB_BORDER_PADDING) + m_CloseIconRed.GetDIBWidth + FixDPI(2))), offsetY + FixDPI(THUMB_BORDER_PADDING) + FixDPI(2)
            m_CloseIconRed.FreeFromDC
        Else
            m_CloseIconGray.AlphaBlendToDC dstSurface.GetSurfaceDC(), 230, offsetX + (m_ThumbWidth - (FixDPI(THUMB_BORDER_PADDING) + m_CloseIconRed.GetDIBWidth + FixDPI(2))), offsetY + FixDPI(THUMB_BORDER_PADDING) + FixDPI(2)
            m_CloseIconGray.FreeFromDC
        End If
        
    End If
    
End Sub

'When the control's size is changed in some way, call this function to perform some internal maintenance tasks,
' and raise an event our parent can deal with.
Public Sub UpdateAgainstTabstripPreferences()
    RaiseEvent PositionChanged
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDIS_Background, "Background", IDE_WHITE
        .LoadThemeColor PDIS_SelectedFill, "SelectedFill", IDE_GRAY
        .LoadThemeColor PDIS_SelectedBorder, "SelectedBorder", IDE_GRAY
        .LoadThemeColor PDIS_UnselectedFill, "UnselectedFill", IDE_GRAY
        .LoadThemeColor PDIS_UnselectedBorder, "UnselectedBorder", IDE_GRAY
        .LoadThemeColor PDIS_Separator, "Separator", IDE_BLACK
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    If ucSupport.ThemeUpdateRequired Then
    
        UpdateColorList
        UserControl.BackColor = m_Colors.RetrieveColor(PDIS_Background, Me.Enabled)
    
        'Reset any resource DIBs, which will force us to regenerate them against new theme settings the
        ' next time we need to render them.
        Set m_CloseIconRed = Nothing
        Set m_CloseIconGray = Nothing
        Set m_CloseIconShadow = Nothing
        Set m_ModifiedIcon = Nothing
        
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        
    End If
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
