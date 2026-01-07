VERSION 5.00
Begin VB.UserControl pdStrip 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   260
   ToolboxBitmap   =   "pdStrip.ctx":0000
End
Attribute VB_Name = "pdStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Owner-Drawn Strip control
'Copyright 2017-2026 by Tanner Helland
'Created: 13/February/17
'Last updated: 13/February/17
'Last update: initial build (split off from the original pdButtonStrip control)
'
'This control is very similar to the pdButtonStrip control, except that control entries are owner-drawn.
' Some of the specific rendering techniques (e.g. hover behavior) have also been tweaked to work better against
' unpredictable strip contents.
'
'At present, this control is only used in the Theme selection dialog, to render the accent color options
' available to the user.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Event Click(ByVal buttonIndex As Long)
Public Event DrawButton(ByVal btnIndex As Long, ByVal btnValue As String, ByVal targetDC As Long, ByVal ptrToRectF As Long)

'These events are provided as a convenience, for hosts who may want to reroute mousewheel events to some other control.
' (In PD, the metadata browser does this.)
Public Event MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Now that this control supports a title caption (which sits above the button itself), we separately track the region of the
' control corresponding to the "buttonstrip" only.
Private m_ButtonStripRect As RECT

'Current button indices
Private m_ButtonIndex As Long
Private m_ButtonHoverIndex As Long
Private m_ButtonMouseDown As Long

'Array of current button entries
Private Type ButtonEntry
    btData As String                'Current button data, as supplied by the user (usually a hint to help with rendering)
    btBounds As RECT                'Boundaries of this button (full clickable area, inclusive - meaning 1px border NOT included)
End Type

Private m_Buttons() As ButtonEntry
Private m_numOfButtons As Long

'Index of which button has the focus.  The user can use arrow keys to move focus between buttons.
Private m_FocusRectActive As Long

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by the control, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDSTRIP_COLOR_LIST
    [_First] = 0
    PDS_Background = 0
    PDS_Caption = 1
    PDS_Border = 2
    [_Last] = 2
    [_Count] = 3
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_Strip
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if it ever becomes relevant.
Public Property Get Caption() As String
    Caption = ucSupport.GetCaptionText()
End Property

Public Property Let Caption(ByRef newCaption As String)
    ucSupport.SetCaptionText newCaption
    PropertyChanged "Caption"
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
    RedrawBackBuffer
End Property

Public Property Get FontSizeCaption() As Single
    FontSizeCaption = ucSupport.GetCaptionFontSize()
End Property

Public Property Let FontSizeCaption(ByVal newSize As Single)
    ucSupport.SetCaptionFontSize newSize
    PropertyChanged "FontSizeCaption"
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

'When the control receives focus, if the focus isn't received via mouse click, display a focus rect around the active button
Private Sub ucSupport_GotFocusAPI()
    
    'If the mouse is *not* over the user control, assume focus was set via keyboard
    If (Not ucSupport.DoIHaveFocus) Then
        m_FocusRectActive = m_ButtonIndex
        RedrawBackBuffer
    End If
    
    RaiseEvent GotFocusAPI
    
End Sub

'When the control loses focus, erase any focus rects it may have active
Private Sub ucSupport_LostFocusAPI()
    
    'If a focus rect has been drawn, remove it now
    If (m_FocusRectActive >= 0) Then
        m_FocusRectActive = -1
        RedrawBackBuffer
    End If
    
    RaiseEvent LostFocusAPI

End Sub

'A few key events are also handled
Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    markEventHandled = False
    
    If (vkCode = VK_RIGHT) Then
        
        'See if a focus rect is already active
        If (m_FocusRectActive >= 0) Then
            m_FocusRectActive = m_FocusRectActive + 1
        Else
            m_FocusRectActive = m_ButtonIndex + 1
        End If
        
        'Bounds-check the new m_FocusRectActive index
        If (m_FocusRectActive >= m_numOfButtons) Then m_FocusRectActive = 0
        
        'Redraw the button strip
        RedrawBackBuffer
        
        markEventHandled = True
        
    ElseIf (vkCode = VK_LEFT) Then
    
        'See if a focus rect is already active
        If (m_FocusRectActive >= 0) Then
            m_FocusRectActive = m_FocusRectActive - 1
        Else
            m_FocusRectActive = m_ButtonIndex - 1
        End If
        
        'Bounds-check the new m_FocusRectActive index
        If (m_FocusRectActive < 0) Then m_FocusRectActive = m_numOfButtons - 1
        
        'Redraw the button strip
        RedrawBackBuffer
        
        markEventHandled = True
        
    'If a focus rect is active, and space is pressed, activate the button with focus
    ElseIf (vkCode = VK_SPACE) Then

        If (m_FocusRectActive >= 0) Then
            ListIndex = m_FocusRectActive
            markEventHandled = True
        End If
        
    End If

End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

'To improve responsiveness, MouseDown is used instead of Click
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
        
    Dim mouseClickIndex As Long
    mouseClickIndex = IsMouseOverButton(x, y)
    
    'Disable any keyboard-generated focus rectangles
    m_FocusRectActive = -1
    
    If Me.Enabled And (mouseClickIndex >= 0) Then
        If (m_ButtonIndex <> mouseClickIndex) Then Me.ListIndex = mouseClickIndex
        m_ButtonMouseDown = mouseClickIndex
    Else
        m_ButtonMouseDown = -1
    End If
    
    RedrawBackBuffer

End Sub

'When the mouse leaves the UC, we must repaint the control (as it's no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_ButtonHoverIndex = -1
    m_ButtonMouseDown = -1
    RedrawBackBuffer
    ucSupport.RequestCursor IDC_DEFAULT
    UpdateCursor -100, -100
End Sub

'When the mouse enters the clickable portion of the UC, we must repaint the hovered button
Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    UpdateCursor x, y
    
    'If the mouse is over the relevant portion of the user control, display the cursor as clickable
    Dim mouseHoverIndex As Long
    mouseHoverIndex = IsMouseOverButton(x, y)
    
    'Only refresh the control if the hover value has changed
    If (mouseHoverIndex <> m_ButtonHoverIndex) Then
    
        m_ButtonHoverIndex = mouseHoverIndex
    
        'If the mouse is not currently hovering a button, set a default arrow cursor and exit
        If (mouseHoverIndex = -1) Then
            ucSupport.RequestCursor IDC_ARROW
            RedrawBackBuffer
        Else
            ucSupport.RequestCursor IDC_HAND
            RedrawBackBuffer
        End If
    
    End If
    
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    m_ButtonMouseDown = -1
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RaiseEvent MouseWheelVertical(Button, Shift, x, y, scrollAmount)
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

'See if the mouse is over the clickable portion of the control
Private Function IsMouseOverButton(ByVal mouseX As Single, ByVal mouseY As Single) As Long
    
    'Search each set of button coords, looking for a match
    Dim i As Long
    For i = 0 To m_numOfButtons - 1
    
        If PDMath.IsPointInRect(mouseX, mouseY, m_Buttons(i).btBounds) Then
            IsMouseOverButton = i
            Exit Function
        End If
    
    Next i
    
    'No match was found; return -1
    IsMouseOverButton = -1

End Function

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

'The most relevant part of this control is this ListIndex property, which just like listboxes, controls which button in the strip
' is currently pressed.
Public Property Get ListIndex() As Long
    ListIndex = m_ButtonIndex
End Property

Public Property Let ListIndex(ByVal newIndex As Long)
    
    'Update our internal value tracker
    If (m_ButtonIndex <> newIndex) Then
    
        m_ButtonIndex = newIndex
        PropertyChanged "ListIndex"
        
        'Redraw the control; it's important to do this *before* raising the associated event, to maintain an impression of max responsiveness
        RedrawBackBuffer
        
        'Notify the user of the change by raising the CLICK event
        RaiseEvent Click(newIndex)
        
    End If
    
End Property

'ListCount is necessary for the command bar's "Randomize" feature
Public Property Get ListCount() As Long
    ListCount = m_numOfButtons
End Property

'To simplify translation handling, this public sub is used to add buttons to the UC.  An optional index can also be specified.
Public Sub AddItem(ByVal srcString As String, Optional ByVal itemIndex As Long = -1)

    'If an index was not requested, force the index to the current number of parameters.
    If (itemIndex = -1) Then itemIndex = m_numOfButtons
    
    'Increase the button count and resize the array to match
    m_numOfButtons = m_numOfButtons + 1
    ReDim Preserve m_Buttons(0 To m_numOfButtons - 1) As ButtonEntry
    
    'Shift all buttons above this one upward, as necessary.
    If (itemIndex < m_numOfButtons - 1) Then
        Dim i As Long
        For i = m_numOfButtons - 1 To itemIndex Step -1
            m_Buttons(i) = m_Buttons(i - 1)
        Next i
    End If
    
    'Copy the new button into place
    m_Buttons(itemIndex).btData = srcString
    
    'Before we can redraw the control, we need to recalculate all button positions - do that now!
    UpdateControlLayout

End Sub

'INITIALIZE control
Private Sub UserControl_Initialize()
    
    m_numOfButtons = 0
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    
    'Request some additional input functionality (custom mouse and key events)
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_RIGHT, VK_LEFT, VK_SPACE
    ucSupport.RequestCaptionSupport
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDSTRIP_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDHistory", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
    'Set various UI trackers to default values.
    m_FocusRectActive = -1
    m_ButtonHoverIndex = -1
    m_ButtonMouseDown = -1
                    
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    Caption = vbNullString
    FontSizeCaption = 12!
    ListIndex = 0
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Caption = .ReadProperty("Caption", vbNullString)
        FontSizeCaption = .ReadProperty("FontSizeCaption", 12#)
        ListIndex = .ReadProperty("ListIndex", 0)
    End With
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub

'Store all associated properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", ucSupport.GetCaptionText, vbNullString
        .WriteProperty "FontSizeCaption", ucSupport.GetCaptionFontSize, 12#
        .WriteProperty "ListIndex", ListIndex, 0
    End With
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Next, determine the positioning of the caption, if present.  (ucSupport.GetCaptionBottom tells us where the
    ' caption text ends vertically.)
    With m_ButtonStripRect
        If ucSupport.IsCaptionActive Then
            .Top = ucSupport.GetCaptionBottom + 2
            .Left = FixDPI(8)
        Else
            .Top = 2
            .Left = 2
        End If
        .Bottom = bHeight - 1
        .Right = bWidth - 3
    End With
    
    'Reset the width/height values to match our newly calculated rect; this simplifies subsequent steps
    bWidth = m_ButtonStripRect.Right - m_ButtonStripRect.Left
    bHeight = m_ButtonStripRect.Bottom - m_ButtonStripRect.Top
    
    'We now need to figure out the size of individual buttons within the strip.  While we could make these proportional
    ' to the text length of each button, I am instead taking the simpler route for now, and making all buttons a uniform size.
    
    'Start by calculating a set size for each button.  We will calculate these as floating-point, to avoid compounded
    ' truncation errors as we move from button to button.
    Dim buttonWidth As Double, buttonHeight As Double
    
    'Button height is easy - assume a 1px border on top and bottom, and give each button access to all space in-between.
    buttonHeight = bHeight - 2
    
    'Button width is trickier.  We have a 1px border around the whole control, and then (n-1) borders on the interior.
    If (m_numOfButtons > 0) Then
        buttonWidth = (bWidth - (m_numOfButtons - 1)) / m_numOfButtons
    Else
        buttonWidth = bWidth
    End If
    
    'Using these values, populate a boundary rect for each button, and store it.  (This makes the render step much faster.)
    If (m_numOfButtons > 0) Then
    
        Dim i As Long
        For i = 0 To m_numOfButtons - 1
        
            With m_Buttons(i).btBounds
                '.Left is calculated as: 1px border for any preceding buttons, plus preceding button widths
                .Left = m_ButtonStripRect.Left + i + (buttonWidth * i)
                .Top = m_ButtonStripRect.Top
                .Bottom = .Top + buttonHeight
            End With
        
        Next i
    
        'Now, we're going to do something odd.  To avoid truncation errors, we're going to dynamically calculate RIGHT bounds
        ' by looping back through the array, and assigning right values to match the left value calculated for the next
        ' button in line.  The final button receives special consideration.
        m_Buttons(m_numOfButtons - 1).btBounds.Right = m_ButtonStripRect.Right
        
        If (m_numOfButtons > 1) Then
            For i = 1 To m_numOfButtons - 1
                m_Buttons(i - 1).btBounds.Right = m_Buttons(i).btBounds.Left
            Next i
        End If
        
    End If
    
    'With all metrics successfully measured, we can now recreate the back buffer
    If ucSupport.AmIVisible Then RedrawBackBuffer
            
End Sub

'Because the control may consist of a non-clickable region (the caption) and a clickable region (the buttonstrip),
' we can't blindly assign a hand cursor to the entire control.
Private Sub UpdateCursor(ByVal x As Single, ByVal y As Single)
    If PDMath.IsPointInRect(x, y, m_ButtonStripRect) Then
        ucSupport.RequestCursor IDC_HAND
    Else
        ucSupport.RequestCursor IDC_DEFAULT
    End If
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDS_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
    
    If PDMain.IsProgramRunning() Then
        
        Dim tmpRectF As RectF
        Dim i As Long
        
        'Because this control is owner-drawn, our owner is responsible for drawing the individual buttons.
        If (m_numOfButtons > 0) Then
            
            For i = 0 To m_numOfButtons - 1
                
                'We shrink the display area by one pixel to account for borders.  This isn't strictly necessary
                ' (as we overpaint borders anyway), but it allows the painter to strictly fit samples inside a
                ' known area, without worrying about edge pixels getting erased.
                With m_Buttons(i).btBounds
                    tmpRectF.Left = .Left
                    tmpRectF.Top = .Top
                    tmpRectF.Width = .Right - .Left
                    tmpRectF.Height = .Bottom - .Top
                End With
                
                RaiseEvent DrawButton(i, m_Buttons(i).btData, bufferDC, VarPtr(tmpRectF))
                
            Next i
        
        End If
        
        'Next, draw a grid around the rendered buttons
        Dim cSurface As pd2DSurface, cPen As pd2DPen
        If (m_numOfButtons > 0) Then
            
            Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, True
            Drawing2D.QuickCreateSolidPen cPen, 1#, m_Colors.RetrieveColor(PDS_Border, Me.Enabled), 100#
        
            For i = 0 To m_numOfButtons - 1
                With m_Buttons(i).btBounds
                    PD2D.DrawRectangleF cSurface, cPen, .Left, .Top, .Right - .Left, .Bottom - .Top
                End With
            Next i
            
            'To improve the quality of rendering the final UI elements, switch to high-quality pixel
            ' coordinate calculations.
            cSurface.SetSurfacePixelOffset P2_PO_Half
            
            'Draw a highlight around the current listindex, if any
            If (m_ButtonIndex >= 0) Then
                
                With m_Buttons(m_ButtonIndex).btBounds
                    cPen.ReleasePen
                    Drawing2D.QuickCreateSolidPen cPen, 4#, vbBlack, 100#
                    PD2D.DrawRectangleF_AbsoluteCoords cSurface, cPen, .Left, .Top, .Right, .Bottom
                End With
                
                With m_Buttons(m_ButtonIndex).btBounds
                    cPen.ReleasePen
                    Drawing2D.QuickCreateSolidPen cPen, 3#, g_Themer.GetGenericUIColor(UI_Accent), 100#
                    PD2D.DrawRectangleF_AbsoluteCoords cSurface, cPen, .Left, .Top, .Right, .Bottom
                End With
                
            End If
            
            'Finally, if one of the buttons is currently hovered, paint it with a chunky, highlighted border
            If (m_ButtonHoverIndex >= 0) Then
                
                With m_Buttons(m_ButtonHoverIndex).btBounds
                    tmpRectF.Left = .Left
                    tmpRectF.Top = .Top
                    tmpRectF.Width = .Right - .Left
                    tmpRectF.Height = .Bottom - .Top
                End With
                
                Dim cOuterPen As pd2DPen
                Drawing2D.QuickCreatePairOfUIPens cOuterPen, cPen, True
                PD2D.DrawRectangleF_FromRectF cSurface, cOuterPen, tmpRectF
                PD2D.DrawRectangleF_FromRectF cSurface, cPen, tmpRectF
                Set cOuterPen = Nothing
                
            End If
            
        End If
        
        Set cSurface = Nothing: Set cPen = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating all button captions against the current language.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    If ucSupport.ThemeUpdateRequired Then
        
        'This control requests quite a few colors from the central themer; update its color cache now
        UpdateColorList
        
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        
        'Update all text managed by the support class (e.g. tooltips)
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        
    End If
    
End Sub

'Before the control is rendered, we need to retrieve all painting colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    
    'Color list retrieval is pretty darn easy - just load each color one at a time, and leave the rest to the color class.
    ' It will build an internal hash table of the colors we request, which makes rendering much faster.
    With m_Colors
        .LoadThemeColor PDS_Background, "Background", IDE_WHITE
        .LoadThemeColor PDS_Caption, "Caption", IDE_GRAY
        .LoadThemeColor PDS_Border, "Border", IDE_GRAY
    End With
    
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
