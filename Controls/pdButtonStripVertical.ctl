VERSION 5.00
Begin VB.UserControl pdButtonStripVertical 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
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
   ScaleHeight     =   51
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   183
   ToolboxBitmap   =   "pdButtonStripVertical.ctx":0000
End
Attribute VB_Name = "pdButtonStripVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Button Strip Vertical" control
'Copyright 2015-2026 by Tanner Helland
'Created: 15/March/15
'Last updated: 29/April/20
'Last update: migrate renderer to pd2D
'
'In a surprise to precisely no one, PhotoDemon has some unique needs when it comes to user controls - needs that
' the intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize" property for
' anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the goal of maintaining
' visual fidelity across the program, while also enabling key features like Unicode support.
'
'A few notes on this button strip control, specifically:
'
' 1) The control supports an arbitrary number of button captions.  Captions are auto-wrapped, but DrawText requires word breaks to do this,
'     so you can't rely on hyphenation for over-long words - plan accordingly!
' 2) High DPI settings are handled automatically.
' 3) A hand cursor is automatically applied, and clicks on individual buttons are returned via the Click event.
' 4) Coloration is automatically handled by PD's internal theming engine.
' 5) When the control receives focus via keyboard, a special focus rect is drawn.  Focus via mouse is conveyed via text glow.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This control really only needs one event raised - Click
Public Event Click(ByVal buttonIndex As Long)

'These events are provided as a convenience, for hosts who may want to reroute mousewheel events to some other control.
' (In PD, the metadata browser does this.)
Public Event MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()
Public Event SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, ByRef newTargetHwnd As Long)

'Rather than use an StdFont container (which requires VB to create redundant font objects), we track font properties manually,
' via dedicated properties.
Private m_FontSize As Single
Private m_FontBold As Boolean

'Now that this control supports a title caption (which sits above the button itself), we separately track the region of the
' control corresponding to the "buttonstrip" only.
Private m_ButtonStripRect As RECT

'Current button indices
Private m_ButtonIndex As Long
Private m_ButtonHoverIndex As Long
Private m_ButtonMouseDown As Long

'Array of current button entries
Private Type ButtonEntry
    btCaptionEn As String           'Current button caption, in its original English
    btCaptionTranslated As String   'Current button caption, translated into the active language (if English is active, this is a copy of btCaptionEn)
    btBounds As RECT                'Boundaries of this button (full clickable area, inclusive - meaning 1px border NOT included)
    btCaptionRect As RECT           'Bounding rect of the caption.  This is dynamically calculated by the UpdateControlLayout function
    btImages As pdDIB               'Optional image(s) to use with the button; this class is ignored if the button has no image
    btImageWidth As Long            'If an image is loaded, these will store the image's width and height
    btImageHeight As Long
    btImageCoords As PointAPI       '(x, y) position of the button image, if any
    btFontSize As Single            'If a button caption fits just fine, this value is 0.  If a translation is active and a button caption must be shrunk,
                                    ' this value will be non-zero, and the button renderer must use it when rendering the button.
    btToolTipText As String         'This control supports per-button tooltips.  This behavior can be overridden by not supplying an index to the
                                    ' AssignTooltip function.
    btToolTipTitle As String        'See above comments for btToolTipText
End Type

Private m_Buttons() As ButtonEntry
Private m_numOfButtons As Long

'Index of which button has the focus.  The user can use arrow keys to move focus between buttons.
Private m_FocusRectActive As Long

'To improve rendering performance, we suspend layout updates until the control is actually visible
Private m_LayoutNeedsUpdate As Boolean

'To prevent over-frequent tooltip updates, we track the last index we received and ignore matching requests
Private m_LastToolTipIndex As Long

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Padding between images (if any) and text.  This is automatically adjusted according to DPI, so set this value as it would be at the
' Windows default of 96 DPI
Private Const IMG_TEXT_PADDING As Long = 8

'Unlike horizontal button strips, the vertical button strip forces its images into a single, continuous alignment.  Because of this,
' the addition of one image causes all BUTTONS to align differently.
Private m_ImagesActive As Boolean, m_ImageSize As Long

'Local list of themable colors.  This list includes all potential colors used by the control, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum BTSV_COLOR_LIST
    [_First] = 0
    BTS_Background = 0
    BTS_SelectedItemFill = 1
    BTS_UnselectedItemFill = 2
    BTS_SelectedItemBorder = 3
    BTS_UnselectedItemBorder = 4
    BTS_SelectedText = 5
    BTS_UnselectedText = 6
    [_Last] = 6
    [_Count] = 7
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_ButtonStripVertical
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
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
    RedrawBackBuffer
End Property

'Font settings are handled via dedicated properties, to avoid the need for an internal VB font object
Public Property Get FontBold() As Boolean
    FontBold = m_FontBold
End Property

Public Property Let FontBold(ByVal newBoldSetting As Boolean)
    If (newBoldSetting <> m_FontBold) Then
        m_FontBold = newBoldSetting
        If PDMain.IsProgramRunning() Then UpdateControlLayout
        PropertyChanged "FontBold"
    End If
End Property

Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    If (newSize <> m_FontSize) Then
        m_FontSize = newSize
        UpdateControlLayout
        PropertyChanged "FontSize"
    End If
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
    If (Not ucSupport.IsMouseInside) Then
        m_FocusRectActive = m_ButtonIndex
        RedrawBackBuffer
    End If
    
    RaiseEvent GotFocusAPI
    
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
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
    
    If (vkCode = VK_DOWN) Then
        
        'Keyboard now takes precedence over mouse
        If (m_ButtonHoverIndex >= 0) Then m_FocusRectActive = m_ButtonHoverIndex
        m_ButtonHoverIndex = -1
        
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
        
    ElseIf (vkCode = VK_UP) Then
        
        'Keyboard now takes precedence over mouse
        If (m_ButtonHoverIndex >= 0) Then m_FocusRectActive = m_ButtonHoverIndex
        m_ButtonHoverIndex = -1
        
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

'To improve responsiveness, MouseDown is used instead of Click
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

    Dim mouseClickIndex As Long
    mouseClickIndex = IsMouseOverButton(x, y)
    
    'Disable any keyboard-generated focus rectangles
    m_FocusRectActive = -1
    
    If Me.Enabled And (mouseClickIndex >= 0) Then
        If m_ButtonIndex <> mouseClickIndex Then
            ListIndex = mouseClickIndex
        End If
        m_ButtonMouseDown = mouseClickIndex
    Else
        m_ButtonMouseDown = -1
    End If
    
    RedrawBackBuffer

End Sub

'When the mouse leaves the UC, we must repaint the caption (as it's no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_ButtonHoverIndex = -1
    m_ButtonMouseDown = -1
    RedrawBackBuffer
    ucSupport.RequestCursor IDC_DEFAULT
    UpdateCursor -100, -100
End Sub

'When the mouse enters the clickable portion of the UC, we must repaint the caption (to reflect its hovered state)
Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    UpdateCursor x, y
    
    'If the mouse is over the relevant portion of the user control, display the cursor as clickable
    Dim mouseHoverIndex As Long
    mouseHoverIndex = IsMouseOverButton(x, y)
    
    'Only refresh the control if the hover value has changed
    If (mouseHoverIndex <> m_ButtonHoverIndex) Then
    
        m_ButtonHoverIndex = mouseHoverIndex
        
        'Synchronize the tooltip accordingly.
        SynchronizeToolTipToIndex m_ButtonHoverIndex
        
        'If the mouse is not currently hovering a button, set a default arrow cursor and exit
        If (mouseHoverIndex = -1) Then
            ucSupport.RequestCursor IDC_DEFAULT
        Else
            ucSupport.RequestCursor IDC_HAND
        End If
        
        RedrawBackBuffer
    
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
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
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
Attribute hWnd.VB_UserMemId = -515
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
    If m_ButtonIndex <> newIndex Then
    
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
    If itemIndex = -1 Then itemIndex = m_numOfButtons
    
    'Increase the button count and resize the array to match
    m_numOfButtons = m_numOfButtons + 1
    ReDim Preserve m_Buttons(0 To m_numOfButtons - 1) As ButtonEntry
    
    'Shift all buttons above this one upward, as necessary.
    If itemIndex < m_numOfButtons - 1 Then
    
        Dim i As Long
        For i = m_numOfButtons - 1 To itemIndex Step -1
            m_Buttons(i) = m_Buttons(i - 1)
        Next i
    
    End If
    
    'Copy the new button into place
    m_Buttons(itemIndex).btCaptionEn = srcString
    
    'We must also determine a translated caption, if any
    If Not (g_Language Is Nothing) Then
    
        If g_Language.TranslationActive Then
            m_Buttons(itemIndex).btCaptionTranslated = g_Language.TranslateMessage(m_Buttons(itemIndex).btCaptionEn)
        Else
            m_Buttons(itemIndex).btCaptionTranslated = m_Buttons(itemIndex).btCaptionEn
        End If
    
    Else
        m_Buttons(itemIndex).btCaptionTranslated = m_Buttons(itemIndex).btCaptionEn
    End If
    
    'Erase any images previously assigned to this button
    With m_Buttons(itemIndex)
        Set .btImages = Nothing
        .btImageWidth = 0
        .btImageHeight = 0
    End With
    
    'Before we can redraw the control, we need to recalculate all button positions - do that now!
    UpdateControlLayout

End Sub

'Assign a DIB to a button entry.  Disabled and hover states are automatically generated.
Public Sub AssignImageToItem(ByVal itemIndex As Long, Optional ByRef resName As String = vbNullString, Optional ByRef srcDIB As pdDIB, Optional ByVal imgWidth As Long = 0, Optional ByVal imgHeight As Long = 0, Optional ByVal resampleAlgorithm As GP_InterpolationMode = GP_IM_HighQualityBicubic, Optional ByVal usePDResamplerInstead As PD_ResamplingFilter = rf_Automatic)
    
    'Load the requested resource DIB, as necessary
    If (imgWidth = 0) Then imgWidth = 32
    If (imgHeight = 0) Then imgHeight = 32
    If (LenB(resName) <> 0) Then IconsAndCursors.LoadResourceToDIB resName, srcDIB, imgWidth, imgHeight, resampleAlgorithm:=resampleAlgorithm, usePDResamplerInstead:=usePDResamplerInstead
    
    'Cache the width and height of the DIB; it serves as our reference measurements for subsequent blt operations.
    ' (We also check for these != 0 to verify that an image was successfully loaded.)
    m_Buttons(itemIndex).btImageWidth = srcDIB.GetDIBWidth
    m_Buttons(itemIndex).btImageHeight = srcDIB.GetDIBHeight
    
    'Create a vertical sprite-sheet DIB, and mark it as having premultiplied alpha
    If (m_Buttons(itemIndex).btImages Is Nothing) Then Set m_Buttons(itemIndex).btImages = New pdDIB
    
    With m_Buttons(itemIndex)
        .btImages.CreateBlank .btImageWidth, .btImageHeight * 3, srcDIB.GetDIBColorDepth, 0, 0
        .btImages.SetInitialAlphaPremultiplicationState True
    
        'Copy this normal-state DIB into place at the top of the sheet
        GDI.BitBltWrapper .btImages.GetDIBDC, 0, 0, .btImageWidth, .btImageHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Next, make a copy of the source DIB.
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateFromExistingDIB srcDIB
        
        'Convert this to a brighter, "glowing" version; we'll use this when rendering a hovered state.
        ScaleDIBRGBValues tmpDIB, UC_HOVER_BRIGHTNESS
        
        'Copy this DIB into position 2, beneath the base DIB
        GDI.BitBltWrapper .btImages.GetDIBDC, 0, .btImageHeight, .btImageWidth, .btImageHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Finally, create a grayscale copy of the original image.  This will serve as the "disabled state" copy.
        tmpDIB.CreateFromExistingDIB srcDIB
        Filters_Layers.GrayscaleDIB tmpDIB
        
        'Place it into position 3, beneath the previous two DIBs
        GDI.BitBltWrapper .btImages.GetDIBDC, 0, .btImageHeight * 2, .btImageWidth, .btImageHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
        
        'Free whatever DIBs we can.  (If the caller passed us the source DIB, we trust them to release it.)
        Set tmpDIB = Nothing
        If (LenB(resName) <> 0) Then Set srcDIB = Nothing
        .btImages.FreeFromDC
        
    End With
    
    'Note that images are now active; this causes alignment changes, so we must reflow the button strip
    m_ImagesActive = True
    If (m_Buttons(itemIndex).btImageWidth > m_ImageSize) Then
        m_ImageSize = m_Buttons(itemIndex).btImageWidth
        UpdateControlLayout
    End If

End Sub

Private Sub ucSupport_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    RaiseEvent SetCustomTabTarget(shiftTabWasPressed, newTargetHwnd)
End Sub

Private Sub ucSupport_VisibilityChange(ByVal newVisibility As Boolean)
    If m_LayoutNeedsUpdate Then UpdateControlLayout
End Sub

'INITIALIZE control
Private Sub UserControl_Initialize()
    
    m_numOfButtons = 0
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    
    'Request some additional input functionality (custom mouse and key events)
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_UP, VK_DOWN, VK_SPACE
    ucSupport.RequestCaptionSupport
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As BTSV_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "ButtonStrip", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
    'Set various UI trackers to default values.
    m_FocusRectActive = -1
    m_ButtonHoverIndex = -1
    m_ButtonMouseDown = -1
    
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    Caption = vbNullString
    FontBold = False
    FontSize = 10
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
        m_FontBold = .ReadProperty("FontBold", False)
        m_FontSize = .ReadProperty("FontSize", 10)
        FontSizeCaption = .ReadProperty("FontSizeCaption", 12!)
        m_ButtonIndex = .ReadProperty("ListIndex", 0)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

'Store all associated properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", ucSupport.GetCaptionText, vbNullString
        .WriteProperty "FontBold", m_FontBold, False
        .WriteProperty "FontSize", m_FontSize, 10
        .WriteProperty "FontSizeCaption", ucSupport.GetCaptionFontSize, 12!
        .WriteProperty "ListIndex", ListIndex, 0
    End With
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub UpdateControlLayout(Optional ByVal forceUpdate As Boolean = False)

    'If this control isn't visible, skip all control layout decisions; we'll handle them before the
    ' control is shown.
    If (Not forceUpdate) And (Not ucSupport.AmIVisible) Then
        m_LayoutNeedsUpdate = True
        Exit Sub
    End If
    
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
            .Top = 1
            .Left = 1
        End If
        .Bottom = bHeight - 1
        .Right = bWidth - 1
    End With
    
    'Reset the width/height values to match our newly calculated rect; this simplifies subsequent steps
    bWidth = m_ButtonStripRect.Right - m_ButtonStripRect.Left
    bHeight = m_ButtonStripRect.Bottom - m_ButtonStripRect.Top
    
    'We now need to figure out the size of individual buttons within the strip.  While we could make these proportional
    ' to the text length of each button, I am instead taking the simpler route for now, and making all buttons a uniform size.
    
    'Start by calculating a set size for each button.  We will calculate these as floating-point, to avoid compounded
    ' truncation errors as we move from button to button.
    Dim buttonWidth As Double, buttonHeight As Double
    
    'Button width is easy - assume a 1px border on top and bottom, and give each button access to all space in-between.
    buttonWidth = bWidth - 2
    
    'Button height is trickier.  We have a 1px border around the whole control, and then (n-1) borders on the interior.
    If m_numOfButtons > 0 Then
        buttonHeight = (bHeight - 2 - (m_numOfButtons - 1)) / m_numOfButtons
    Else
        buttonHeight = bHeight - 2
    End If
    
    'Using these values, populate a boundary rect for each button, and store it.  (This makes the render step much faster.)
    Dim i As Long
    For i = 0 To m_numOfButtons - 1
    
        With m_Buttons(i).btBounds
            '.Top is calculated as: 1px top border, plus 1px border for any preceding buttons, plus preceding button heights
            .Top = m_ButtonStripRect.Top + 1 + i + (buttonHeight * i)
            .Left = m_ButtonStripRect.Left + 1
            .Right = .Left + buttonWidth
        End With
    
    Next i
    
    'Now, we're going to do something odd.  To avoid truncation errors, we're going to dynamically calculate BOTTOM bounds
    ' by looping back through the array, and assigning bottom values to match the top value calculated for the next
    ' button in line.  The final button receives special consideration.
    If (m_numOfButtons > 0) Then
    
        m_Buttons(m_numOfButtons - 1).btBounds.Bottom = m_ButtonStripRect.Bottom - 2
        
        If (m_numOfButtons > 1) Then
            For i = 1 To m_numOfButtons - 1
                m_Buttons(i - 1).btBounds.Bottom = m_Buttons(i).btBounds.Top - 2
            Next i
        End If
        
    End If
    
    'Each button now has its boundaries precisely calculated.  Next, we want to precalculate all text positioning inside
    ' each button.  Because text positioning varies by both caption, and the presence of images, we are also going to
    ' pre-cache these values, to further reduce the amount of work we need to do in the render loop.
    Dim strWidth As Long, strHeight As Long
    
    'Rather than create and manage our own font object(s), we borrow font objects from the global PD font cache.
    Dim tmpFont As pdFont
    
    For i = 0 To m_numOfButtons - 1
    
        'Reset font size for this button
        m_Buttons(i).btFontSize = 0
        
        'Calculate the width of this button
        buttonWidth = m_Buttons(i).btBounds.Right - m_Buttons(i).btBounds.Left
        
        'If a button has an image, we have to alter its sizing somewhat.  To make sure word-wrap is calculated correctly,
        ' remove the width of the image, plus padding, in advance.
        If m_ImagesActive Then
            buttonWidth = buttonWidth - (m_ImageSize + FixDPI(IMG_TEXT_PADDING) * 2)
        End If
        
        'Retrieve the expected size of the string, in pixels
        Set tmpFont = Fonts.GetMatchingUIFont(m_FontSize, m_FontBold)
        strWidth = tmpFont.GetWidthOfString(m_Buttons(i).btCaptionTranslated)
        
        'If the string is too long for its containing button, activate word wrap and measure again
        If strWidth > buttonWidth Then
            
            strWidth = buttonWidth
            strHeight = tmpFont.GetHeightOfWordwrapString(m_Buttons(i).btCaptionTranslated, strWidth)
            
            'As a failsafe for ultra-long captions, restrict their size to the button size.  Truncation will (necessarily) occur.
            If (strHeight > buttonHeight) Then
                strHeight = buttonHeight
                
            'As a second failsafe, if word-wrapping didn't solve the problem (because the text is a single word, for example, as is common
            ' in German), we will forcibly set a smaller font size for this caption alone.
            ElseIf tmpFont.GetHeightOfWordwrapString(m_Buttons(i).btCaptionTranslated, strWidth) = tmpFont.GetHeightOfString(m_Buttons(i).btCaptionTranslated) Then
                m_Buttons(i).btFontSize = tmpFont.GetMaxFontSizeToFitStringWidth(m_Buttons(i).btCaptionTranslated, buttonWidth, m_FontSize)
                Set tmpFont = Fonts.GetMatchingUIFont(m_Buttons(i).btFontSize, m_FontBold)
                strHeight = tmpFont.GetHeightOfString(m_Buttons(i).btCaptionTranslated)
            End If
            
        Else
            strHeight = tmpFont.GetHeightOfString(m_Buttons(i).btCaptionTranslated)
        End If
        
        'Release our copy of this global PD UI font
        Set tmpFont = Nothing
        
        'Use the size of the string, the size of the button's image (if any), and the size of the button itself to determine
        ' optimal painting position (using top-left alignment).
        With m_Buttons(i)
        
            'No image...
            If Not m_ImagesActive Then
                .btCaptionRect.Left = .btBounds.Left
            
            'Image...
            Else
                .btCaptionRect.Left = .btBounds.Left + m_ImageSize + FixDPI(IMG_TEXT_PADDING) * 2
                
            End If
            
            .btCaptionRect.Top = .btBounds.Top + (buttonHeight - strHeight) \ 2
            .btCaptionRect.Right = .btBounds.Right
            .btCaptionRect.Bottom = .btBounds.Bottom
        
            'Calculate a position for button images, if any
            If m_ImagesActive Then
                .btImageCoords.x = .btBounds.Left + FixDPI(IMG_TEXT_PADDING)
                .btImageCoords.y = .btBounds.Top + (buttonHeight - m_ImageSize) \ 2
            End If
        
        End With
        
    Next i
    
    'With all metrics successfully measured, we can now recreate the back buffer
    m_LayoutNeedsUpdate = False
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
    bufferDC = ucSupport.GetBackBufferDC(True)
    If (bufferDC = 0) Then Exit Sub
    
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight

    'NOTE: if a title caption exists, it has already been drawn.  We just need to draw the clickable button portion.
    
    'To improve rendering performance, we cache all colors locally prior to rendering
    Dim btnColorBackground As Long
    Dim btnColorSelectedBorder As Long, btnColorSelectedFill As Long
    Dim btnColorSelectedBorderHover As Long, btnColorSelectedFillHover As Long
    Dim btnColorUnselectedBorder As Long, btnColorUnselectedFill As Long
    Dim btnColorUnselectedBorderHover As Long, btnColorUnselectedFillHover As Long
    Dim fontColorSelected As Long, fontColorSelectedHover As Long
    Dim fontColorUnselected As Long, fontColorUnselectedHover As Long
    
    Dim curColor As Long
    Dim isButtonSelected As Boolean, isButtonHovered As Boolean
    Dim enabledState As Boolean
    enabledState = Me.Enabled
    
    btnColorBackground = m_Colors.RetrieveColor(BTS_Background, enabledState, False, False)
    btnColorUnselectedBorder = m_Colors.RetrieveColor(BTS_UnselectedItemBorder, enabledState, False, False)
    btnColorUnselectedFill = m_Colors.RetrieveColor(BTS_UnselectedItemFill, enabledState, False, False)
    btnColorUnselectedBorderHover = m_Colors.RetrieveColor(BTS_UnselectedItemBorder, enabledState, False, True)
    btnColorUnselectedFillHover = m_Colors.RetrieveColor(BTS_UnselectedItemFill, enabledState, False, True)
    btnColorSelectedBorder = m_Colors.RetrieveColor(BTS_SelectedItemBorder, enabledState, False, False)
    btnColorSelectedFill = m_Colors.RetrieveColor(BTS_SelectedItemFill, enabledState, False, False)
    btnColorSelectedBorderHover = m_Colors.RetrieveColor(BTS_SelectedItemBorder, enabledState, False, True)
    btnColorSelectedFillHover = m_Colors.RetrieveColor(BTS_SelectedItemFill, enabledState, False, True)
    fontColorSelected = m_Colors.RetrieveColor(BTS_SelectedText, enabledState, False, False)
    fontColorSelectedHover = m_Colors.RetrieveColor(BTS_SelectedText, enabledState, False, True)
    fontColorUnselected = m_Colors.RetrieveColor(BTS_UnselectedText, enabledState, False, False)
    fontColorUnselectedHover = m_Colors.RetrieveColor(BTS_UnselectedText, enabledState, False, True)
    
    'Each individual button is rendered in turn.  (0-button strips are not currently supported.)
    If ((m_numOfButtons > 0) And PDMain.IsProgramRunning()) Then
    
        'This control doesn't maintain its own fonts; instead, it borrows it from the public PD font cache, as necessary
        Dim tmpFont As pdFont
            
        'pd2D simplifies many rendering tasks
        Dim cSurface As pd2DSurface
        Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundDC bufferDC
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.SetSurfacePixelOffset P2_PO_Normal
        
        Dim cPen As pd2DPen
        Set cPen = New pd2DPen
        cPen.SetPenWidth 1!
        cPen.SetPenLineJoin P2_LJ_Miter
        
        Dim cBrush As pd2DBrush
        Set cBrush = New pd2DBrush
        
        'Start by filling the desired background color, then rendering a single-pixel unselected border around the control.
        ' (The border will be overwritten with Selected or Hovered borders, as necessary.)
        With m_ButtonStripRect
            cBrush.SetBrushColor btnColorBackground
            PD2D.FillRectangleI_AbsoluteCoords cSurface, cBrush, .Left, .Top, .Right - 1, .Bottom - 1
            cPen.SetPenColor btnColorUnselectedBorder
            PD2D.DrawRectangleI_AbsoluteCoords cSurface, cPen, .Left, .Top, .Right - 1, .Bottom - 1
        End With
        
        Dim i As Long
        For i = 0 To m_numOfButtons - 1
            
            If enabledState Then
                isButtonSelected = (i = m_ButtonIndex)
                isButtonHovered = (i = m_ButtonHoverIndex)
            Else
                isButtonSelected = False
                isButtonHovered = False
            End If
            
            With m_Buttons(i)
                
                'Fill the current button with its target fill color
                If isButtonSelected Then curColor = btnColorSelectedFill Else curColor = btnColorUnselectedFill
                cBrush.SetBrushColor curColor
                PD2D.FillRectangleI cSurface, cBrush, .btBounds.Left, .btBounds.Top, .btBounds.Right - .btBounds.Left, .btBounds.Bottom - .btBounds.Top + 1
                
                'For performance reasons, we only render each button's right-side border at this stage, and we always start
                ' with the inactive border color.
                If i < (m_numOfButtons - 1) Then
                    cPen.SetPenColor btnColorUnselectedBorder
                    PD2D.DrawLineI cSurface, cPen, m_ButtonStripRect.Left, .btBounds.Bottom + 1, .btBounds.Right, .btBounds.Bottom + 1
                End If
                
                'Active/hover changes are only rendered if the control is enabled
                If enabledState Then
                
                    'If this is the selected button (.ListIndex), paint its border with a special color.
                    ' (Note that we skip this step if the button is hovered, because the hover rect is drawn LAST to ensure that it overlaps
                    '  the surrounding buttons correctly.)
                    If isButtonSelected Then
                        If (Not isButtonHovered) Then
                            cPen.SetPenColor btnColorSelectedBorder
                            PD2D.DrawRectangleI_AbsoluteCoords cSurface, cPen, .btBounds.Left - 1, .btBounds.Top - 1, .btBounds.Right, .btBounds.Bottom + 1
                        End If
                    End If
                    
                End If
                
                'Paint the button's caption, if one exists
                If (LenB(.btCaptionTranslated) <> 0) Then
                
                    If isButtonSelected Then
                        If isButtonHovered Then curColor = fontColorSelectedHover Else curColor = fontColorSelected
                    Else
                        If isButtonHovered Then curColor = fontColorUnselectedHover Else curColor = fontColorUnselected
                    End If
                    
                    'Borrow a relevant UI font from the public UI font cache, then render the
                    ' button caption using the clipping rect we already calculated in previous steps.
                    
                    '(Remember that a font size of "0" means that text fits inside this button at the control's default font size)
                    If .btFontSize = 0 Then
                        Set tmpFont = Fonts.GetMatchingUIFont(m_FontSize, m_FontBold)
                        
                    'Text does not fit the button area; use the custom font size we calculated in a previous step
                    Else
                        Set tmpFont = Fonts.GetMatchingUIFont(.btFontSize, m_FontBold)
                    End If
                    
                    'Render the text onto the button
                    tmpFont.SetFontColor curColor
                    tmpFont.AttachToDC bufferDC
                    tmpFont.SetTextAlignment vbLeftJustify
                    tmpFont.DrawCenteredTextToRect .btCaptionTranslated, .btCaptionRect
                    tmpFont.ReleaseFromDC
                    
                End If
                
                'Paint the button image, if any, while branching for enabled/disabled/hovered variants
                If (Not .btImages Is Nothing) Then
                    
                    'Determine which image from the spritesheet to use.  (This is just a pixel offset.)
                    Dim pxOffset As Long
                    If enabledState Then
                        If isButtonHovered Then pxOffset = .btImageHeight Else pxOffset = 0
                    Else
                        pxOffset = .btImageHeight * 2
                    End If
                    
                    .btImages.AlphaBlendToDCEx bufferDC, .btImageCoords.x, .btImageCoords.y, .btImageWidth, .btImageHeight, 0, pxOffset, .btImageWidth, .btImageHeight
                    .btImages.FreeFromDC
                    
                End If
                
            End With
        
        'This button has been rendered successfully.  Move on to the next one.
        Next i
        
        'The hover rect (if any) is drawn last; because it's chunkier than normal borders, we must ensure that it overlaps
        ' neighboring buttons correctly.
        If ((m_ButtonHoverIndex >= 0) Or (m_FocusRectActive >= 0) Or ucSupport.DoIHaveFocus) Then
        
            'Color changes when the active button is hovered, or when we have keyboard focus and
            ' the user is attempting to change button via arrow keys.
            curColor = btnColorSelectedBorderHover
            
            Dim targetIndex As Long
            If (m_ButtonHoverIndex >= 0) Then
                targetIndex = m_ButtonHoverIndex
            ElseIf (m_FocusRectActive >= 0) Then
                targetIndex = m_FocusRectActive
            Else
                targetIndex = m_ButtonIndex
            End If
            
            With m_Buttons(targetIndex).btBounds
                cPen.SetPenWidth 3!
                cPen.SetPenColor curColor
                PD2D.DrawRectangleI_AbsoluteCoords cSurface, cPen, .Left - 1, .Top - 1, .Right, .Bottom + 1
            End With
            
        End If
        
    End If
    
    Set cSurface = Nothing
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating all button captions against the current language.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    If ucSupport.ThemeUpdateRequired Then
        
        'Determine if translations are active.  If they are, retrieve translated captions for all buttons within the control.
        If PDMain.IsProgramRunning() Then
            
            'See if translations are necessary.
            Dim isTranslationActive As Boolean
                
            If Not (g_Language Is Nothing) Then
                isTranslationActive = g_Language.TranslationActive()
            Else
                isTranslationActive = False
            End If
            
            'Apply the new translations, if any.
            Dim i As Long
            For i = 0 To m_numOfButtons - 1
                If isTranslationActive Then
                    m_Buttons(i).btCaptionTranslated = g_Language.TranslateMessage(m_Buttons(i).btCaptionEn)
                Else
                    m_Buttons(i).btCaptionTranslated = m_Buttons(i).btCaptionEn
                End If
            Next i
            
        End If
        
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
        .LoadThemeColor BTS_Background, "Background", IDE_WHITE
        .LoadThemeColor BTS_SelectedItemFill, "SelectedItemFill", IDE_BLUE
        .LoadThemeColor BTS_UnselectedItemFill, "UnselectedItemFill", IDE_WHITE
        .LoadThemeColor BTS_SelectedItemBorder, "SelectedItemBorder", IDE_BLUE
        .LoadThemeColor BTS_UnselectedItemBorder, "UnselectedItemBorder", IDE_BLUE
        .LoadThemeColor BTS_SelectedText, "SelectedText", IDE_WHITE
        .LoadThemeColor BTS_UnselectedText, "UnselectedText", IDE_GRAY
    End With
    
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal toolTipIndex As Long = -1)
    
    'If toolTipIndex = -1, all buttons receive the same tooltip
    If toolTipIndex = -1 Then
    
        Dim i As Long
        For i = LBound(m_Buttons) To UBound(m_Buttons)
            With m_Buttons(i)
                .btToolTipText = newTooltip
                .btToolTipTitle = newTooltipTitle
            End With
        Next i
        
        'Update the index to 0, so the subsequent call to synchronizeToolTipToIndex doesn't fail.
        toolTipIndex = 0
        
    'If an index is specified, each button is allowed its own tooltip.  This can be used to set tooltips for all buttons but one,
    ' for example.
    Else
        
        If (toolTipIndex >= LBound(m_Buttons)) And (toolTipIndex <= UBound(m_Buttons)) Then
        
            With m_Buttons(toolTipIndex)
                .btToolTipText = newTooltip
                .btToolTipTitle = newTooltipTitle
            End With
            
        End If
                
    End If
    
    'Synchronize the tooltip object to the new tooltip.  (This is now handled manually, during mouse events.)
        
End Sub

Private Sub SynchronizeToolTipToIndex(Optional ByVal srcIndex As Long = 0)

    'Ignore invalid index requests
    If (srcIndex >= LBound(m_Buttons)) And (srcIndex <= UBound(m_Buttons)) And (srcIndex <> m_LastToolTipIndex) Then
    
        'Manually sync the tooltip now
        ucSupport.AssignTooltip Me.ContainerHwnd, m_Buttons(srcIndex).btToolTipText, m_Buttons(srcIndex).btToolTipTitle
        
        'Update the cached last index value
        m_LastToolTipIndex = srcIndex
    
    End If

End Sub
