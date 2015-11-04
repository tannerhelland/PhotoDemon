VERSION 5.00
Begin VB.UserControl buttonStripVertical 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
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
   ToolboxBitmap   =   "buttonStripVertical.ctx":0000
End
Attribute VB_Name = "buttonStripVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "Button Strip Vertical" control
'Copyright 2015-2015 by Tanner Helland
'Created: 15/March/15
'Last updated: 04/November/15
'Last update: convert to master usercontrol support class; switch to spitesheets for button images
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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

'Rather than use an StdFont container (which requires VB to create redundant font objects), we track font properties manually,
' via dedicated properties.
Private m_FontSize As Single
Private m_FontBold As Boolean

'Current button indices
Private m_ButtonIndex As Long
Private m_ButtonHoverIndex As Long

'Array of current button entries
Private Type buttonEntry
    btCaptionEn As String           'Current button caption, in its original English
    btCaptionTranslated As String   'Current button caption, translated into the active language (if English is active, this is a copy of btCaptionEn)
    btBounds As RECT                'Boundaries of this button (full clickable area, inclusive - meaning 1px border NOT included)
    btCaptionRect As RECT           'Bounding rect of the caption.  This is dynamically calculated by the UpdateControlLayout function
    btImages As pdDIB               'Optional image(s) to use with the button; this class is ignored if the button has no image
    btImageWidth As Long            'If an image is loaded, these will store the image's width and height
    btImageHeight As Long
    btImageCoords As POINTAPI       '(x, y) position of the button image, if any
    btFontSize As Single            'If a button caption fits just fine, this value is 0.  If a translation is active and a button caption must be shrunk,
                                    ' this value will be non-zero, and the button renderer must use it when rendering the button.
    btToolTipText As String         'This control supports per-button tooltips.  This behavior can be overridden by not supplying an index to the
                                    ' AssignTooltip function.
    btToolTipTitle As String        'See above comments for btToolTipText
    btToolTipIcon As TT_ICON_TYPE   'See above comments for btToolTipText
End Type

Private m_Buttons() As buttonEntry
Private m_numOfButtons As Long

'Index of which button has the focus.  The user can use arrow keys to move focus between buttons.
Private m_FocusRectActive As Long

'To prevent over-frequent tooltip updates, we track the last index we received and ignore matching requests
Private m_LastToolTipIndex As Long

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Padding between images (if any) and text.  This is automatically adjusted according to DPI, so set this value as it would be at the
' Windows default of 96 DPI
Private Const IMG_TEXT_PADDING As Long = 8

'Unlike horizontal button strips, the vertical button strip forces its images into a single, continuous alignment.  Because of this,
' the addition of one image causes all BUTTONS to align differently.
Private m_ImagesActive As Boolean, m_ImageSize As Long

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
    If newBoldSetting <> m_FontBold Then
        m_FontBold = newBoldSetting
        UpdateControlLayout
    End If
End Property

Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    If newSize <> m_FontSize Then
        m_FontSize = newSize
        UpdateControlLayout
    End If
End Property

'When the control receives focus, if the focus isn't received via mouse click, display a focus rect around the active button
Private Sub ucSupport_GotFocusAPI()
    
    'If the mouse is *not* over the user control, assume focus was set via keyboard
    If Not ucSupport.DoIHaveFocus Then
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

    If (vkCode = VK_DOWN) Then
        
        'See if a focus rect is already active
        If (m_FocusRectActive >= 0) Then
            m_FocusRectActive = m_FocusRectActive + 1
        Else
            m_FocusRectActive = m_ButtonIndex + 1
        End If
        
        'Bounds-check the new m_FocusRectActive index
        If m_FocusRectActive >= m_numOfButtons Then m_FocusRectActive = 0
        
        'Redraw the button strip
        RedrawBackBuffer
        
    ElseIf (vkCode = VK_UP) Then
    
        'See if a focus rect is already active
        If (m_FocusRectActive >= 0) Then
            m_FocusRectActive = m_FocusRectActive - 1
        Else
            m_FocusRectActive = m_ButtonIndex - 1
        End If
        
        'Bounds-check the new m_FocusRectActive index
        If m_FocusRectActive < 0 Then m_FocusRectActive = m_numOfButtons - 1
        
        'Redraw the button strip
        RedrawBackBuffer
        
    'If a focus rect is active, and space is pressed, activate the button with focus
    ElseIf (vkCode = VK_SPACE) Then

        If m_FocusRectActive >= 0 Then ListIndex = m_FocusRectActive
        
    End If

End Sub

Private Sub ucSupport_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    RaiseEvent MouseWheelVertical(Button, Shift, x, y, scrollAmount)
End Sub

'To improve responsiveness, MouseDown is used instead of Click
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    Dim mouseClickIndex As Long
    mouseClickIndex = isMouseOverButton(x, y)
    
    'Disable any keyboard-generated focus rectangles
    m_FocusRectActive = -1
    
    If Me.Enabled And (mouseClickIndex >= 0) Then
        If m_ButtonIndex <> mouseClickIndex Then
            ListIndex = mouseClickIndex
        End If
    End If

End Sub

'When the mouse leaves the UC, we must repaint the caption (as it's no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_ButtonHoverIndex = -1
    RedrawBackBuffer
    ucSupport.RequestCursor IDC_DEFAULT
End Sub

'When the mouse enters the clickable portion of the UC, we must repaint the caption (to reflect its hovered state)
Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'If the mouse is over the relevant portion of the user control, display the cursor as clickable
    Dim mouseHoverIndex As Long
    mouseHoverIndex = isMouseOverButton(x, y)
    
    'Only refresh the control if the hover value has changed
    If mouseHoverIndex <> m_ButtonHoverIndex Then
    
        m_ButtonHoverIndex = mouseHoverIndex
        
        'Synchronize the tooltip accordingly.
        SynchronizeToolTipToIndex m_ButtonHoverIndex
        
        'If the mouse is not currently hovering a button, set a default arrow cursor and exit
        If mouseHoverIndex = -1 Then
            ucSupport.RequestCursor IDC_DEFAULT
            RedrawBackBuffer
        Else
            ucSupport.RequestCursor IDC_HAND
            RedrawBackBuffer
        End If
    
    End If
    
End Sub

'See if the mouse is over the clickable portion of the control
Private Function isMouseOverButton(ByVal mouseX As Single, ByVal mouseY As Single) As Long
    
    'Search each set of button coords, looking for a match
    Dim i As Long
    For i = 0 To m_numOfButtons - 1
    
        If Math_Functions.isPointInRect(mouseX, mouseY, m_Buttons(i).btBounds) Then
            isMouseOverButton = i
            Exit Function
        End If
    
    Next i
    
    'No match was found; return -1
    isMouseOverButton = -1

End Function

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    UpdateControlLayout
    RedrawBackBuffer
End Sub

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
    ReDim Preserve m_Buttons(0 To m_numOfButtons - 1) As buttonEntry
    
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
    
        If g_Language.translationActive Then
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
Public Sub AssignImageToItem(ByVal itemIndex As Long, Optional ByVal resName As String = "", Optional ByRef srcDIB As pdDIB)
    
    'Load the requested resource DIB, as necessary
    If Len(resName) <> 0 Then loadResourceToDIB resName, srcDIB
    
    'Cache the width and height of the DIB; it serves as our reference measurements for subsequent blt operations.
    ' (We also check for these != 0 to verify that an image was successfully loaded.)
    m_Buttons(itemIndex).btImageWidth = srcDIB.getDIBWidth
    m_Buttons(itemIndex).btImageHeight = srcDIB.getDIBHeight
    
    'Create a vertical sprite-sheet DIB, and mark it as having premultiplied alpha
    If m_Buttons(itemIndex).btImages Is Nothing Then Set m_Buttons(itemIndex).btImages = New pdDIB
    
    With m_Buttons(itemIndex)
        .btImages.createBlank .btImageWidth, .btImageHeight * 3, srcDIB.getDIBColorDepth, 0, 0
        .btImages.setInitialAlphaPremultiplicationState True
    
        'Copy this normal-state DIB into place at the top of the sheet
        BitBlt .btImages.getDIBDC, 0, 0, .btImageWidth, .btImageHeight, srcDIB.getDIBDC, 0, 0, vbSrcCopy
        
        'Next, make a copy of the source DIB.
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.createFromExistingDIB srcDIB
        
        'Convert this to a brighter, "glowing" version; we'll use this when rendering a hovered state.
        ScaleDIBRGBValues tmpDIB, UC_HOVER_BRIGHTNESS, True
        
        'Copy this DIB into position #2, beneath the base DIB
        BitBlt .btImages.getDIBDC, 0, .btImageHeight, .btImageWidth, .btImageHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
        
        'Finally, create a grayscale copy of the original image.  This will serve as the "disabled state" copy.
        tmpDIB.createFromExistingDIB srcDIB
        GrayscaleDIB tmpDIB, True
        
        'Place it into position #3, beneath the previous two DIBs
        BitBlt .btImages.getDIBDC, 0, .btImageHeight * 2, .btImageWidth, .btImageHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
        
        'Free whatever DIBs we can.  (If the caller passed us the source DIB, we trust them to release it.)
        Set tmpDIB = Nothing
        If Len(resName) <> 0 Then Set srcDIB = Nothing
    End With
    
    'Note that images are now active; this causes alignment changes, so we must reflow the button strip
    m_ImagesActive = True
    If m_Buttons(itemIndex).btImageWidth > m_ImageSize Then
        m_ImageSize = m_Buttons(itemIndex).btImageWidth
        UpdateControlLayout
    End If

End Sub

'INITIALIZE control
Private Sub UserControl_Initialize()
    
    m_numOfButtons = 0
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd
    
    'Request some additional input functionality (custom mouse and key events)
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_UP, VK_DOWN, VK_SPACE
    
    'In design mode, initialize a base theming class, so our paint functions don't fail
    If g_Themer Is Nothing Then Set g_Themer = New pdVisualThemes
    
    m_FocusRectActive = -1
    m_ButtonHoverIndex = -1
    m_LastToolTipIndex = -1
                    
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    ListIndex = 0
    m_FontBold = False
    m_FontSize = 10
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        ListIndex = .ReadProperty("ListIndex", 0)
        FontBold = .ReadProperty("FontBold", False)
        FontSize = .ReadProperty("FontSize", 10)
    End With

End Sub

Private Sub UserControl_Resize()
    If Not g_IsProgramRunning Then ucSupport.RequestRepaint True
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub UpdateControlLayout()

    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
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
            .Top = 1 + i + (buttonHeight * i)
            .Left = 1
            .Right = .Left + buttonWidth
        End With
    
    Next i
    
    'Now, we're going to do something odd.  To avoid truncation errors, we're going to dynamically calculate BOTTOM bounds
    ' by looping back through the array, and assigning bottom values to match the top value calculated for the next
    ' button in line.  The final button receives special consideration.
    If m_numOfButtons > 0 Then
    
        m_Buttons(m_numOfButtons - 1).btBounds.Bottom = bHeight - 2
        
        If m_numOfButtons > 1 Then
        
            For i = 1 To m_numOfButtons - 1
                m_Buttons(i - 1).btBounds.Bottom = m_Buttons(i).btBounds.Top - 2
            Next i
        
        End If
        
    End If
    
    'Each button now has its boundaries precisely calculated.  Next, we want to precalculate all text positioning inside
    ' each button.  Because text positioning varies by both caption, and the presence of images, we are also going to
    ' pre-cache these values, to further reduce the amount of work we need to do in the render loop.
    Dim tmpPoint As POINTAPI
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
        Set tmpFont = Font_Management.GetMatchingUIFont(m_FontSize, m_FontBold)
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
                Set tmpFont = Font_Management.GetMatchingUIFont(m_Buttons(i).btFontSize, m_FontBold)
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
    If ucSupport.AmIVisible Then RedrawBackBuffer
            
End Sub

'Store all associated properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ListIndex", ListIndex, 0
        .WriteProperty "FontBold", m_FontBold, False
        .WriteProperty "FontSize", m_FontSize, 10
    End With
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long, bWidth As Long, bHeight As Long
    bufferDC = ucSupport.GetBackBufferDC(True)
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Colors used throughout this paint function are determined primarily control enablement
    Dim btnColorActiveBorder As Long, btnColorActiveFill As Long, btnColorHoverBorder As Long
    Dim btnColorInactiveBorder As Long, btnColorInactiveFill As Long
    Dim fontColorActive As Long, fontColorInactive As Long, fontColorHover As Long
    Dim curColor As Long
    
    If Me.Enabled Then
    
        'All colors are determined by PD's central themer
        btnColorInactiveBorder = g_Themer.GetThemeColor(PDTC_GRAY_DEFAULT)
        btnColorInactiveFill = g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT)
        btnColorActiveBorder = g_Themer.GetThemeColor(PDTC_ACCENT_SHADOW)
        btnColorActiveFill = g_Themer.GetThemeColor(PDTC_ACCENT_DEFAULT)
        btnColorHoverBorder = g_Themer.GetThemeColor(PDTC_ACCENT_DEFAULT)
        
        fontColorInactive = g_Themer.GetThemeColor(PDTC_TEXT_DEFAULT)
        fontColorActive = g_Themer.GetThemeColor(PDTC_TEXT_INVERT)
        fontColorHover = g_Themer.GetThemeColor(PDTC_TEXT_HYPERLINK)
        
    Else
    
        btnColorInactiveBorder = g_Themer.GetThemeColor(PDTC_DISABLED)
        btnColorInactiveFill = g_Themer.GetThemeColor(PDTC_BACKGROUND_DEFAULT)
        btnColorActiveBorder = g_Themer.GetThemeColor(PDTC_DISABLED)
        btnColorActiveFill = g_Themer.GetThemeColor(PDTC_DISABLED)
        btnColorHoverBorder = g_Themer.GetThemeColor(PDTC_DISABLED)
        
        fontColorInactive = g_Themer.GetThemeColor(PDTC_DISABLED)
        fontColorActive = g_Themer.GetThemeColor(PDTC_TEXT_INVERT)
        fontColorHover = g_Themer.GetThemeColor(PDTC_DISABLED)
        
    End If
    
    'A single-pixel border is always drawn around the control
    GDI_Plus.GDIPlusDrawRectOutlineToDC bufferDC, 0, 0, bWidth - 1, bHeight - 1, btnColorInactiveBorder, 255, 1
    
    'This control doesn't maintain its own fonts; instead, it borrows it from the public PD font cache, as necessary
    Dim tmpFont As pdFont
    
    'Next, each individual button is rendered in turn.
    If m_numOfButtons > 0 Then
    
        Dim i As Long
        For i = 0 To m_numOfButtons - 1
        
            With m_Buttons(i)
            
                'Fill the current button with its target fill color
                If i = m_ButtonIndex Then
                    curColor = btnColorActiveFill
                Else
                    curColor = btnColorInactiveFill
                End If
                
                GDI_Plus.GDIPlusFillRectToDC bufferDC, .btBounds.Left, .btBounds.Top, .btBounds.Right - .btBounds.Left, .btBounds.Bottom - .btBounds.Top + 1, curColor
                
                'For performance reasons, we only render bottom borders
                If i < (m_numOfButtons - 1) Then
                    GDI_Plus.GDIPlusDrawLineToDC bufferDC, 0, .btBounds.Bottom + 1, bWidth, .btBounds.Bottom + 1, btnColorInactiveBorder, 255, 1
                End If
                
                'Disable the next block of rendering if the control is disabled.
                If Me.Enabled Then
                
                    'If this is the active button, paint it with a special border.
                    If i = m_ButtonIndex Then
                        GDI_Plus.GDIPlusDrawRectOutlineToDC bufferDC, .btBounds.Left - 1, .btBounds.Top - 1, .btBounds.Right, .btBounds.Bottom + 1, btnColorActiveBorder, 255, 1
                    
                    'If this control is hovered by the mouse, paint it with an extra-thick border
                    ElseIf (i = m_ButtonHoverIndex) Then
                        GDI_Plus.GDIPlusDrawRectOutlineToDC bufferDC, .btBounds.Left, .btBounds.Top, .btBounds.Right, .btBounds.Bottom + 1, btnColorHoverBorder, 255, 2, False, LineJoinMiter
                    
                    End If
                    
                    'If this button has received focus via keyboard, paint it with a special interior border
                    If i = m_FocusRectActive Then
                        GDI_Plus.GDIPlusDrawRectOutlineToDC bufferDC, .btBounds.Left + 2, .btBounds.Top + 2, .btBounds.Right - 2, .btBounds.Bottom - 2, btnColorActiveBorder, 255, 1
                    End If
                    
                End If
                
                'Paint the caption
                If Len(.btCaptionTranslated) <> 0 Then
                
                    If i = m_ButtonIndex Then
                        curColor = fontColorActive
                    Else
                        If i = m_ButtonHoverIndex Then
                            curColor = fontColorHover
                        Else
                            curColor = fontColorInactive
                        End If
                    End If
                    
                    'Borrow a relevant UI font from the public UI font cache, then render the button caption using the clipping
                    ' rect we already calculated in previous steps.
                    
                    'Text fits just fine, so use the control font size
                    If .btFontSize = 0 Then
                        Set tmpFont = Font_Management.GetMatchingUIFont(m_FontSize, m_FontBold)
                        tmpFont.SetFontColor curColor
                        tmpFont.AttachToDC bufferDC
                        tmpFont.SetTextAlignment vbLeftJustify
                        tmpFont.DrawCenteredTextToRect .btCaptionTranslated, .btCaptionRect
                        tmpFont.ReleaseFromDC
                        
                    'Text does not fit the button area; use the custom font size we calculated in a previous step
                    Else
                        
                        Set tmpFont = Font_Management.GetMatchingUIFont(.btFontSize, m_FontBold)
                        tmpFont.SetFontColor curColor
                        tmpFont.AttachToDC bufferDC
                        tmpFont.SetTextAlignment vbLeftJustify
                        tmpFont.DrawCenteredTextToRect .btCaptionTranslated, .btCaptionRect
                        tmpFont.ReleaseFromDC
                        
                    End If
                    
                End If
                
                'Paint the button image, if any
                If Not (.btImages Is Nothing) Then
                    
                    If Me.Enabled Then
                    
                        If i = m_ButtonHoverIndex Then
                            .btImages.alphaBlendToDCEx bufferDC, .btImageCoords.x, .btImageCoords.y, .btImageWidth, .btImageHeight, 0, .btImageHeight, .btImageWidth, .btImageHeight
                        Else
                            .btImages.alphaBlendToDCEx bufferDC, .btImageCoords.x, .btImageCoords.y, .btImageWidth, .btImageHeight, 0, 0, .btImageWidth, .btImageHeight
                        End If
                        
                    Else
                        .btImages.alphaBlendToDCEx bufferDC, .btImageCoords.x, .btImageCoords.y, .btImageWidth, .btImageHeight, 0, .btImageHeight * 2, .btImageWidth, .btImageHeight
                    End If
                    
                End If
                
            End With
        
        Next i
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog,
' and/or retranslating all button captions against the current language.
Public Sub UpdateAgainstCurrentTheme()
    
    'Determine if translations are active.  If they are, retrieve translated captions for all buttons within the control.
    If g_IsProgramRunning Then
        
        'See if translations are necessary.
        Dim isTranslationActive As Boolean
            
        If Not (g_Language Is Nothing) Then
            If g_Language.translationActive Then
                isTranslationActive = True
            Else
                isTranslationActive = False
            End If
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
    
    'Update all text managed by the support class (e.g. tooltips)
    If g_IsProgramRunning Then ucSupport.UpdateAgainstThemeAndLanguage
        
    'Because translations can change text layout, we need to recalculate font metrics prior to redrawing the button
    UpdateControlLayout
    
End Sub

'Due to complex interactions between user controls and PD's translation engine, tooltips require this dedicated function.
' (IMPORTANT NOTE: the tooltip class will handle translations automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE, Optional ByVal toolTipIndex As Long = -1)
    
    'If toolTipIndex = -1, all buttons receive the same tooltip
    If toolTipIndex = -1 Then
    
        Dim i As Long
        For i = LBound(m_Buttons) To UBound(m_Buttons)
            With m_Buttons(i)
                .btToolTipText = newTooltip
                .btToolTipTitle = newTooltipTitle
                .btToolTipIcon = newTooltipIcon
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
                .btToolTipIcon = newTooltipIcon
            End With
            
        End If
                
    End If
    
    'Synchronize the tooltip object to the new tooltip.  (This is now handled manually, during mouse events.)
        
End Sub

Private Sub SynchronizeToolTipToIndex(Optional ByVal srcIndex As Long = 0)

    'Ignore invalid index requests
    If (srcIndex >= LBound(m_Buttons)) And (srcIndex <= UBound(m_Buttons)) And (srcIndex <> m_LastToolTipIndex) Then
    
        'Manually sync the tooltip now
        ucSupport.AssignTooltip Me.ContainerHwnd, m_Buttons(srcIndex).btToolTipText, m_Buttons(srcIndex).btToolTipTitle, m_Buttons(srcIndex).btToolTipIcon
        
        'Update the cached last index value
        m_LastToolTipIndex = srcIndex
    
    End If

End Sub
