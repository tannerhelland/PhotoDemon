VERSION 5.00
Begin VB.UserControl pdHistory 
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
   ToolboxBitmap   =   "pdHistory.ctx":0000
End
Attribute VB_Name = "pdHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Automatic History control
'Copyright 2016-2026 by Tanner Helland
'Created: 16/October/16
'Last updated: 21/March/18
'Last update: add support for keyboard nav
'
'This control is currently used in the color selection dialog.  It provides a semi-owner-drawn mechanism
' for displaying an interactive "history" of items selected by the user (in the color selection dialog, for example,
' this is a list of colors selected by the user).  The user can click any item in the history to make it the
' "current" value.
'
'This control is best used in places where access to past values is useful, but only if it's available immediately
' (unlike a menu, where the user must navigate through discrete "layers" of an interface).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Event HistoryItemClicked(ByVal histIndex As Long, ByVal histValue As String)
Public Event HistoryDoesntExist(ByVal histIndex As Long, ByRef histValue As String)
Public Event HistoryItemMouseOver(ByVal histIndex As Long, ByVal histValue As String)
Public Event DrawHistoryItem(ByVal histIndex As Long, ByVal histValue As String, ByVal targetDC As Long, ByVal ptrToRectF As Long)
Public Event CustomWindowMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'To simplify rendering, we pre-calculate a rectangle for the "history" area of the control.  (Individual items
' within the history control can be resolved on-the-fly).  This rect is calculated by UpdateControlLayout,
' and it must be recalculated if the control size changes.
Private m_HistoryRect As RectF

Private Type PD_HistoryItem
    ItemString As String
    ItemRect As RectF
End Type

Private m_HistoryItems() As PD_HistoryItem
Private m_HistoryCount As Long

'Individual history items within the history rectangle are resolved by (x, y) position.  Note that the number of rows
' is controlled by user property; columns are automatically inferred from that value.
Private m_NumHistoryRows As Long, m_NumHistoryColumns As Long

'If a history item is hovered, this will be set to some value >= 0
Private m_HistoryItemHovered As Long

'If the control has focus, the currently selected index is stored here; -1 means "nothing has been selected"
Private m_LastItemClicked As Long

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDHISTORY_COLOR_LIST
    [_First] = 0
    PDH_Background = 0
    PDH_Caption = 1
    PDH_Border = 2
    [_Last] = 2
    [_Count] = 3
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'If we were able to load previous settings from file, this will be set to TRUE
Private m_SavedHistoryExists As Boolean

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_History
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
    RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

Public Property Get FontSize() As Single
    FontSize = ucSupport.GetCaptionFontSize()
End Property

Public Property Let FontSize(ByVal newSize As Single)
    ucSupport.SetCaptionFontSize newSize
    PropertyChanged "FontSize"
End Property

Public Property Get HistoryRows() As Long
    HistoryRows = m_NumHistoryRows
End Property

Public Property Let HistoryRows(ByVal newRows As Long)
    If (newRows <> m_NumHistoryRows) Then
        m_NumHistoryRows = newRows
        If ((Not ucSupport Is Nothing) And PDMain.IsProgramRunning()) Then
            If ucSupport.AmIVisible Then UpdateControlLayout
        End If
    End If
End Property

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'To save/restore settings persistently, the entire history collection can be retrieved as a single string
Public Function GetHistoryAsString() As String

    If (m_HistoryCount > 0) Then
    
        Dim cXML As pdSerialize
        Set cXML = New pdSerialize
        
        With cXML
            
            .AddParam "pdHistoryCount", m_HistoryCount
            
            Dim i As Long
            For i = 0 To m_HistoryCount - 1
                .AddParam "pdHistoryItem" & ":" & CStr(i), m_HistoryItems(i).ItemString
            Next i
            
        End With
        
        GetHistoryAsString = cXML.GetParamString
        
    Else
        GetHistoryAsString = vbNullString
    End If

End Function

Public Sub SetHistoryFromString(ByRef srcString As String)
    
    'Note that a saved history exists and the caller has attempted to load it.  This spares us from asking our owner
    ' for placeholder entries.
    m_SavedHistoryExists = True
    
    Dim i As Long
                
    If (LenB(srcString) <> 0) Then
    
        Dim cXML As pdSerialize
        Set cXML = New pdSerialize
        
        With cXML
        
            .SetParamString srcString
                
            Dim savedHistoryCount As Long
            savedHistoryCount = .GetLong("pdHistoryCount", 0)
            
            'If the saved history count is larger than the current history count (for whatever reason), load the full collection
            ' from file.  It's possible that a screen resize or other event may increase the number of history items can display
            ' on the screen, and we want to have all saved entries available for that if possible.
            If (savedHistoryCount > m_HistoryCount) Then ReDim Preserve m_HistoryItems(0 To savedHistoryCount) As PD_HistoryItem
            
            'If the saved history count is non-zero, load as many items from file as we can.
            If (savedHistoryCount > 0) Then
                
                Dim entryName As String
                
                For i = 0 To savedHistoryCount - 1
                    
                    entryName = "pdHistoryItem" & ":" & CStr(i)
                    If cXML.DoesParamExist(entryName) Then
                        m_HistoryItems(i).ItemString = .GetString(entryName, vbNullString)
                        If (LenB(m_HistoryItems(i).ItemString) = 0) Then RaiseEvent HistoryDoesntExist(i, m_HistoryItems(i).ItemString)
                    Else
                        RaiseEvent HistoryDoesntExist(i, m_HistoryItems(i).ItemString)
                    End If
                    
                Next i
            
            Else
                If (m_HistoryCount > 0) Then
                    For i = 0 To m_HistoryCount - 1
                        RaiseEvent HistoryDoesntExist(i, m_HistoryItems(i).ItemString)
                    Next i
                End If
            End If
        End With
    
    'If a saved history doesn't exist, give our owner a chance to supply their own default values.
    Else
        If (m_HistoryCount > 0) Then
            For i = 0 To m_HistoryCount - 1
                RaiseEvent HistoryDoesntExist(i, m_HistoryItems(i).ItemString)
            Next i
        End If
    End If
    
End Sub

'When the user wants to permanently push a new item into the history, they can use this function.
Public Sub PushNewHistoryItem(ByVal newItemValue As String, Optional ByVal lookForExistingMatch As Boolean = True, Optional ByVal redrawImmediately As Boolean = False)
    
    If (m_HistoryCount = 0) Then
        m_HistoryCount = 1
        ReDim m_HistoryItems(0) As PD_HistoryItem
    End If
    
    'If the user doesn't want duplicates in the list, we can look for them in advance, and remove
    ' them from the list.
    Dim i As Long
    
    Dim terminalValue As Long
    terminalValue = m_HistoryCount - 1
    
    If lookForExistingMatch Then
        For i = 0 To m_HistoryCount - 1
            If Strings.StringsEqual(newItemValue, m_HistoryItems(i).ItemString, False) Then
                terminalValue = i
                Exit For
            End If
        Next i
    End If
    
    'Shift everything in the list down
    For i = terminalValue To 1 Step -1
        m_HistoryItems(i).ItemString = m_HistoryItems(i - 1).ItemString
    Next i
    
    'Insert the new history item at the start
    m_HistoryItems(0).ItemString = newItemValue
    
    'NOTE: before pdHistory was used in the main color selector area, this was set to immediately
    ' invalidate and repaint the entire control (which is somewhat time-consuming, depending on
    ' what kind of history is being rendered).  I have since changed this to simply post a paint
    ' message to the back of the queue, and have not noticed any problems... yet.  Add "True" to
    ' the line below to restore the old behavior,
    If redrawImmediately Then Me.RequestRedraw

End Sub

'Some history items can be notified of new history events from far-away parts of the program.  PD handles these
' via window messages (as they're async-ish, and VB plays nicely with them).
Public Sub RequestCustomSubclassing(ByVal msgID As Long, Optional ByVal msgIsInternalToPD As Boolean = True)
    ucSupport.SubclassCustomMessage msgID, msgIsInternalToPD
End Sub

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

'If our parent control needs a redraw for some reason, it can request one here
Public Sub RequestRedraw(Optional ByVal paintImmediately As Boolean = False)
    RedrawBackBuffer paintImmediately
End Sub

Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    RaiseEvent CustomWindowMessage(wMsg, wParam, lParam, bHandled, lReturn)
End Sub

Private Sub ucSupport_GotFocusAPI()
    m_LastItemClicked = 0
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    markEventHandled = False
        
    If (vkCode = VK_LEFT) Then
        m_LastItemClicked = m_LastItemClicked - 1
        If (m_LastItemClicked < 0) Then m_LastItemClicked = m_HistoryCount - 1
        markEventHandled = True
    ElseIf (vkCode = VK_RIGHT) Then
        m_LastItemClicked = m_LastItemClicked + 1
        If (m_LastItemClicked >= m_HistoryCount) Then m_LastItemClicked = 0
        markEventHandled = True
    ElseIf (vkCode = VK_UP) Then
        m_LastItemClicked = m_LastItemClicked - m_NumHistoryColumns
        If (m_LastItemClicked < 0) Then
            m_LastItemClicked = m_LastItemClicked + m_HistoryCount
            If (m_LastItemClicked >= m_HistoryCount) Then m_LastItemClicked = m_HistoryCount - 1
        End If
        markEventHandled = True
    ElseIf (vkCode = VK_DOWN) Then
        m_LastItemClicked = m_LastItemClicked + m_NumHistoryColumns
        If (m_LastItemClicked >= m_HistoryCount) Then
            m_LastItemClicked = m_LastItemClicked - m_HistoryCount
            If (m_LastItemClicked < 0) Then m_LastItemClicked = 0
        End If
        markEventHandled = True
    ElseIf (vkCode = VK_SPACE) Then
        If (m_LastItemClicked >= 0) And (m_LastItemClicked < m_HistoryCount) Then
            RaiseEvent HistoryItemClicked(m_LastItemClicked, m_HistoryItems(m_LastItemClicked).ItemString)
            RedrawBackBuffer
        End If
    End If
    
    If markEventHandled Then RedrawBackBuffer

End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_LostFocusAPI()
    m_LastItemClicked = -1
    RedrawBackBuffer
    RaiseEvent LostFocusAPI
End Sub

'Only left clicks raise Click() events
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    If Me.Enabled And ((Button And pdLeftButton) <> 0) Then
        
        'Start by seeing if the mouse is inside the history portion of the control
        m_LastItemClicked = GetHistoryItemUnderMouse(x, y)
        
        If ((m_LastItemClicked >= 0) And (m_LastItemClicked < m_HistoryCount)) Then
            RaiseEvent HistoryItemClicked(m_LastItemClicked, m_HistoryItems(m_LastItemClicked).ItemString)
            RedrawBackBuffer
        End If
        
    End If
    
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    ucSupport.RequestCursor IDC_HAND
    RedrawBackBuffer
End Sub

'When the mouse leaves the UC, we must repaint the button (as it's no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_HistoryItemHovered = -1
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    Dim oldHoverCheck As Long
    oldHoverCheck = m_HistoryItemHovered
    m_HistoryItemHovered = GetHistoryItemUnderMouse(x, y)
    
    If (m_HistoryItemHovered >= 0) Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
    If (oldHoverCheck <> m_HistoryItemHovered) Then
        If ((m_HistoryItemHovered >= 0) And (m_HistoryItemHovered < m_HistoryCount)) Then RaiseEvent HistoryItemMouseOver(m_HistoryItemHovered, m_HistoryItems(m_HistoryItemHovered).ItemString)
        RedrawBackBuffer
    End If
    
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Private Sub UserControl_Initialize()
    
    m_HistoryCount = 0
    m_HistoryItemHovered = -1
    m_LastItemClicked = -1
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN, VK_SPACE
    ucSupport.RequestCaptionSupport
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDHISTORY_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDHistory", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    Caption = vbNullString
    FontSize = 12
    HistoryRows = 1
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Caption = .ReadProperty("Caption", vbNullString)
        FontSize = .ReadProperty("FontSize", 12)
        HistoryRows = .ReadProperty("HistoryRows", 1)
    End With
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", ucSupport.GetCaptionText, vbNullString
        .WriteProperty "FontSize", ucSupport.GetCaptionFontSize, 12
        .WriteProperty "HistoryRows", m_NumHistoryRows, 1
    End With
End Sub

Private Function GetHistoryItemUnderMouse(ByVal srcX As Single, ByVal srcY As Single) As Long
    
    GetHistoryItemUnderMouse = -1
    
    'First, shortcut the function by seeing if the mouse is even inside the history area.  (If a caption is in use,
    ' this may not be true.)
    If PDMath.IsPointInRectF(srcX, srcY, m_HistoryRect) Then
    
        If (m_HistoryCount > 0) Then
            
            Dim i As Long
            For i = 0 To m_HistoryCount - 1
                If PDMath.IsPointInRectF(srcX, srcY, m_HistoryItems(i).ItemRect) Then
                    GetHistoryItemUnderMouse = i
                    Exit For
                End If
            Next i
            
        End If
    
    End If

End Function

'Call this layout calculator whenever the control size changes.  This is particularly important for this control,
' as history item rects are pre-calculated ahead of time.
Private Sub UpdateControlLayout()

    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'The first thing we want to do is calculate an available area for the entire history section of the control.
    ' (Spacing rules are the same as all other captioned controls.)
    If ucSupport.IsCaptionActive Then
        
        'The brush area is placed relative to the caption
        With m_HistoryRect
            .Left = FixDPI(8)
            .Top = ucSupport.GetCaptionBottom + 2
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    'If there's no caption, allow the clickable portion to fill the entire control
    Else
        
        With m_HistoryRect
            .Left = 1
            .Top = 1
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    End If
    
    'Next, calculate how many history items we can fit inside the current control area.  We want each history item
    ' to be a perfect square, so the number of items we can support is limited to how many perfectly square items
    ' we can fit in the supplied area (while respecting the user-settable "number of rows" property).
    If (m_NumHistoryRows = 0) Then m_NumHistoryRows = 1
    
    Dim newHistoryCount As Long
    newHistoryCount = ((m_HistoryRect.Width - 1) \ (m_HistoryRect.Height \ m_NumHistoryRows)) * m_NumHistoryRows
    
    If (newHistoryCount <> 0) Then
        
        Dim i As Long
        
        Dim oldHistoryCount As Long: oldHistoryCount = m_HistoryCount
        m_HistoryCount = newHistoryCount
        ReDim Preserve m_HistoryItems(0 To m_HistoryCount - 1) As PD_HistoryItem
        
        'If the new "available history display size" is larger than the old size, we may have more history items than we
        ' have stored values.  Ask our owner to supply new default values now.
        Dim tmpString As String
        If (newHistoryCount > oldHistoryCount) Then
            For i = oldHistoryCount To newHistoryCount - 1
                RaiseEvent HistoryDoesntExist(i, tmpString)
                m_HistoryItems(i).ItemString = tmpString
            Next i
        Else
        
            'If no history exists at all, ask the caller to supply a full list of items.
            If (Not m_SavedHistoryExists) Then
                For i = 0 To newHistoryCount - 1
                    RaiseEvent HistoryDoesntExist(i, tmpString)
                    m_HistoryItems(i).ItemString = tmpString
                Next i
            End If
        
        End If
        
        'Generate a column count from the total history count
        m_NumHistoryColumns = m_HistoryCount \ m_NumHistoryRows
        
        'Next, we want to crop the history area width to an even integer multiple of the current history count
        Dim defaultHeight As Long
        defaultHeight = m_HistoryRect.Height \ m_NumHistoryRows
        m_HistoryRect.Width = (m_NumHistoryColumns * defaultHeight) + 1
        
        Dim vOffset As Long: vOffset = m_HistoryRect.Top
        Dim hOffset As Long: hOffset = m_HistoryRect.Left
        
        'With the history area correctly identified, we can now calculate rects for each individual history item.
        ' (These are pre-calculated to improve rendering performance.)
        For i = 0 To m_HistoryCount - 1
        
            With m_HistoryItems(i).ItemRect
                .Top = vOffset
                .Left = hOffset
                .Height = defaultHeight
                .Width = defaultHeight
                
                hOffset = hOffset + .Width
                
                'If a history box gets pushed past the edge of the control, move it down to the next row
                If (hOffset > (m_HistoryRect.Left + m_HistoryRect.Width - 2)) Then
                    hOffset = m_HistoryRect.Left
                    vOffset = vOffset + defaultHeight
                End If
                
            End With
        
        Next i
        
    End If
        
    'No other special preparation is required for this control, so proceed with recreating the back buffer
    RedrawBackBuffer
            
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDH_Background, "Background", IDE_WHITE
        .LoadThemeColor PDH_Caption, "Caption", IDE_GRAY
        .LoadThemeColor PDH_Border, "Border", IDE_GRAY
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer(Optional ByVal paintImmediately As Boolean = False)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDH_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
    
    If PDMain.IsProgramRunning() Then
        
        Dim i As Long
        
        'Because this control is owner-drawn, our owner is responsible for drawing the individual history samples.
        If (m_HistoryCount > 0) Then
            
            Dim tmpRectF As RectF
            
            For i = 0 To m_HistoryCount - 1
                
                'We shrink the display area by one pixel to account for borders.  This isn't strictly necessary
                ' (as we overpaint borders anyway), but it allows the painter to strictly fit samples inside a
                ' known area, without worrying about edge pixels getting erased.
                tmpRectF.Left = m_HistoryItems(i).ItemRect.Left + 1
                tmpRectF.Top = m_HistoryItems(i).ItemRect.Top + 1
                tmpRectF.Width = m_HistoryItems(i).ItemRect.Width - 1
                tmpRectF.Height = m_HistoryItems(i).ItemRect.Height - 1
                
                RaiseEvent DrawHistoryItem(i, m_HistoryItems(i).ItemString, bufferDC, VarPtr(tmpRectF))
                
            Next i
            
        End If
        
        'Next, draw a grid around the rendered history items
        Dim cSurface As pd2DSurface, cPen As pd2DPen
        If (m_HistoryCount > 0) And (bufferDC <> 0) Then
            
            Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, True
            Drawing2D.QuickCreateSolidPen cPen, 1#, m_Colors.RetrieveColor(PDH_Border, Me.Enabled), 100#
        
            For i = 0 To m_HistoryCount - 1
                With m_HistoryItems(i).ItemRect
                    PD2D.DrawRectangleF cSurface, cPen, .Left, .Top, .Width, .Height
                End With
            Next i
            
            'Finally, if one of the history items is currently hovered or selected, paint it with a
            ' chunky, highlighted border
            Dim cOuterPen As pd2DPen
            Drawing2D.QuickCreatePairOfUIPens cOuterPen, cPen, True
            
            'Last-clicked item is only highlighted if the control has focus
            If ucSupport.DoIHaveFocus Then
                If (m_LastItemClicked >= 0) Then
                    PD2D.DrawRectangleF_FromRectF cSurface, cOuterPen, m_HistoryItems(m_LastItemClicked).ItemRect
                    PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_HistoryItems(m_LastItemClicked).ItemRect
                End If
            End If
            
            'Hovered entries are always highlighted
            If (m_HistoryItemHovered >= 0) Then
                PD2D.DrawRectangleF_FromRectF cSurface, cOuterPen, m_HistoryItems(m_HistoryItemHovered).ItemRect
                PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_HistoryItems(m_HistoryItemHovered).ItemRect
            End If
                
            Set cOuterPen = Nothing
            
        End If
        
        Set cSurface = Nothing: Set cPen = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint paintImmediately
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
