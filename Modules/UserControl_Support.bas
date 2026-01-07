Attribute VB_Name = "UserControls"
'***************************************************************************
'Helper functions for various PhotoDemon UCs
'Copyright 2014-2026 by Tanner Helland
'Created: 06/February/14
'Last updated: 20/August/15
'Last update: start migrating various UC-inspecific functions here
'
'Many of PD's custom user controls share similar functionality.  Rather than duplicate that functionality across
' all controls, I've tried to collect reusable functions here.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'There are many different PhotoDemon UI controls.  New ones must be added to this enum.
Public Enum PD_ControlType
    pdct_Accelerator = 0
    pdct_BrushSelector = 1
    pdct_Button = 2
    pdct_ButtonStrip = 3
    pdct_ButtonStripVertical = 4
    pdct_ButtonToolbox = 5
    pdct_Canvas = 6
    pdct_CanvasView = 7
    pdct_CheckBox = 8
    pdct_ColorDepth = 9
    pdct_ColorSelector = 10
    pdct_ColorVariants = 11
    pdct_ColorWheel = 12
    pdct_CommandBar = 13
    pdct_CommandBarMini = 14
    pdct_Container = 15
    pdct_Download = 16
    pdct_DropDown = 17
    pdct_DropDownFont = 18
    pdct_FxPreviewCtl = 19
    pdct_GradientSelector = 20
    pdct_History = 21
    pdct_Hyperlink = 22
    pdct_ImageStrip = 23
    pdct_Label = 24
    pdct_LayerList = 25
    pdct_LayerListInner = 26
    pdct_ListBox = 27
    pdct_ListBoxOD = 28
    pdct_ListBoxView = 29
    pdct_ListBoxViewOD = 30
    pdct_MetadataExport = 31
    pdct_Navigator = 32
    pdct_NavigatorInner = 33
    pdct_NewOld = 34
    pdct_PaletteUI = 35
    pdct_PenSelector = 36
    pdct_PictureBox = 37
    pdct_PictureBoxInteractive = 38
    pdct_Preview = 39
    pdct_ProgressBar = 40
    pdct_RadioButton = 41
    pdct_RandomizeUI = 42
    pdct_Resize = 43
    pdct_Ruler = 44
    pdct_ScrollBar = 45
    pdct_SearchBar = 46
    pdct_Slider = 47
    pdct_SliderStandalone = 48
    pdct_Spinner = 49
    pdct_StatusBar = 50
    pdct_Strip = 51
    pdct_TextBox = 52
    pdct_Title = 53
    pdct_TreeviewOD = 54
    pdct_TreeviewViewOD = 55
End Enum

#If False Then
    Private Const pdct_Accelerator = 0, pdct_BrushSelector = 1, pdct_Button = 2, pdct_ButtonStrip = 3, pdct_ButtonStripVertical = 4, pdct_ButtonToolbox = 5, pdct_Canvas = 6, pdct_CanvasView = 7, pdct_CheckBox = 8, pdct_ColorDepth = 9
    Private Const pdct_ColorSelector = 10, pdct_ColorVariants = 11, pdct_ColorWheel = 12, pdct_CommandBar = 13, pdct_CommandBarMini = 14, pdct_Container = 15, pdct_Download = 16, pdct_DropDown = 17, pdct_DropDownFont = 18, pdct_FxPreviewCtl = 19
    Private Const pdct_GradientSelector = 20, pdct_History = 21, pdct_Hyperlink = 22, pdct_ImageStrip = 23, pdct_Label = 24, pdct_LayerList = 25, pdct_LayerListInner = 26, pdct_ListBox = 27, pdct_ListBoxOD = 28, pdct_ListBoxView = 29
    Private Const pdct_ListBoxViewOD = 30, pdct_MetadataExport = 31, pdct_Navigator = 32, pdct_NavigatorInner = 33, pdct_PaletteUI = 35, pdct_PenSelector = 36, pdct_PictureBox = 37, pdct_PictureBoxInteractive = 38, pdct_Preview = 39
    Private Const pdct_ProgressBar = 40, pdct_RadioButton = 41, pdct_RandomizeUI = 42, pdct_Resize = 43, pdct_Ruler = 44, pdct_ScrollBar = 45, pdct_SearchBar = 46, pdct_Slider = 47, pdct_SliderStandalone = 48, pdct_Spinner = 49
    Private Const pdct_StatusBar = 50, pdct_Strip = 51, pdct_TextBox = 52, pdct_Title = 53, pdct_TreeviewOD = 54, pdct_TreeviewViewOD = 55
#End If

'User control text *THAT IS DISPLAYED TO THE USER* needs to be translated.  This module provides a helper
' function that caches shared control text, and returns it based on this enum.  This provides a nice
' perf boost for users in other locales, but this should only be used for text that is inherent to PD's
' user controls, and will appear multiple times in the same session.  (There is a fixed startup cost
' for generating these translations, and if a translation is unlikely to be used *every* session,
' translate it locally - not here.)
Public Enum PD_UserControlText
    pduct_AnimationRepeatToggle
    pduct_CommandBarPresetList
    pduct_CommandBarRandom
    pduct_CommandBarRedo
    pduct_CommandBarReset
    pduct_CommandBarSavePreset
    pduct_CommandBarUndo
    pduct_FlyoutLockTitle
    pduct_FlyoutLockTooltip
    pduct_Randomize
End Enum

Public Type PD_ListItem
    textEn As String
    textTranslated As String
    itemTop As Long
    itemHeight As Long
    isSeparator As Boolean
End Type

Public Type PD_TreeItem
    textEn As String
    itemTop As Long         'Rendering rect top, in pixels, assuming no collapses.  Must be y-offset at render time.
    ItemRect As RectF       'Rendering rect of the full list item, in pixels.
    controlRect As RectF    'If a node has children, its expand/collapse control is triggerable from this rect
    captionRect As RectF    'Rendering rect of the caption only, in pixels.
    itemID As String
    parentID As String
    numParents As Long      'Calculated automatically; required to determine rendering position.
    isCollapsed As Boolean              'Persistently tracks collapse state *for parent nodes only*
    isCollapsedThisRender As Boolean    'Tracks collapse state for *all nodes* to accelerate rendering
    hasChildren As Boolean  'Not supplied by caller; inferred by treeview automatically
End Type

'At times, PD may need to post custom messages to all application windows (e.g. theme changes may eventually be implemented
' like this).  Do not call PostMessage directly, as it sends messages to the thread's message queue; instead, call the
' PostPDMessage() function below, which asynchronously relays the request to registered windows via SendNotifyMessage.
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Current list of registered windows and the custom messages they want to receive.

'This spares us from having to enumerate all windows, or worse, blast all windows
' in the system with our internal messages.  (At present, these are naive lists
' because PD uses so few of them, but in the future, we could look at a hash table.
' I've deliberately made the list interactions structure-agnostic to simplify future
' improvements.)
Private m_windowList() As Long, m_wMsgList() As Long
Private m_windowMsgCount As Long
Private Const INITIAL_WINDOW_MESSAGE_LIST_SIZE As Long = 16&

'Current list of shared GDI brushes; this spares us from creating unique brushes for every edit box instance
Private Type SharedGDIBrush
    brushColor As Long
    brushHandle As Long
    numOfOwners As Long
End Type

Private m_numOfSharedBrushes As Long
Private m_SharedBrushes() As SharedGDIBrush
Private Const INIT_SIZE_OF_BRUSH_CACHE As Long = 4&

'Current list of shared GDI fonts; this spares us from creating unique font for every edit box instance.
' (Note that we don't use a pdFontCollection object for this, as these fonts are immutable.  If we accidentally
'  destroy one before its matching edit box is freed, we will crash and burn, so a different handling technique
'  is requred.)
Private Type SharedGDIFont
    FontSize As Single
    fontHandle As Long
    numOfOwners As Long
End Type

Private m_numOfSharedFonts As Long
Private m_SharedFonts() As SharedGDIFont
Private Const INIT_SIZE_OF_FONT_CACHE As Long = 4&

'As part of broad optimization efforts in the 7.0 release, this module now tracks how many custom PD controls we're managing at
' any given time.  Use this for leak-detection and resource counting.  For example: each ucSupport-managed PD control uses two
' GDI objects: one DIB and one persistent DC for the control's backbuffer (all controls are double-buffered).  Use this to
' figure out how many of the program's GDI objects are being used by UCs, and how many are being created and used elsewhere.
Private m_PDControlCount As Long

'Dropdown boxes (and similar controls, like flyout panels) are problematic, because we have to
' play weird window ownership games to ensure that the dropdowns appear "above" or "outside"
' VB windows, as necessary.  As such, this function is notified whenever a listbox (or flyout)
' is raised, and the hWnd is cached so we can kill that window as necessary.
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal srcColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Enum AnimateWindowFlags
    AW_ACTIVATE = &H20000   'Activates the window. Do not use this value with AW_HIDE.
    AW_BLEND = &H80000      'Uses a fade effect. This flag can be used only if hwnd is a top-level window.
    AW_CENTER = &H10&       'Makes the window appear to collapse inward if AW_HIDE is used or expand outward if the AW_HIDE is not used. The various direction flags have no effect.
    AW_HIDE = &H10000       'Hides the window. By default, the window is shown.
    AW_HOR_POSITIVE = &H1&  'Animates the window from left to right. This flag can be used with roll or slide animation. It is ignored when used with AW_CENTER or AW_BLEND.
    AW_HOR_NEGATIVE = &H2&  'Animates the window from right to left. This flag can be used with roll or slide animation. It is ignored when used with AW_CENTER or AW_BLEND.
    AW_SLIDE = &H40000      'Uses slide animation. By default, roll animation is used. This flag is ignored when used with AW_CENTER.
    AW_VER_POSITIVE = &H4&  'Animates the window from top to bottom. This flag can be used with roll or slide animation. It is ignored when used with AW_CENTER or AW_BLEND.
    AW_VER_NEGATIVE = &H8&  'Animates the window from bottom to top. This flag can be used with roll or slide animation. It is ignored when used with AW_CENTER or AW_BLEND.
End Enum

#If False Then
    Private Const AW_ACTIVATE = &H20000, AW_BLEND = &H80000, AW_CENTER = &H10&, AW_HIDE = &H10000, AW_HOR_POSITIVE = &H1&, AW_HOR_NEGATIVE = &H2&, AW_SLIDE = &H40000, AW_VER_POSITIVE = &H4&, AW_VER_NEGATIVE = &H8&
#End If

Private Declare Function AnimateWindow Lib "user32" (ByVal hWnd As Long, ByVal dwTime As Long, ByVal dwFlags As AnimateWindowFlags) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal ptrToRect As Long, ByVal bErase As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private m_CurrentDropDownHWnd As Long, m_CurrentDropDownListHWnd As Long
Private m_CurrentFlyoutParentHWnd As Long, m_CurrentFlyoutPanelHWnd As Long
Private m_FlyoutRef As pdFlyout

'To better manage resources, we also track how many API windows we've created/destroyed during this session
Private m_APIWindowsCreated As Long, m_APIWindowsDestroyed As Long

'Same goes for timers and certain other object types
Private m_TimersCreated As Long, m_TimersDestroyed As Long

'Common control text shared across multiple instances can be cached in this dictionary, and accessed
' via the related public function (GetCommonTranslation).
Private m_CommonTranslations As pdDictionary

'Because there can only be one visible tooltip at a time, this support module is a great place to handle them.  Requests for new
' tooltips automatically unload old ones, although user controls still need to request tooltip hiding when they lose focus and/or
' are unloaded.

'To optimize tooltip positioning, we determine the tooltip edge closest to the mouse hover position (with the assumption that
' the user's eyes will be pointed at or near the cursor, so the closer a tooltip is to that, the better - without obscuring
' the control underneath, obviously)
Private Enum TT_SIDE
    TTS_Top = 0
    TTS_Right = 1
    TTS_Bottom = 2
    TTS_Left = 3
End Enum

#If False Then
    Private Const TTS_Top = 0, TTS_Right = 1, TTS_Bottom = 2, TTS_Left = 3
#End If

'All tooltip sizes are in pixels at 96 DPI ("100%" in Windows).  Other DPIs will automatically be handled at run-time,
' as necessary.
Private Const PD_TT_EXTERNAL_PADDING As Long = 2
Private Const PD_TT_INTERNAL_PADDING As Long = 6
Private Const PD_TT_MAX_WIDTH As Long = 450         'Tips larger than this will be word-wrapped to fit.
Private Const PD_TT_TITLE_PADDING As Long = 4       'Pixels between the tip title (if any) and caption

Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSENDCHANGING As Long = &H400
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_HIDEWINDOW As Long = &H80
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_NOZORDER As Long = &H4
Private Const WS_EX_NOACTIVATE As Long = &H8000000
Private Const WS_EX_TOOLWINDOW As Long = &H80

Private m_TTActive As Boolean, m_TTOwner As Long, m_TTHwnd As Long
Private m_TTWindowStyleHasBeenSet As Boolean, m_OriginalTTWindowBits As Long, m_OriginalTTWindowBitsEx As Long
Private m_TTRectCopy As RectL, m_LastTTPosition As TT_SIDE

Private Type TT_Sort
    ttSide As TT_SIDE
    ttDistance As Single
End Type

'Tooltips are hidden based on a timer.  If a new tooltip is requested before the timer expires, we simply move the existing window
' into place, rather than animating it.
Private m_TimerEventSink As pdUCEventSink
Private m_InitialTTTimerTime As Double

'Iterate through all sibling controls in our container, and if one is capable of receiving focus, activate it.  I had *really* hoped
' to bypass this kind of manual handling by using WM_NEXTDLGCTL, but I failed to get it working reliably with all types of VB windows.
' I'm honestly not sure whether VB even uses that message, or whether it uses some internal mechanism for focus tracking; the latter
' might explain some of the bugginess involved in VB focus handling.
'
'Returns: TRUE if focus was moved; FALSE if it was left behind.
Public Function ForwardFocusToNewControl(ByRef sourceControl As Object, ByVal focusDirectionForward As Boolean) As Boolean

    'When sited on a UC, we may not be able to iterate a controls collection.  Simply exit if this occurs.
    ' (Note that this is only a temporary solution; I'm working on a better one, where UC's cascade tab handling to their parent
    '  until an acceptable handler is reached.)
    On Error GoTo ParentHasNoControls

    'If the user has deactivated tab support, or we are invisible/disabled, ignore this completely
    If (sourceControl.Extender.TabStop And sourceControl.Extender.Visible And sourceControl.Enabled) Then
        
        'Iterate through all controls in the container, looking for the next TabStop index
        Dim myIndex As Long
        myIndex = sourceControl.Extender.TabIndex
        
        Dim newIndex As Long
        Const MAX_INDEX As Long = 99999
        
        'Forward and back focus checks require different search strategies
        If focusDirectionForward Then newIndex = MAX_INDEX Else newIndex = myIndex
        
        Dim Ctl As Control, targetControl As Control
        For Each Ctl In sourceControl.Parent.Controls
        
        'Some controls may not have a TabStop property.  That's okay - just keep iterating if it happens.
        On Error GoTo 0
        On Error GoTo NextControlCheck
            
            'Hypothetically, our error handler should remove the need for this kind of check.  That said, I prefer to handle the
            ' non-focusable objects like this, although this requires any outside user to complete the list with their own potentially
            ' non-focusable controls.  Not ideal, but I don't know a good way (short of error handling) to see whether a VB object
            ' is focusable.
            If IsControlFocusable(Ctl) Then
            
                'Ignore controls whose TabStop property is False, who are not visible, or who are disabled
                If (Ctl.TabStop And Ctl.Visible And Ctl.Enabled) Then
                        
                    If focusDirectionForward Then
                    
                        'Check the tab index of this control.  We're looking for the lowest tab index that is also larger than our tab index.
                        If (Ctl.TabIndex > myIndex) And (Ctl.TabIndex < newIndex) Then
                            newIndex = Ctl.TabIndex
                            Set targetControl = Ctl
                        End If
                        
                    Else
                    
                        'Check the tab index of this control.  We're looking for the highest tab index that is also larger than our tab index.
                        If (Ctl.TabIndex > newIndex) Then
                            newIndex = Ctl.TabIndex
                            Set targetControl = Ctl
                        End If
                    
                    End If
    
                End If
                
            End If
            
NextControlCheck:
        Next

        'When moving focus forward, we now have one of two possibilites:
        ' 1) NewIndex represents the tab index of a valid control whose index is higher than us.
        ' 2) NewIndex is still MAX_INDEX, because no control with a valid tab index was found.
        
        'When moving focus backward, we also have two possibilities:
        ' 1) NewIndex represents the tab index of a valid control whose index is higher than us.  (Required if Shift+Tab will push the
        '     TabIndex below 0.)
        ' 2) NewIndex is still MY_INDEX, because no control with a valid tab index was found.
        
        'Handle case 2 now.
        If (focusDirectionForward And (newIndex = MAX_INDEX)) Or (Not focusDirectionForward) Then
            
            If focusDirectionForward Then
                newIndex = myIndex
            Else
                newIndex = -1
            End If
            
            'Some controls may not have a TabStop property.  That's okay - just keep iterating if it happens.
            On Error GoTo 0
            On Error GoTo NextControlCheck2
            
            'If our control is last in line for tabstops, we need to now find the LOWEST tab index to forward focus to.
            For Each Ctl In sourceControl.Parent.Controls
                
                'Hypothetically, our error handler should remove the need for this kind of check.  That said, I prefer to handle the
                ' non-focusable objects like this, although this requires any outside user to complete the list with their own potentially
                ' non-focusable controls.  Not ideal, but I don't know a good way (short of error handling) to see whether a VB object
                ' is focusable.
                If IsControlFocusable(Ctl) Then
                    
                    'Ignore controls whose TabStop property is False, who are not visible, or who are disabled
                    If (Ctl.TabStop And Ctl.Visible And Ctl.Enabled) Then
                            
                        If focusDirectionForward Then
                        
                            'Check the tab index of this control.  We're looking for the lowest valid tab index.
                            If (Ctl.TabIndex < myIndex) And (Ctl.TabIndex < newIndex) Then
                                newIndex = Ctl.TabIndex
                                Set targetControl = Ctl
                            End If
                            
                        Else
                        
                            'Check the tab index of this control.  We're looking for the lowest valid tab index, if one exists.
                            If (Ctl.TabIndex < myIndex) And (Ctl.TabIndex > newIndex) Then
                                newIndex = Ctl.TabIndex
                                Set targetControl = Ctl
                            End If
                        
                        End If
                    
                    End If
                    
                End If
                
NextControlCheck2:
            Next
        
        End If
        
        If (Not focusDirectionForward) Then
            If newIndex = -1 Then newIndex = myIndex
        End If
        
        'Regardless of focus direction, we once again have one of two possibilites.
        ' 1) NewIndex represents the tab index of the next valid control in VB's tab order.
        ' 2) NewIndex = our index, because no control with a valid tab index was found.
        
        'SetFocus can fail under a variety of circumstances, so error handling is still required
        On Error GoTo 0
        On Error GoTo NoFocusRecipient
        
        'Ignore the second case completely, as tab should have no effect
        If (newIndex <> myIndex) Then
            If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI targetControl.hWnd
        
NoFocusRecipient:
        
        End If
        
    End If
    
    ForwardFocusToNewControl = True
    
    Exit Function
    
ParentHasNoControls:

    ForwardFocusToNewControl = False

End Function

'If an object of type Control is capable of receiving focus, this will return TRUE.  We use this to make sure focus-setting
' doesn't break by accidentally setting focus to something we shouldn't.
Public Function IsControlFocusable(ByRef Ctl As Control) As Boolean

    If Not (TypeOf Ctl Is Timer) And Not (TypeOf Ctl Is Line) And Not (TypeOf Ctl Is Image) And Not (TypeOf Ctl Is pdLabel) And Not (TypeOf Ctl Is pdAccelerator) And Not (TypeOf Ctl Is pdDownload) Then
        IsControlFocusable = True
    Else
        IsControlFocusable = False
    End If

End Function

'PD's various user controls sometimes like to share data via custom window messages.
' Instead of calling PostMessage directly, use this wrapper function, which may perform
' additional maintenance.
Public Sub PostPDMessage(ByVal wMsg As Long, Optional ByVal wParam As Long = 0&, Optional ByVal lParam As Long = 0&, Optional ByVal usePostMessageInstead As Boolean = False)
    
    Dim pmReturn As Long
    pmReturn = 1
    
    'Enumerate all matching, non-zero windows, and post the requested message, without waiting for a response.
    Dim i As Long
    For i = 0 To m_windowMsgCount - 1
        If (m_wMsgList(i) = wMsg) Then
            If (m_windowList(i) <> 0) Then
                If usePostMessageInstead Then
                    pmReturn = pmReturn And PostMessage(m_windowList(i), wMsg, wParam, lParam)
                Else
                    pmReturn = pmReturn And SendNotifyMessage(m_windowList(i), wMsg, wParam, lParam)
                End If
            End If
        End If
    Next i
    
    If (pmReturn = 0) Then PDDebug.LogAction "PostPDMessage was unable to post message ID #" & wMsg & " to one or more windows."
    
End Sub

'Rather than blast all windows with manually raised messages, PD maintains a list of hWnds and registered message requests.
' Add windows and/or messages via this function, and when the messages need to be raised (via PostPDMessage(), above),
' the function will automatically notify all registered recipients.
Public Sub AddMessageRecipient(ByVal targetHWnd As Long, ByVal wMsg As Long)
    
    'Prep the storage structure, as necessary.
    If (m_windowMsgCount = 0) Then
        ReDim m_windowList(0 To INITIAL_WINDOW_MESSAGE_LIST_SIZE - 1) As Long
        ReDim m_wMsgList(0 To INITIAL_WINDOW_MESSAGE_LIST_SIZE - 1) As Long
    End If
    
    If (m_windowMsgCount > UBound(m_windowList)) Then
        ReDim Preserve m_windowList(0 To (UBound(m_windowList) * 2 + 1)) As Long
        ReDim Preserve m_wMsgList(0 To (UBound(m_wMsgList) * 2 + 1)) As Long
    End If
    
    m_windowList(m_windowMsgCount) = targetHWnd
    m_wMsgList(m_windowMsgCount) = wMsg
    
    m_windowMsgCount = m_windowMsgCount + 1
    
End Sub

Public Sub RemoveMessageRecipient(ByVal targetHWnd As Long)
    
    'Rather then condensing the list, we simply set all corresponding window entries to zero.
    Dim i As Long
    For i = 0 To m_windowMsgCount - 1
        If (m_windowList(i) = targetHWnd) Then
            m_windowList(i) = 0
            m_wMsgList(i) = 0
        End If
    Next i
    
End Sub

Public Sub NotifyAPIWindowCreated()
    m_APIWindowsCreated = m_APIWindowsCreated + 1
End Sub

Public Sub NotifyAPIWindowDestroyed()
    m_APIWindowsDestroyed = m_APIWindowsDestroyed + 1
End Sub

Public Function GetAPIWindowCount(Optional ByRef windowsCreated As Long, Optional ByRef windowsDestroyed As Long) As Long
    windowsCreated = m_APIWindowsCreated
    windowsDestroyed = m_APIWindowsDestroyed
    GetAPIWindowCount = (windowsCreated - windowsDestroyed)
End Function

Public Sub NotifyTimerCreated()
    m_TimersCreated = m_TimersCreated + 1
End Sub

Public Sub NotifyTimerDestroyed()
    m_TimersDestroyed = m_TimersDestroyed + 1
End Sub

Public Function GetTimerCount(Optional ByRef timersCreated As Long, Optional ByRef timersDestroyed As Long) As Long
    timersCreated = m_TimersCreated
    timersDestroyed = m_TimersDestroyed
    GetTimerCount = (timersCreated - timersDestroyed)
End Function

'Edit boxes can all share the same background brush (as they are all themed identically).  Call this function instead
' of creating your own brush for every text box instance.
Public Function GetSharedGDIBrush(ByVal requestedColor As Long) As Long
    
    'First things first: if this is the first brush requested by the system, initialize the shared brush list
    If (m_numOfSharedBrushes = 0) Then ReDim m_SharedBrushes(0 To INIT_SIZE_OF_BRUSH_CACHE - 1) As SharedGDIBrush
    
    'Next, look for the requested color in our current cache.  If it exists, we don't want to recreate it.
    Dim brushExists As Boolean, brushIndex As Long, i As Long
    brushExists = False
    
    If (m_numOfSharedBrushes > 0) Then
            
        For i = 0 To m_numOfSharedBrushes - 1
            If (m_SharedBrushes(i).brushColor = requestedColor) And (m_SharedBrushes(i).brushHandle <> 0) Then
            
                'As a failsafe for black brushes, make sure the owner count is valid too
                If m_SharedBrushes(i).numOfOwners > 0 Then
                    brushExists = True
                    brushIndex = i
                    Exit For
                End If
            
            End If
        Next i
            
    End If
    
    'If we found the brush in our cache, increment the owner count, and return the handle immediately
    If brushExists Then
        m_SharedBrushes(brushIndex).numOfOwners = m_SharedBrushes(brushIndex).numOfOwners + 1
        GetSharedGDIBrush = m_SharedBrushes(brushIndex).brushHandle
    
    'If the brush doesn't exist, create it anew
    Else
    
        'If the cache is too small, resize it
        If (m_numOfSharedBrushes > UBound(m_SharedBrushes)) Then ReDim Preserve m_SharedBrushes(0 To m_numOfSharedBrushes * 2 - 1) As SharedGDIBrush
        
        'Update the cache entry with new stats (including the created brush)
        m_SharedBrushes(m_numOfSharedBrushes).brushColor = requestedColor
        m_SharedBrushes(m_numOfSharedBrushes).numOfOwners = 1
        m_SharedBrushes(m_numOfSharedBrushes).brushHandle = CreateSolidBrush(requestedColor)
        
        'Return the newly created brush handle, increment the brush count, and exit
        GetSharedGDIBrush = m_SharedBrushes(m_numOfSharedBrushes).brushHandle
        m_numOfSharedBrushes = m_numOfSharedBrushes + 1
        
    End If
    
End Function

Public Sub ReleaseSharedGDIBrushByHandle(ByVal requestedHandle As Long)

    If (m_numOfSharedBrushes = 0) Then
        Debug.Print "FYI: UserControls.ReleaseSharedGDIBrushByHandle() received a release request, but no shared brushes exist."
        Exit Sub
    Else
        Dim i As Long
        For i = 0 To m_numOfSharedBrushes - 1
            If (m_SharedBrushes(i).brushHandle = requestedHandle) Then
                m_SharedBrushes(i).numOfOwners = m_SharedBrushes(i).numOfOwners - 1
                If (m_SharedBrushes(i).numOfOwners = 0) Then
                    DeleteObject m_SharedBrushes(i).brushHandle
                    m_SharedBrushes(i).brushHandle = 0
                    m_SharedBrushes(i).brushColor = 0
                End If
                Exit For
            End If
        Next i
    End If

End Sub

'Edit boxes can all share the same rendering font (as they are all themed identically).  Call this function instead
' of creating your own hFont for every text box instance.
Public Function GetSharedGDIFont(ByVal requestedSize As Single) As Long
    
    'First things first: if this is the first font requested by the system, initialize the shared font list
    If (m_numOfSharedFonts = 0) Then
        ReDim m_SharedFonts(0 To INIT_SIZE_OF_FONT_CACHE - 1) As SharedGDIFont
    End If
    
    'Next, look for the requested size in our current cache.  If it exists, we don't want to recreate it.
    Dim fontExists As Boolean, fontIndex As Long, i As Long
    fontExists = False
    
    If (m_numOfSharedFonts > 0) Then
            
        For i = 0 To m_numOfSharedFonts - 1
            If (m_SharedFonts(i).FontSize = requestedSize) Then
            
                'As a failsafe, make sure the owner count is valid too
                If (m_SharedFonts(i).numOfOwners > 0) And (m_SharedFonts(i).fontHandle <> 0) Then
                    fontExists = True
                    fontIndex = i
                    Exit For
                End If
            
            End If
        Next i
            
    End If
    
    'If we found the right font in our cache, increment the owner count, and return the handle immediately
    If fontExists Then
        m_SharedFonts(fontIndex).numOfOwners = m_SharedFonts(fontIndex).numOfOwners + 1
        GetSharedGDIFont = m_SharedFonts(fontIndex).fontHandle
    
    'If the font doesn't exist, create it anew
    Else
    
        'If the cache is too small, resize it
        If (m_numOfSharedFonts > UBound(m_SharedFonts)) Then ReDim Preserve m_SharedFonts(0 To m_numOfSharedFonts * 2 - 1) As SharedGDIFont
        
        'Font creation is cumbersome, but PD provides some helper functions to simplify it
        Dim tmpLogFont As LOGFONTW
        Fonts.FillLogFontW_Basic tmpLogFont, Fonts.GetUIFontName(), False, False, False, False
        Fonts.FillLogFontW_Size tmpLogFont, requestedSize, fu_Point
        Fonts.FillLogFontW_Quality tmpLogFont, TextRenderingHintClearTypeGridFit
        
        'Update the cache entry with new stats (including the created font)
        m_SharedFonts(m_numOfSharedFonts).FontSize = requestedSize
        m_SharedFonts(m_numOfSharedFonts).numOfOwners = 1
        If (Not Fonts.CreateGDIFont(tmpLogFont, m_SharedFonts(m_numOfSharedFonts).fontHandle)) Then
            PDDebug.LogAction "WARNING!  UserControls.GetSharedGDIFont() failed to create a new UI font handle."
        End If
        
        'Return the newly created font handle, increment the font count, and exit
        GetSharedGDIFont = m_SharedFonts(m_numOfSharedFonts).fontHandle
        m_numOfSharedFonts = m_numOfSharedFonts + 1
    
    End If
    
End Function

Public Sub ReleaseSharedGDIFontByHandle(ByVal requestedHandle As Long)

    If (m_numOfSharedFonts = 0) Then
        Debug.Print "FYI: UserControls.ReleaseSharedGDIFontByHandle() received a release request, but no shared fonts exist."
        Exit Sub
    Else
        Dim i As Long
        For i = 0 To m_numOfSharedFonts - 1
            If (m_SharedFonts(i).fontHandle = requestedHandle) Then
                m_SharedFonts(i).numOfOwners = m_SharedFonts(i).numOfOwners - 1
                If (m_SharedFonts(i).numOfOwners = 0) Then
                    Fonts.DeleteGDIFont m_SharedFonts(i).fontHandle
                    m_SharedFonts(i).fontHandle = 0
                    m_SharedFonts(i).FontSize = 0
                End If
                Exit For
            End If
        Next i
    End If

End Sub

'You can cache common (shared) translations here.  There is no separate Set function - just call this
' function to retrieve common translations, and it will automatically translate and store novel requests.
' (Note that you still need to use a dummy translation line somewhere to ensure the translated text is
' caught by PD's translation file generator.)
Public Function GetCommonTranslation(ByVal textKey As PD_UserControlText) As String
    
    If (m_CommonTranslations Is Nothing) Then GenerateCommonTranslations
    If m_CommonTranslations.DoesKeyExist(textKey) Then
        GetCommonTranslation = m_CommonTranslations.GetEntry_String(textKey)
    
    'Failsafe only; translations should never go missing!
    Else
        PDDebug.LogAction "WARNING: shared translation missing #" & textKey
    End If
    
End Function

'NOTE: this function should only be called once, on-demand, if a common translation is missing.
' It will auto-generate all common translations and cache them in a pdDictionary object.
Private Sub GenerateCommonTranslations()
    
    'Reset the shared translation object
    Set m_CommonTranslations = New pdDictionary
    
    'Generate all common translations
    
    'Animation controls appear many places now
    m_CommonTranslations.AddEntry pduct_AnimationRepeatToggle, g_Language.TranslateMessage("Toggle between 1x and repeating previews")
    
    'Command bars have a lot of tooltip text, owing to their ubiquity
    m_CommonTranslations.AddEntry pduct_CommandBarPresetList, g_Language.TranslateMessage("Previously saved presets can be selected here.  You can save the current settings as a new preset by clicking the Save Preset button on the right.")
    m_CommonTranslations.AddEntry pduct_CommandBarRandom, g_Language.TranslateMessage("Randomly select new settings for this tool.  This is helpful for exploring how different settings affect the image.")
    m_CommonTranslations.AddEntry pduct_CommandBarRedo, g_Language.TranslateMessage("Redo (fast-forward to a later state)")
    m_CommonTranslations.AddEntry pduct_CommandBarReset, g_Language.TranslateMessage("Reset all settings to their default values.")
    m_CommonTranslations.AddEntry pduct_CommandBarSavePreset, g_Language.TranslateMessage("Save the current settings as a new preset.")
    m_CommonTranslations.AddEntry pduct_CommandBarUndo, g_Language.TranslateMessage("Undo (rewind to an earlier state)")
    
    'Flyout panels share a common "lock this panel" explanation tooltip
    m_CommonTranslations.AddEntry pduct_FlyoutLockTitle, g_Language.TranslateMessage("Pin this panel open")
    m_CommonTranslations.AddEntry pduct_FlyoutLockTooltip, g_Language.TranslateMessage("Toolbox panels close automatically, but you can pin one to keep it open.  (Pinned panels still close when switching tools or opening new panels.)")
    
    'PD's built-in "randomize" control displays a tooltip for its "dice" button
    m_CommonTranslations.AddEntry pduct_Randomize, g_Language.TranslateMessage("Generate a new random number seed.")
    
End Sub

'If the active language changes, call this function to reset any shared translations
Public Sub ResetCommonTranslations()
    Set m_CommonTranslations = Nothing
End Sub

Public Sub ThemeFlyoutControls(ByRef cmdFlyoutLock As Variant)
    
    'Flyout lock controls use the same behavior across all instances
    Dim buttonSize As Long
    buttonSize = Interface.FixDPI(18)
    
    Dim i As Long
    For i = cmdFlyoutLock.lBound To cmdFlyoutLock.UBound
        cmdFlyoutLock(i).AssignImage "push_pin", Nothing, buttonSize, buttonSize
        cmdFlyoutLock(i).AssignTooltip UserControls.GetCommonTranslation(pduct_FlyoutLockTooltip), UserControls.GetCommonTranslation(pduct_FlyoutLockTitle)
        cmdFlyoutLock(i).Value = False
    Next i
    
End Sub

Public Function GetNameOfControlType(ByVal ctlType As PD_ControlType) As String
    
    Select Case ctlType
    
        Case pdct_Accelerator
            GetNameOfControlType = "pdAccelerator"
        Case pdct_BrushSelector
            GetNameOfControlType = "pdBrushSelector"
        Case pdct_Button
            GetNameOfControlType = "pdButton"
        Case pdct_ButtonStrip
            GetNameOfControlType = "pdButtonStrip"
        Case pdct_ButtonStripVertical
            GetNameOfControlType = "pdButtonStripVertical"
        Case pdct_ButtonToolbox
            GetNameOfControlType = "pdButtonToolbox"
        Case pdct_Canvas
            GetNameOfControlType = "pdCanvas"
        Case pdct_CanvasView
            GetNameOfControlType = "pdCanvasView"
        Case pdct_CheckBox
            GetNameOfControlType = "pdCheckBox"
        Case pdct_ColorDepth
            GetNameOfControlType = "pdColorDepth"
        Case pdct_ColorSelector
            GetNameOfControlType = "pdColorSelector"
        Case pdct_ColorVariants
            GetNameOfControlType = "pdColorVariants"
        Case pdct_ColorWheel
            GetNameOfControlType = "pdColorWheel"
        Case pdct_CommandBar
            GetNameOfControlType = "pdCommandBar"
        Case pdct_CommandBarMini
            GetNameOfControlType = "pdCommandBarMini"
        Case pdct_Container
            GetNameOfControlType = "pdContainer"
        Case pdct_Download
            GetNameOfControlType = "pdDownload"
        Case pdct_DropDown
            GetNameOfControlType = "pdDropDown"
        Case pdct_DropDownFont
            GetNameOfControlType = "pdDropDownFont"
        Case pdct_FxPreviewCtl
            GetNameOfControlType = "pdFxPreviewCtl"
        Case pdct_GradientSelector
            GetNameOfControlType = "pdGradientSelector"
        Case pdct_History
            GetNameOfControlType = "pdHistory"
        Case pdct_Hyperlink
            GetNameOfControlType = "pdHyperlink"
        Case pdct_ImageStrip
            GetNameOfControlType = "pdImageStrip"
        Case pdct_Label
            GetNameOfControlType = "pdLabel"
        Case pdct_LayerList
            GetNameOfControlType = "pdLayerList"
        Case pdct_LayerListInner
            GetNameOfControlType = "pdLayerListInner"
        Case pdct_ListBox
            GetNameOfControlType = "pdListBox"
        Case pdct_ListBoxOD
            GetNameOfControlType = "pdListBoxOD"
        Case pdct_ListBoxView
            GetNameOfControlType = "pdListBoxView"
        Case pdct_ListBoxViewOD
            GetNameOfControlType = "pdListBoxViewOD"
        Case pdct_MetadataExport
            GetNameOfControlType = "pdMetadataExport"
        Case pdct_Navigator
            GetNameOfControlType = "pdNavigator"
        Case pdct_NavigatorInner
            GetNameOfControlType = "pdNavigatorInner"
        Case pdct_NewOld
            GetNameOfControlType = "pdNewOld"
        Case pdct_PaletteUI
            GetNameOfControlType = "pdPaletteUI"
        Case pdct_PenSelector
            GetNameOfControlType = "pdPenSelector"
        Case pdct_PictureBox
            GetNameOfControlType = "pdPictureBox"
        Case pdct_PictureBoxInteractive
            GetNameOfControlType = "pdPictureBoxInteractive"
        Case pdct_Preview
            GetNameOfControlType = "pdPreview"
        Case pdct_ProgressBar
            GetNameOfControlType = "pdProgressBar"
        Case pdct_RadioButton
            GetNameOfControlType = "pdRadioButton"
        Case pdct_RandomizeUI
            GetNameOfControlType = "pdRandomizeUI"
        Case pdct_Resize
            GetNameOfControlType = "pdResize"
        Case pdct_Ruler
            GetNameOfControlType = "pdRuler"
        Case pdct_ScrollBar
            GetNameOfControlType = "pdScrollBar"
        Case pdct_SearchBar
            GetNameOfControlType = "pdSearchBar"
        Case pdct_Slider
            GetNameOfControlType = "pdSlider"
        Case pdct_SliderStandalone
            GetNameOfControlType = "pdSliderStandalone"
        Case pdct_Spinner
            GetNameOfControlType = "pdSpinner"
        Case pdct_StatusBar
            GetNameOfControlType = "pdStatusBar"
        Case pdct_Strip
            GetNameOfControlType = "pdStrip"
        Case pdct_TextBox
            GetNameOfControlType = "pdTextBox"
        Case pdct_Title
            GetNameOfControlType = "pdTitle"
        Case pdct_TreeviewOD
            GetNameOfControlType = "pdTreeviewOD"
        Case pdct_TreeviewViewOD
            GetNameOfControlType = "pdTreeviewViewOD"
            
    End Select
    
End Function

'Whenever a ucSupport instance is registered by a custom PD usercontrol, this function is called, and our running UC count
' is incremented.
Public Sub IncrementPDControlCount()
    m_PDControlCount = m_PDControlCount + 1
End Sub

Public Sub DecrementPDControlCount()
    m_PDControlCount = m_PDControlCount - 1
End Sub

Public Function GetPDControlCount() As Long
    GetPDControlCount = m_PDControlCount
End Function

'Whenever a dropdown raises its list box, call this function to set some program-wide flags.
' Subsequent focus events will also notify us, and we will kill the list box as necessary.
Public Sub NotifyDropDownChangeState(ByVal dropDownHWnd As Long, ByVal dropDownListHWnd As Long, ByVal newState As Boolean)
    
    If newState Then
        m_CurrentDropDownHWnd = dropDownHWnd
        m_CurrentDropDownListHWnd = dropDownListHWnd
    Else
        m_CurrentDropDownHWnd = 0
        m_CurrentDropDownListHWnd = 0
    End If

End Sub

'Whenever a pdPanel object raises a flyout panel, call this function to set some program-wide flags.
' Subsequent focus events will also notify us, and we will kill the flyout as necessary.
Public Sub NotifyFlyoutChangeState(ByVal flyoutParentHWnd As Long, ByVal flyoutPanelHWnd As Long, ByRef flyoutManager As pdFlyout, ByVal newState As Boolean)
    
    If newState Then
        m_CurrentFlyoutParentHWnd = flyoutParentHWnd
        m_CurrentFlyoutPanelHWnd = flyoutPanelHWnd
        Set m_FlyoutRef = flyoutManager
    Else
        m_CurrentFlyoutParentHWnd = 0
        m_CurrentFlyoutPanelHWnd = 0
        Set m_FlyoutRef = Nothing
    End If

End Sub

'Whenever a PD control loses or receives focus, we receive a corresponding notification
Public Sub PDControlReceivedFocus(ByVal controlHWnd As Long)
    
    'If a dropdown window is still active, hide it now
    HideOpenDropdowns controlHWnd
    
    'Do the same for flyout panels, but note that they're a little more complex because
    ' a single panel can host multiple (nested) controls, so we only close the flyout if focus
    ' has shifted to a control *not* hosted on the panel (or hosted on something hosted on the panel).
    HideOpenFlyouts controlHWnd
    
End Sub

'If a dropdown is open, this will release it
Private Sub HideOpenDropdowns(Optional ByVal hWndResponsible As Long = 0)

    If (m_CurrentDropDownHWnd <> 0) Or (m_CurrentDropDownListHWnd <> 0) Then
    
        If (m_CurrentDropDownHWnd <> hWndResponsible) And (m_CurrentDropDownListHWnd <> hWndResponsible) Then
            SetParent m_CurrentDropDownListHWnd, m_CurrentDropDownHWnd
            g_WindowManager.SetVisibilityByHWnd m_CurrentDropDownListHWnd, False
            m_CurrentDropDownHWnd = 0
            m_CurrentDropDownListHWnd = 0
        End If
    
    End If
    
End Sub

Public Sub HideOpenFlyouts(Optional ByVal hWndResponsible As Long = 0&)

    If (m_CurrentFlyoutParentHWnd <> 0) Or (m_CurrentFlyoutPanelHWnd <> 0) Then
        
        'Iterate the hWndResponsible and see if it *is* the panel or *shares* the panel parent.
        If (m_CurrentFlyoutPanelHWnd <> hWndResponsible) And (m_CurrentFlyoutParentHWnd <> hWndResponsible) Then
            
            'If the responsible hWnd is 0, it means we must hide the flyout immediately
            Dim targetSharesParentOrOwner As Boolean
            If (hWndResponsible <> 0) Then
                
                'Start testing parent controls of the control that now has focus and the flyout.
                ' (If they have the same parent, we'll leave the flyout open.)
                Dim testhWnd As Long
                testhWnd = GetParent(hWndResponsible)
                Do While (testhWnd <> 0) And (testhWnd <> hWndResponsible)
                    
                    'Found a match!  Flag and exit.
                    If (testhWnd = m_CurrentFlyoutPanelHWnd) Or (testhWnd = m_CurrentFlyoutParentHWnd) Then
                        targetSharesParentOrOwner = True
                        Exit Do
                    End If
                    
                    'No parent listed; either the target has no parent or its parent is the desktop.
                    If (testhWnd = 0) Or (testhWnd = GetDesktopWindow()) Then
                        targetSharesParentOrOwner = False
                        Exit Do
                    End If
                    
                    'No failure but no match; find the next parent in line
                    testhWnd = GetParent(testhWnd)
                    
                Loop
                
            'Skip ahead immediately; we need to hide the flyout regardless
            Else
                targetSharesParentOrOwner = False
            End If
            
            'If the new focused object does not share the same parent or owner as the current flyout, hide the flyout
            If (Not targetSharesParentOrOwner) Then
                
                'Normally, when a flyout panel's parent toolbox loses focus, we deactivate the flyout.
                ' However, the user can choose to "lock" a flyout in the open position.  In this state,
                ' we only hide the flyout for mandatory cases (like unloading the toolbox).
                If (Not m_FlyoutRef Is Nothing) Then
                    If (m_CurrentFlyoutPanelHWnd = m_FlyoutRef.GetLockedHWnd()) And (hWndResponsible <> 0) Then Exit Sub
                End If
                
                'If we have a reference to a flyout object, let it handle closure
                If (Not m_FlyoutRef Is Nothing) Then
                    m_FlyoutRef.HideFlyout
                
                'If we don't have a reference, something went awry - hide the flyout manually
                Else
                    SetParent m_CurrentFlyoutPanelHWnd, m_CurrentFlyoutParentHWnd
                    g_WindowManager.SetVisibilityByHWnd m_CurrentFlyoutPanelHWnd, False
                End If
                
                Set m_FlyoutRef = Nothing
                m_CurrentFlyoutPanelHWnd = 0
                m_CurrentFlyoutParentHWnd = 0
                
            End If
        
        '/new focus target is flyout panel or flyout panel parent
        End If
    
    '/no flyout hWnd tracked
    End If
    
End Sub

Public Sub PDControlLostFocus(ByVal controlHWnd As Long)
    
    'If this control raised a tooltip (and said tooltip is still active), unload it now
    If (controlHWnd = m_TTOwner) Then HideUCTooltip
    
End Sub

'When an object requests a tooltip, they need to pass a number of additional parameters (like the window rect, which is used to
' ideally position the tooltip).  Logic similar to pdDropDown is used to display the tooltip.
Public Sub ShowUCTooltip(ByVal ownerHwnd As Long, ByRef srcControlRect As RectL, ByVal mouseX As Single, ByVal mouseY As Single, ByRef ttCaption As String, ByRef ttTitle As String)
    
    On Error GoTo UnexpectedTTTrouble
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'We run into trouble when displaying a tooltip when a dropdown box is active.  (The two windows
    ' compete for top-level status, causing weird issues depending on Windows version.)  Rather than
    ' mess with this, we simply suspend tooltip display while a dropdown is active.
    If (m_CurrentDropDownHWnd <> 0) Or (m_CurrentDropDownListHWnd <> 0) Then Exit Sub
    
    'Alternatively, you could also close any drop-down boxes if they are still open; the following line
    ' accomplishes that.
    'HideOpenDropdowns
    
    'If a tooltip is currently active, suspend the release timer (because we're just going to "snap" the current
    ' tooltip window into place, rather than waiting for an animation).
    Dim ttAlreadyVisible As Boolean
    ttAlreadyVisible = ((Not m_TimerEventSink Is Nothing) And m_TTActive)
    If ttAlreadyVisible Then m_TimerEventSink.StopTTTimer
    
    'As a failsafe, make sure the requested tooltip's owner window is the same as our current window.
    ' (If it isn't, we want to perform full positioning calculations.)
    ttAlreadyVisible = ttAlreadyVisible And (m_TTOwner = ownerHwnd)
    m_TTOwner = ownerHwnd
    
    'Next, figure out how big the tooltip needs to be.  This is fairly cumbersome, as we need to fit both the caption
    ' and the title (if any), with text contents auto-wrapped if the tooltip is too long.
    Dim ttCaptionWidth As Long, ttCaptionHeight As Long
    Dim ttFont As pdFont
    Set ttFont = Fonts.GetMatchingUIFont(10)
    
    Dim availableWidth As Long
    availableWidth = Interface.FixDPI(PD_TT_MAX_WIDTH)
    
    'Tooltips can include linebreaks, so all size detection needs to be multiline-aware.
    Dim dtRect As RectL
    ttFont.GetBoundaryRectOfMultilineString ttCaption, availableWidth, dtRect, True
    ttCaptionWidth = dtRect.Right - dtRect.Left + 1
    
    If (ttCaptionWidth > availableWidth) Then
        
        'If the caption contains linebreaks, we can simply wordwrap
        If (InStr(1, ttCaption, vbCrLf, vbBinaryCompare) <> 0) Then
            ttCaptionWidth = availableWidth
            ttCaptionHeight = ttFont.GetHeightOfWordwrapString(ttCaption, ttCaptionWidth)
            
        'If the caption does not contain linebreaks, we need to make it as wide as possible
        Else
            ttCaptionHeight = dtRect.Bottom - dtRect.Top
        End If
        
    Else
        ttCaptionHeight = dtRect.Bottom - dtRect.Top
    End If
        
    'We now have a precise width/height measurement for the tooltip caption.
    ' Repeat the steps for the tooltip title, if any.
    Dim ttTitleWidth As Long, ttTitleHeight As Long
    If (LenB(ttTitle) > 0) Then
    
        Set ttFont = Fonts.GetMatchingUIFont(10, True)
        ttTitleWidth = ttFont.GetWidthOfString(ttTitle) + 1
        
        If (ttTitleWidth > availableWidth) Then
            ttTitleWidth = availableWidth
            ttTitleHeight = ttFont.GetHeightOfWordwrapString(ttTitle, ttTitleWidth)
        Else
            ttTitleHeight = ttFont.GetHeightOfString(ttTitle)
        End If
        
    Else
        ttTitleWidth = 0
        ttTitleHeight = 0
    End If
    
    'All font calculations use PD's shared UI font cache, so we need to free our fonts when we're done with them.
    Set ttFont = Nothing
        
    'With all string sizes calculated, we can now calculate a total size for the tooltip, including padding, borders,
    ' and spacing between the caption and title (if any).
    Dim internalPadding As Long
    internalPadding = Interface.FixDPI(PD_TT_INTERNAL_PADDING)
    
    Dim ttRect As RectF
    With ttRect
        .Width = internalPadding * 2 + PDMath.Max2Int(ttCaptionWidth, ttTitleWidth)
        If (ttTitleHeight > 0) Then
            .Height = internalPadding * 2 + ttCaptionHeight + ttTitleHeight + Interface.FixDPI(PD_TT_TITLE_PADDING)
        Else
            .Height = internalPadding * 2 + ttCaptionHeight
        End If
    End With
    
    'With our tooltip size correctly calculated, we now need to determine tooltip position.
    ' (If the tooltip is already visible, we'll skip this step, and instead attempt to re-use our current
    '  tooltip position.)
    If ttAlreadyVisible Then
        
        'Maintain the tooltip's current position as closely as possible
        ttRect.Left = m_TTRectCopy.Left
        ttRect.Top = m_TTRectCopy.Top
        
        'Based on the current tooltip position (remember: the tooltip is already visible),
        ' ensure that the tooltip doesn't overlap the underlying control window
        Select Case m_LastTTPosition
            Case TTS_Top
                ttRect.Top = srcControlRect.Top - (PD_TT_EXTERNAL_PADDING + ttRect.Height)
            Case TTS_Bottom
                ttRect.Top = srcControlRect.Bottom + PD_TT_EXTERNAL_PADDING + Interface.FixDPI(12)
            Case TTS_Right
                ttRect.Left = srcControlRect.Right + PD_TT_EXTERNAL_PADDING
            Case TTS_Left
                ttRect.Left = srcControlRect.Left - (PD_TT_EXTERNAL_PADDING + ttRect.Width)
        End Select
        
        'Note that we still need to ensure the tooltip does *not* lie off-screen.  We will handle this in a
        ' subsequent step.
        
    End If
    
    'Our goal is to position the tooltip as close to the mouse pointer as possible, while also positioning
    ' it outside the control rectangle (so that we don't obscure the control's contents.)
    
    'Start by figuring out which edge is closest to the current mouse position.  The passed mouse x/y ratios
    ' make this simple. (Each mouse value is a value [0, 1] instead of a hard-coded coordinate.)
    Dim mouseScreenPos As PointAPI
    mouseScreenPos.x = mouseX
    mouseScreenPos.y = mouseY
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.GetClientToScreen ownerHwnd, mouseScreenPos
    
    Dim ttPositions(0 To 3) As TT_Sort
    ttPositions(0).ttDistance = srcControlRect.Right - mouseScreenPos.x
    ttPositions(0).ttSide = TTS_Right
    ttPositions(1).ttDistance = mouseScreenPos.x - srcControlRect.Left
    ttPositions(1).ttSide = TTS_Left
    ttPositions(2).ttDistance = srcControlRect.Bottom - mouseScreenPos.y
    ttPositions(2).ttSide = TTS_Bottom
    ttPositions(3).ttDistance = mouseScreenPos.y - srcControlRect.Top
    ttPositions(3).ttSide = TTS_Top

    'Sort all distances from least to most.  We want to position the tooltip as close to the cursor as
    ' physically possible, but if that position is unavailable (due to lying off-screen), we'll try the
    ' next-closest position.
    Dim i As Long, j As Long, tmpSort As TT_Sort
    For i = 0 To 3
    For j = 0 To 3
        If (ttPositions(j).ttDistance < ttPositions(i).ttDistance) And (j > i) Then
            tmpSort = ttPositions(i)
            ttPositions(i) = ttPositions(j)
            ttPositions(j) = tmpSort
        End If
    Next j
    Next i
    
    'If the tooltip is already visible, shift its current position to the "top" of the list.
    ' (If the newly resized tooltip lies off-screen in its current position, we will still need to
    '  move through the list to find a better position for it - we just want to ensure that its
    '  current position is tried *first*.)
    If ttAlreadyVisible Then
    
        For i = 1 To 3
            
            If (ttPositions(i).ttSide = m_LastTTPosition) Then
                
                tmpSort = ttPositions(0)
                ttPositions(0) = ttPositions(i)
                
                If (i = 1) Then
                    ttPositions(1) = tmpSort
                Else
                    For j = 1 To i - 1
                        ttPositions(j) = ttPositions(j + 1)
                    Next j
                    ttPositions(i) = tmpSort
                End If
                
                Exit For
            
            End If
            
        Next i
    
    End If
    
    'If the tooltip is already visible, we want to "skip" our first attempt at positioning (because we're
    ' already using the tooltip's current position) - however, we still want to perform failsafe checks against
    ' things like "is the tooltip partially off-screen".
    Dim skipFirstPosition As Boolean
    skipFirstPosition = ttAlreadyVisible
    
    'By default, attempt to position the tooltip at the position nearest the cursor.
    ' (If this fails, we'll try again at the next-nearest position.)
    Dim ttIndex As Long
    ttIndex = 0
        
    Do
        
        m_LastTTPosition = ttPositions(ttIndex).ttSide
        
        'If we try all four tooltip positions, and all four fail, resort to the nearest position.
        ' (This should never happen, but better safe than sorry!)
        Dim failsafePosition As Boolean
        failsafePosition = (ttIndex > 3)
        If failsafePosition Then ttIndex = 0
        
        'Based on the current position (top/bottom/right/left), figure out where the top/left position
        ' of the tooltip should lie.
        If (Not skipFirstPosition) Then
        
            Select Case ttPositions(ttIndex).ttSide
                Case TTS_Top
                    ttRect.Top = srcControlRect.Top - (PD_TT_EXTERNAL_PADDING + ttRect.Height)
                
                'When positioning a tooltip on the bottom of a control, we need to add extra padding to account for the cursor.
                ' (Most cursors have their hotspot at the *top* of the cursor, so the cursor itself may extend below the control,
                '  impeding view of the tooltip.)
                Case TTS_Bottom
                    ttRect.Top = srcControlRect.Bottom + PD_TT_EXTERNAL_PADDING + Interface.FixDPI(12)
                Case TTS_Right
                    ttRect.Left = srcControlRect.Right + PD_TT_EXTERNAL_PADDING
                Case TTS_Left
                    ttRect.Left = srcControlRect.Left - (PD_TT_EXTERNAL_PADDING + ttRect.Width)
            End Select
            
        Else
            skipFirstPosition = False
        End If
    
        'Next, make sure that the tooltip lies on-screen.  (For this to work, we need to know the
        ' current screen dimensions; pdDisplays is used for this.)
        Dim hMonitor As Long
        hMonitor = g_Displays.GetHMonitorFromRectL(srcControlRect)
        
        Dim monitorRect As RectL
        g_Displays.GetDisplayByHandle(hMonitor).GetWorkingRect monitorRect
        
        Dim positionFailed As Boolean: positionFailed = False
        Select Case ttPositions(ttIndex).ttSide
            Case TTS_Top
                positionFailed = (ttRect.Top < monitorRect.Top)
            Case TTS_Bottom
                positionFailed = ((ttRect.Top + ttRect.Height) > monitorRect.Bottom)
            Case TTS_Right
                positionFailed = ((ttRect.Left + ttRect.Width) > monitorRect.Right)
            Case TTS_Left
                positionFailed = (ttRect.Left < monitorRect.Left)
        End Select
        
        If positionFailed And (Not failsafePosition) Then ttIndex = ttIndex + 1
        
    'Attempt to position on another side, as necessary
    Loop While positionFailed And (Not failsafePosition)
        
    'The tooltip's primary dimension has been properly set.  Next, calculate its secondary dimension.
    ' (Ideally, the secondary dimension is centered relative to the mouse hover position.  If this results in an
    ' off-screen tooltip, we automatically nudge it on-screen.)  Note that we can skip the position calculation
    ' if the tooltip is already visible - but we still need to check for "is it off-screen?"
    Select Case ttPositions(ttIndex).ttSide
        Case TTS_Top, TTS_Bottom
            If (Not ttAlreadyVisible) Then ttRect.Left = mouseScreenPos.x - ttRect.Width \ 2
            If (ttRect.Left < monitorRect.Left) Then ttRect.Left = monitorRect.Left
            If (ttRect.Left + ttRect.Width > monitorRect.Right) Then ttRect.Left = monitorRect.Right - ttRect.Width
            
        Case TTS_Right, TTS_Left
            If (Not ttAlreadyVisible) Then ttRect.Top = mouseScreenPos.y - ttRect.Height \ 2
            If (ttRect.Top < monitorRect.Top) Then ttRect.Top = monitorRect.Top
            If (ttRect.Top + ttRect.Height > monitorRect.Bottom) Then ttRect.Top = monitorRect.Bottom - ttRect.Height
            
    End Select
    
    'We have now calculated the tooltip position.  Time to display it!
    
    'The first time we raise the tooltip form, we want to cache its current window longs.  (We must restore these before
    ' unloading the form, or VB's built-in teardown functions will crash and burn.)
    If (m_TTHwnd = 0) Then
        
        Load tool_Tooltip
        m_TTHwnd = tool_Tooltip.hWnd
    
        If (Not m_TTWindowStyleHasBeenSet) Then
            m_TTWindowStyleHasBeenSet = True
            m_OriginalTTWindowBits = g_WindowManager.GetWindowLongWrapper(m_TTHwnd)
            m_OriginalTTWindowBitsEx = g_WindowManager.GetWindowLongWrapper(m_TTHwnd, True)
        End If
    
        'Overwrite VB's default window bits to ensure that the tooltip form behaves like a tooltip window.  Of particular
        ' importance is the WS_EX_NOACTIVATE option, to ensure that the tooltip does *not* receive focus.
        Const WS_POPUP As Long = &H80000000
        g_WindowManager.SetWindowLongWrapper m_TTHwnd, WS_POPUP, False, False, True
        g_WindowManager.SetWindowLongWrapper m_TTHwnd, WS_EX_NOACTIVATE Or WS_EX_TOOLWINDOW, False, True, True
        
        'Notify the window of the frame changes
        SetWindowPos m_TTHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER
        
    End If
    
    'Move the tooltip window into position *but do not display it* just yet.
    With ttRect
        SetWindowPos m_TTHwnd, 0&, .Left, .Top, .Width, .Height, SWP_NOACTIVATE
    End With
    
    'Cache the tooltip's display rect.  When the tooltip disappears, we will manually invalidate windows
    ' beneath it (only on certain OS + theme combinations; Aero handles this correctly).
    With m_TTRectCopy
        .Left = ttRect.Left
        .Top = ttRect.Top
        .Right = ttRect.Left + ttRect.Width
        .Bottom = ttRect.Top + ttRect.Height
    End With
    
    'As the last step before showing the tooltip, we need to notify the tooltip form of the tooltip caption
    ' and/or title.  It will cache these values and prepare internal rendering structs to match.
    tool_Tooltip.NotifyTooltipSettings ttCaption, ttTitle, PD_TT_INTERNAL_PADDING, PD_TT_TITLE_PADDING
    
    'We are finally ready to display the tooltip!
    If ttAlreadyVisible Then
        With ttRect
            SetWindowPos m_TTHwnd, 0&, .Left, .Top, .Width, .Height, SWP_SHOWWINDOW Or SWP_NOACTIVATE
        End With
    Else
        AnimateWindow m_TTHwnd, 150&, AW_BLEND
    End If
    
    m_TTActive = True
    
    Exit Sub
    
UnexpectedTTTrouble:
    PDDebug.LogAction "WARNING!  UserControls.ShowUCTooltip failed because of Err # " & Err.Number & ", " & Err.Description
    
End Sub

Public Sub HideUCTooltip(Optional ByVal hideImmediately As Boolean = False, Optional ByVal useAnimation As Boolean = True)
    
    If m_TTActive Then
        
        If hideImmediately Then
            HideTTImmediately useAnimation
        Else
        
            'Note the current time, then start the tooltip hide countdown
            If (m_TimerEventSink Is Nothing) Then
                Set m_TimerEventSink = New pdUCEventSink
            Else
                m_TimerEventSink.StopTTTimer
            End If
            
            m_InitialTTTimerTime = Timer
            m_TimerEventSink.StartTTTimer 100
            
        End If
        
    End If
        
End Sub

Public Sub TTTimerFired()
    
    'If enough time has passed, hide the tooltip and release the countdown timer
    If (Abs(Timer - m_InitialTTTimerTime) >= 0.5) Then HideTTImmediately
    
End Sub

Private Sub HideTTImmediately(Optional ByVal useAnimation As Boolean = True)

    If (Not m_TimerEventSink Is Nothing) Then m_TimerEventSink.StopTTTimer
    
    If m_TTActive And (m_TTHwnd <> 0) Then
        
        'Hide (but do not unload!) the tooltip window.  Animations can be suspended if there are interaction concerns
        ' (typically if the mouse is over the tooltip window area).
        If useAnimation Then
            AnimateWindow m_TTHwnd, 150&, AW_HIDE Or AW_BLEND
        Else
            g_WindowManager.SetVisibilityByHWnd m_TTHwnd, False
        End If
        
        'If Aero theming is not active, hiding the tooltip may cause windows beneath the current one to render incorrectly.
        If (OS.IsVistaOrLater And (Not g_WindowManager.IsDWMCompositionEnabled)) Then
            InvalidateRect 0&, VarPtr(m_TTRectCopy), 0&
        End If
        
    End If
    
    m_TTOwner = 0
    m_TTActive = False
        
End Sub

Public Function IsTooltipActive(ByVal ownerHwnd As Long) As Boolean
    IsTooltipActive = (m_TTOwner = ownerHwnd)
End Function

Public Sub NotifyTooltipThemeChange()

    'If the tooltip isn't active, ignore this event; the tooltip will automatically grab theme settings
    ' when it is first invoked.
    If (m_TTHwnd <> 0) Then tool_Tooltip.UpdateAgainstCurrentTheme

End Sub

'Do not call this function until the program is going down.  VB is very unhappy about changing window longs on the fly,
' so we only do it once, when the tooltip form is first raised.  After that, we keep the form in memory as-is, and do not
' touch its window longs again until the window is released.
Public Sub FinalTooltipUnload()
    
    'If a release timer is already active, release it immediately
    If (Not m_TimerEventSink Is Nothing) Then
        m_TimerEventSink.StopTTTimer
        Set m_TimerEventSink = Nothing
    End If
    
    If (m_TTHwnd <> 0) Then
    
        'Before doing anything else, ensure the window is invisible
        g_WindowManager.SetVisibilityByHWnd m_TTHwnd, False
        
        'Restore the original VB window bits; this ensures that teardown happens correctly
        If (m_OriginalTTWindowBits <> 0) Then g_WindowManager.SetWindowLongWrapper m_TTHwnd, m_OriginalTTWindowBits, , , True
        If (m_OriginalTTWindowBitsEx <> 0) Then g_WindowManager.SetWindowLongWrapper m_TTHwnd, m_OriginalTTWindowBits, , True, True
        
        'Windows caches window longs; ensure that our changes are applied immediately
        SetWindowPos m_TTHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_HIDEWINDOW Or SWP_NOSENDCHANGING
        
        'With all original settings restored, we can safely unload the tooltip window
        Unload tool_Tooltip
        Set tool_Tooltip = Nothing
        
        m_TTHwnd = 0
        
    End If
    
End Sub
