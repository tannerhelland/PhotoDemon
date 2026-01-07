VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PD Object List Manager
'Copyright 2017-2026 by Tanner Helland
'Created: 18/August/17
'Last updated: 27/October/21
'Last update: catch (and report) window maximize/restore events, in case we add availability to more tool windows
'
'This class is exclusively used by the NavKey module.  Look there for implementation details.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'If you want to dump tab order to the debug window, set this to TRUE.
Private Const DISPLAY_DEBUG_TABORDER_DATA As Boolean = False

'SetWindowPos flags
Private Enum SWP_FLAGS
    SWP_ASYNCWINDOWPOS = &H4000
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOSENDCHANGING = &H400
    SWP_NOSIZE = &H1
    SWP_HIDEWINDOW = &H80
    SWP_SHOWWINDOW = &H40
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = &H20
    SWP_NOCOPYBITS = &H100
End Enum

'MINMAXINFO struct
Private Type MinMaxInfo
    ptReserved As PointAPI
    ptMaxSize As PointAPI
    ptMaxPosition As PointAPI
    ptMinTrackSize As PointAPI
    ptMaxTrackSize As PointAPI
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByVal ptrToPointList As Long, ByVal numPoints As Long) As Long
Private Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As SWP_FLAGS)

'The hWnd of the parent (the form hosting this collection)
Private m_ParentHWnd As Long

'The WindowRect *in screen coords* of the parent hWnd, at load-time
Private m_ParentRectScreen As winRect, m_ParentRectClient As winRect

'Handle auto-resize of child controls for this window?
Private m_AutoResize As Boolean

'Because this class already tracks window relationships (particularly controls against their parent forms),
' it provides a natural place to generate run-time object names.  PD uses these to generate unique control names
' that localizers can use for translation purposes (e.g. "toolbar_Layers.layerpanel_Layers.lblLayerSettings(2)")
Private m_parentName As String

'hWnds of child controls
Private Const INIT_NUM_OF_CONTROLS As Long = 16

'As part of subclassing the parent control, we want to detect ENTER and ESC keypresses.  Dialogs with
' command bars will automatically assign those keypresses to OK/Cancel.  As of 9.0, we'll also handle
' min/max messages (since users can now resize some windows!).
Private Const WM_ENTERSIZEMOVE As Long = &H231
Private Const WM_EXITSIZEMOVE As Long = &H232
Private Const WM_GETMINMAXINFO As Long = &H24
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_SIZE As Long = &H5

'We may not need to use a custom type for this - an hWnd may ultimately be enough - but during debugging,
' it's nice to store some extra bits of information
Private Type PD_Control
    hWnd As Long                'Actual control hWnd
    ctlRectLoadTime As winRect  'The control's rectangle *when it was loaded* (used to auto-calculate tab order)
    ctlName As String           'Control name (helpful for debugging)
    ctlType As PD_ControlType   'Internal PD user control enum
    ctlCanGetFocus As Boolean   'Some controls register here but do *not* receive focus; we skip them when assigning tab order
End Type

Private m_Controls() As PD_Control
Private m_numOfControls As Long

'Some control indices are stored specially; this is useful for subsequent loops
Private m_CommandBarIndex As Long, m_PreviewIndex As Long

'Some windows (animation effects, at present) may implement custom preview methods that don't utilize
' a traditional pdPreviewFX control.  These windows can specify a different control to use as the
' "anchor" during resizing events; as long as that control provides PD's traditional position interface,
' it'll work just fine.
'
'If this value is zero, no special window exists; rely on m_PreviewIndex above, instead.
Private m_SpecialPreviewHWnd As Long

'When tab is pressed, we sort the current control list according to each control's top-left position.
' This sorted list can then be used to easily determine the next/prev control in order.
Private Type PD_ControlSort
    hWnd As Long
    sortKey As Long
    x As Long
    y As Long
End Type

Private m_ControlsSort() As PD_ControlSort
Private m_numOfControlsSort As Long

'Subclasser for intercepting window messages
Implements ISubclass

'Start a subclasser against the target hWnd
Friend Sub SetParentHWnd(ByVal hWnd As Long, ByVal handleAutoResize As Boolean, Optional ByVal hWndCustomAnchor As Long = 0, Optional ByRef parentName As String = vbNullString)
    
    If (PDMain.IsProgramRunning() And (hWnd <> 0)) Then
    
        m_ParentHWnd = hWnd
        GetWindowRect m_ParentHWnd, VarPtr(m_ParentRectScreen)
        GetClientRect m_ParentHWnd, VarPtr(m_ParentRectClient)
        m_AutoResize = handleAutoResize
        m_parentName = parentName
        
        'Some dialogs use custom preview windows; they need to notify us of this, so we can anchor
        ' that control specially
        m_SpecialPreviewHWnd = hWndCustomAnchor
        
        'We also subclass the parent control so we can capture ENTER and ESCAPE keypresses (to facilitate
        ' keyboard navigation)
        VBHacks.StartSubclassing hWnd, Me
        
    End If
    
End Sub

Friend Function GetParentHWnd() As Long
    GetParentHWnd = m_ParentHWnd
End Function

Friend Function DoesHWndExist(ByVal srcHWnd As Long) As Boolean
    
    DoesHWndExist = False
    
    If (m_numOfControls > 0) Then
        Dim i As Long
        For i = 0 To m_numOfControls - 1
            If (m_Controls(i).hWnd = srcHWnd) Then
                DoesHWndExist = True
                Exit For
            End If
        Next i
    End If
    
End Function

Friend Function DoesTypeOfControlExist(ByVal srcType As PD_ControlType) As Boolean
    
    DoesTypeOfControlExist = False
    
    If (m_numOfControls > 0) Then
        Dim i As Long
        For i = 0 To m_numOfControls - 1
            If (m_Controls(i).ctlType = srcType) Then
                DoesTypeOfControlExist = True
                Exit For
            End If
        Next i
    End If
    
End Function

Friend Function GetFirstHWndForType(ByVal srcType As PD_ControlType) As Long

    GetFirstHWndForType = 0
    
    If (m_numOfControls > 0) Then
        Dim i As Long
        For i = 0 To m_numOfControls - 1
            If (m_Controls(i).ctlType = srcType) Then
                GetFirstHWndForType = m_Controls(i).hWnd
                Exit For
            End If
        Next i
    End If

End Function

Friend Function GetParentName() As String
    GetParentName = m_parentName
End Function

'Load individual controls using this function; we'll retrieve whatever we need from them
Friend Sub NotifyChildControl(ByRef childControl As Object, ByVal canReceiveFocus As Boolean)

    'Make sure we have room for this hWnd
    If (m_numOfControls = 0) Then
        ReDim m_Controls(0 To INIT_NUM_OF_CONTROLS - 1) As PD_Control
    Else
        If (m_numOfControls > UBound(m_Controls)) Then ReDim Preserve m_Controls(0 To m_numOfControls * 2 - 1) As PD_Control
    End If
    
    'Store all load-time values
    With m_Controls(m_numOfControls)
        .hWnd = childControl.hWnd
        GetWindowRect .hWnd, VarPtr(.ctlRectLoadTime)
        MapWindowPoints 0&, m_ParentHWnd, VarPtr(.ctlRectLoadTime), 2&
        If DISPLAY_DEBUG_TABORDER_DATA Then .ctlName = childControl.GetControlName()
        .ctlType = childControl.GetControlType()
        .ctlCanGetFocus = canReceiveFocus
        If (.ctlType = pdct_CommandBar) Or (.ctlType = pdct_CommandBarMini) Then m_CommandBarIndex = m_numOfControls
        If (.ctlType = pdct_FxPreviewCtl) Then m_PreviewIndex = m_numOfControls
        If (m_SpecialPreviewHWnd <> 0) And (.hWnd = m_SpecialPreviewHWnd) Then m_PreviewIndex = m_numOfControls
    End With
    
    m_numOfControls = m_numOfControls + 1
    
End Sub

'Tab keypresses only require an incoming hWnd; we'll figure out the rest.
Friend Sub NotifyTabKey(ByVal srcHWnd As Long, ByVal shiftTabPressed As Boolean)
    
    'Before doing anything else, ask the source control if it wants to specify a custom tab-key target.
    ' If it does, we'll use its suggestion instead of determining our own.
    Dim newTargetHwnd As Long: newTargetHwnd = 0
    If shiftTabPressed Then
        UserControls.PostPDMessage WM_PD_SHIFT_TAB_KEY_TARGET, srcHWnd, VarPtr(newTargetHwnd)
    Else
        UserControls.PostPDMessage WM_PD_TAB_KEY_TARGET, srcHWnd, VarPtr(newTargetHwnd)
    End If
    
    If (newTargetHwnd = 0) Then
    
        'First thing we need to do is make a list of all visible/enabled controls on the form.
        ' (Invisible/disabled controls can't receive focus, obviously!)
        ReDim m_ControlsSort(0 To m_numOfControls) As PD_ControlSort
        m_numOfControlsSort = 0
        
        'We also need the width of the current screen, so we can sort coordinates one-dimensionally.
        Dim screenWidth As Long
        screenWidth = g_Displays.GetDesktopWidth()
        
        Dim currentControlIndex As Long
        currentControlIndex = -1
        
        Dim i As Long, tmpRect As winRect
        For i = 0 To m_numOfControls - 1
        
            'Only visible + enabled controls are eligible for receiving focus
            If (IsWindowVisible(m_Controls(i).hWnd) <> 0) And (IsWindowEnabled(m_Controls(i).hWnd)) Then
                
                'Some controls are tracked by this class for auto-resize purposes, but we don't want them
                ' involved in tab key events because they are just containers (or they aren't meant to
                ' receive focus, like labels.)
                If m_Controls(i).ctlCanGetFocus Then
                
                    'This control is visible, enabled, and interactive, making it a candidate for receiving focus.
                    ' Retrieve its coordinates.
                    g_WindowManager.GetWindowRect_API m_Controls(i).hWnd, tmpRect
                    
                    'I haven't decided if it is worthwhile to map coordinates to a new coordinate space prior to
                    ' determining order.  (I don't think it matters, but it's possible I haven't fully considered the math!)
                    
                    'For now, convert the top-left corner of the rect into a single-dimension variable
                    With m_ControlsSort(m_numOfControlsSort)
                        .hWnd = m_Controls(i).hWnd
                        .x = tmpRect.x1
                        .y = tmpRect.y1
                        .sortKey = GetSortKeyFromRect(tmpRect, screenWidth)
                    End With
                    
                    'If this control is the one that supplied the tabkey, note it now
                    If (srcHWnd = m_Controls(i).hWnd) Then currentControlIndex = m_numOfControlsSort
                    
                    m_numOfControlsSort = m_numOfControlsSort + 1
                    
                End If
                
            End If
        Next i
        
        'We now have a list of all valid tab recipients on this form.  Hopefully our source control was included;
        ' if it wasn't (for whatever reason), retrieve its last-known position and use that instead.
        If (currentControlIndex = -1) Then
            
            g_WindowManager.GetWindowRect_API srcHWnd, tmpRect
            
            With m_ControlsSort(m_numOfControlsSort)
                .hWnd = srcHWnd
                .x = tmpRect.x1
                .y = tmpRect.y1
                .sortKey = GetSortKeyFromRect(tmpRect, screenWidth)
            End With
            
            m_numOfControlsSort = m_numOfControlsSort + 1
            
        End If
        
        'Our list of "valid" window targets is now guaranteed to include the source window that triggered this tab press
        ' in the first place!  (We need it in the list, obviously, so we know which control(s) surround it in the tab order.)
        
        'Next, we need to sort the list by its sortKey property.  This list is guaranteed to be small, so we shouldn't
        ' need a fancy sort.  An in-place insertion sort (as used elsewhere in the project) should be more than sufficient.
        If (m_numOfControlsSort > 1) Then
        
            Dim j As Long, loopBound As Long, tmpRef As PD_ControlSort
            loopBound = m_numOfControlsSort - 1
            
            'Loop through all entries in the stack, sorting them as we go
            For i = 0 To loopBound
                For j = 0 To loopBound
                    If (m_ControlsSort(i).sortKey < m_ControlsSort(j).sortKey) Then
                        tmpRef = m_ControlsSort(i)
                        m_ControlsSort(i) = m_ControlsSort(j)
                        m_ControlsSort(j) = tmpRef
                    End If
                Next j
            Next i
            
        'If there is only one (or zero) valid tab key recipient(s) on this dialog, skip the sort step, obviously.
        Else
        
        End If
        
        'Now that our list is sorted, we need to once again find the source window's hWnd.
        For i = 0 To m_numOfControlsSort - 1
            If (m_ControlsSort(i).hWnd = srcHWnd) Then
                currentControlIndex = i
                Exit For
            End If
        Next i
        
        'With a sorted list of controls, finding the next/previous control is easy!
        Dim targetIndex As Long
        
        If shiftTabPressed Then
            targetIndex = currentControlIndex - 1
            If (targetIndex < 0) Then targetIndex = m_numOfControlsSort - 1
        Else
            targetIndex = currentControlIndex + 1
            If (targetIndex >= m_numOfControlsSort) Then targetIndex = 0
        End If
        
        'During debug sessions, it can be helpful to print window details to the immediate window
        If DISPLAY_DEBUG_TABORDER_DATA Then
            Debug.Print "---------"
            For i = 0 To m_numOfControlsSort - 1
                If (i = targetIndex) Then
                    Debug.Print "> " & GetControlNameFromHWnd(m_ControlsSort(i).hWnd), m_ControlsSort(i).sortKey
                ElseIf (i = currentControlIndex) Then
                    Debug.Print "* " & GetControlNameFromHWnd(m_ControlsSort(i).hWnd), m_ControlsSort(i).sortKey
                Else
                    Debug.Print GetControlNameFromHWnd(m_ControlsSort(i).hWnd), m_ControlsSort(i).sortKey
                End If
            Next i
            Debug.Print "FYI, target control is " & GetControlNameFromHWnd(m_ControlsSort(targetIndex).hWnd)
        End If
        
        newTargetHwnd = m_ControlsSort(targetIndex).hWnd
        
    End If
        
    'Some controls require special focus notifications (e.g. spinners, which default to
    ' the edit box receiving focus via tab key - NOT the spin control).  Check for this now.
    Dim useSpecialFocusEvent As Long: useSpecialFocusEvent = 0
    UserControls.PostPDMessage WM_PD_FOCUS_FROM_TAB_KEY, newTargetHwnd, VarPtr(useSpecialFocusEvent)
    
    'Finally, apply the focus change!  (If the previous step succeeded, the caller must set
    ' useSpecialFocusEvent to a non-zero value - this means they handled the focus event
    ' internally, so we don't need to handle it for them.)
    If (useSpecialFocusEvent = 0) And (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI newTargetHwnd
    
End Sub

'Return a sort key for a given control's window rectangle (e.g. the control's coordinates and dimensions,
' in screen coordinate space).  This function will convert those four parameters into a single, easily
' sortable key.
Private Function GetSortKeyFromRect(ByRef srcRect As winRect, ByVal screenWidth As Long) As Long
    
    'On-screen controls can be complicated to sort, because their positions are often "nudged" a few pixels
    ' in any given direction to make things "look nice".  (Typically, this is done to keep elements
    ' center-aligned vertically along a shared horizontal axis, like the buttons and dropdowns on a
    ' command bar.)
    
    'To make sure we intelligently sort controls, we have to modify their rect in various ways.  I'll try to
    ' explain the settings as we go.
    
    'First, let's deal with "normal-sized" on-screen elements.  90+% of PD controls fit into this class.
    ' These controls are typically less than 80 pixels in height (multiplied by the current system DPI).
    ' These controls are the most likely ones to be "nudged" into aesthetically pleasing positions
    ' relative to neighboring controls.  When sorting them, we sort them along their centerline instead
    ' of by their top-left corner.
    Const CONTROL_MAX_HEIGHT As Long = 80
    
    Dim sortX As Long, sortY As Long
    If ((srcRect.y2 - srcRect.y1) <= CONTROL_MAX_HEIGHT) Then
        
        sortX = srcRect.x1
        sortY = srcRect.y1 + ((srcRect.y2 - srcRect.y1) \ 2)
        
    'If this control is taller than our pixel threshold, it is probably a very large element like a preview box
    ' or a custom-built control (like the Curves dialog).  Sort it by its top-left corner only.
    Else
    
        sortX = srcRect.x1
        sortY = srcRect.y1
    
    End If
    
    'When sorting on-screen elements, we clamp y-coordinates to their nearest multiple of 12.  This helps
    ' address the case where nearby elements vary in top-position by only a few pixels, as is necessary
    ' to maintain an appearance of center-line alignment (e.g. again refer to command bars, where the
    ' "preset names" dropdown sits just slightly below the neighboring buttons, because it is slightly
    ' shorter vertically than the button images).  To ensure such controls are still sorted in normal
    ' LTR order, instead of being treated as "on the next line", we clamp their controls to a fixed
    ' multiple of 8.
    Const MAX_VERTICAL_DISTANCE As Long = 12
    GetSortKeyFromRect = sortX + (Int((sortY + (MAX_VERTICAL_DISTANCE \ 2)) \ MAX_VERTICAL_DISTANCE) * MAX_VERTICAL_DISTANCE * screenWidth)
    
End Function

Private Function GetControlNameFromHWnd(ByVal srcHWnd As Long) As String

    Dim i As Long
    For i = 0 To m_numOfControls - 1
        If (srcHWnd = m_Controls(i).hWnd) Then
            GetControlNameFromHWnd = m_Controls(i).ctlName & " (" & m_Controls(i).ctlType & ")"
            Exit For
        End If
    Next i

End Function

Friend Sub PrintDebugList()
    If (m_numOfControls > 0) Then
        Dim i As Long
        For i = 0 To m_numOfControls - 1
            Debug.Print m_Controls(i).ctlName, m_Controls(i).hWnd, UserControls.GetNameOfControlType(m_Controls(i).ctlType)
        Next i
    End If
End Sub

Private Sub PerformAutoResize()
    
    'Start by retrieving the current rect of the window in question
    Dim curParentRect As winRect, curParentClientRect As winRect
    Dim curParentClientWidth As Long, curParentClientHeight As Long
    GetWindowRect m_ParentHWnd, VarPtr(curParentRect)
    GetClientRect m_ParentHWnd, VarPtr(curParentClientRect)
    curParentClientWidth = curParentClientRect.x2 - curParentClientRect.x1
    curParentClientHeight = curParentClientRect.y2 - curParentClientRect.y1
    
    'We now need to iterate all child controls and right-align all of them EXCEPT:
    ' - previewFX controls (these get resized to fill)
    ' - command bars (these are bottom-aligned)
    ' - controls sitting on command bars (these are aligned automatically by the command bar)
    ' - controls sitting inside a container *other* than the parent form (parent controls will auto-align them)
    Dim i As Long, newLeft As Long, newTop As Long, newWidth As Long, newHeight As Long
    For i = 0 To m_numOfControls - 1
        
        With m_Controls(i)
            
            'Handle command bar specially (bottom-aligned)
            If (i = m_CommandBarIndex) Then
                SetWindowPos .hWnd, 0, 0, curParentClientHeight - (m_ParentRectClient.y2 - .ctlRectLoadTime.y1), curParentClientWidth, curParentClientHeight, SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
            
            'Preview boxes are the only control that gets resized in both directions
            ElseIf (i = m_PreviewIndex) Then
            
                'Resize to an identical proportional amount of the dialog
                newWidth = curParentClientWidth - ((m_ParentRectClient.x2 - m_ParentRectClient.x1) - (.ctlRectLoadTime.x2 - .ctlRectLoadTime.x1))
                newHeight = curParentClientHeight - ((m_ParentRectClient.y2 - m_ParentRectClient.y1) - (.ctlRectLoadTime.y2 - .ctlRectLoadTime.y1))
                SetWindowPos .hWnd, 0, .ctlRectLoadTime.x1, .ctlRectLoadTime.y1, newWidth, newHeight, SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
                
            'any other controls WHOSE PARENT IS THE FORM ITSELF get right-aligned
            Else
            
                'See if this control sits below the command bar; if it does, ignore it.
                If (m_CommandBarIndex >= 0) Then
                    
                    If (.ctlRectLoadTime.y1 < m_Controls(m_CommandBarIndex).ctlRectLoadTime.y1) Then
                        
                        'Only move controls whose parent is the form itself; anything else is a nested control whose position
                        ' will be handled automatically when we move its parent.
                        If (GetParent(.hWnd) = m_ParentHWnd) Then
                            
                            'Next, some esoteric dialogs (e.g. curves) place some settings *beneath* the
                            ' preview box.  These require horizontal resizing to match whatever we've done
                            'to the preview box.
                            Dim changeHorizontalSize As Boolean
                            If (m_PreviewIndex >= 0) Then changeHorizontalSize = (.ctlRectLoadTime.x1 < (m_Controls(m_PreviewIndex).ctlRectLoadTime.x2 - 8))   '8px is an arbitrary failsafe for tight-fitting controls
                            
                            If changeHorizontalSize Then
                                
                                'Height stays the same.
                                
                                'Figure out preview window width (which may not have been calculated yet, since it
                                ' depends on the order of control enumeration according to VB)
                                newWidth = curParentClientWidth - ((m_ParentRectClient.x2 - m_ParentRectClient.x1) - (m_Controls(m_PreviewIndex).ctlRectLoadTime.x2 - m_Controls(m_PreviewIndex).ctlRectLoadTime.x1))
                                
                                'Left only gets changed if it is *not* aligned with the preview object
                                If (.ctlRectLoadTime.x1 > (m_Controls(m_PreviewIndex).ctlRectLoadTime.x1 + 8)) Then
                                    newLeft = (.ctlRectLoadTime.x1 * (newWidth + m_Controls(m_PreviewIndex).ctlRectLoadTime.x1)) / m_Controls(m_PreviewIndex).ctlRectLoadTime.x2
                                Else
                                    newLeft = .ctlRectLoadTime.x1
                                End If
                                
                                'Top gets moved to an identical offset from the parent window's *bottom*
                                newTop = curParentClientHeight - (m_ParentRectClient.y2 - .ctlRectLoadTime.y1)
                                
                                'Width is modified as a fraction of the original preview control's width
                                ' (relative to the new preview control width); this allows complex arrangements
                                ' to *mostly* look good - see e.g. the Effects > Animation menu.
                                newWidth = ((.ctlRectLoadTime.x2 - .ctlRectLoadTime.x1) / (m_Controls(m_PreviewIndex).ctlRectLoadTime.x2 - m_Controls(m_PreviewIndex).ctlRectLoadTime.x1)) * newWidth
                                
                                'Apply the changes!
                                SetWindowPos .hWnd, 0, newLeft, newTop, newWidth, .ctlRectLoadTime.y2 - .ctlRectLoadTime.y1, SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
                                
                            'Change left parameter only; top/width/height can remain as-is.
                            Else
                            
                                newLeft = curParentClientWidth - (m_ParentRectClient.x2 - .ctlRectLoadTime.x1)
                                
                                'We have two options here: we can either "reflow" items vertically
                                ' (which will place items at a proportionally equivalent vertical position)
                                ' ...or we can just leave items at their original top position.  It's hard
                                ' to unequivocally call one or the other "better", since it really depends
                                ' on the dialog.  For now, I default to the simpler option (leaving items
                                ' in their current position).
                                
                                'Leave in current position:
                                newTop = .ctlRectLoadTime.y1
                                
                                'Place in proportionate position (to new size):
                                'newTop = (.ctlRectLoadTime.y1 / (m_ParentRectClient.y2 - m_ParentRectClient.y1)) * curParentClientHeight
                                
                                SetWindowPos .hWnd, 0, newLeft, newTop, .ctlRectLoadTime.x2 - .ctlRectLoadTime.x1, .ctlRectLoadTime.y2 - .ctlRectLoadTime.y1, SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
                                
                            End If
                            
                        End If
                        
                    End If
                
                'Failsafe check to ensure a command-bar was found
                Else
                    PDDebug.LogAction "WARNING: pdObjectList.PerformAutoResize didn't find a command-bar!"
                End If
                
            End If
        
        End With
        
    Next i
    
End Sub

'You *must* pass lParam; this is a pointer to a MINMAXINFO struct which holds the actual size information
Private Sub HandleMinMaxMsg(ByVal lParam As Long)
    
    Dim copyOfStruct As MinMaxInfo
    CopyMemoryStrict VarPtr(copyOfStruct), lParam, LenB(copyOfStruct)
    
    'We don't enforce a max size - only a min size
    copyOfStruct.ptMinTrackSize.x = PDMath.Min2Int(Interface.FixDPI(800), m_ParentRectScreen.x2 - m_ParentRectScreen.x1)
    copyOfStruct.ptMinTrackSize.y = PDMath.Min2Int(Interface.FixDPI(600), m_ParentRectScreen.y2 - m_ParentRectScreen.y1)
    CopyMemoryStrict lParam, VarPtr(copyOfStruct), LenB(copyOfStruct)
    
End Sub

Private Sub Class_Initialize()
    m_CommandBarIndex = -1
    m_PreviewIndex = -1
End Sub

'WM_EXITSIZEMOVE: set query-able flags upon resize start/finish so child controls can query it
Private Function HandleMsg_WM_EXITSIZEMOVE(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long, ByRef eatMsg As Boolean) As Long
    
    Interface.SetDialogResizeFlag False
        
    'Immediately notify the preview control, if any, of the finished resize event; it will use
    ' this to immediately generate a new preview.
    If (m_PreviewIndex >= 0) Then
        
        'If this is a normal preview control, send the message directly to it
        If (m_SpecialPreviewHWnd = 0) Then
            SendNotifyMessage m_Controls(m_PreviewIndex).hWnd, WM_PD_DIALOG_RESIZE_FINISHED, 0&, 0&
        
        'Otherwise, we need to notify the arbitrary control specified by the user
        Else
            SendNotifyMessage m_SpecialPreviewHWnd, WM_PD_DIALOG_RESIZE_FINISHED, 0&, 0&
        End If
        
    End If
        
End Function

'WM_GETMINMAXINFO: Windows with auto-resizing get "free" minimum size handling (we don't let users shrink
' them below their original size, or 800x600, whichever is smaller).
Private Function HandleMsg_WM_GETMINMAXINFO(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long, ByRef eatMsg As Boolean) As Long
    If m_AutoResize Then
        If IsWindowVisible(hWnd) Then
            HandleMinMaxMsg lParam
            HandleMsg_WM_GETMINMAXINFO = 0
            eatMsg = True
        End If
    End If
End Function

'WM_KEYDOWN: Enter and Escape are handled automatically
Private Function HandleMsg_WM_KEYDOWN(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long, ByRef eatMsg As Boolean) As Long

    'The only keypresses we currently care about are ENTER and ESCAPE
    If (wParam = pdnk_Enter) Or (wParam = pdnk_Escape) Or (wParam = pdnk_Space) Then
    
        'See if this form 1) is a raised dialog, and 2) contains a command bar
        If Interface.IsModalDialogActive() Then
        
            If Me.DoesTypeOfControlExist(pdct_CommandBar) Then
            
                'It does!  Grab the hWnd and forward the relevant window message to it
                SendNotifyMessage Me.GetFirstHWndForType(pdct_CommandBar), WM_PD_DIALOG_NAVKEY, wParam, 0&
                HandleMsg_WM_KEYDOWN = 0&
                eatMsg = True
            
            'If a command bar doesn't exist, look for a "mini command bar" instead
            ElseIf Me.DoesTypeOfControlExist(pdct_CommandBarMini) Then
                SendNotifyMessage Me.GetFirstHWndForType(pdct_CommandBarMini), WM_PD_DIALOG_NAVKEY, wParam, 0&
                HandleMsg_WM_KEYDOWN = 0&
                eatMsg = True
                
            'No command bar exists on this form, which is fine - this could be a toolpanel, for example.
            ' As such, there's nothing we need to do.
            End If
        
        End If
        
    End If
    
End Function

'WM_SIZE: Windows with auto-resizing are also handled specially
Private Function HandleMsg_WM_SIZE(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long, ByRef eatMsg As Boolean) As Long

    If m_AutoResize Then
        If IsWindowVisible(hWnd) Then PerformAutoResize
    End If
    
    'Maximize and Restore events also require special handling
    Const SIZE_MAXIMIZED As Long = 2, SIZE_RESTORE As Long = 0
    If (wParam = SIZE_MAXIMIZED) Or (wParam = SIZE_RESTORE) Then
    
        'Immediately notify the preview control, if any, of the finished resize event; it will use
        ' this to immediately generate a new preview.
        If (m_PreviewIndex >= 0) Then
            
            'If this is a normal preview control, send the message directly to it
            If (m_SpecialPreviewHWnd = 0) Then
                SendNotifyMessage m_Controls(m_PreviewIndex).hWnd, WM_PD_DIALOG_RESIZE_FINISHED, 0&, 0&
            
            'Otherwise, we need to notify the arbitrary control specified by the user
            Else
                SendNotifyMessage m_SpecialPreviewHWnd, WM_PD_DIALOG_RESIZE_FINISHED, 0&, 0&
            End If
            
        End If
        
    End If
    
End Function

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    
    Dim eatMsg As Boolean: eatMsg = False
    
    'Enter and Escape are handled automatically
    If (uiMsg = WM_KEYDOWN) Then
        ISubclass_WindowMsg = HandleMsg_WM_KEYDOWN(hWnd, uiMsg, wParam, lParam, dwRefData, eatMsg)
    
    'Windows with auto-resizing are also handled specially
    ElseIf (uiMsg = WM_SIZE) Then
        ISubclass_WindowMsg = HandleMsg_WM_SIZE(hWnd, uiMsg, wParam, lParam, dwRefData, eatMsg)
        
    'Windows with auto-resizing also get "free" minimum size handling (we don't let users shrink them below
    ' their original size, or 800x600, whichever is smaller).
    ElseIf (uiMsg = WM_GETMINMAXINFO) Then
        ISubclass_WindowMsg = HandleMsg_WM_GETMINMAXINFO(hWnd, uiMsg, wParam, lParam, dwRefData, eatMsg)
        
    'Set query-able flags upon resize start/finish so our child controls can query it
    ElseIf (uiMsg = WM_ENTERSIZEMOVE) Then
        Interface.SetDialogResizeFlag True
        
    ElseIf (uiMsg = WM_EXITSIZEMOVE) Then
        ISubclass_WindowMsg = HandleMsg_WM_EXITSIZEMOVE(hWnd, uiMsg, wParam, lParam, dwRefData, eatMsg)
     
    'Failsafe window destruction check
    ElseIf (uiMsg = WM_NCDESTROY) Then
        EndSubclassing
    End If
    
    If (Not eatMsg) Then ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
End Function

Private Sub EndSubclassing()
    If (m_ParentHWnd <> 0) Then
        VBHacks.StopSubclassing m_ParentHWnd, Me
        m_ParentHWnd = 0
    End If
End Sub

Private Sub Class_Terminate()
    EndSubclassing
End Sub
