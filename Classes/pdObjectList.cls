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
'Copyright 2017-2017 by Tanner Helland
'Created: 18/August/17
'Last updated: 24/August/17
'Last update: finish work on automatic tab-order sorting
'
'This class is exclusively used by the NavKey module.  Look there for implementation details.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'If you want to dump tab order to the debug window, set this to TRUE.
Private Const DISPLAY_DEBUG_TABORDER_DATA As Boolean = False

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'The hWnd of the parent (the form hosting this collection)
Private m_ParentHWnd As Long

'hWnds of child controls
Private Const INIT_NUM_OF_CONTROLS As Long = 16

'As part of subclassing the parent control, we want to detect ENTER and ESC keypresses.  Dialogs with
' command bars will automatically assign those keypresses to OK/Cancel.
Private Const WM_KEYDOWN As Long = &H100

'We may not need to use a custom type for this - an hWnd may ultimately be enough - but during debugging,
' it's nice to store some extra bits of information
Private Type PD_Control
    hWnd As Long
    ctlName As String
    ctlType As PD_ControlType
End Type

Private m_Controls() As PD_Control
Private m_numOfControls As Long

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

Friend Sub SetParentHWnd(ByVal hWnd As Long)
    
    If (MainModule.IsProgramRunning() And (hWnd <> 0)) Then
    
        m_ParentHWnd = hWnd
        
        'We also subclass the parent control so we can capture ENTER and ESCAPE keypresses (to facilitate
        ' keyboard navigation)
        VBHacks.StartSubclassing hWnd, Me
        
    End If
    
End Sub

Friend Function GetParentHWnd() As Long
    GetParentHWnd = m_ParentHWnd
End Function

Friend Function DoesHWndExist(ByVal srcHwnd As Long) As Boolean
    
    DoesHWndExist = False
    
    If (m_numOfControls > 0) Then
        Dim i As Long
        For i = 0 To m_numOfControls - 1
            If (m_Controls(i).hWnd = srcHwnd) Then
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

'Load individual controls using this function; we'll retrieve whatever we need from them
Friend Sub NotifyChildControl(ByRef childControl As Object)

    'Make sure we have room for this hWnd
    If (m_numOfControls = 0) Then
        ReDim m_Controls(0 To INIT_NUM_OF_CONTROLS - 1) As PD_Control
    Else
        If (m_numOfControls > UBound(m_Controls)) Then ReDim Preserve m_Controls(0 To m_numOfControls * 2 - 1) As PD_Control
    End If
    
    With m_Controls(m_numOfControls)
        .hWnd = childControl.hWnd
        .ctlName = childControl.GetControlName()
        .ctlType = childControl.GetControlType()
    End With
    
    m_numOfControls = m_numOfControls + 1
    
End Sub

'Tab keypresses only require an incoming hWnd; we'll figure out the rest.
Friend Sub NotifyTabKey(ByVal srcHwnd As Long, ByVal shiftTabPressed As Boolean)
    
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
            
            'The command bar (and mini command bar) are a special case.  They are managed by this class,
            ' because we need to detect them for OK/Cancel purposes, but we don't want them receiving
            ' focus during KeyPress events.  (Instead, focus should go to one of their child controls.)
            If (m_Controls(i).ctlType <> pdct_CommandBar) And (m_Controls(i).ctlType <> pdct_CommandBarMini) Then
            
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
                If (srcHwnd = m_Controls(i).hWnd) Then currentControlIndex = m_numOfControlsSort
                
                m_numOfControlsSort = m_numOfControlsSort + 1
                
            End If
            
        End If
    Next i
    
    'We now have a list of all valid tab recipients on this form.  Hopefully our source control was included;
    ' if it wasn't (for whatever reason), retrieve its last-known position and use that instead.
    If (currentControlIndex = -1) Then
        
        g_WindowManager.GetWindowRect_API srcHwnd, tmpRect
        
        With m_ControlsSort(m_numOfControlsSort)
            .hWnd = srcHwnd
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
        If (m_ControlsSort(i).hWnd = srcHwnd) Then
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
    
    'Finally, apply the focus change!
    g_WindowManager.SetFocusAPI m_ControlsSort(targetIndex).hWnd

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
        
    'If this control is taller than 80 pixels, it is probably a very large element like a preview box
    ' or a custom-built control (like the Curves dialog).  Sort it by its top-left corner only.
    Else
    
        sortX = srcRect.x1
        sortY = srcRect.y1
    
    End If
    
    'When sorting on-screen elements, we clamp y-coordinates to their nearest multiple of 8.  This helps
    ' address the case where nearby elements vary in top-position by only a few pixels, as is necessary
    ' to maintain an appearance of center-line alignment (e.g. again refer to command bars, where the
    ' "preset names" dropdown sits just slightly below the neighboring buttons, because it is slightly
    ' shorter vertically than the button images).  To ensure such controls are still sorted in normal
    ' LTR order, instead of being treated as "on the next line", we clamp their controls to a fixed
    ' multiple of 8.
    Const MAX_VERTICAL_DISTANCE As Long = 8
    GetSortKeyFromRect = sortX + (Int((sortY + (MAX_VERTICAL_DISTANCE \ 2)) \ MAX_VERTICAL_DISTANCE) * MAX_VERTICAL_DISTANCE * screenWidth)
    
End Function

Private Function GetControlNameFromHWnd(ByVal srcHwnd As Long) As String

    Dim i As Long
    For i = 0 To m_numOfControls - 1
        If (srcHwnd = m_Controls(i).hWnd) Then
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

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    
    Dim eatMsg As Boolean: eatMsg = False
    
    If (uiMsg = WM_KEYDOWN) Then
    
        'The only keypresses we currently care about are ENTER and ESCAPE
        If (wParam = pdnk_Enter) Or (wParam = pdnk_Escape) Then
        
            'See if this form 1) is a raised dialog, and 2) contains a command bar
            If Interface.IsModalDialogActive() Then
            
                If Me.DoesTypeOfControlExist(pdct_CommandBar) Then
                
                    'It does!  Grab the hWnd and forward the relevant window message to it
                    SendNotifyMessage Me.GetFirstHWndForType(pdct_CommandBar), WM_PD_DIALOG_NAVKEY, wParam, 0&
                    ISubclass_WindowMsg = 0&
                    eatMsg = True
                
                'If a command bar doesn't exist, look for a "mini command bar" instead
                ElseIf Me.DoesTypeOfControlExist(pdct_CommandBarMini) Then
                    SendNotifyMessage Me.GetFirstHWndForType(pdct_CommandBarMini), WM_PD_DIALOG_NAVKEY, wParam, 0&
                    ISubclass_WindowMsg = 0&
                    eatMsg = True
                    
                'No command bar exists on this form, which is fine - this could be a toolpanel, for example.
                ' As such, there's nothing we need to do.
                End If
            
            End If
            
        End If
        
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

