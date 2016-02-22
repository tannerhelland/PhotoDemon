Attribute VB_Name = "UserControl_Support"
'***************************************************************************
'Helper functions for various PhotoDemon UCs
'Copyright 2014-2016 by Tanner Helland
'Created: 06/February/14
'Last updated: 20/August/15
'Last update: start migrating various UC-inspecific functions here
'
'Many of PD's custom user controls share similar functionality.  Rather than duplicate that functionality across
' all controls, I've tried to collect reusable functions here.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Public Type PD_LISTITEM
    textEn As String
    textTranslated As String
    itemTop As Long
    itemHeight As Long
    isSeparator As Boolean
End Type

'The tab key presents a particular problem when intermixing API windows and VB controls.  VB (obviously) ignores the
' API windows entirely, and PD further complicates this by sometimes mixing API windows and VB controls on the same UC.
' To avoid this disaster, we manage our own tab key presses using an automated system that sorts controls from top-left
' to bottom-right, and automatically figures out tab order from there.
'
'To make sure the automated system works correctly, some controls actually raise a TabPress event, which their parent
' UC can use to cycle control focus within the UC.  When focus is tabbed-out from the last control on the UC, the UC
' itself can then notify the master TabHandler to pass focus to an entirely new control.
Public Enum PDUC_TAB_BEHAVIOR
    TabDefaultBehavior = 0
    TabRaiseEvent = 1
End Enum

#If False Then
    Private Const TabDefaultBehavior = 0, TabRaiseEvent = 1
#End If

'At times, PD may need to post custom messages to all application windows (e.g. theme changes may eventually be implemented
' like this).  Do not call PostMessage directly, as it sends messages to the thread's message queue; instead, call the
' PostPDMessage() function below, which asynchronously relays the request to registered windows via SendNotifyMessage.
Private Declare Function PostMessage Lib "user32" Alias "PostMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendNotifyMessage Lib "user32" Alias "SendNotifyMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Current list of registered windows and the custom messages they want to receive.  This spares us from having to enumerate
' all windows, or worse, blast all windows in the system with our internal messages.  (At present, these are naive lists
' because PD uses so few of them, but in the future, we could look at a hash table.  I've deliberately made the list
' interactions structure-agnostic to simplify future improvements.)
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

'Sometimes controls need unique ID values.  This module will provide a non-zero unique ID via the GetUniqueControlID() function.
Private m_UniqueIDTracker As Long

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
    If sourceControl.Extender.TabStop And sourceControl.Extender.Visible And sourceControl.Enabled Then
        
        'Iterate through all controls in the container, looking for the next TabStop index
        Dim myIndex As Long
        myIndex = sourceControl.Extender.TabIndex
        
        Dim newIndex As Long
        Const MAX_INDEX As Long = 99999
        
        'Forward and back focus checks require different search strategies
        If focusDirectionForward Then
            newIndex = MAX_INDEX
        Else
            newIndex = myIndex
        End If
        
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
                If Ctl.TabStop And Ctl.Visible And Ctl.Enabled Then
                        
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
                    If Ctl.TabStop And Ctl.Visible And Ctl.Enabled Then
                            
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
        If newIndex <> myIndex Then
            targetControl.SetFocus
        
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

    If Not (TypeOf Ctl Is Timer) And Not (TypeOf Ctl Is Line) And Not (TypeOf Ctl Is pdLabel) And Not (TypeOf Ctl Is Frame) And Not (TypeOf Ctl Is Shape) And Not (TypeOf Ctl Is Image) And Not (TypeOf Ctl Is pdAccelerator) And Not (TypeOf Ctl Is ShellPipe) And Not (TypeOf Ctl Is pdDownload) Then
        IsControlFocusable = True
    Else
        IsControlFocusable = False
    End If

End Function

'PD's various user controls sometimes like to share data via custom window messages.  Instead of calling PostMessage directly,
' use this wrapper function, which may perform additional maintenance.
Public Sub PostPDMessage(ByVal wMsg As Long, Optional ByVal wParam As Long = 0&, Optional ByVal lParam As Long = 0&)
    
    Dim pmReturn As Long
    pmReturn = 1
    
    'Enumerate all matching, non-zero windows, and post the requested message, without waiting for a response.
    Dim i As Long
    For i = 0 To m_windowMsgCount - 1
        If (m_wMsgList(i) = wMsg) Then
            If (m_windowList(i) <> 0) Then
                pmReturn = pmReturn And SendNotifyMessage(m_windowList(i), wMsg, wParam, lParam)
            End If
        End If
    Next i
    
    #If DEBUGMODE = 1 Then
        If pmReturn = 0 Then
            pdDebug.LogAction "PostPDMessage was unable to post message ID #" & wMsg & " to one or more windows."
        End If
    #End If
    
End Sub

'Rather than blast all windows with manually raised messages, PD maintains a list of hWnds and registered message requests.
' Add windows and/or messages via this function, and when the messages need to be raised (via PostPDMessage(), above),
' the function will automatically notify all registered recipients.
Public Sub AddMessageRecipient(ByVal targetHwnd As Long, ByVal wMsg As Long)
    
    'Prep the storage structure, as necessary.
    If m_windowMsgCount = 0 Then
        ReDim m_windowList(0 To INITIAL_WINDOW_MESSAGE_LIST_SIZE - 1) As Long
        ReDim m_wMsgList(0 To INITIAL_WINDOW_MESSAGE_LIST_SIZE - 1) As Long
    ElseIf m_windowMsgCount > UBound(m_windowList) Then
        ReDim m_windowList(0 To (UBound(m_windowList) * 2 + 1)) As Long
        ReDim m_wMsgList(0 To (UBound(m_wMsgList) * 2 + 1)) As Long
    End If
    
    m_windowList(m_windowMsgCount) = targetHwnd
    m_wMsgList(m_windowMsgCount) = wMsg
    
    m_windowMsgCount = m_windowMsgCount + 1
    
End Sub

Public Sub RemoveMessageRecipient(ByVal targetHwnd As Long)
    
    'Rather then condensing the list, we simply set all corresponding window entries to zero.
    Dim i As Long
    For i = 0 To m_windowMsgCount - 1
        If m_windowList(i) = targetHwnd Then
            m_windowList(i) = 0
            m_wMsgList(i) = 0
        End If
    Next i
    
End Sub

'Edit boxes can all share the same background brush (as they are all themed identically).  Call this function instead
' of creating your own brush for every text box instance.
Public Function GetSharedGDIBrush(ByVal requestedColor As Long) As Long
    
    'First things first: if this is the first brush requested by the system, initialize the shared brush list
    If m_numOfSharedBrushes = 0 Then
        ReDim m_SharedBrushes(0 To INIT_SIZE_OF_BRUSH_CACHE - 1) As SharedGDIBrush
    End If
    
    'Next, look for the requested color in our current cache.  If it exists, we don't want to recreate it.
    Dim brushExists As Boolean, brushIndex As Long, i As Long
    brushExists = False
    
    If m_numOfSharedBrushes > 0 Then
            
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
        If m_numOfSharedBrushes > UBound(m_SharedBrushes) Then ReDim Preserve m_SharedBrushes(0 To m_numOfSharedBrushes * 2 - 1) As SharedGDIBrush
        
        'Update the cache entry with new stats (including the created brush)
        m_SharedBrushes(m_numOfSharedBrushes).brushColor = requestedColor
        m_SharedBrushes(m_numOfSharedBrushes).numOfOwners = 1
        m_SharedBrushes(m_numOfSharedBrushes).brushHandle = CreateSolidBrush(requestedColor)
        
        'Return the newly created brush handle, increment the brush count, and exit
        GetSharedGDIBrush = m_SharedBrushes(m_numOfSharedBrushes).brushHandle
        m_numOfSharedBrushes = m_numOfSharedBrushes + 1
    
    End If
    
End Function

'Edit boxes can all share the same background brush (as they are all themed identically).  Call this function after
' an edit box is unloaded, so we can free the shared brush accordingly.  (If it's more convenient, you can also use
' the ReleaseSharedGDIBrushByHandle version of this function, below.)
Public Sub ReleaseSharedGDIBrushByColor(ByVal requestedColor As Long)

    'If the cache is empty, ignore this request
    If m_numOfSharedBrushes = 0 Then
        Debug.Print "FYI: UserControl_Support.ReleaseSharedGDIBrush() received a release request, but no shared brushes exist."
        Exit Sub
    
    'If the cache is non-empty, find the matching brush and decrement its count.
    Else
    
        Dim i As Long
        For i = 0 To m_numOfSharedBrushes - 1
            
            If m_SharedBrushes(i).brushColor = requestedColor Then
                m_SharedBrushes(i).numOfOwners = m_SharedBrushes(i).numOfOwners - 1
                
                'Brushes with a count of 0 are immediately killed.
                If m_SharedBrushes(i).numOfOwners = 0 Then
                    DeleteObject m_SharedBrushes(i).brushHandle
                    m_SharedBrushes(i).brushHandle = 0
                    m_SharedBrushes(i).brushColor = 0
                End If
                
                Exit For
            End If
            
        Next i
        
    End If

End Sub

Public Sub ReleaseSharedGDIBrushByHandle(ByVal requestedHandle As Long)

    If m_numOfSharedBrushes = 0 Then
        Debug.Print "FYI: UserControl_Support.ReleaseSharedGDIBrushByHandle() received a release request, but no shared brushes exist."
        Exit Sub
    Else
        Dim i As Long
        For i = 0 To m_numOfSharedBrushes - 1
            If m_SharedBrushes(i).brushHandle = requestedHandle Then
                m_SharedBrushes(i).numOfOwners = m_SharedBrushes(i).numOfOwners - 1
                If m_SharedBrushes(i).numOfOwners = 0 Then
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
    If m_numOfSharedFonts = 0 Then
        ReDim m_SharedFonts(0 To INIT_SIZE_OF_FONT_CACHE - 1) As SharedGDIFont
    End If
    
    'Next, look for the requested size in our current cache.  If it exists, we don't want to recreate it.
    Dim fontExists As Boolean, fontIndex As Long, i As Long
    fontExists = False
    
    If m_numOfSharedFonts > 0 Then
            
        For i = 0 To m_numOfSharedFonts - 1
            If m_SharedFonts(i).FontSize = requestedSize Then
            
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
        If m_numOfSharedFonts > UBound(m_SharedFonts) Then ReDim Preserve m_SharedFonts(0 To m_numOfSharedFonts * 2 - 1) As SharedGDIFont
        
        'Font creation is cumbersome, but PD provides some helper functions to simplify it
        Dim tmpLogFont As LOGFONTW
        Font_Management.FillLogFontW_Basic tmpLogFont, g_InterfaceFont, False, False, False, False
        Font_Management.FillLogFontW_Size tmpLogFont, requestedSize, pdfu_Point
        Font_Management.FillLogFontW_Quality tmpLogFont, TextRenderingHintClearTypeGridFit
        
        'Update the cache entry with new stats (including the created font)
        m_SharedFonts(m_numOfSharedFonts).FontSize = requestedSize
        m_SharedFonts(m_numOfSharedFonts).numOfOwners = 1
        If Not Font_Management.CreateGDIFont(tmpLogFont, m_SharedFonts(m_numOfSharedFonts).fontHandle) Then
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  UserControl_Support.GetSharedGDIFont() failed to create a new UI font handle."
            #End If
        End If
        
        'Return the newly created font handle, increment the font count, and exit
        GetSharedGDIFont = m_SharedFonts(m_numOfSharedFonts).fontHandle
        m_numOfSharedFonts = m_numOfSharedFonts + 1
    
    End If
    
End Function

Public Sub ReleaseSharedGDIFontByHandle(ByVal requestedHandle As Long)

    If m_numOfSharedFonts = 0 Then
        Debug.Print "FYI: UserControl_Support.ReleaseSharedGDIFontByHandle() received a release request, but no shared fonts exist."
        Exit Sub
    Else
        Dim i As Long
        For i = 0 To m_numOfSharedFonts - 1
            If m_SharedFonts(i).fontHandle = requestedHandle Then
                m_SharedFonts(i).numOfOwners = m_SharedFonts(i).numOfOwners - 1
                If m_SharedFonts(i).numOfOwners = 0 Then
                    Font_Management.DeleteGDIFont m_SharedFonts(i).fontHandle
                    m_SharedFonts(i).fontHandle = 0
                    m_SharedFonts(i).FontSize = 0
                End If
                Exit For
            End If
        Next i
    End If

End Sub

'Return a unique, non-zero control ID.  Limited to the size of a VB Long (32-bytes), so don't call more than ~4 billion times.
Public Function GetUniqueControlID() As Long
    
    If m_UniqueIDTracker = LONG_MAX Then
        m_UniqueIDTracker = -1 * LONG_MAX
    Else
        m_UniqueIDTracker = m_UniqueIDTracker + 1
    End If
    
    GetUniqueControlID = m_UniqueIDTracker
    
End Function
