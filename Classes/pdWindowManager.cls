VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdWindowManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Window Manager class
'Copyright ©2012-2013 by Tanner Helland
'Created: 02/October/13
'Last updated: 04/October/13
'Last update: added tracking for child windows; got transparency working for toolboxes!
'
'PhotoDemon first began as a single-pane, single-image editor.  About a decade ago, I rewrote it as an MDI project, and in
' 2013, it became time to rewrite its window method once again - this time, to a true multi-window project, separate from
' MDI.  (MDI has a lot of interaction problems, not to mention a host of known bugs that will never be fixed because
' Microsoft is - rightfully - moving away from the MDI model.)
'
'One unfortunate result of this change is that PhotoDemon must now be responsible for much of its own window handling.  In
' particular, child windows must move with the parent window, and all document windows must be managed intelligently.
'
'Thus this class.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Opacity of toolbox windows when a child window is active
Private Const INACTIVE_TOOLBOX_OPACITY As Byte = 170

'Subclassed messages related to window movement
Private Const WM_SYSCOMMAND As Long = &H112
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOVING As Long = &H216
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_NCMOUSEMOVE As Long = &HA0
Private Const WM_SIZING As Long = &H214

'Subclassed messages for detecting when PD is losing focus to another app, or regaining it from another app
Private Const WM_ACTIVATE      As Long = &H6
Private Const WM_ACTIVATEAPP   As Long = &H1C
Private Const WA_INACTIVE      As Long = 0
Private Const WA_ACTIVE        As Long = 1
Private Const WA_CLICKACTIVE   As Long = 2

'System message that show up as additional params when subclassing WM_SYSCOMMAND
Private Const SC_MAXIMIZE As Long = &HF030&
Private Const SC_RESTORE As Long = &HF120&

'Constant for tracking MouseLeave events
Private Const TME_LEAVE = &H2&

'Constants for setting window transparency
Private Const LWA_ALPHA As Long = &H2&
Private Const WS_EX_LAYERED As Long = &H80000
Private Const GWL_EXSTYLE As Long = (-20)
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOREDRAW As Long = &H8
Private Const SWP_NOSENDCHANGING As Long = &H400
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_ASYNCWINDOWPOS As Long = &H4000
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_HIDEWINDOW As Long = &H80

'Most API calls handle window position and movement in terms of a rect-type variable
Private Type winRect
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

'Each window that the manager tracks stores some basic metadata in a variable of this type.
Private Type pdWindowData
    formReference As Form
    hWnd As Long
    hasBeenMadeTransparent As Boolean
    offsetX As Long
    offsetY As Long
    subclassingReleased As Boolean
    typeOfWindow As pdWindowType
    windowName As String
    lastPosition As winRect
    lastWindowState As Long
    lastWindowBorder As Long
End Type

'This class tracks a lot of different window types.  Each gets special treatment, so we track their type as follows:
Public Enum pdWindowType
    MAIN_WINDOW = 0
    TOOLBOX_WINDOW = 1
    IMAGE_WINDOW = 2
End Enum

#If False Then
    Private Const MAIN_WINDOW = 0, TOOLBOX_WINDOW = 1, IMAGE_WINDOW = 2
#End If

'API calls for window positioning and tracking
Private Declare Function ScreenToClient Lib "user32" (ByVal hndWindow As Long, ByRef scrPoint As POINTAPI) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
      
'API calls for changing toolbar window opacity
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal targetHwnd As Long, ByVal TransColor As Long, ByVal newAlpha As Byte, ByVal fFlags As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal targetHwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal targetHwnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Private Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long

'Child window tracking
Private pdWindows() As pdWindowData
Private numOfPDWindows As Long

'Primary window tracking
Private mainWindowHwnd As Long, mainWindowIndex As Long

'When PD loses focus, we automatically revert all top-level toolbox windows to normal style.  When PD regains focus, we restore them as top-most.
' Activation/deactivation messages are received multiple times, but we only want to do activation/deactivation stuff once, so our current state
' is tracked by this variable.
Private isPDDeactivated As Boolean

'We must subclass various window types in order to manage them properly
Private cSubclass As cSelfSubHookCallback

'XML handling (used to save/load window locations) is handled through a specialized class
Private xmlEngine As pdXML
Private windowDataPath As String

'If a window had location data previously stored, this function will retrieve that data and move the window into place.
' If window data cannot be found, the form will be moved to the requested pseudo-location supplied by the second parameter.
Public Function restoreWindowLocation(ByVal windowIndex As Long, ByVal preferredLocationIfNecessary As Long) As Boolean

    'Start by looking for this form's location data in the XML engine.
    
    'If an entry is found, restore the window to that location.
    If xmlEngine.doesTagExist("windowEntry", "id", pdWindows(windowIndex).windowName) Then
    
        'Retrieve this window's location data from the XML file.
        With pdWindows(windowIndex)
        
            'Window rect values
            .lastPosition.x1 = CLng(xmlEngine.getUniqueTag_String("windowLeft", 0, , "windowEntry", "id", .windowName))
            .lastPosition.y1 = CLng(xmlEngine.getUniqueTag_String("windowTop", 0, , "windowEntry", "id", .windowName))
            .lastPosition.x2 = CLng(xmlEngine.getUniqueTag_String("windowRight", 0, , "windowEntry", "id", .windowName))
            .lastPosition.y2 = CLng(xmlEngine.getUniqueTag_String("windowBottom", 0, , "windowEntry", "id", .windowName))
            
            'Window and border state
            .lastWindowState = CLng(xmlEngine.getUniqueTag_String("windowState", 0, , "windowEntry", "id", .windowName))
            .lastWindowBorder = CLng(xmlEngine.getUniqueTag_String("windowStyle", 0, , "windowEntry", "id", .windowName))
            
            'Apply those values to the window, starting with border style, then state, then coordinates
            .formReference.BorderStyle = .lastWindowBorder
            .formReference.WindowState = .lastWindowState
            
            MoveWindow .hWnd, .lastPosition.x1, .lastPosition.y1, .lastPosition.x2 - .lastPosition.x1, .lastPosition.y2 - .lastPosition.y1, 1
            
        End With
        
    'Window data was not found.  Move the window to a new location using the preferred location param.
    Else
    
        requestIdealPosition pdWindows(windowIndex).hWnd, preferredLocationIfNecessary
    
    End If

End Function

'Load previous window locations from file.
Public Function loadAllWindowLocations() As Boolean
    
    If FileExist(windowDataPath) Then
            
        'Attempt to load and validate the relevant preset file; if we can't, create a new, blank XML object
        If (Not xmlEngine.loadXMLFile(windowDataPath)) Or Not (xmlEngine.isPDDataType("Window locations")) Then
            Debug.Print "No window location data found.  A new window location file has been created."
            resetXMLData
        End If
            
    Else
        resetXMLData
    End If
    
    'We don't actually load window locations now.  Now that the XML data is happily inside our XML engine, we load window data from it
    ' on-demand as windows are added to the window manager.

End Function

'Write the current locations of all windows to the XML engine.  (These will be used to restore the window location on subsequent loads.)
Public Function saveAllWindowLocations() As Boolean
    
    Dim i As Long
    For i = 0 To numOfPDWindows - 1
    
        'We only save locations for the main window and toolbox windows - not image windows!
        If (pdWindows(i).typeOfWindow = MAIN_WINDOW) Or (pdWindows(i).typeOfWindow = TOOLBOX_WINDOW) Then
    
            'Check for this window preset in the file.  If it does not exist, add it now.
            If Not xmlEngine.doesTagExist("windowEntry", "id", pdWindows(i).windowName) Then
            
                xmlEngine.writeTagWithAttribute "windowEntry", "id", pdWindows(i).windowName, "", True
                xmlEngine.closeTag "windowEntry"
                xmlEngine.writeBlankLine
            
            End If
            
            'Write this window's location data into the XML file.
            With pdWindows(i)
            
                'Window rect values
                xmlEngine.updateTag "windowLeft", .lastPosition.x1, "windowEntry", "id", .windowName
                xmlEngine.updateTag "windowTop", .lastPosition.y1, "windowEntry", "id", .windowName
                xmlEngine.updateTag "windowRight", .lastPosition.x2, "windowEntry", "id", .windowName
                xmlEngine.updateTag "windowBottom", .lastPosition.y2, "windowEntry", "id", .windowName
                
                'Window and border state
                xmlEngine.updateTag "windowState", .lastWindowState, "windowEntry", "id", .windowName
                xmlEngine.updateTag "windowStyle", .lastWindowBorder, "windowEntry", "id", .windowName
                
            End With
        
        End If
    
    Next i
    
    'Write the data out to file
    xmlEngine.writeXMLToFile windowDataPath

End Function

'Reset the XML engine.  Note that the XML object SHOULD ALREADY BE INSTANTIATED before calling this function.
Private Function resetXMLData()

    xmlEngine.prepareNewXML "Window locations"
    xmlEngine.writeBlankLine
    xmlEngine.writeComment "Everything past this point is window location data for various PhotoDemon dialogs."
    xmlEngine.writeBlankLine
    
End Function

'This class provides some helper functions to intelligently position child windows, typically when the program is run for the
' first time and custom positions have not yet been established.  Constants are used to specify where the window should be placed.
Public Sub requestIdealPosition(ByVal childHwnd As Long, Optional ByVal desiredPosition As Long = 1)

    'Before doing anything else, get the current size and location of child and parent windows
    Dim childRect As winRect, parentRect As winRect, parentClientRect As winRect
    GetWindowRect childHwnd, childRect
    GetWindowRect mainWindowHwnd, parentRect
    GetClientRect mainWindowHwnd, parentClientRect

    'By comparing the parent window's client area and actual area, we can determine how much is "chrome", e.g. window borders, etc.
    Dim parentChromeWidth As Long, parentChromeHeight As Long
    parentChromeWidth = (parentRect.x2 - parentRect.x1) - parentClientRect.x2
    parentChromeHeight = (parentRect.y2 - parentRect.y1) - parentClientRect.y2
    
    'Right now, we only support two "desired positions": top-left and top-right corners.  I may add more after additional testing.
    Select Case desiredPosition
    
        'Do nothing
        Case 0
    
        'Top-left
        Case 1
            MoveWindow childHwnd, parentRect.x1 + parentChromeWidth \ 2, parentRect.y1 + parentChromeHeight, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
            
        'Top-right
        Case 2
            MoveWindow childHwnd, parentRect.x2 - parentChromeWidth \ 2 - (childRect.x2 - childRect.x1), parentRect.y1 + parentChromeHeight, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
    
    End Select

End Sub

'The window manager keeps a persistent list of last-known locations for every window it accesses.  Window locations are updated when:
' - They're first loaded (and no previous position has been stored)
' - They're moved
' - They're closed
'Any of those actions just need to send a request to this sub, which will handle the actual location storage process.
Private Sub updateStoredWindowLocation(ByVal indexInWindowArray As Long)

    Dim newWindowRect As winRect
    
    With pdWindows(indexInWindowArray)
    
        GetWindowRect .hWnd, newWindowRect
        .lastPosition = newWindowRect
        .lastWindowBorder = .formReference.BorderStyle
        .lastWindowState = .formReference.WindowState
    
    End With

End Sub

'The primary PhotoDemon form must register its hWnd, so we can track its movement and move any children windows accordingly.
Public Sub registerParentForm(ByRef parentForm As Form)

    'Has this window already been registered?  If so, don't re-register it.
    Dim i As Long
    For i = 0 To numOfPDWindows
        If pdWindows(i).hWnd = parentForm.hWnd Then Exit Sub
    Next i
    
    'Add this window to the collection, and cache its hWnd (because we reference the main window handle frequently)
    Set pdWindows(numOfPDWindows).formReference = parentForm
    pdWindows(numOfPDWindows).hWnd = parentForm.hWnd
    mainWindowHwnd = parentForm.hWnd
    mainWindowIndex = numOfPDWindows
    pdWindows(numOfPDWindows).typeOfWindow = MAIN_WINDOW
    pdWindows(numOfPDWindows).windowName = parentForm.Name
    
    'Resize the window array for future additions
    numOfPDWindows = numOfPDWindows + 1
    ReDim Preserve pdWindows(0 To numOfPDWindows) As pdWindowData

    'Subclass the main window
    cSubclass.ssc_Subclass mainWindowHwnd, , , Me
    cSubclass.ssc_AddMsg mainWindowHwnd, MSG_AFTER, WM_MOVING, WM_SYSCOMMAND, WM_MOUSEMOVE, WM_NCMOUSEMOVE, WM_SIZING
    
    'In the future, it may be helpful to know when PD loses focus.  I am leaving this code here "just in case"
    'cSubclass.ssc_AddMsg parentHwnd, MSG_AFTER, WM_ACTIVATEAPP, WM_ACTIVATE
    
    'Look for previous location data in the window location file.  If said data exists, load it and move the window to that location.
    restoreWindowLocation numOfPDWindows - 1, 0
    
End Sub

'Each child window must be registered through this class, using its hWnd.  This class will then use that hWnd to do things
' like move child windows proportionally to their parent.
Public Sub registerChildForm(ByRef childForm As Form, ByVal whatWindowType As pdWindowType, Optional ByVal preferredLocationIfNecessary As Long = 1)

    'Has this window already been registered?  If so, don't re-register it.
    Dim i As Long
    For i = 0 To numOfPDWindows
        If pdWindows(i).hWnd = childForm.hWnd Then Exit Sub
    Next i
    
    'Add this window to the collection
    Set pdWindows(numOfPDWindows).formReference = childForm
    pdWindows(numOfPDWindows).hWnd = childForm.hWnd
    
    'Calculate the initial offset between this window and the main window, and store that offset
    Dim childRect As winRect, parentRect As winRect
    GetWindowRect pdWindows(numOfPDWindows).hWnd, childRect
    GetWindowRect mainWindowHwnd, parentRect
    
    With pdWindows(numOfPDWindows)
        .offsetX = childRect.x1 - parentRect.x1
        .offsetY = childRect.y1 - parentRect.y1
        .hasBeenMadeTransparent = False
        .typeOfWindow = whatWindowType
        .windowName = childForm.Name
    End With
    
    'Resize the window array for future additions
    numOfPDWindows = numOfPDWindows + 1
    ReDim Preserve pdWindows(0 To numOfPDWindows) As pdWindowData
    
    'Subclass this window, and add a user param of 64 + "index in array".  We can use this value to recognize this hWnd as belonging
    ' to a child window, and quickly look it up in our window tracking array.
    cSubclass.ssc_Subclass pdWindows(numOfPDWindows - 1).hWnd, 64 + (numOfPDWindows - 1), , Me
    cSubclass.ssc_AddMsg pdWindows(numOfPDWindows - 1).hWnd, MSG_AFTER, WM_MOVING, WM_MOUSEMOVE, WM_NCMOUSEMOVE
    
    'Move the window into place using either last-known location data, or the supplied "preferred location"
    restoreWindowLocation numOfPDWindows - 1, preferredLocationIfNecessary
    
    'Finally, update this window's offset (because it may have been moved by the above function)
    With pdWindows(numOfPDWindows - 1)
        GetWindowRect .hWnd, childRect
        .offsetX = childRect.x1 - parentRect.x1
        .offsetY = childRect.y1 - parentRect.y1
    End With
    
End Sub

'When a window is unloaded, call this function so that we can stop subclassing in a safe and predictable way, and also track the
' last-known location of this window.
Public Sub unregisterForm(ByRef srcForm As Form)

    'Find this window in the array.  (Note that if it cannot be found, nothing will happen.)
    Dim i As Long
    For i = 0 To numOfPDWindows
        If pdWindows(i).hWnd = srcForm.hWnd Then
            
            'Update the last-known window location before the form closes
            updateStoredWindowLocation i
            
            'Release this window's subclassing
            pdWindows(i).subclassingReleased = True
            cSubclass.ssc_UnSubclass pdWindows(i).hWnd
            pdWindows(i).hWnd = 0
            
            'Remove our reference to this form (if we don't, it will remain in memory)
            Set pdWindows(i).formReference = Nothing
            
        End If
    Next i

End Sub

'When all child windows have been unloaded, we need to reassign toolboxes as children of the main form.  (Otherwise, they will get unloaded!)
Public Sub allImageWindowsUnloaded()

    Dim i As Long
    For i = 0 To numOfPDWindows - 1
        If pdWindows(i).typeOfWindow = TOOLBOX_WINDOW Then
            SetWindowLong pdWindows(i).hWnd, -8, mainWindowHwnd
            SetWindowPos pdWindows(i).hWnd, -2, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOSENDCHANGING Or SWP_SHOWWINDOW
        End If
    Next i
    
End Sub

'When an image window receives focus, we dynamically reassign the toolboxes as children of that window, so that they stay on top of it.
Public Sub notifyChildReceivedFocus(ByRef childForm As Form)

    'Make all toolboxes a child of this form - this keeps them on top!  Clever, but not necessarily advisable...
    Dim i As Long
    For i = 0 To numOfPDWindows - 1
        If pdWindows(i).typeOfWindow = TOOLBOX_WINDOW Then
            SetWindowLong pdWindows(i).hWnd, -8, childForm.hWnd
            SetWindowPos pdWindows(i).hWnd, -2, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOSENDCHANGING Or SWP_SHOWWINDOW
        End If
    Next i
    
    'Redraw the child as well
    SetWindowPos childForm.hWnd, -2, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOSENDCHANGING Or SWP_SHOWWINDOW

End Sub

Private Sub Class_Initialize()

    'Reset all toolbar tracking variables
    numOfPDWindows = 0
    ReDim pdWindows(0) As pdWindowData
        
    'Prepare the subclasser
    Set cSubclass = New cSelfSubHookCallback
    
    'Prepare the XML handler, and retrieve window location data from file (if it exists)
    Set xmlEngine = New pdXML
    windowDataPath = g_UserPreferences.getPresetPath & "Program_WindowLocations.xml"
    loadAllWindowLocations

End Sub

Private Sub Class_Terminate()

    'Release all subclassing
    cSubclass.ssc_Terminate
    
    'Write window location data out to file.
    saveAllWindowLocations

End Sub

'Set a toolbar window's translucency.  We do this to make them semi-transparent when they do not have the mouse over them.
Private Sub SetWindowTranslucency(ByVal srcHwnd As Long, Optional ByVal nTransparency As Long = 255)

    'SetWindowLong and SetLayeredWindowAttributes are API functions, see MSDN for details
    Dim tmpAttributes As Long
    tmpAttributes = GetWindowLong(srcHwnd, GWL_EXSTYLE)
    
    If nTransparency = 255 Then
        SetLayeredWindowAttributes srcHwnd, 0, nTransparency, LWA_ALPHA
        SetWindowLong srcHwnd, GWL_EXSTYLE, tmpAttributes And Not WS_EX_LAYERED
    Else
        SetWindowLong srcHwnd, GWL_EXSTYLE, tmpAttributes Or WS_EX_LAYERED
        SetLayeredWindowAttributes srcHwnd, 0, nTransparency, LWA_ALPHA
    End If
       
End Sub

'All events subclassed by this window are processed here.
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
    Dim curWinRect As winRect
    Dim curWinPoint As POINTAPI
    
    Dim i As Long
    
    'Before doing anything else, check the user-supplied param.  If it is >= 64, we know this is a child (toolbar or image) window.
    If lParamUser >= 64 Then
    
        'Retrieve the window's index in our master tracking array
        Dim winIndex As Long
        winIndex = lParamUser - 64
    
        'Check child window messages
        Select Case uMsg
        
            'The window is currently being moved
            Case WM_MOVING
                
                'lParam contains a pointer to this window's rect location.  Retrieve it now.
                Call CopyMemory(curWinRect, ByVal lParam, Len(curWinRect))
                
                'Calculate this window's rectangle as an offset of its parent.
                Dim parentRect As winRect
                GetWindowRect mainWindowHwnd, parentRect
                
                'Store the new offset values
                pdWindows(winIndex).offsetX = curWinRect.x1 - parentRect.x1
                pdWindows(winIndex).offsetY = curWinRect.y1 - parentRect.y1
                
            'The mouse is moving inside this child window.
            Case WM_MOUSEMOVE, WM_NCMOUSEMOVE
            
                Select Case pdWindows(winIndex).typeOfWindow
                
                    'If this is a toolbox window, and it was previously made transparent, restore its opacity now.
                    Case TOOLBOX_WINDOW
                        If pdWindows(winIndex).hasBeenMadeTransparent Then
                            pdWindows(winIndex).hasBeenMadeTransparent = False
                            SetWindowTranslucency lng_hWnd, 255
                        End If
                    
                    'If this is an image window, and we have non-transparent toolbox windows, change that now.
                    Case IMAGE_WINDOW
                        For i = 0 To numOfPDWindows - 1
                            If (pdWindows(i).typeOfWindow = TOOLBOX_WINDOW) And (Not pdWindows(i).hasBeenMadeTransparent) Then
                                pdWindows(i).hasBeenMadeTransparent = True
                                SetWindowTranslucency pdWindows(i).hWnd, INACTIVE_TOOLBOX_OPACITY
                            End If
                        Next i
                
                End Select
                
        End Select
    
    'The user param is not >= 64, meaning it references the primary window (FormMain).  Move all child windows to match the
    ' parent form's new position.
    Else
    
        Dim childRect As winRect
        
        'Check parent window messages
        Select Case uMsg
        
            'The parent window is currently being moved or resized
            Case WM_MOVING, WM_SIZING
                
                'lParam contains a pointer to the window's rect location.  Retrieve it now.
                Call CopyMemory(curWinRect, ByVal lParam, Len(curWinRect))
                
                'Move each child window, while retaining the current offset
                For i = 0 To numOfPDWindows - 1
                    GetWindowRect pdWindows(i).hWnd, childRect
                    MoveWindow pdWindows(i).hWnd, curWinRect.x1 + pdWindows(i).offsetX, curWinRect.y1 + pdWindows(i).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                Next i
                
                'Remember the main window's location
                updateStoredWindowLocation mainWindowIndex
            
            'The parent window has just been maximized
            Case WM_SYSCOMMAND
            
                If (wParam = SC_MAXIMIZE) Or (wParam = SC_RESTORE) Then
                
                    'Manually retrieve the new window position.
                    GetWindowRect lng_hWnd, curWinRect
                    
                    'Move each child window, while retaining the current offset
                    For i = 0 To numOfPDWindows - 1
                        GetWindowRect pdWindows(i).hWnd, childRect
                        MoveWindow pdWindows(i).hWnd, curWinRect.x1 + pdWindows(i).offsetX, curWinRect.y1 + pdWindows(i).offsetY, childRect.x2 - childRect.x1, childRect.y2 - childRect.y1, 1
                    Next i
                    
                End If
                
            'Mouse has been moved.  If toolboxes are floating, make them translucent.
            Case WM_MOUSEMOVE, WM_NCMOUSEMOVE
                For i = 0 To numOfPDWindows - 1
                    If (pdWindows(i).typeOfWindow = TOOLBOX_WINDOW) And (Not pdWindows(i).hasBeenMadeTransparent) Then
                        pdWindows(i).hasBeenMadeTransparent = True
                        SetWindowTranslucency pdWindows(i).hWnd, INACTIVE_TOOLBOX_OPACITY
                    End If
                Next i
                
            'The user is switching to or from PD and another application
            Case WM_ACTIVATEAPP
            
                'wParam contains a notification of whether we are gaining or losing focus.
                Select Case wParam
                
                    'PD is losing focus
                    Case WA_INACTIVE
                        
                    'PD is gaining focus
                    Case Else
                        
                End Select
                        
        End Select
    
    End If




' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub


