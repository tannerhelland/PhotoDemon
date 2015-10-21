VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Unicode Tooltip class
'Copyright 2014-2015 by Tanner Helland
'Created: 09/January/15
'Last updated: 09/January/15
'Last update: initial build
'
'PD previously used a 3rd-party tooltip class for handling tooltips.  Unfortunately, the class had a lot of bugs
' (including simple mistakes like swapping ByVal and ByRef incorrectly), and it doesn't support Unicode.  As part of
' rewriting all common controls as Unicode-aware, tooltips also needed to be covered - hence this class!
'
'Still, many thanks to the original version of this code, which served PD well for several years.  You can obtain a
' copy of that code (which served as a valuable reference) from the following link (good as of January '15):
' http://www.vb-helper.com/howto_multi_line_tooltip.html
'
'This class operates strictly on hWnds, not VB objects, so please adjust code accordingly when making use of this class.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Tooltips are actually their own window class, so we interact with them using standard window creation/destruction interfaces
Private Const WM_USER As Long = &H400
Private Const TTS_ALWAYSTIP As Long = &H1
Private Const TTS_NOPREFIX As Long = &H2

'As with many structs, there are different TOOLINFO layouts for different OSes.  We limit ours to the XP+ variety.
Private Type TOOLINFO_COMPILED
    cbSize As Long
    uFlags As TT_Flags
    hWnd As Long
    uId As Long
    toolRect As RECTL
    hInst As Long
    lpszTextPointer As Long
    lParam As Long
    lpReserved As Long
End Type

Private Type TOOLINFO_IDE
    cbSize As Long
    uFlags As TT_Flags
    hWnd As Long
    uId As Long
    toolRect As RECTL
    hInst As Long
    lpszTextPointer As Long
    lParam As Long
End Type


'Like other windows, tooltips have many styles available
Private Enum TT_Flags
    TTF_IDISHWND = &H1
    TTF_CENTERTIP = &H2
    TTF_RTLREADING = &H4
    TTF_SUBCLASS = &H10
    TTF_TRACK = &H20
    TTF_ABSOLUTE = &H80
    TTF_TRANSPARENT = &H100
    TTF_DI_SETITEM = &H8000&
    TTF_PARSELINKS = &H1000&
End Enum

#If False Then
    Private Const TTF_IDISHWND = &H1, TTF_CENTERTIP = &H2, TTF_RTLREADING = &H4, TTF_SUBCLASS = &H10, TTF_TRACK = &H20
    Private Const TTF_ABSOLUTE = &H80, TTF_TRANSPARENT = &H100, TTF_DI_SETITEM = &H8000&, TTF_PARSELINKS = &H1000&
#End If

Private Enum TT_DelayTime
    TTDT_AUTOMATIC = 0
    TTDT_RESHOW = 1
    TTDT_AUTOPOP = 2
    TTDT_INITIAL = 3
End Enum

#If False Then
    Private Const TTDT_AUTOMATIC = 0, TTDT_RESHOW = 1, TTDT_AUTOPOP = 2, TTDT_INITIAL = 3
#End If

Private Enum ttDelayTimeConstants
    ttDelayDefault = TTDT_AUTOMATIC '= 0
    ttDelayInitial = TTDT_INITIAL '= 3
    ttDelayShow = TTDT_AUTOPOP '= 2
    ttDelayReshow = TTDT_RESHOW '= 1
    ttDelayMask = 3
End Enum

#If False Then
    Private Const ttDelayDefault = TTDT_AUTOMATIC, ttDelayInitial = TTDT_INITIAL, ttDelayShow = TTDT_AUTOPOP, ttDelayReshow = TTDT_RESHOW, ttDelayMask = 3
#End If

Private Enum ttMarginConstants
    ttMarginLeft = 0
    ttMarginTop = 1
    ttMarginRight = 2
    ttMarginBottom = 3
End Enum

#If False Then
    Private Const ttMarginLeft = 0, ttMarginTop = 1, ttMarginRight = 2, ttMarginBottom = 3
#End If

Private Enum TT_Msgs
    TTM_ACTIVATE = (WM_USER + 1)
    TTM_SETDELAYTIME = (WM_USER + 3)
    TTM_RELAYEVENT = (WM_USER + 7)
    TTM_GETTOOLCOUNT = (WM_USER + 13)
    TTM_WINDOWFROMPOINT = (WM_USER + 16)
    TTM_TRACKACTIVATE = (WM_USER + 17)
    TTM_TRACKPOSITION = (WM_USER + 18)
    TTM_SETTIPBKCOLOR = (WM_USER + 19)
    TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
    TTM_GETDELAYTIME = (WM_USER + 21)
    TTM_GETTIPBKCOLOR = (WM_USER + 22)
    TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
    TTM_SETMAXTIPWIDTH = (WM_USER + 24)
    TTM_GETMAXTIPWIDTH = (WM_USER + 25)
    TTM_SETMARGIN = (WM_USER + 26)
    TTM_GETMARGIN = (WM_USER + 27)
    TTM_POP = (WM_USER + 28)
    TTM_UPDATE = (WM_USER + 29)
    TTM_SETTITLE = (WM_USER + 33)
    TTM_ADDTOOL = (WM_USER + 50)
    TTM_DELTOOL = (WM_USER + 51)
    TTM_NEWTOOLRECT = (WM_USER + 52)
    TTM_GETTOOLINFO = (WM_USER + 53)
    TTM_SETTOOLINFO = (WM_USER + 54)
    TTM_HITTEST = (WM_USER + 55)
    TTM_GETTEXT = (WM_USER + 56)
    TTM_UPDATETIPTEXT = (WM_USER + 57)
    TTM_ENUMTOOLS = (WM_USER + 58)
    TTM_GETCURRENTTOOL = (WM_USER + 59)
End Enum

#If False Then
    Private Const TTM_ACTIVATE = (WM_USER + 1), TTM_SETDELAYTIME = (WM_USER + 3), TTM_RELAYEVENT = (WM_USER + 7), TTM_GETTOOLCOUNT = (WM_USER + 13)
    Private Const TTM_WINDOWFROMPOINT = (WM_USER + 16), TTM_ADDTOOL = (WM_USER + 50), TTM_DELTOOL = (WM_USER + 51), TTM_NEWTOOLRECT = (WM_USER + 52)
    Private Const TTM_GETTOOLINFO = (WM_USER + 53), TTM_SETTOOLINFO = (WM_USER + 54), TTM_HITTEST = (WM_USER + 55), TTM_GETTEXT = (WM_USER + 56)
    Private Const TTM_UPDATETIPTEXT = (WM_USER + 57), TTM_ENUMTOOLS = (WM_USER + 58), TTM_GETCURRENTTOOL = (WM_USER + 59), TTM_TRACKACTIVATE = (WM_USER + 17)
    Private Const TTM_TRACKPOSITION = (WM_USER + 18), TTM_SETTIPBKCOLOR = (WM_USER + 19), TTM_SETTIPTEXTCOLOR = (WM_USER + 20), TTM_GETDELAYTIME = (WM_USER + 21)
    Private Const TTM_GETTIPBKCOLOR = (WM_USER + 22), TTM_GETTIPTEXTCOLOR = (WM_USER + 23), TTM_SETMAXTIPWIDTH = (WM_USER + 24), TTM_GETMAXTIPWIDTH = (WM_USER + 25)
    Private Const TTM_SETMARGIN = (WM_USER + 26), TTM_GETMARGIN = (WM_USER + 27), TTM_POP = (WM_USER + 28), TTM_UPDATE = (WM_USER + 29), TTM_SETTITLE = (WM_USER + 33)
#End If

Private Enum TT_Notifications
    TTN_FIRST = -520&
    TTN_LAST = -549&
    TTN_NEEDTEXT = (TTN_FIRST - 10)
    TTN_SHOW = (TTN_FIRST - 1)
    TTN_POP = (TTN_FIRST - 2)
End Enum

#If False Then
    Private Const TTN_FIRST = -520&, TTN_LAST = -549&, TTN_NEEDTEXT = (TTN_FIRST - 10), TTN_SHOW = (TTN_FIRST - 1), TTN_POP = (TTN_FIRST - 2)
#End If

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

'Added by Tanner: due to complexities with always-on-top windows, we must also make tooltips always-on-top
Private Const WS_EX_TOPMOST As Long = &H8&

'pdToolTip now supports icons
Public Enum TT_ICON_TYPE
    TTI_NONE = 0
    TTI_INFO = 1
    TTI_WARNING = 2
    TTI_ERROR = 3
    TTI_INFO_LARGE = 4
    TTI_WARNING_LARGE = 5
    TTI_ERROR_LARGE = 6
End Enum

#If False Then
    Private Const TTI_NONE = 0, TTI_INFO = 1, TTI_WARNING = 2, TTI_ERROR = 3, TTI_INFO_LARGE = 4, TTI_WARNING_LARGE = 5, TTI_ERROR_LARGE = 6
#End If

'Tooltips are managed using a custom collection
Private Type pdTooltipItem
    childHwnd As Long
    parentHwnd As Long
    toolHwnd As Long
    textStringEn As String
    titleStringEn As String
    textStringTranslated As String
    titleStringTranslated As String
    toolIcon As TT_ICON_TYPE
End Type

Private m_NumOfTooltips As Long
Private m_ToolCollection() As pdTooltipItem

'When creating a tooltip via this class, there are two things to note:
' 1) The parent hWnd (required) owns the tooltip window.  This is relevant for z-order and cleanup, among other details.
' 2) The child hWnd (required) determines the rect that actually raises the tooltip.
'
'Tooltip window destruction will happen automatically when the parent is killed, but this class performs its own manual
' clean-up if terminated first.
'
'Also, to eliminate the need for too many different handlers, this function will accept any parent window for any given tooltip.
' You are not required to use the same parent for each request.
'
'Setting a blank tooltip will result in destruction of the tool window.  (Note that there is no penalty to doing this, as the
' tooltip can easily be recreated by calling this function again with Len(textString) > 0)
Public Function setTooltip(ByVal childWindowHwnd As Long, ByVal parentWindowHwnd As Long, ByVal textString As String, Optional ByVal titleString As String = "", Optional ByVal toolTipIcon As TT_ICON_TYPE = TTI_NONE) As Boolean
    
    'Check for invalid requests
    If (childWindowHwnd = 0) Or (parentWindowHwnd = 0) Then
        Debug.Print "WARNING! setTooltip received a 0 hWnd for either a child or parent.  Don't do that!"
        setTooltip = False
        Exit Function
    End If
    
    Dim toolIndex As Long
    
    'Start by looking for the requested hWnd in the array.  If the hWnd already exists, we can update the existing tool window
    ' instead of creating a new one.
    toolIndex = findTooltipByChildHwnd(childWindowHwnd)
    
    'If the tooltip does not already exist, create it now.  (Note that the tooltip may have been created earlier, but destroyed.
    ' This is also a valid "Does Not Exist" state.)
    Dim toolDoesNotExist As Boolean
    If toolIndex = -1 Then
        toolDoesNotExist = True
    Else
        If m_ToolCollection(toolIndex).toolHwnd = 0 Then toolDoesNotExist = True
    End If
    
    'See if a translation is currently active.  If it is, we will auto-translate both the tooltip text and title automatically.
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
    
    If toolDoesNotExist Then
        
        'Assign a tool index, as necessary
        If toolIndex = -1 Then
            
            toolIndex = m_NumOfTooltips
            
            'Increment the tool counter and resize the array as necessary
            m_NumOfTooltips = m_NumOfTooltips + 1
            If m_NumOfTooltips > UBound(m_ToolCollection) Then ReDim Preserve m_ToolCollection(0 To m_NumOfTooltips * 2 - 1) As pdTooltipItem
            
        End If
        
        'Fill the current tool collection entry with the passed parameters
        With m_ToolCollection(toolIndex)
            .childHwnd = childWindowHwnd
            .parentHwnd = parentWindowHwnd
            .textStringEn = textString
            .titleStringEn = titleString
            
            If isTranslationActive Then
                .textStringTranslated = g_Language.TranslateMessage(.textStringEn)
                .titleStringTranslated = g_Language.TranslateMessage(.titleStringEn)
            Else
                .textStringTranslated = .textStringEn
                .titleStringTranslated = .titleStringEn
            End If
            
            .toolIcon = toolTipIcon
        End With
                
        'Create the tooltip window now
        m_ToolCollection(toolIndex).toolHwnd = CreateWindowEx(0, "tooltips_class32", vbNullString, TTS_ALWAYSTIP Or WS_EX_TOPMOST Or TTS_NOPREFIX, 0, 0, 0, 0, m_ToolCollection(toolIndex).parentHwnd, 0, App.hInstance, ByVal 0&)
        
        'If the window couldn't be created, bail
        If m_ToolCollection(toolIndex).toolHwnd = 0 Then
            Debug.Print "WARNING!  Tooltip window couldn't be created.  Exiting setTooltip."
            setTooltip = False
            Exit Function
        End If
        
        'We are now going to set some default parameters for the tooltip window.  These do not vary across PD.
        
        'Set a uniform tooltip width and delay time
        SendMessage m_ToolCollection(toolIndex).toolHwnd, TTM_SETMAXTIPWIDTH, 0, ByVal PD_MAX_TOOLTIP_WIDTH
        SendMessage m_ToolCollection(toolIndex).toolHwnd, TTM_SETDELAYTIME, (ttDelayShow And ttDelayMask), ByVal 20000&
        
    'The tool does exist, so all we need to do is update the text and/or title parameters
    Else
    
        'Fill the current tool collection entry with the passed parameters
        With m_ToolCollection(toolIndex)
            .textStringEn = textString
            .titleStringEn = titleString
            
            If isTranslationActive Then
                .textStringTranslated = g_Language.TranslateMessage(.textStringEn)
                .titleStringTranslated = g_Language.TranslateMessage(.titleStringEn)
            Else
                .textStringTranslated = .textStringEn
                .titleStringTranslated = .titleStringEn
            End If
            
        End With
    
    End If
    
    'Tooltips are set using the TOOLINFO struct.  Construct one now.
    ' Note that the compiled .exe can take advantage of some unique features (such as automatic parsing of links).
    ' As such, we have to provide two code paths here: one for the IDE, and one for the compiled .exe.
    Dim currentToolInfo_Compiled As TOOLINFO_COMPILED, currentToolInfo_IDE As TOOLINFO_IDE
    
    If g_IsProgramCompiled Then
    
        With currentToolInfo_Compiled
            .cbSize = LenB(currentToolInfo_Compiled)
            .hWnd = m_ToolCollection(toolIndex).childHwnd
            .uId = m_ToolCollection(toolIndex).childHwnd
            .uFlags = TTF_SUBCLASS Or TTF_IDISHWND Or TTF_PARSELINKS
            .lpszTextPointer = StrPtr(m_ToolCollection(toolIndex).textStringTranslated)
        End With
    
    Else
        
        With currentToolInfo_IDE
            .cbSize = LenB(currentToolInfo_IDE)
            .hWnd = m_ToolCollection(toolIndex).childHwnd
            .uId = m_ToolCollection(toolIndex).childHwnd
            .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
            .lpszTextPointer = StrPtr(m_ToolCollection(toolIndex).textStringTranslated)
        End With
        
    End If
        
    'If a title string or icon has been specified, create them now.  (Both are set by the same message.)
    ' IMPORTANT CAVEAT!  Title and icon are ignored if the tooltip itself is blank.  This behavior is set by Windows,
    ' and cannot be changed short of rendering tooltips manually.
    
    'Titles are limited to 100 chars, including null terminator, so ensure that now as a failsafe
    If Len(m_ToolCollection(toolIndex).titleStringTranslated) > 99 Then m_ToolCollection(toolIndex).titleStringTranslated = Left$(m_ToolCollection(toolIndex).titleStringTranslated, 99)
    
    'Null strings are not valid; check for this case, and if found, replace with a blank string.
    ' (MSDN specifically states that a valid pointer *must* be passed; see http://msdn.microsoft.com/en-us/library/windows/desktop/bb760414%28v=vs.85%29.aspx)
    If m_ToolCollection(toolIndex).titleStringTranslated = vbNullString Then m_ToolCollection(toolIndex).titleStringTranslated = ""
    
    'Apply the title and icon settings, if any
    SendMessage m_ToolCollection(toolIndex).toolHwnd, TTM_SETTITLE, m_ToolCollection(toolIndex).toolIcon, ByVal StrPtr(m_ToolCollection(toolIndex).titleStringTranslated)
    
    'We now have two choices: adding the tooltip text anew, or simply updating it.
    Dim toolMsgSuccess As Boolean
    
    'The tooltip does not exist; create it from scratch
    If toolDoesNotExist Then
        
        If g_IsProgramCompiled Then
            toolMsgSuccess = CBool(SendMessage(m_ToolCollection(toolIndex).toolHwnd, TTM_ADDTOOL, 0&, currentToolInfo_Compiled))
        Else
            toolMsgSuccess = CBool(SendMessage(m_ToolCollection(toolIndex).toolHwnd, TTM_ADDTOOL, 0&, currentToolInfo_IDE))
        End If
    
    'The tooltip does exist; simply update it
    Else
        
        If g_IsProgramCompiled Then
            toolMsgSuccess = CBool(SendMessage(m_ToolCollection(toolIndex).toolHwnd, TTM_UPDATETIPTEXT, 0&, currentToolInfo_Compiled))
        Else
            toolMsgSuccess = CBool(SendMessage(m_ToolCollection(toolIndex).toolHwnd, TTM_UPDATETIPTEXT, 0&, currentToolInfo_IDE))
        End If
        
    End If
    
    'We don't have much use for this return value at present, but it doesn't hurt to return it
    setTooltip = (toolMsgSuccess <> 0)
        
End Function

'Look for a tooltip in the current collection.  If found, the tooltip's index will be returned.
' If not found, -1 will be returned.
Private Function findTooltipByChildHwnd(ByVal targetHwnd As Long) As Long
    
    'Check for empty hWnds
    If targetHwnd = 0 Then
        Debug.Print "WARNING! findTooltipByChildHwnd was given 0 as an hWnd.  Don't do that!"
        findTooltipByChildHwnd = -1
        Exit Function
    End If
    
    'Search for the target hWnd in our current collection
    findTooltipByChildHwnd = -1
    
    Dim i As Long
    For i = 0 To m_NumOfTooltips
        If m_ToolCollection(i).childHwnd = targetHwnd Then
            findTooltipByChildHwnd = i
            Exit For
        End If
    Next i
    
End Function

Private Sub Class_Initialize()
    m_NumOfTooltips = 0
    ReDim m_ToolCollection(0) As pdTooltipItem
End Sub

'When this class is terminated, any created tooltip windows are manually destroyed.  This isn't strictly necessary,
' but it doesn't hurt!
Private Sub Class_Terminate()
    
    Dim i As Long
    If m_NumOfTooltips > 0 Then
    
        For i = 0 To m_NumOfTooltips - 1
            If m_ToolCollection(i).toolHwnd <> 0 Then DestroyWindow m_ToolCollection(i).toolHwnd
        Next i
    
    End If
    
End Sub

'If you want to forcibly kill a child's hWnd, use this function to do so.
Public Sub killTooltip(ByVal targetHwnd As Long)

    Dim i As Long
    If m_NumOfTooltips > 0 Then
    
        For i = 0 To m_NumOfTooltips - 1
            If m_ToolCollection(i).childHwnd = targetHwnd Then
                
                With m_ToolCollection(i)
                    
                    DestroyWindow .toolHwnd
                    .toolHwnd = 0
                    .textStringEn = ""
                    .titleStringEn = ""
                    .textStringTranslated = ""
                    .titleStringTranslated = ""
                    .toolIcon = TTI_NONE
                    
                End With
                
                'Note that we leave the child and parent hWnds in the collection.  This shortcuts the process of adding
                ' a tooltip back to that hWnd in the future, if the caller so desires.
                
            End If
        Next i
    
    End If

End Sub

'If the theme or language is changed at run-time, the caller should call this function.  It will refresh all tooltips in
' its collection against the current language and theme, without requiring any special input or parameters.
Public Sub UpdateAgainstCurrentTheme()

    'This function is simple: just iterate through all tooltips, and re-add them to the collection.  This will automatically
    ' trigger re-translation as necessary.
    Dim i As Long
    For i = 0 To m_NumOfTooltips - 1
        
        With m_ToolCollection(i)
            If (.toolHwnd <> 0) Then setTooltip .childHwnd, .parentHwnd, .textStringEn, .titleStringEn, .toolIcon
        End With
        
    Next i

End Sub
