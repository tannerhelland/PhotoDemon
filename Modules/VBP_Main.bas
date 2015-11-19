Attribute VB_Name = "modMain"
'Note: this file has been modified for use within PhotoDemon.

'This module is required for theming via embedded manifest.  Many thanks to LaVolpe for the automated tool that coincides
' with this fine piece of code.  Download it yourself at: http://www.vbforums.com/showthread.php?t=606736

Option Explicit

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type

'As of September 2015, reordering the list of files in the master VBP has caused unpredictable crashes when PD closes.
' I've spent two days bisecting commits and I can conclusively nail it down to
' https://github.com/tannerhelland/PhotoDemon/commit/293de1ba4f2d5bc3102304d0263af624e93b6093
'
'I eventually solved the problem by manually unloading all global class instances in a specific order, rather than
' leaving it to VB, but during testing, I found it helpful to suppress the default Windows crash dialog.  In case this
' ever proves useful in the future, I'll leave the declaration here.
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private Const SEM_NOOPENFILEERRORBOX = &H8000&

Private hMod As Long

'PhotoDemon starts here.  Main() is necessary as a start point (vs a form) to make sure that theming is implemented
' correctly.  Note that this code is irrelevant within the IDE.
Public Sub Main()

    Dim iccex As InitCommonControlsExStruct
    
    'For descriptions of these constants, visit: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx
    'Const ICC_ANIMATE_CLASS As Long = &H80&
    Const ICC_BAR_CLASSES As Long = &H4&
    'Const ICC_COOL_CLASSES As Long = &H400&
    'Const ICC_DATE_CLASSES As Long = &H100&
    'Const ICC_HOTKEY_CLASS As Long = &H40&
    'Const ICC_INTERNET_CLASSES As Long = &H800&
    'Const ICC_LINK_CLASS As Long = &H8000&
    'Const ICC_LISTVIEW_CLASSES As Long = &H1&
    'Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&
    'Const ICC_PAGESCROLLER_CLASS As Long = &H1000&
    'Const ICC_PROGRESS_CLASS As Long = &H20&
    'Const ICC_TAB_CLASSES As Long = &H8&
    'Const ICC_TREEVIEW_CLASSES As Long = &H2&
    'Const ICC_UPDOWN_CLASS As Long = &H10&
    'Const ICC_USEREX_CLASSES As Long = &H200&
    Const ICC_STANDARD_CLASSES As Long = &H4000&
    Const ICC_WIN95_CLASSES As Long = &HFF&
    'Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all values above

    With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_STANDARD_CLASSES Or ICC_BAR_CLASSES Or ICC_WIN95_CLASSES
    End With
    
    'InitCommonControlsEx requires IEv3 or above, which shouldn't be a problem on any modern system.  But just to be
    ' safe, use On Error Resume Next.
    On Error Resume Next
    
    'The following block of code prevents XP crashes when VB usercontrols are present in a project (as they are in PhotoDemon)
    hMod = LoadLibrary("shell32.dll")
    InitCommonControlsEx iccex
    
    'If an error occurs, attempt to initiate the Win9x version
    If Err Then
        InitCommonControls ' try Win9x version
        Err.Clear
    End If
    
    On Error GoTo 0
    
    'Because Ambient.UserMode does not report IDE behavior properly, we use our own UserMode tracker.  Many thanks to
    ' Kroc of camendesign.com for suggesting this fix.
    g_IsProgramRunning = True
    
    'FormMain can now be loaded.  It will handle the rest of the load process.
    Load FormMain
    
End Sub

Public Sub finalShutdown()
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "finalShutdown() reached.  Forcibly unloading FormMain..."
    #End If
    
    Set FormMain = Nothing
    
    'Release FreeImage (if available)
    If g_FreeImageHandle <> 0 Then
    
        FreeLibrary g_FreeImageHandle
        g_ImageFormats.FreeImageEnabled = False
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "FreeImage released"
        #End If
        
    End If
    
    'Release zLib (if available)
    If g_ZLibEnabled Then
    
        Plugin_zLib_Interface.releaseZLib
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "zLib released"
        #End If
    
    End If
    
    'Release GDIPlus (if applicable)
    If g_ImageFormats.GDIPlusEnabled Then
        
        releaseGDIPlus
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "GDI+ released"
        #End If
    
    End If
    
    g_IsProgramRunning = False
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Manually unloading all remaining public class instances..."
    #End If
    
    Set g_RecentFiles = Nothing
    Set g_RecentMacros = Nothing
    Set g_Themer = Nothing
    Set g_Displays = Nothing
    Set g_CheckerboardPattern = Nothing
    Set g_Zoom = Nothing
    Set g_WindowManager = Nothing
    
    Dim i As Long
    For i = LBound(pdImages) To UBound(pdImages)
        Set pdImages(i) = Nothing
    Next i
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Everything we can physically unload has been forcibly unloaded.  Releasing final library reference..."
    #End If
    
    'If the shell32 library was loaded successfully, once FormMain is closed, we need to unload the library handle.
    If hMod <> 0 Then FreeLibrary hMod
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "All human-written code complete.  Shutting down pdDebug and exiting gracefully."
        pdDebug.TerminateDebugger
        Set pdDebug = Nothing
    #End If
    
    'We have now terminated everything we can physically terminate.
    
    'Suppress any crashes caused by VB herself (which are possible, unfortunately), then let the program go...
    SetErrorMode SEM_NOGPFAULTERRORBOX
    
End Sub
