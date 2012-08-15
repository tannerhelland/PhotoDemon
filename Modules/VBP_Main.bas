Attribute VB_Name = "modMain"
'Note: this file has been modified for use within PhotoDemon.

'This module is required for theming via embedded manifest.  Many thanks to LaVolpe for the automated tool that coincides
' with this fine piece of code.  Download it yourself at: http://www.vbforums.com/showthread.php?t=606736

Private Declare Function LoadLibraryA Lib "kernel32.dll" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type

'PhotoDemon starts here.  Main() is necessary as a start point (vs a form) to make sure that theming is implemented
' correctly.  Note that this code is irrelevant within the IDE.
Public Sub Main()

    Dim iccex As InitCommonControlsExStruct
    
    'For descriptions of these constants, visit: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx
    Const ICC_ANIMATE_CLASS As Long = &H80&
    Const ICC_BAR_CLASSES As Long = &H4&
    Const ICC_COOL_CLASSES As Long = &H400&
    Const ICC_DATE_CLASSES As Long = &H100&
    Const ICC_HOTKEY_CLASS As Long = &H40&
    Const ICC_INTERNET_CLASSES As Long = &H800&
    Const ICC_LINK_CLASS As Long = &H8000&
    Const ICC_LISTVIEW_CLASSES As Long = &H1&
    Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&
    Const ICC_PAGESCROLLER_CLASS As Long = &H1000&
    Const ICC_PROGRESS_CLASS As Long = &H20&
    Const ICC_TAB_CLASSES As Long = &H8&
    Const ICC_TREEVIEW_CLASSES As Long = &H2&
    Const ICC_UPDOWN_CLASS As Long = &H10&
    Const ICC_USEREX_CLASSES As Long = &H200&
    Const ICC_STANDARD_CLASSES As Long = &H4000&
    Const ICC_WIN95_CLASSES As Long = &HFF&
    Const ICC_ALL_CLASSES As Long = &HFDFF& ' combination of all values above

    With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_STANDARD_CLASSES Or ICC_BAR_CLASSES Or ICC_WIN95_CLASSES
    End With
    
    On Error Resume Next ' error? InitCommonControlsEx requires IEv3 or above
    
    Dim hMod As Long
    hMod = LoadLibraryA("shell32.dll") ' patch to prevent XP crashes when VB usercontrols present
    InitCommonControlsEx iccex
    
    If Err Then
        InitCommonControls ' try Win9x version
        Err.Clear
    End If
    
    On Error GoTo 0
    
    Load FormMain
    
    If hMod Then FreeLibrary hMod

End Sub
