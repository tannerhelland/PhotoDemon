VERSION 5.00
Begin VB.Form FormPluginManager 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Third-party libraries"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10815
   DrawStyle       =   5  'Transparent
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   721
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdListBox lstPlugins 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   9340
   End
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   6375
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   1323
   End
   Begin PhotoDemon.pdButton cmdReset 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      Caption         =   "Reset all library options"
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6135
      Index           =   0
      Left            =   3000
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10821
      Begin PhotoDemon.pdListBoxOD lstOverview 
         Height          =   5175
         Left            =   105
         TabIndex        =   4
         Top             =   840
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   9128
         Caption         =   "library details:"
      End
      Begin PhotoDemon.pdHyperlink hypPluginFolder 
         Height          =   300
         Left            =   1800
         TabIndex        =   5
         Top             =   0
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   529
         FontSize        =   12
         RaiseClickEvent =   -1  'True
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   300
         Index           =   0
         Left            =   120
         Top             =   420
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         Caption         =   "library status:"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPluginStatus 
         Height          =   300
         Index           =   0
         Left            =   1800
         Top             =   420
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   529
         Caption         =   "GOOD"
         FontSize        =   12
         ForeColor       =   47369
         UseCustomForeColor=   -1  'True
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   300
         Index           =   1
         Left            =   120
         Top             =   0
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         Caption         =   "library folder:"
         FontSize        =   12
         ForeColor       =   4210752
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   6135
      Index           =   1
      Left            =   3000
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10821
      Begin PhotoDemon.pdLabel lblAdditionalInfo 
         Height          =   1935
         Left            =   480
         Top             =   3960
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3413
         Caption         =   ""
         Layout          =   1
      End
      Begin PhotoDemon.pdButtonStrip btsDisablePlugin 
         Height          =   1095
         Left            =   360
         TabIndex        =   3
         Top             =   2520
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1931
         Caption         =   "forcibly disable:"
         FontSizeCaption =   11
      End
      Begin PhotoDemon.pdHyperlink hypHomepage 
         Height          =   270
         Left            =   1680
         TabIndex        =   6
         Top             =   600
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   476
         Caption         =   "homepage"
         FontSize        =   11
         ForeColor       =   12611633
         URL             =   "http://freeimage.sourceforge.net/"
      End
      Begin PhotoDemon.pdHyperlink hypLicense 
         Height          =   270
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   476
         Caption         =   "license"
         FontSize        =   11
         ForeColor       =   12611633
         URL             =   "http://freeimage.sourceforge.net/freeimage-license.txt"
      End
      Begin PhotoDemon.pdLabel lblPluginTitle 
         Height          =   285
         Left            =   120
         Top             =   15
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   503
         Caption         =   "%1 summary"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   270
         Index           =   2
         Left            =   360
         Top             =   1560
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   476
         Caption         =   "expected version:"
         FontSize        =   11
         ForeColor       =   4210752
         Layout          =   2
      End
      Begin PhotoDemon.pdLabel lblPluginVersion 
         Height          =   270
         Left            =   2400
         Top             =   2040
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   476
         Caption         =   "XX.XX.XX"
         FontSize        =   11
         ForeColor       =   49152
         UseCustomForeColor=   -1  'True
      End
      Begin PhotoDemon.pdLabel lblPluginExpectedVersion 
         Height          =   270
         Left            =   2400
         Top             =   1560
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   476
         Caption         =   "XX.XX.XX"
         FontSize        =   11
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   270
         Index           =   3
         Left            =   360
         Top             =   2040
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   476
         Caption         =   "version found:"
         FontSize        =   11
         ForeColor       =   4210752
         Layout          =   2
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   270
         Index           =   0
         Left            =   360
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   476
         Caption         =   "homepage:"
         FontSize        =   11
         ForeColor       =   4210752
         Layout          =   2
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   270
         Index           =   1
         Left            =   360
         Top             =   1080
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   476
         Caption         =   "license:"
         FontSize        =   11
         ForeColor       =   4210752
         Layout          =   2
      End
   End
End
Attribute VB_Name = "FormPluginManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Plugin Manager
'Copyright 2012-2026 by Tanner Helland
'Created: 21/December/12
'Last updated: 22/August/23
'Last update: minor tweaks as more libraries are moved to a "download on-demand" implementation
'
'I've considered merging this form with the main Tools > Options dialog, but that dialog
' is already cluttered and I'd prefer that average users don't interact with this dialog at all.
' So this dialog exists as a standalone UI, and it should really be used only if there are problems.
'
'As of April '16, this dialog should never need to be updated against new libraries.  All library
' information is dynamically pulled from the PluginManager module, so simply update that module
' with the new library's information, and this dialog will pull those changes at run-time.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'These arrays will contain the full version strings of our various plugins, and the expected version strings
Private m_LibraryVersion() As String

'If the user presses "cancel", we need to restore the previous enabled/disabled values
Private m_LibraryEnabled() As Boolean

'We need to distinguish between the user clicking on the "disable plugin" button strip, and programmatically
' changing the button strip to reflect the current setting.
Private m_IgnoreButtonStripEvents As Boolean

'Height of each library overview block on the front-page list
Private Const OVERVIEW_ITEM_HEIGHT As Long = 26

'Font object for owner-drawn listbox rendering, plus a few related UI measurements (cached at load time)
Private m_ListBoxFont As pdFont
Private m_ListBoxMaxWidth As Long

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDPM_COLOR_LIST
    [_First] = 0
    PDPM_GoodText = 0
    PDPM_BadText = 1
    [_Last] = 1
    [_Count] = 2
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Private Sub btsDisablePlugin_Click(ByVal buttonIndex As Long)
    If m_IgnoreButtonStripEvents Then Exit Sub
    m_LibraryEnabled(lstPlugins.ListIndex - 1) = (btsDisablePlugin.ListIndex = 0)
    UpdateLibraryLabels
End Sub

Private Sub cmdBarMini_OKClick()
    
    Message "Saving preferences..."
    
    'Hide this form
    Me.Visible = False
    
    'Remember the current container the user is viewing
    UserPrefs.SetPref_Long "Plugins", "Last Plugin Preferences Page", lstPlugins.ListIndex
    
    'Look for any changes to plugin settings
    Dim settingsChanged As Boolean: settingsChanged = False
    
    Dim i As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        If (PluginManager.IsPluginCurrentlyEnabled(i) <> m_LibraryEnabled(i)) Then
            PluginManager.SetPluginEnablement i, m_LibraryEnabled(i)
            PluginManager.SetPluginAllowed i, m_LibraryEnabled(i)
            settingsChanged = True
        End If
    Next i
    
    'If the user has changed any plugin enable/disable settings, a number of things must be refreshed program-wide
    If settingsChanged Then
        PluginManager.InitializePluginManager
        PluginManager.LoadPluginGroup True
        PluginManager.LoadPluginGroup False
        IconsAndCursors.ApplyAllMenuIcons
        IconsAndCursors.ResetMenuIcons
        ImageFormats.GenerateInputFormats
        ImageFormats.GenerateOutputFormats
    End If
    
    'Because plugin data is stored in the core user prefs file, force a write-to-file op now
    UserPrefs.ForceWriteToFile
    
    Message "Plugin options saved."
    
End Sub

'RESET all plugin options
Private Sub cmdReset_Click()

    'Set current container to zero
    UserPrefs.SetPref_Long "Plugins", "Last Plugin Preferences Page", 0
    
    'Enable all plugins if possible
    Dim i As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        PluginManager.SetPluginAllowed i, True
    Next i
    
    'Reload all plugins (which will also refresh all plugin-related settings)
    PluginManager.InitializePluginManager
    PluginManager.LoadPluginGroup True
    PluginManager.LoadPluginGroup False
    
    'Reload the dialog
    LoadAllPluginSettings
    
End Sub

'LOAD the form
Private Sub Form_Load()
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDPM_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDPluginManager", colorCount
    UpdateColorList
    
    'Prepare a custom font object for the owner-drawn overview listbox
    Set m_ListBoxFont = New pdFont
    m_ListBoxFont.SetFontBold False
    m_ListBoxFont.SetFontSize 10
    m_ListBoxFont.CreateFontObject
    m_ListBoxFont.SetTextAlignment vbLeftJustify
    
    'Populate the left-hand list box with all relevant plugins
    lstOverview.ListItemHeight = Interface.FixDPI(OVERVIEW_ITEM_HEIGHT)
    lstOverview.SetAutomaticRedraws False
    
    lstPlugins.Clear
    lstPlugins.AddItem "Overview", 0, True
    
    Dim i As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        lstPlugins.AddItem PluginManager.GetPluginName(i), i + 1
        lstOverview.AddItem PluginManager.GetPluginName(i), i
    Next i
    
    lstOverview.SetAutomaticRedraws True, True
    lstPlugins.ListIndex = 0
    
    'Provide a convenient link to the library folder
    Dim shortPathToLibs As String
    shortPathToLibs = PluginManager.GetPluginPath()
    If (InStr(1, shortPathToLibs, UserPrefs.GetProgramPath(), vbTextCompare) = 1) Then
        shortPathToLibs = "[PhotoDemon folder]\" & Right$(shortPathToLibs, Len(shortPathToLibs) - Len(UserPrefs.GetProgramPath()))
    End If
    
    hypPluginFolder.Caption = shortPathToLibs
    
    'Find the longest plugin name; we need this to know how to layout text in the
    ' Overview panel
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        If (m_ListBoxFont.GetWidthOfString(PluginManager.GetPluginName(i)) > m_ListBoxMaxWidth) Then m_ListBoxMaxWidth = m_ListBoxFont.GetWidthOfString(PluginManager.GetPluginName(i))
    Next i
    
    m_IgnoreButtonStripEvents = True
    btsDisablePlugin.AddItem "no", 0
    btsDisablePlugin.AddItem "yes", 1
    btsDisablePlugin.ListIndex = 0
    m_IgnoreButtonStripEvents = False
    
    'Load all user-editable settings from the preferences file, and populate all plugin information
    LoadAllPluginSettings
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me
    
    'Adjust the positioning of some labels to account for translation widths
    Dim maxPosition As Long
    maxPosition = lblSubheader(0).GetLeft + lblSubheader(0).GetWidth
    If (lblSubheader(1).GetLeft + lblSubheader(1).GetWidth) > maxPosition Then maxPosition = lblSubheader(1).GetLeft + lblSubheader(1).GetWidth
    hypHomepage.SetLeft maxPosition + FixDPI(12)
    hypLicense.SetLeft hypHomepage.GetLeft
    
    maxPosition = lblSubheader(2).GetLeft + lblSubheader(2).GetWidth
    If (lblSubheader(3).GetLeft + lblSubheader(3).GetWidth) > maxPosition Then maxPosition = lblSubheader(3).GetLeft + lblSubheader(3).GetWidth
    lblPluginExpectedVersion.SetLeft maxPosition + FixDPI(12)
    lblPluginVersion.SetLeft lblPluginExpectedVersion.GetLeft
    
    hypHomepage.SetWidth picContainer(1).GetWidth - hypHomepage.GetLeft
    hypLicense.SetWidth picContainer(1).GetWidth - hypLicense.GetLeft
    lblPluginExpectedVersion.SetWidth picContainer(1).GetWidth - lblPluginExpectedVersion.GetLeft
    lblPluginVersion.SetWidth picContainer(1).GetWidth - lblPluginVersion.GetLeft
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'When the dialog is first launched, use this to populate the dialog with any settings the user may have modified
Private Sub LoadAllPluginSettings()
    
    'Remember which plugins the user has enabled or disabled.  (We store these locally, instead of directly invoking
    ' the plugin manager, so that no changes are applied if the user clicks "Cancel".)
    ReDim m_LibraryEnabled(0 To PluginManager.GetNumOfPlugins - 1) As Boolean
    Dim i As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        m_LibraryEnabled(i) = PluginManager.IsPluginCurrentlyEnabled(i)
    Next i
    
    'Now, check version numbers of each plugin.  This is more complicated than it needs to be, on account of
    ' each plugin having its own unique mechanism for version-checking, but I have wrapped these various functions
    ' inside fairly standard wrapper calls.
    CollectAllVersionNumbers
    
    'We now have a collection of version numbers for our various plugins.  Let's use those to populate our
    ' "good/bad" labels for each plugin.
    UpdateLibraryLabels
    
    'Enable the last container the user selected
    lstPlugins.ListIndex = UserPrefs.GetPref_Long("Plugins", "Last Plugin Preferences Page", 0)
    LibraryChanged
    
End Sub

'Assuming version numbers have been successfully retrieved, this function can be called to update the
' green/red plugin label display on the main panel.
Private Sub UpdateLibraryLabels()
    
    Dim pluginStatus As Boolean: pluginStatus = True
    
    Dim i As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        pluginStatus = pluginStatus And CheckLibraryStateUI(i)
    Next i
    
    If pluginStatus Then
        lblPluginStatus(0).ForeColor = m_Colors.RetrieveColor(PDPM_GoodText)
        lblPluginStatus(0).Caption = LCase$(g_Language.TranslateMessage("GOOD"))
    Else
        lblPluginStatus(0).ForeColor = m_Colors.RetrieveColor(PDPM_BadText)
        lblPluginStatus(0).Caption = g_Language.TranslateMessage("problems detected")
    End If
        
End Sub

'Retrieve all relevant plugin version numbers and store them in the m_LibraryVersion() array
Private Sub CollectAllVersionNumbers()
    
    ReDim m_LibraryVersion(0 To PluginManager.GetNumOfPlugins - 1) As String
    
    'Start by querying the plugin file's metadata for version information.  This only works for some plugins,
    ' unfortunately, but we'll manually fill in outliers afterward.
    Dim i As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        If PluginManager.IsPluginCurrentlyInstalled(i) Then
            m_LibraryVersion(i) = PluginManager.GetPluginVersion(i)
        Else
            m_LibraryVersion(i) = vbNullString
        End If
    Next i
    
    'Remove trailing build numbers from certain version strings.
    Dim dotPos As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        If (i <> CCP_ExifTool) And (i <> CCP_libdeflate) And (i <> CCP_libavif) And (i <> CCP_resvg) And (i <> CCP_libjxl) And (i <> CCP_CharLS) Then
            If (LenB(m_LibraryVersion(i)) <> 0) Then
                dotPos = InStrRev(m_LibraryVersion(i), ".", -1, vbBinaryCompare)
                If (dotPos <> 0) Then m_LibraryVersion(i) = Left$(m_LibraryVersion(i), dotPos - 1)
            Else
                m_LibraryVersion(i) = g_Language.TranslateMessage("none")
            End If
        End If
    Next i
    
End Sub

'Given a plugin's availability, expected version, and index on this form, populate the relevant labels associated with it.
' This function will return TRUE if the plugin is in good status, FALSE if it isn't (for any reason)
Private Function CheckLibraryStateUI(ByVal pluginID As PD_PluginCore, Optional ByRef dstStateString As String = vbNullString, Optional ByRef dstStateUIColor As Long = vbBlack) As Boolean
    
    'Is this plugin present on the machine?
    If PluginManager.IsPluginCurrentlyInstalled(pluginID) Then
        
        'Failsafe check for premature accesses during loading
        If Not VBHacks.IsArrayInitialized(m_LibraryEnabled) Then Exit Function
        If (pluginID > UBound(m_LibraryEnabled)) Then Exit Function
        
        'If present, has it been forcibly disabled?  (Note that we use our internal enablement tracker for this,
        ' to reflect any changes the user has just made.)
        If m_LibraryEnabled(pluginID) Then
            
            'If this plugin is present and enabled, does its version match what we expect?
            If Strings.StringsEqual(m_LibraryVersion(pluginID), PluginManager.ExpectedPluginVersion(pluginID), False) Then
                dstStateString = g_Language.TranslateMessage("installed and up to date")
                dstStateUIColor = m_Colors.RetrieveColor(PDPM_GoodText)
                CheckLibraryStateUI = True
                
            'Version mismatch
            Else
                dstStateString = g_Language.TranslateMessage("installed, but version is unexpected")
                dstStateUIColor = m_Colors.RetrieveColor(PDPM_GoodText)
                CheckLibraryStateUI = True
            End If
            
        'Plugin is disabled
        Else
            
            'If this is as simple as an XP compatibility issue (more prevalent now that PD supports
            ' a variety of modern image formats that can't be built in XP-compatible ways),
            ' let the user know
            If (Not OS.IsVistaOrLater) And PluginManager.IsPluginUnavailableOnXP(pluginID) Then
                dstStateString = g_Language.TranslateMessage("incompatible with Windows XP")
            ElseIf (Not OS.IsWin10OrLater) And PluginManager.IsPluginUnavailableOnWin7(pluginID) Then
                dstStateString = g_Language.TranslateMessage("incompatible with Windows 7")
            Else
                dstStateString = g_Language.TranslateMessage("installed, but disabled by user")
            End If
            
            dstStateUIColor = m_Colors.RetrieveColor(PDPM_BadText)
            CheckLibraryStateUI = False
            
        End If
        
    'Plugin is not present on the machine.  For some libraries, this is problematic (e.g. critical libraries like lcms).
    ' For optional libraries (like libavif) this is fine - the plugin will be downloaded if a user performs an action
    ' that requires it.
    Else
        
        dstStateString = g_Language.TranslateMessage("not installed")
        
        'If this plugin doesn't ship with PD, leave it marked as OK
        If PluginManager.IsPluginAvailableOnDemand(pluginID) Then
            dstStateUIColor = m_Colors.RetrieveColor(PDPM_GoodText)
            CheckLibraryStateUI = True
        Else
            dstStateUIColor = m_Colors.RetrieveColor(PDPM_BadText)
            CheckLibraryStateUI = False
        End If
        
    End If
    
End Function

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDPM_GoodText, "PluginOK", RGB(0, 255, 0)
    m_Colors.LoadThemeColor PDPM_BadText, "PluginError", RGB(255, 0, 0)
End Sub

Private Sub hypPluginFolder_Click()
    Dim filePath As String, shellCommand As String
    filePath = PluginManager.GetPluginPath()
    shellCommand = "explorer.exe """ & filePath & """"
    Shell shellCommand, vbNormalFocus
End Sub

Private Sub lstOverview_Click()
    If (lstOverview.ListIndex >= 0) Then lstPlugins.ListIndex = lstOverview.ListIndex + 1
End Sub

Private Sub lstOverview_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)
    
    If (bufferDC = 0) Then Exit Sub
    
    Dim ItemRect As RectF
    CopyMemoryStrict VarPtr(ItemRect), ptrToRectF, 16
    
    Dim xPadding As Long, yPadding As Long
    xPadding = ItemRect.Left + Interface.FixDPI(8)
    yPadding = ItemRect.Top + Interface.FixDPI(4)
    
    'Always paint the plugin name in standard colors
    m_ListBoxFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextReadOnly)
    m_ListBoxFont.AttachToDC bufferDC
    m_ListBoxFont.FastRenderText xPadding, yPadding, PluginManager.GetPluginName(itemIndex)
    
    'Next, retrieve plugin status (and an associated color)
    Dim stateString As String, stateColor As Long
    CheckLibraryStateUI itemIndex, stateString, stateColor
    m_ListBoxFont.SetFontColor stateColor
    m_ListBoxFont.FastRenderText xPadding + m_ListBoxMaxWidth + xPadding, yPadding, stateString
    m_ListBoxFont.ReleaseFromDC
    
End Sub

'When a new plugin is selected, display only the relevant plugin panel
Private Sub lstPlugins_Click()
    LibraryChanged
End Sub

Private Sub LibraryChanged()

    'Display the overview panel
    If (lstPlugins.ListIndex = 0) Then
        picContainer(0).Visible = True
        picContainer(1).Visible = False
    
    'Display the plugin-specific panel, including populating a bunch of run-time text
    Else
        picContainer(0).Visible = False
        picContainer(1).Visible = True
        
        Dim pluginIndex As PD_PluginCore, pluginName As String
        pluginIndex = lstPlugins.ListIndex - 1
        pluginName = PluginManager.GetPluginName(pluginIndex)
        
        lblPluginTitle.Caption = g_Language.TranslateMessage("%1 summary", pluginName)
        lblPluginExpectedVersion.Caption = PluginManager.ExpectedPluginVersion(pluginIndex)
        
        If PluginManager.IsPluginCurrentlyInstalled(pluginIndex) Then
            lblPluginVersion.Caption = m_LibraryVersion(pluginIndex)
            If Strings.StringsEqual(m_LibraryVersion(pluginIndex), PluginManager.ExpectedPluginVersion(pluginIndex), False) Then
                lblPluginVersion.ForeColor = m_Colors.RetrieveColor(PDPM_GoodText)
            Else
                lblPluginVersion.ForeColor = m_Colors.RetrieveColor(PDPM_BadText)
            End If
        Else
            
            'On-demand libraries are not a problem if they simply haven't been installed yet
            If PluginManager.IsPluginAvailableOnDemand(pluginIndex) Then
                lblPluginVersion.Caption = g_Language.TranslateMessage("not installed, but available on-demand")
                lblPluginVersion.ForeColor = m_Colors.RetrieveColor(PDPM_GoodText)
            Else
                lblPluginVersion.Caption = g_Language.TranslateMessage("missing")
                lblPluginVersion.ForeColor = m_Colors.RetrieveColor(PDPM_BadText)
            End If
            
        End If
        
        hypHomepage.Caption = PluginManager.GetPluginHomepage(pluginIndex)
        hypHomepage.URL = PluginManager.GetPluginHomepage(pluginIndex)
        hypLicense.Caption = PluginManager.GetPluginLicenseName(pluginIndex)
        hypLicense.URL = PluginManager.GetPluginLicenseURL(pluginIndex)
        
        'Some plugins support optional explanation text.
        If PluginManager.IsPluginAvailableOnDemand(pluginIndex) And (Not PluginManager.IsPluginCurrentlyInstalled(pluginIndex)) Then
            
            Dim actionTarget As String
            If (pluginIndex = CCP_libavif) Then
                actionTarget = g_Language.TranslateMessage("an AVIF image")
            ElseIf (pluginIndex = CCP_libjxl) Then
                actionTarget = g_Language.TranslateMessage("a JPEG XL image")
            End If
            
            Dim additionalInfo As pdString
            Set additionalInfo = New pdString
            
            additionalInfo.AppendLine g_Language.TranslateMessage("This library does not ship with PhotoDemon.")
            additionalInfo.AppendLineBreak
            additionalInfo.Append g_Language.TranslateMessage("If you interact with %1, PhotoDemon will offer to download and configure this library for you.", actionTarget)
            lblAdditionalInfo.Caption = additionalInfo.ToString()
            
            lblAdditionalInfo.Visible = True
            
        Else
            lblAdditionalInfo.Visible = False
        End If
        
        m_IgnoreButtonStripEvents = True
        If m_LibraryEnabled(pluginIndex) Then btsDisablePlugin.ListIndex = 0 Else btsDisablePlugin.ListIndex = 1
        m_IgnoreButtonStripEvents = False
        
        'Mandatory libraries cannot be disabled
        btsDisablePlugin.Enabled = (Not PluginManager.IsPluginHighPriority(pluginIndex))
        
    End If
    
End Sub
