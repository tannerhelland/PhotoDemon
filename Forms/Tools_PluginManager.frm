VERSION 5.00
Begin VB.Form FormPluginManager 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Plugin Manager"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
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
      TabIndex        =   5
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
      Caption         =   "Reset all plugin options"
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   5895
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   7695
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButtonStrip btsDisablePlugin 
         Height          =   1095
         Left            =   360
         TabIndex        =   4
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
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   5895
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   7695
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   300
         Index           =   0
         Left            =   120
         Top             =   15
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   529
         Caption         =   "current plugin status:"
         FontSize        =   12
         ForeColor       =   4210752
         Layout          =   2
      End
      Begin PhotoDemon.pdLabel lblPluginStatus 
         Height          =   285
         Left            =   2460
         Top             =   15
         Width           =   690
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "GOOD"
         FontSize        =   12
         ForeColor       =   47369
         Layout          =   2
         UseCustomForeColor=   -1  'True
      End
      Begin PhotoDemon.pdLabel lblInterfaceTitle 
         Height          =   285
         Index           =   0
         Left            =   240
         Top             =   720
         Width           =   705
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "status:"
         FontSize        =   11
         ForeColor       =   4210752
         Layout          =   2
      End
      Begin PhotoDemon.pdLabel lblStatus 
         Height          =   285
         Index           =   0
         Left            =   1080
         Top             =   720
         Width           =   3540
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "installed, enabled, and up to date"
         FontSize        =   11
         ForeColor       =   49152
         UseCustomForeColor=   -1  'True
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
'Copyright 2012-2017 by Tanner Helland
'Created: 21/December/12
'Last updated: 20/April/16
'Last update: overhaul the dialog so that it never needs to be updated against new plugins.  (Instead, all settings
'             are dynamically pulled from the PluginManager module, and a matching UI is generated at run-time.)
'
'I've considered merging this form with the main Tools > Options dialog, but that dialog is already cluttered,
' and I really prefer that users don't mess around with plugin settings.  So this dialog exists as a standalone UI,
' and it should really be used only if there are problems.
'
'As of April '16, this dialog should never need to be updated against new plugins.  All plugin information is
' dynamically pulled from the PluginManager module, so simply update that module with a new plugin's information,
' and this dialog will pull the changes at run-time.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'These arrays will contain the full version strings of our various plugins, and the expected version strings
Private m_PluginVersion() As String

'If the user presses "cancel", we need to restore the previous enabled/disabled values
Private m_PluginEnabled() As Boolean

'We need to distinguish between the user clicking on the "disable plugin" button strip, and programmatically
' changing the button strip to reflect the current setting.
Private m_IgnoreButtonStripEvents As Boolean

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
    m_PluginEnabled(lstPlugins.ListIndex - 1) = CBool(btsDisablePlugin.ListIndex = 0)
    UpdatePluginLabels
End Sub

'Upon cancellation, we don't propagate any local changes back to the plugin manager
Private Sub cmdBarMini_CancelClick()
    
End Sub

Private Sub cmdBarMini_OKClick()
    
    Message "Saving plugin options..."
    
    'Hide this form
    Me.Visible = False
    
    'Remember the current container the user is viewing
    g_UserPreferences.StartBatchPreferenceMode
    g_UserPreferences.SetPref_Long "Plugins", "Last Plugin Preferences Page", lstPlugins.ListIndex
    
    'Look for any changes to plugin settings
    Dim settingsChanged As Boolean: settingsChanged = False
    
    Dim i As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        If (PluginManager.IsPluginCurrentlyEnabled(i) <> m_PluginEnabled(i)) Then
            PluginManager.SetPluginEnablement i, m_PluginEnabled(i)
            PluginManager.SetPluginAllowed i, m_PluginEnabled(i)
            settingsChanged = True
        End If
    Next i
    
    'If the user has changed any plugin enable/disable settings, a number of things must be refreshed program-wide
    If settingsChanged Then
        PluginManager.InitializePluginManager
        PluginManager.LoadPluginGroup True
        PluginManager.LoadPluginGroup False
        ApplyAllMenuIcons
        IconsAndCursors.ResetMenuIcons
        g_ImageFormats.GenerateInputFormats
        g_ImageFormats.GenerateOutputFormats
    End If
    
    'End batch preference update mode, which will force a write-to-file operation
    g_UserPreferences.EndBatchPreferenceMode
    
    Message "Plugin options saved."
    
End Sub

'RESET all plugin options
Private Sub cmdReset_Click()

    'Set current container to zero
    g_UserPreferences.SetPref_Long "Plugins", "Last Plugin Preferences Page", 0
    
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
    
    'Populate the left-hand list box with all relevant plugins
    lstPlugins.Clear
    lstPlugins.AddItem "Overview", 0, True
    
    Dim i As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        lstPlugins.AddItem PluginManager.GetPluginName(i), i + 1
    Next i
    
    lstPlugins.ListIndex = 0
    
    'Dynamically generate all text on the overview page
    Dim maxWidth As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        
        'Load two new label instances
        If (i > 0) Then
            Load lblInterfaceTitle(i): lblInterfaceTitle(i).Visible = True
            Load lblStatus(i): lblStatus(i).Visible = True
        End If
        
        'Assign title captions and position accordingly
        lblInterfaceTitle(i).Caption = PluginManager.GetPluginName(i) & ":"
        If lblInterfaceTitle(i).GetWidth > maxWidth Then maxWidth = lblInterfaceTitle(i).GetWidth
        
        'Align the top position of each status label with its matching title label
        If (i > 0) Then lblInterfaceTitle(i).SetTop lblInterfaceTitle(i - 1).GetTop + lblInterfaceTitle(i - 1).GetHeight + FixDPI(12)
        lblStatus(i).SetTop lblInterfaceTitle(i).GetTop
        
    Next i
    
    'Left-align all status labels to the same position
    maxWidth = maxWidth + lblInterfaceTitle(0).GetLeft + FixDPI(8)
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        lblStatus(i).SetLeft maxWidth
        lblStatus(i).SetWidth picContainer(0).GetWidth - lblStatus(i).GetLeft
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
    ReDim m_PluginEnabled(0 To PluginManager.GetNumOfPlugins - 1) As Boolean
    Dim i As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        m_PluginEnabled(i) = PluginManager.IsPluginCurrentlyEnabled(i)
    Next i
    
    'Start batch preference processing mode.
    g_UserPreferences.StartBatchPreferenceMode
    
    'Now, check version numbers of each plugin.  This is more complicated than it needs to be, on account of
    ' each plugin having its own unique mechanism for version-checking, but I have wrapped these various functions
    ' inside fairly standard wrapper calls.
    CollectAllVersionNumbers
    
    'We now have a collection of version numbers for our various plugins.  Let's use those to populate our
    ' "good/bad" labels for each plugin.
    UpdatePluginLabels
    
    'Enable the last container the user selected
    lstPlugins.ListIndex = g_UserPreferences.GetPref_Long("Plugins", "Last Plugin Preferences Page", 0)
    PluginChanged
    
    'End batch preference mode
    g_UserPreferences.EndBatchPreferenceMode
    
End Sub

'Assuming version numbers have been successfully retrieved, this function can be called to update the
' green/red plugin label display on the main panel.
Private Sub UpdatePluginLabels()
    
    Dim pluginStatus As Boolean: pluginStatus = True
    
    Dim i As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        pluginStatus = pluginStatus And PopPluginLabel(i)
    Next i
    
    If pluginStatus Then
        lblPluginStatus.ForeColor = m_Colors.RetrieveColor(PDPM_GoodText)
        lblPluginStatus.Caption = UCase(g_Language.TranslateMessage("GOOD"))
    Else
        lblPluginStatus.ForeColor = m_Colors.RetrieveColor(PDPM_BadText)
        lblPluginStatus.Caption = g_Language.TranslateMessage("problems detected")
    End If
        
End Sub

'Retrieve all relevant plugin version numbers and store them in the m_PluginVersion() array
Private Sub CollectAllVersionNumbers()
    
    ReDim m_PluginVersion(0 To PluginManager.GetNumOfPlugins - 1) As String
    
    'Start by querying the plugin file's metadata for version information.  This only works for some plugins,
    ' unfortunately, but we'll manually fill in outliers afterward.
    Dim i As Long
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        If PluginManager.IsPluginCurrentlyInstalled(i) Then
            m_PluginVersion(i) = PluginManager.GetPluginVersion(i)
        Else
            m_PluginVersion(i) = vbNullString
        End If
    Next i
    
    'Remove trailing build numbers from version strings as necessary.  (Note: ExifTool is ignored, as it does not
    ' actually report a build number)
    For i = 0 To PluginManager.GetNumOfPlugins - 1
        If (i <> CCP_ExifTool) Then
            If Len(m_PluginVersion(i)) <> 0 Then
                StripOffExtension m_PluginVersion(i)
            Else
                m_PluginVersion(i) = g_Language.TranslateMessage("none")
            End If
        End If
    Next i
    
End Sub

'Given a plugin's availability, expected version, and index on this form, populate the relevant labels associated with it.
' This function will return TRUE if the plugin is in good status, FALSE if it isn't (for any reason)
Private Function PopPluginLabel(ByVal pluginID As CORE_PLUGINS) As Boolean
    
    'Is this plugin present on the machine?
    If PluginManager.IsPluginCurrentlyInstalled(pluginID) Then
    
        'If present, has it been forcibly disabled?  (Note that we use our internal enablement tracker for this,
        ' to reflect any changes the user has just made.)
        If m_PluginEnabled(pluginID) Then
            
            'If this plugin is present and enabled, does its version match what we expect?
            If StrComp(m_PluginVersion(pluginID), PluginManager.ExpectedPluginVersion(pluginID), vbBinaryCompare) = 0 Then
                lblStatus(pluginID).Caption = g_Language.TranslateMessage("installed and up to date")
                lblStatus(pluginID).ForeColor = m_Colors.RetrieveColor(PDPM_GoodText)
                PopPluginLabel = True
                
            'Version mismatch
            Else
                lblStatus(pluginID).Caption = g_Language.TranslateMessage("installed, but version is unexpected")
                lblStatus(pluginID).ForeColor = m_Colors.RetrieveColor(PDPM_BadText)
                PopPluginLabel = False
            End If
            
        'Plugin is disabled
        Else
            lblStatus(pluginID).Caption = g_Language.TranslateMessage("installed, but disabled by user")
            lblStatus(pluginID).ForeColor = m_Colors.RetrieveColor(PDPM_BadText)
            PopPluginLabel = False
        End If
        
    'Plugin is not present on the machine
    Else
        lblStatus(pluginID).Caption = g_Language.TranslateMessage("not installed")
        lblStatus(pluginID).ForeColor = m_Colors.RetrieveColor(PDPM_BadText)
        PopPluginLabel = False
    End If
    
End Function

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDPM_GoodText, "PluginOK", RGB(0, 255, 0)
    m_Colors.LoadThemeColor PDPM_BadText, "PluginError", RGB(255, 0, 0)
End Sub

'When a new plugin is selected, display only the relevant plugin panel
Private Sub lstPlugins_Click()
    PluginChanged
End Sub

Private Sub PluginChanged()

    'Display the overview panel
    If (lstPlugins.ListIndex = 0) Then
        picContainer(0).Visible = True
        picContainer(1).Visible = False
    
    'Display the plugin-specific panel, including populating a bunch of run-time text
    Else
        picContainer(0).Visible = False
        picContainer(1).Visible = True
        
        Dim pluginIndex As CORE_PLUGINS, pluginName As String
        pluginIndex = lstPlugins.ListIndex - 1
        pluginName = PluginManager.GetPluginName(pluginIndex)
        
        lblPluginTitle.Caption = g_Language.TranslateMessage("%1 summary", pluginName)
        lblPluginExpectedVersion.Caption = PluginManager.ExpectedPluginVersion(pluginIndex)
        
        If PluginManager.IsPluginCurrentlyInstalled(pluginIndex) Then
            lblPluginVersion.Caption = m_PluginVersion(pluginIndex)
            If StrComp(m_PluginVersion(pluginIndex), PluginManager.ExpectedPluginVersion(pluginIndex), vbBinaryCompare) = 0 Then
                lblPluginVersion.ForeColor = m_Colors.RetrieveColor(PDPM_GoodText)
            Else
                lblPluginVersion.ForeColor = m_Colors.RetrieveColor(PDPM_BadText)
            End If
        Else
            lblPluginVersion.Caption = g_Language.TranslateMessage("missing")
            lblPluginVersion.ForeColor = m_Colors.RetrieveColor(PDPM_BadText)
        End If
        
        hypHomepage.Caption = PluginManager.GetPluginHomepage(pluginIndex)
        hypHomepage.URL = PluginManager.GetPluginHomepage(pluginIndex)
        hypLicense.Caption = PluginManager.GetPluginLicenseName(pluginIndex)
        hypLicense.URL = PluginManager.GetPluginLicenseURL(pluginIndex)
        
        m_IgnoreButtonStripEvents = True
        If m_PluginEnabled(pluginIndex) Then btsDisablePlugin.ListIndex = 0 Else btsDisablePlugin.ListIndex = 1
        m_IgnoreButtonStripEvents = False
        
    End If
    
End Sub
