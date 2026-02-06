VERSION 5.00
Begin VB.Form options_ColorManagement 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8295
   ControlBox      =   0   'False
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
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   553
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdCheckBox chkColorManagement 
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   556
      Caption         =   "use embedded ICC profiles, when available"
   End
   Begin PhotoDemon.pdCheckBox chkColorManagement 
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   556
      Caption         =   "use black point compensation"
   End
   Begin PhotoDemon.pdDropDown cboDisplayRenderIntent 
      Height          =   735
      Left            =   180
      TabIndex        =   2
      Top             =   3240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1296
      Caption         =   "display rendering intent:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdButton cmdColorProfilePath 
      Height          =   375
      Left            =   7380
      TabIndex        =   3
      Top             =   2760
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   661
      Caption         =   "..."
   End
   Begin PhotoDemon.pdDropDown cboDisplays 
      Height          =   690
      Left            =   780
      TabIndex        =   4
      Top             =   1590
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   1217
      Caption         =   "available displays:"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdTextBox txtColorProfilePath 
      Height          =   315
      Left            =   900
      TabIndex        =   5
      Top             =   2790
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   556
      Text            =   "(none)"
   End
   Begin PhotoDemon.pdRadioButton optColorManagement 
      Height          =   330
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   480
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "turn off display color management"
      Value           =   -1  'True
   End
   Begin PhotoDemon.pdRadioButton optColorManagement 
      Height          =   330
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   840
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "use the current system profiles for each display"
   End
   Begin PhotoDemon.pdLabel lblColorManagement 
      Height          =   240
      Index           =   2
      Left            =   780
      Top             =   2430
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   503
      Caption         =   "color profile for this display:"
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   503
      Caption         =   "display policies"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdRadioButton optColorManagement 
      Height          =   330
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   1200
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "use custom profiles for each display"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   24
      Left            =   0
      Top             =   4560
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   503
      Caption         =   "file policies"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdCheckBox chkColorManagement 
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   5400
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   556
      Caption         =   "use format-specific color management data, when available"
   End
End
Attribute VB_Name = "options_ColorManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options > Color Management panel
'Copyright 2002-2026 by Tanner Helland
'Created: 8/November/02
'Last updated: 02/April/25
'Last update: split this panel into a standalone form
'
'This form contains a single subpanel worth of program options.  At run-time, it is dynamically
' made a child of FormOptions.  It will only be loaded if/when the user interacts with this category.
'
'All Tools > Options child panels contain some mandatory public functions, including ones for loading
' and saving user preferences, as well as validating any UI elements where the user can enter
' custom values.  (A reset-style function is *not* required; this is automatically handled by
' FormOptions.)
'
'This form, like all Tools > Options panels, interacts heavily with the UserPrefs module.
' (That module is responsible for all low-level preference reading/writing.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Whenever the Color and Transparency -> Color Management -> Monitor combo box is changed, load the relevant color profile
' path from the preferences file (if one exists)
Private Sub cboDisplays_Click()

    'One of the difficulties with tracking multiple monitors is that the user can attach/detach them at will.
    
    'Prior to v7.0, PD used HMONITOR handles to track displays, using the reasoning from this article:
    ' http://www.microsoft.com/msj/0697/monitor/monitor.aspx
    '...specifically the line, "A physical device has the same HMONITOR value throughout its lifetime,
    ' even across changes to display settings, as long as it remains a part of the desktop."
    
    'This worked "well enough", as long as the user never disconnected the display monitor only to attach
    ' it again at some point in the future (as is common with second monitors and a laptop, for example).
    
    'In 7.0, this system was upgraded to use monitor serial numbers, and only fall back to the HMONITOR
    ' if a serial number (or EDID) doesn't exist.
    
    Dim uniqueDisplayID As String
    If (Not g_Displays.Displays(cboDisplays.ListIndex) Is Nothing) Then
        uniqueDisplayID = g_Displays.Displays(cboDisplays.ListIndex).GetUniqueDescriptor
        Dim tmpXML As pdXML
        Set tmpXML = New pdXML
        uniqueDisplayID = tmpXML.GetXMLSafeTagName(uniqueDisplayID)
    End If
    
    'Use that to retrieve a stored color profile (if any)
    Dim profilePath As String
    profilePath = UserPrefs.GetPref_String("ColorManagement", "DisplayProfile_" & uniqueDisplayID, "(none)")
    
    'If the returned value is "(none)", translate that into the user's language before displaying; otherwise, display
    ' whatever path we retrieved.
    If Strings.StringsEqual(profilePath, "(none)", False) Then
        txtColorProfilePath.Text = g_Language.TranslateMessage("(none)")
    Else
        txtColorProfilePath.Text = profilePath
    End If
    
End Sub

'Allow the user to select a new color profile for the attached monitor.  Because this text box is re-used for multiple
' settings, save any changes to file immediately, rather than waiting for the user to click OK.
Private Sub cmdColorProfilePath_Click()

    'Disable user input until the dialog closes
    Interface.DisableUserInput
    
    Dim sFile As String
    
    'Get the last color profile path from the preferences file
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Paths", "Color Profile", vbNullString)
    
    'If no color profile path was found, populate it with the default system color profile path
    If (LenB(tempPathString) = 0) Then tempPathString = GetSystemColorFolder()
    
    'Prepare a common dialog filter list with extensions of known profile types
    Dim cdFilter As String
    cdFilter = g_Language.TranslateMessage("ICC profile") & " (.icc, .icm)|*.icc;*.icm"
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Please select a color profile")
    
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    
    If openDialog.GetOpenFileName(sFile, , True, False, cdFilter, 1, tempPathString, cdTitle, ".icc", FormOptions.hWnd) Then
        
        'Save this new directory as the default path for future usage
        Dim listPath As String
        listPath = Files.FileGetPath(sFile)
        UserPrefs.SetPref_String "Paths", "Color Profile", listPath
        
        'Set the text box to match this color profile, and save the resulting preference out to file.
        txtColorProfilePath = sFile
        
        Dim uniqueMonID As String
        If (Not g_Displays.Displays(cboDisplays.ListIndex) Is Nothing) Then
            uniqueMonID = g_Displays.Displays(cboDisplays.ListIndex).GetUniqueDescriptor
            Dim tmpXML As pdXML
            Set tmpXML = New pdXML
            uniqueMonID = tmpXML.GetXMLSafeTagName(uniqueMonID)
        End If
        
        UserPrefs.SetPref_String "ColorManagement", "DisplayProfile_" & uniqueMonID, sFile
        
        'If the "user custom color profiles" option button isn't selected, mark it now
        If (Not optColorManagement(2).Value) Then optColorManagement(2).Value = True
        
    End If
    
    'Re-enable user input
    Interface.EnableUserInput

End Sub

Private Sub Form_Load()

    'Color-management prefs
    Interface.PopulateRenderingIntentDropDown cboDisplayRenderIntent, ColorManagement.GetDisplayRenderingIntentPref()
    
    'Load a list of all available displays
    cboDisplays.Clear
    
    Dim mainDisplay As String, secondaryDisplay As String
    mainDisplay = g_Language.TranslateMessage("Primary display:")
    secondaryDisplay = g_Language.TranslateMessage("Secondary display:")
    
    'Failsafe only
    If (g_Displays Is Nothing) Then Exit Sub
    
    Dim primaryIndex As Long, displayEntry As String
    If (g_Displays.GetDisplayCount > 0) Then
        
        Dim i As Long
        For i = 0 To g_Displays.GetDisplayCount - 1
        
            displayEntry = vbNullString
            
            'Explicitly label the primary monitor
            If g_Displays.Displays(i).IsPrimary Then
                displayEntry = mainDisplay
                primaryIndex = i
            Else
                displayEntry = secondaryDisplay
            End If
            
            'Add the monitor's physical size
            displayEntry = displayEntry & " " & g_Displays.Displays(i).GetMonitorSizeAsString
            
            'Add the monitor's name
            displayEntry = displayEntry & " " & g_Displays.Displays(i).GetBestMonitorName
            
            'Add the monitor's native resolution
            displayEntry = displayEntry & " (" & g_Displays.Displays(i).GetMonitorResolutionAsString & ")"
                            
            'Display this monitor in the list
            cboDisplays.AddItem displayEntry, i
            
        Next i
        
    Else
        primaryIndex = 0
        cboDisplays.AddItem "Unknown display", 0
    End If
    
    'Display the primary monitor by default; this will also trigger a load of the matching
    ' custom profile, if one exists.
    cboDisplays.ListIndex = primaryIndex
    
    optColorManagement(0).AssignTooltip "Turning off display color management can provide a small performance boost.  If your display is not currently configured for color management, use this setting."
    optColorManagement(1).AssignTooltip "This setting is the best choice for most users.  If you have no idea what color management is, use this setting.  If you have correctly configured a display profile via the Windows Control Panel, also use this setting."
    optColorManagement(2).AssignTooltip "To configure custom color profiles on a per-monitor basis, please use this setting."
    
    cboDisplays.AssignTooltip "Please specify a color profile for each monitor currently attached to the system.  Note that the text in parentheses is the display adapter driving the named monitor."
    cmdColorProfilePath.AssignTooltip "Click this button to bring up a ""browse for color profile"" dialog."
    
    cboDisplayRenderIntent.AssignTooltip "If you do not know what this setting controls, set it to ""Perceptual"".  Perceptual rendering intent is the best choice for most users."
    chkColorManagement(0).AssignTooltip "BPC is primarily relevant in colorimetric rendering intents, where it helps preserve detail in dark (shadow) regions of images.  For most workflows, BPC should be turned ON."
    
    chkColorManagement(1).Value = ColorManagement.UseEmbeddedICCProfiles()
    chkColorManagement(1).AssignTooltip "Embedded ICC profiles improve color fidelity.  Even if this setting is turned off, PhotoDemon may still use ICC profiles for some tasks (like handling CMYK data)."
    chkColorManagement(2).Value = ColorManagement.UseEmbeddedLegacyProfiles()
    chkColorManagement(2).AssignTooltip "Some image formats support both ICC profiles and their own color management solutions.  PhotoDemon always prefers ICC profiles, but when none are embedded, other color management approaches can be tried."
    
End Sub

Public Sub LoadUserPreferences()

    'Color-management preferences
    optColorManagement(ColorManagement.GetDisplayColorManagementPreference()).Value = True
    chkColorManagement(0).Value = ColorManagement.GetDisplayBPC()
    cboDisplayRenderIntent.ListIndex = ColorManagement.GetDisplayRenderingIntentPref()
    ' (note: monitor display preferences are also here, but they are retrieved auto-magically
    '  when the display dropdown listindex changes)
    
End Sub

Public Sub SaveUserPreferences()

    If optColorManagement(0).Value Then
        ColorManagement.SetDisplayColorManagementPreference DCM_NoManagement
    ElseIf optColorManagement(1).Value Then
        ColorManagement.SetDisplayColorManagementPreference DCM_SystemProfile
    Else
        ColorManagement.SetDisplayColorManagementPreference DCM_CustomProfile
    End If
    
    ColorManagement.SetDisplayBPC chkColorManagement(0).Value
    ColorManagement.SetDisplayRenderingIntentPref cboDisplayRenderIntent.ListIndex
    UserPrefs.SetPref_Boolean "ColorManagement", "allow-icc-profiles", chkColorManagement(1).Value
    UserPrefs.SetPref_Boolean "ColorManagement", "allow-legacy-profiles", chkColorManagement(2).Value
    ColorManagement.UpdateColorManagementPreferences
    
    'Changes to color preferences require us to re-cache any working-space-to-screen transform data.
    CacheDisplayCMMData
    ColorManagement.CheckParentMonitor False, True
    
End Sub

'Upon calling, validate all input.  Return FALSE if validation on 1+ controls fails.
Public Function ValidateAllInput() As Boolean
    
    ValidateAllInput = True
    
    Dim eControl As Object
    For Each eControl In Me.Controls
        
        'Most UI elements on this dialog are idiot-proof, but spin controls (including those embedded
        ' in slider controls) are an exception.
        If (TypeOf eControl Is pdSlider) Or (TypeOf eControl Is pdSpinner) Then
            
            'Finally, ask the control to validate itself
            If (Not eControl.IsValid) Then
                ValidateAllInput = False
                Exit For
            End If
            
        End If
    Next eControl
    
End Function

'This function is called at least once, immediately following Form_Load(),
' but it can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    Interface.ApplyThemeAndTranslations Me
End Sub

