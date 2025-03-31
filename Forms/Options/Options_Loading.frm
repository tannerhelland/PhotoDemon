VERSION 5.00
Begin VB.Form options_Loading 
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
   Begin PhotoDemon.pdButtonStrip btsMultiInstance 
      Height          =   975
      Left            =   150
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1720
      Caption         =   "when images arrive from an external source (like Windows Explorer):"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdCheckBox chkToneMapping 
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   2040
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "display tone mapping options when importing HDR and RAW images"
   End
   Begin PhotoDemon.pdCheckBox chkLoadingOrientation 
      Height          =   330
      Left            =   180
      TabIndex        =   2
      Top             =   5040
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "obey auto-rotate instructions inside image files"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   9
      Left            =   0
      Top             =   4680
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   503
      Caption         =   "orientation"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   10
      Left            =   0
      Top             =   1680
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   503
      Caption         =   "high-dynamic range (HDR) images"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   12
      Left            =   0
      Top             =   2640
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   503
      Caption         =   "metadata"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdCheckBox chkMetadataBinary 
      Height          =   330
      Left            =   180
      TabIndex        =   3
      Top             =   4080
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "forcibly extract binary-type tags as Base64 (slow)"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkMetadataJPEG 
      Height          =   330
      Left            =   180
      TabIndex        =   4
      Top             =   3360
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "estimate original JPEG quality settings"
   End
   Begin PhotoDemon.pdCheckBox chkMetadataUnknown 
      Height          =   330
      Left            =   180
      TabIndex        =   5
      Top             =   3720
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "extract unknown tags"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkMetadataDuplicates 
      Height          =   330
      Left            =   180
      TabIndex        =   6
      Top             =   3000
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "automatically hide duplicate tags"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   11
      Left            =   0
      Top             =   0
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   503
      Caption         =   "app instances"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   22
      Left            =   0
      Top             =   5520
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   503
      Caption         =   "session restore"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdCheckBox chkSystemReboots 
      Height          =   330
      Left            =   180
      TabIndex        =   7
      Top             =   5880
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "automatically restore sessions interrupted by system updates or reboots"
   End
End
Attribute VB_Name = "options_Loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    'Loading prefs
    chkToneMapping.AssignTooltip "HDR and RAW images contain more colors than PC screens can physically display.  Before displaying such images, a tone mapping operation must be applied to the original image data."
    btsMultiInstance.AddItem "load into this instance", 0
    btsMultiInstance.AddItem "load into a new PhotoDemon instance", 1
    chkMetadataDuplicates.AssignTooltip "Older cameras and photo-editing software may not embed metadata correctly, leading to multiple metadata copies within a single file.  PhotoDemon can automatically resolve duplicate entries for you."
    chkMetadataJPEG.AssignTooltip "The JPEG format does not provide a way to store JPEG quality settings inside image files.  PhotoDemon can work around this by inferring quality settings from other metadata (like quantization tables)."
    chkMetadataUnknown.AssignTooltip "Some camera manufacturers store proprietary metadata tags inside image files.  These tags are not generally useful to humans, but PhotoDemon can attempt to extract them anyway."
    chkMetadataBinary.AssignTooltip "By default, large binary tags (like image thumbnails) are not processed.  Instead, PhotoDemon simply reports the size of the embedded data.  If you require this data, PhotoDemon can manually convert it to Base64 for further analysis."
    chkLoadingOrientation.AssignTooltip "Most digital photos include rotation instructions (EXIF orientation metadata), which PhotoDemon will use to automatically rotate photos.  Some older smartphones and cameras may not write these instructions correctly, so if your photos are being imported sideways or upside-down, you can try disabling the auto-rotate feature."
    chkSystemReboots.AssignTooltip "If your PC reboots while PhotoDemon is running, PhotoDemon can automatically restore your previous session."
    
End Sub

Public Sub LoadUserPreferences()
    
    'Loading preferences
    If UserPrefs.GetPref_Boolean("Loading", "Single Instance", False) Then btsMultiInstance.ListIndex = 0 Else btsMultiInstance.ListIndex = 1
    chkToneMapping.Value = UserPrefs.GetPref_Boolean("Loading", "Tone Mapping Prompt", True)
    chkMetadataDuplicates.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Hide Duplicates", True)
    chkMetadataJPEG.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Estimate JPEG", True)
    chkMetadataUnknown.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Extract Unknown", False)
    chkMetadataBinary.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Extract Binary", False)
    chkLoadingOrientation.Value = UserPrefs.GetPref_Boolean("Loading", "EXIF Auto Rotate", True)
    chkSystemReboots.Value = UserPrefs.GetPref_Boolean("Loading", "RestoreAfterReboot", False)
    
End Sub

Public Sub SaveUserPreferences()

    UserPrefs.SetPref_Boolean "Loading", "Single Instance", (btsMultiInstance.ListIndex = 0)
    
    UserPrefs.SetPref_Boolean "Loading", "Tone Mapping Prompt", chkToneMapping.Value
    
    UserPrefs.SetPref_Boolean "Loading", "Metadata Hide Duplicates", chkMetadataDuplicates.Value
    UserPrefs.SetPref_Boolean "Loading", "Metadata Estimate JPEG", chkMetadataJPEG.Value
    UserPrefs.SetPref_Boolean "Loading", "Metadata Extract Binary", chkMetadataBinary.Value
    UserPrefs.SetPref_Boolean "Loading", "Metadata Extract Unknown", chkMetadataUnknown.Value
    
    UserPrefs.SetPref_Boolean "Loading", "ExifAutoRotate", chkLoadingOrientation.Value
    
    'Restore after reboot behavior requires an immediate API to de/activate
    UserPrefs.SetPref_Boolean "Loading", "RestoreAfterReboot", chkSystemReboots.Value
    OS.SetRestartRestoreBehavior chkSystemReboots.Value
    
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
