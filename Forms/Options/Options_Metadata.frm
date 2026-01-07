VERSION 5.00
Begin VB.Form options_Metadata 
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
   Begin PhotoDemon.pdCheckBox chkLoadingOrientation 
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Top             =   2280
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "obey auto-rotate instructions inside image files"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   0
      Top             =   1920
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   503
      Caption         =   "orientation"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   503
      Caption         =   "loading"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdCheckBox chkMetadataBinary 
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   1440
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "forcibly extract binary-type tags as Base64 (slow)"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkMetadataJPEG 
      Height          =   330
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "estimate original JPEG quality settings"
   End
   Begin PhotoDemon.pdCheckBox chkMetadataUnknown 
      Height          =   330
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "extract unknown tags"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkMetadataDuplicates 
      Height          =   330
      Left            =   180
      TabIndex        =   4
      Top             =   360
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      Caption         =   "automatically hide duplicate tags"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   4
      Left            =   0
      Top             =   2760
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   503
      Caption         =   "saving"
      FontSize        =   12
      ForeColor       =   5263440
   End
   Begin PhotoDemon.pdCheckBox chkMetadataListPD 
      Height          =   375
      Left            =   180
      TabIndex        =   5
      Top             =   3120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      Caption         =   "list PhotoDemon as the last-used editing software"
   End
End
Attribute VB_Name = "options_Metadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Tools > Options > Metadata panel
'Copyright 2002-2026 by Tanner Helland
'Created: 8/November/02
'Last updated: 08/April/25
'Last update: combine metadata options from other panels into this standalone panel
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

Public Sub LoadUserPreferences()
    
    chkMetadataDuplicates.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Hide Duplicates", True)
    chkMetadataJPEG.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Estimate JPEG", True)
    chkMetadataUnknown.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Extract Unknown", False)
    chkMetadataBinary.Value = UserPrefs.GetPref_Boolean("Loading", "Metadata Extract Binary", False)
    chkLoadingOrientation.Value = UserPrefs.GetPref_Boolean("Loading", "ExifAutoRotate", True)
    chkMetadataListPD.Value = UserPrefs.GetPref_Boolean("Saving", "MetadataListPD", True)
    
End Sub

Public Sub SaveUserPreferences()

    UserPrefs.SetPref_Boolean "Loading", "Metadata Hide Duplicates", chkMetadataDuplicates.Value
    UserPrefs.SetPref_Boolean "Loading", "Metadata Estimate JPEG", chkMetadataJPEG.Value
    UserPrefs.SetPref_Boolean "Loading", "Metadata Extract Binary", chkMetadataBinary.Value
    UserPrefs.SetPref_Boolean "Loading", "Metadata Extract Unknown", chkMetadataUnknown.Value
    
    UserPrefs.SetPref_Boolean "Loading", "ExifAutoRotate", chkLoadingOrientation.Value
    
    UserPrefs.SetPref_Boolean "Saving", "MetadataListPD", chkMetadataListPD.Value
    
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
    
    chkMetadataDuplicates.AssignTooltip "Older cameras and photo-editing software may not embed metadata correctly, leading to multiple metadata copies within a single file.  PhotoDemon can automatically resolve duplicate entries for you."
    chkMetadataJPEG.AssignTooltip "The JPEG format does not provide a way to store JPEG quality settings inside image files.  PhotoDemon can work around this by inferring quality settings from other metadata (like quantization tables)."
    chkMetadataUnknown.AssignTooltip "Some camera manufacturers store proprietary metadata tags inside image files.  These tags are not generally useful to humans, but PhotoDemon can attempt to extract them anyway."
    chkMetadataBinary.AssignTooltip "By default, large binary tags (like image thumbnails) are not processed.  Instead, PhotoDemon simply reports the size of the embedded data.  If you require this data, PhotoDemon can manually convert it to Base64 for further analysis."
    chkLoadingOrientation.AssignTooltip "Most digital photos include rotation instructions (EXIF orientation metadata), which PhotoDemon will use to automatically rotate photos.  Some older smartphones and cameras may not write these instructions correctly, so if your photos are being imported sideways or upside-down, you can try disabling the auto-rotate feature."
    
    chkMetadataListPD.AssignTooltip "The EXIF specification asks programs to correctly identify themselves as the software of origin when exporting image files.  For increased privacy, you can suspend this behavior."
    
    Interface.ApplyThemeAndTranslations Me
    
End Sub
