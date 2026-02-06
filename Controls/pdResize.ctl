VERSION 5.00
Begin VB.UserControl pdResize 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   ControlContainer=   -1  'True
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   187
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   552
   ToolboxBitmap   =   "pdResize.ctx":0000
   Begin PhotoDemon.pdButtonToolbox cmdAspectRatio 
      Height          =   630
      Left            =   7590
      TabIndex        =   2
      Top             =   165
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1111
      DontHighlightDownState=   -1  'True
      StickyToggle    =   -1  'True
   End
   Begin PhotoDemon.pdDropDown cmbResolution 
      Height          =   360
      Left            =   3825
      TabIndex        =   6
      Top             =   1200
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   635
      FontSize        =   11
   End
   Begin PhotoDemon.pdDropDown cmbHeightUnit 
      Height          =   360
      Left            =   3825
      TabIndex        =   3
      Top             =   600
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   635
      FontSize        =   11
   End
   Begin PhotoDemon.pdDropDown cmbWidthUnit 
      Height          =   360
      Left            =   3825
      TabIndex        =   4
      Top             =   0
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   635
      FontSize        =   11
   End
   Begin PhotoDemon.pdSpinner tudWidth 
      Height          =   375
      Left            =   2130
      TabIndex        =   0
      Top             =   30
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      DefaultValue    =   1
      Min             =   1
      Max             =   32767
      Value           =   1
      FontSize        =   11
   End
   Begin PhotoDemon.pdSpinner tudHeight 
      Height          =   375
      Left            =   2130
      TabIndex        =   1
      Top             =   630
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      DefaultValue    =   1
      Min             =   1
      Max             =   32767
      Value           =   1
      FontSize        =   11
   End
   Begin PhotoDemon.pdSpinner tudResolution 
      Height          =   375
      Left            =   2130
      TabIndex        =   5
      Top             =   1230
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      DefaultValue    =   96
      Min             =   1
      Max             =   32767
      SigDigits       =   2
      Value           =   1
      FontSize        =   11
   End
   Begin PhotoDemon.pdLabel lblResolution 
      Height          =   285
      Left            =   0
      Top             =   1230
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "resolution"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblDimensions 
      Height          =   285
      Index           =   0
      Left            =   0
      Top             =   1800
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "dimensions"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblDimensions 
      Height          =   285
      Index           =   1
      Left            =   2130
      Top             =   1800
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   503
      Caption         =   "            "
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblAspectRatio 
      Height          =   285
      Index           =   1
      Left            =   2130
      Top             =   2370
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   503
      Caption         =   "            "
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblWidth 
      Height          =   285
      Left            =   0
      Top             =   30
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "width"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblHeight 
      Height          =   285
      Left            =   0
      Top             =   630
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "height"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblAspectRatio 
      Height          =   285
      Index           =   0
      Left            =   0
      Top             =   2370
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "aspect ratio"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "pdResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Image Resize User Control
'Copyright 2001-2026 by Tanner Helland
'Created: 12/June/01
'Last updated: 18/August/21
'Last update: display in-memory size of the image as dimensions are changed (vs orig. size, when appropriate)
'
'Many tools in PD relate to resizing: image size, canvas size, (soon) layer size, content-aware rescaling,
' perhaps a more advanced autocrop tool, plus dedicated resize options in the batch converter...
'
'Rather than develop custom resize UIs for all these scenarios, I finally asked myself: why not use a single
' resize-centric user control?  As an added bonus, that would allow me to update the user control to extend
' new capabilities to all of PD's resize tools.
'
'Thus this UC was born.  All resize-related dialogs in the project now use it, and any new features can now
' automatically propagate across them.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This object provides a single raised event:
' - Change (which triggers when a size value is updated)
Public Event Change(ByVal newWidthPixels As Double, ByVal newHeightPixels As Double)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Store a copy of the original width/height values we are passed
Private m_initWidth As Long, m_initHeight As Long

'Store a copy of the initial DPI value we are passed
Private m_initDPI As Double

'Used for maintaining ratios when the check box is clicked
Private m_wRatio As Double, m_hRatio As Double

'Used to prevent infinite recursion as updates to one text box force updates to the other
Private m_allowedToUpdateWidth As Boolean, m_allowedToUpdateHeight As Boolean

'Used to prevent infinite recursion as updates to one pixel init dropdown forces updates to the other
Private m_unitSyncingSuspended As Boolean

'When switching from one type of measurement to another, we must convert measurement values accordingly.
' This function is used to store the previous measurement method, which helps the conversion function know how to switch values.
Private m_previousUnitOfMeasurement As PD_MeasurementUnit

'Similar to changing measurement units for width/height, the user can also switch between PPI and PPCM for resolution.
' However, because this does not offer percent and measurement values, we track it separately.
Private m_previousUnitOfResolution As PD_ResolutionUnit

'"Unknown size mode" is used in the batch conversion dialog, as we don't know the size of images in advance
' (so if the user selects PERCENT resizing, we can't give them exact dimensions).
Private m_UnknownSizeMode As Boolean

'If percentage measurements are disabled, this will be set to TRUE.
Private m_PercentDisabled As Boolean

'If, at load-time, the owner wants us to default to inch/cm instead of pixels, this will be TRUE.
Private m_DefaultToRealWorldUnits As Boolean

'To minimize the frequency of raised _Change events, we only raise events if something has actually changed.
Private m_LastImgWidthPixels As Long, m_LastImgHeightPixels As Long

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDRESIZE_COLOR_LIST
    [_First] = 0
    PDR_Background = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_Resize
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function

'In most instances (e.g. Image > Resize), this control should default to PIXELS as the measurement.
' But in rare instances (e.g. Import > PDF), we want to default to real-world units (inch or cm,
' with PD choosing the appropriate one for the current user locale).
Public Property Get DefaultToRealWorldUnits() As Boolean
    DefaultToRealWorldUnits = m_DefaultToRealWorldUnits
End Property

Public Property Let DefaultToRealWorldUnits(ByVal newValue As Boolean)
    m_DefaultToRealWorldUnits = newValue
    PropertyChanged "DefaultToRealWorldUnits"
End Property

'If the owner does not want percentage available as an option, set this property to TRUE.
Public Property Get DisablePercentOption() As Boolean
    DisablePercentOption = m_PercentDisabled
End Property

Public Property Let DisablePercentOption(newMode As Boolean)
    m_PercentDisabled = newMode
    PopulateDropdowns
    PropertyChanged "DisablePercentOption"
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

'If any text value is NOT valid, this will return FALSE
Public Function IsValid(Optional ByVal showErrors As Boolean = True) As Boolean
    
    IsValid = True
    
    'If the current text value is not valid, highlight the problem and optionally display an error message box
    If Not tudWidth.IsValid(showErrors) Then
        IsValid = False
        Exit Function
    End If
    
    If Not tudHeight.IsValid(showErrors) Then
        IsValid = False
        Exit Function
    End If
    
    If Not tudResolution.IsValid(showErrors) Then
        IsValid = False
        Exit Function
    End If
    
    Dim pxWidthFloat As Double, pxHeightFloat As Double
    pxWidthFloat = ConvertUnitToPixels(GetCurrentWidthUnit, tudWidth, GetResolutionAsPPI(), m_initWidth)
    pxHeightFloat = ConvertUnitToPixels(GetCurrentWidthUnit, tudHeight, GetResolutionAsPPI(), m_initHeight)
    
    If (pxWidthFloat < 1) Or (pxHeightFloat < 1) Then
        IsValid = False
        If showErrors Then PDMsgBox "The final width and height measurements must be between 1 and %1 pixels.", vbInformation Or vbOKOnly, "Invalid measurement", CStr(PD_MAX_IMAGE_DIMENSION)
        Exit Function
    End If
    
    If (pxWidthFloat > PD_MAX_IMAGE_DIMENSION) Or (pxHeightFloat > PD_MAX_IMAGE_DIMENSION) Then
        IsValid = False
        If showErrors Then PDMsgBox "The final width and height measurements must be between 1 and %1 pixels.", vbInformation Or vbOKOnly, "Invalid measurement", CStr(PD_MAX_IMAGE_DIMENSION)
        Exit Function
    End If
    
End Function

'For batch conversions, we can't display an exact size when using PERCENT mode (as the exact size will vary according
' to the original image dimensions).  Use this mode to enable/disable that feature.
Public Property Get UnknownSizeMode() As Boolean
    UnknownSizeMode = m_UnknownSizeMode
End Property

Public Property Let UnknownSizeMode(newMode As Boolean)
    m_UnknownSizeMode = newMode
    PropertyChanged "UnknownSizeMode"
End Property

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

'User has changed the Resolution (PPI) measurement drop-down
Private Sub cmbResolution_Click()
    
    Select Case cmbResolution.ListIndex
    
        'Current unit is PPI
        Case ru_PPI
        
            'If the user previously had PPCM selected, convert the resolution now
            If (m_previousUnitOfResolution = ru_PPCM) Then
                If tudResolution.IsValid(False) Then
                    tudResolution = Units.GetCMFromInches(tudResolution)
                Else
                    tudResolution = Units.GetCMFromInches(m_initDPI)
                End If
            End If
            
            If (m_initDPI <> 0) Then tudResolution.DefaultValue = m_initDPI Else tudResolution.DefaultValue = 96
            
        'Current unit is PPCM
        Case ru_PPCM
        
            'If the user previously had PPI selected, convert the resolution now
            If (m_previousUnitOfResolution = ru_PPI) Then
                If tudResolution.IsValid(False) Then
                    tudResolution = Units.GetInchesFromCM(tudResolution)
                Else
                    tudResolution = m_initDPI
                End If
            End If
            
            If (m_initDPI <> 0) Then tudResolution.DefaultValue = GetInchesFromCM(m_initDPI) Else tudResolution.DefaultValue = GetInchesFromCM(96#)
    
    End Select
    
    'Note that changing the resolution unit will automatically trigger an updateAspectRatio call, so we don't need to do it here.
    
    'Store the current unit so future resolution changes know how to convert their values
    m_previousUnitOfResolution = cmbResolution.ListIndex
    
End Sub

'jcButton requires an hWnd for the parent control for subclassing purposes
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Lock aspect ratio can be set/retrieved by the owning dialog
Public Property Get AspectRatioLock() As Boolean
    AspectRatioLock = cmdAspectRatio.Value
End Property

Public Property Let AspectRatioLock(newSetting As Boolean)
    cmdAspectRatio.Value = newSetting
    SyncDimensions True
End Property

'Width and height in pixels can be set/retrieved from these properties.  Note that if the current text value for either dimension
' is invalid, this function will simply return the image's original width/height (in pixels, obviously).
Public Property Get ResizeWidthInPixels() As Long
    ResizeWidthInPixels = ConvertUnitToPixels(GetCurrentWidthUnit, tudWidth, GetResolutionAsPPI(), m_initWidth)
End Property

Public Property Let ResizeWidthInPixels(ByVal newWidth As Long)
    If m_PercentDisabled Then
        cmbWidthUnit.ListIndex = mu_Pixels - 1
    Else
        cmbWidthUnit.ListIndex = mu_Pixels
    End If
    m_unitSyncingSuspended = True
    tudWidth = newWidth
    m_unitSyncingSuspended = False
    SyncDimensions True
End Property

Public Property Get ResizeHeightInPixels() As Long
    ResizeHeightInPixels = ConvertUnitToPixels(GetCurrentHeightUnit, tudHeight, GetResolutionAsPPI(), m_initHeight)
End Property

Public Property Let ResizeHeightInPixels(ByVal newHeight As Long)
    If m_PercentDisabled Then
        cmbWidthUnit.ListIndex = mu_Pixels - 1
    Else
        cmbWidthUnit.ListIndex = mu_Pixels
    End If
    m_unitSyncingSuspended = True
    tudHeight = newHeight
    m_unitSyncingSuspended = False
    SyncDimensions False
End Property

'Width and height can be set/retrieved from these properties. IMPORTANT NOTE: these functions will return width/height
' per the current unit of measurement, so make sure to also read (and process) the .UnitOfMeasurement property!
' If the current text value for either dimension is invalid, this function will simply return the image's original width/height.
Public Property Get ResizeWidth() As Double
    If tudWidth.IsValid(False) Then
        ResizeWidth = tudWidth
    Else
        ResizeWidth = ConvertOtherUnitToPixels(GetCurrentWidthUnit, m_initWidth, GetResolutionAsPPI(), m_initWidth)
    End If
End Property

Public Property Let ResizeWidth(newWidth As Double)
    tudWidth = newWidth
    SyncDimensions True
End Property

Public Property Get ResizeHeight() As Double
    If tudHeight.IsValid(False) Then
        ResizeHeight = tudHeight
    Else
        ResizeHeight = ConvertOtherUnitToPixels(GetCurrentHeightUnit, m_initHeight, GetResolutionAsPPI(), m_initHeight)
    End If
End Property

Public Property Let ResizeHeight(newHeight As Double)
    tudHeight = newHeight
    SyncDimensions False
End Property

'Resolution can be set/retrieved via this property.  Note that if the current text value for resolution is invalid,
' this function will simply return the image's original resolution.
Public Property Get ResizeDPIAsPPI() As Double
    If tudResolution.IsValid(False) Then
        ResizeDPIAsPPI = GetResolutionAsPPI()
    Else
        ResizeDPIAsPPI = m_initDPI
    End If
End Property

Public Property Let ResizeDPIAsPPI(ByVal newDPI As Double)
    tudResolution = newDPI
    SyncDimensions True
End Property

Public Property Get ResizeDPI() As Double
    If tudResolution.IsValid(False) Then
        ResizeDPI = tudResolution.Value
    Else
        ResizeDPI = m_initDPI
    End If
End Property

Public Property Let ResizeDPI(ByVal newDPI As Double)
    tudResolution.Value = newDPI
    SyncDimensions True
End Property

'The current unit of measurement can also be retrieved.  Note that these values are kept in sync for both width/height.
Public Property Get UnitOfMeasurement() As PD_MeasurementUnit
    UnitOfMeasurement = GetCurrentWidthUnit
End Property

Public Property Let UnitOfMeasurement(newUnit As PD_MeasurementUnit)
    
    If m_PercentDisabled Then
        cmbWidthUnit.ListIndex = newUnit - 1
    Else
        cmbWidthUnit.ListIndex = newUnit
    End If
    
    'As a failsafe, make sure significant digits and everything else are properly synchronized.
    ' (This is necessary because the .ListIndex assignment, above, won't trigger a _Click event unless the
    '  new measurement differs from the old measurement.)
    ConvertUnitsToNewValue m_previousUnitOfMeasurement, newUnit
    
End Property

'The current unit of resolution (e.g. PPI).
Public Property Get UnitOfResolution() As PD_ResolutionUnit
    UnitOfResolution = cmbResolution.ListIndex
End Property

Public Property Let UnitOfResolution(newUnit As PD_ResolutionUnit)
    cmbResolution.ListIndex = newUnit
End Property

Private Sub cmbHeightUnit_Click()

    If m_unitSyncingSuspended Then Exit Sub
    
    'Suspend automatic synchronzation
    m_unitSyncingSuspended = True
    
    'Make the width drop-down match this one
    cmbWidthUnit.ListIndex = cmbHeightUnit.ListIndex
    
    'Convert the current measurements to the new ones.
    ConvertUnitsToNewValue m_previousUnitOfMeasurement, GetCurrentHeightUnit
    
    'Mark the new unit as the previous unit of measurement.  Future unit conversions will rely on this value to know how
    ' to convert their values.  (We must store this separately, because clicking a combo box will instantly change the
    ' ListIndex, erasing the previous value.)
    m_previousUnitOfMeasurement = GetCurrentHeightUnit
    
    'Restore automatic synchronization
    m_unitSyncingSuspended = False
    
    'Perform a final synchronization
    SyncDimensions False

End Sub

Private Sub cmbWidthUnit_Click()
    
    If m_unitSyncingSuspended Then Exit Sub

    'Suspend automatic synchronzation
    m_unitSyncingSuspended = True
    
    'Make the height drop-down match this one
    cmbHeightUnit.ListIndex = cmbWidthUnit.ListIndex
    
    'Convert the current measurements to the new ones.
    ConvertUnitsToNewValue m_previousUnitOfMeasurement, GetCurrentWidthUnit
    
    'Mark this as the previous unit of measurement.  Future unit conversions will rely on this value to know how to convert the present values.
    m_previousUnitOfMeasurement = GetCurrentWidthUnit
    
    'Restore automatic synchronization
    m_unitSyncingSuspended = False
    
    'Perform a final synchronization
    SyncDimensions True

End Sub

'Whenever the user switches to a new unit of measurement, we must convert all text box values (and possible limits and/or
' significant digits) to match.  Use this function to do so.
Private Sub ConvertUnitsToNewValue(ByVal oldUnit As PD_MeasurementUnit, ByVal newUnit As PD_MeasurementUnit)
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'Start by retrieving the old values in pixel measurements
    Dim imgWidthPixels As Double, imgHeightPixels As Double
    
    'If the current width or height value is invalid, note that the convertUnitToPixels function will
    ' simply return the original dimension.
    imgWidthPixels = ConvertUnitToPixels(oldUnit, tudWidth, GetResolutionAsPPI(), m_initWidth)
    imgHeightPixels = ConvertUnitToPixels(oldUnit, tudHeight, GetResolutionAsPPI(), m_initHeight)
    
    'Use those pixel measurements to retrieve new values in the desired unit of measurement
    Dim newWidth As Double, newHeight As Double
    newWidth = ConvertPixelToOtherUnit(newUnit, imgWidthPixels, GetResolutionAsPPI(), m_initWidth)
    newHeight = ConvertPixelToOtherUnit(newUnit, imgHeightPixels, GetResolutionAsPPI(), m_initHeight)
        
    'Depending on the unit of measurement, change the significant digits and upper limit of the text up/down boxes
    Select Case newUnit
    
        Case mu_Percent
            tudWidth.SigDigits = 1
            tudWidth.Min = 0.1
            tudWidth.Max = 3200#
            tudWidth.DefaultValue = 100#
            tudHeight.DefaultValue = 100#
            
        Case mu_Pixels
            tudWidth.SigDigits = 0
            tudWidth.Min = 1
            tudWidth.Max = PD_MAX_IMAGE_DIMENSION
            If (m_initWidth <> 0) Then tudWidth.DefaultValue = m_initWidth Else tudWidth.DefaultValue = g_Displays.GetDesktopWidth
            If (m_initHeight <> 0) Then tudHeight.DefaultValue = m_initHeight Else tudHeight.DefaultValue = g_Displays.GetDesktopHeight
            
        Case Else
            tudWidth.SigDigits = 2
            tudWidth.Min = 0.01
            If (newUnit = mu_Millimeters) Or (newUnit = mu_Points) Or (newUnit = mu_Picas) Then tudWidth.Max = CDbl(PD_MAX_IMAGE_DIMENSION * 10) Else tudWidth.Max = CDbl(PD_MAX_IMAGE_DIMENSION)
            If (m_initWidth <> 0) Then tudWidth.DefaultValue = ConvertPixelToOtherUnit(newUnit, m_initWidth, GetResolutionAsPPI(), m_initWidth) Else tudWidth.DefaultValue = ConvertPixelToOtherUnit(newUnit, g_Displays.GetDesktopWidth, GetResolutionAsPPI(), g_Displays.GetDesktopWidth)
            If (m_initHeight <> 0) Then tudHeight.DefaultValue = ConvertPixelToOtherUnit(newUnit, m_initHeight, GetResolutionAsPPI(), m_initHeight) Else tudHeight.DefaultValue = ConvertPixelToOtherUnit(newUnit, g_Displays.GetDesktopHeight, GetResolutionAsPPI(), g_Displays.GetDesktopHeight)
            
    End Select
    
    'As the height and width boxes will always match, simply mirror the tudBox limits
    tudHeight.Min = tudWidth.Min
    tudHeight.Max = tudWidth.Max
    tudHeight.SigDigits = tudWidth.SigDigits
    
    'Copy the new values to their respective text boxes
    tudWidth.Value = newWidth
    tudHeight.Value = newHeight
    
End Sub

Private Sub cmdAspectRatio_Click(ByVal Shift As ShiftConstants)
    SyncDimensions True
End Sub

'Before using this control, dialogs MUST call this function to notify the control of the initial width/height values
' they want to use.  We cannot do this automatically as some dialogs determine this by the current image's dimensions
' (e.g. resize) while others may do it when no images are loaded (e.g. batch process).
Public Sub SetInitialDimensions(ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal srcDPI As Double = 96#)

    'Store local copies
    m_initWidth = srcWidth
    m_initHeight = srcHeight
    m_initDPI = srcDPI
    
    'To prevent aspect ratio changes to one box resulting in recursion-type changes to the other, we only
    ' allow one box at a time to be updated.
    m_allowedToUpdateWidth = True
    m_allowedToUpdateHeight = True
    
    'Establish aspect ratios
    m_wRatio = m_initWidth / m_initHeight
    m_hRatio = m_initHeight / m_initWidth
    
    'Display the initial width/height
    m_unitSyncingSuspended = True
    tudWidth.Value = srcWidth
    tudHeight.Value = srcHeight
    tudResolution.Value = srcDPI
    m_unitSyncingSuspended = False
    
    'Set the "previous unit of measurement" to equal pixels, as that's always how we begin
    m_previousUnitOfMeasurement = mu_Pixels
    
    'Set the "previous unit of resolution" to equal PPI, as that is PD's default
    m_previousUnitOfResolution = ru_PPI
    cmbResolution.ListIndex = ru_PPI
    
    'If the user wants this control to default to "real world units", swap to inches or cm (as appropriate)
    If Me.DefaultToRealWorldUnits Then
        If Units.LocaleUsesMetric() Then
            UserControl.cmbWidthUnit.ListIndex = mu_Centimeters
            UserControl.cmbHeightUnit.ListIndex = mu_Centimeters
        Else
            UserControl.cmbWidthUnit.ListIndex = mu_Inches
            UserControl.cmbHeightUnit.ListIndex = mu_Inches
        End If
    End If
    
End Sub

'Changing the resolution text box is a bit different than changing the width/height ones.  This box never changes the current
' width and height.  The only thing it immediately changes is the pixel size label, and it only changes that if the current
' width/height unit is inches or cm.
Private Sub tudResolution_Change()
    If (tudResolution.IsValid(False)) Then UpdateAspectRatio
End Sub

Private Sub UserControl_Initialize()
    
    m_allowedToUpdateWidth = True
    m_allowedToUpdateHeight = True
    tudWidth.Max = PD_MAX_IMAGE_DIMENSION
    tudHeight.Max = PD_MAX_IMAGE_DIMENSION
    
    'Populate all dropdowns
    PopulateDropdowns
    
    'Default all interface elements to pixels
    ConvertUnitsToNewValue mu_Pixels, mu_Pixels
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, False
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDRESIZE_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDResize", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
End Sub

Private Sub UserControl_InitProperties()
    UnknownSizeMode = False
    DisablePercentOption = False
    DefaultToRealWorldUnits = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UnknownSizeMode = .ReadProperty("UnknownSizeMode", False)
        DisablePercentOption = .ReadProperty("DisablePercentOption", False)
        DefaultToRealWorldUnits = .ReadProperty("DefaultToRealWorldUnits", False)
    End With
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_Show()

    'Translate various bits of UI text at run-time
    If PDMain.IsProgramRunning() Then
        
        'Add tooltips to the controls that natively support them
        cmdAspectRatio.AssignTooltip "Preserve aspect ratio (sometimes called Constrain Proportions).  Use this option to resize an image while keeping the width and height in sync.", "Preserve aspect ratio"
        cmbWidthUnit.AssignTooltip "Change the unit of measurement used to resize the image."
        cmbHeightUnit.AssignTooltip "Change the unit of measurement used to resize the image."
        cmbResolution.AssignTooltip "Change the unit of measurement used for image resolution (pixel density)."
                
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "UnknownSizeMode", m_UnknownSizeMode, False
        .WriteProperty "DisablePercentOption", m_PercentDisabled, False
        .WriteProperty "DefaultToRealWorldUnits", m_DefaultToRealWorldUnits, False
    End With
End Sub

'If "Lock Image Aspect Ratio" is selected, these two routines keep all values in sync
Private Sub tudHeight_Change()
    If (Not m_unitSyncingSuspended) And (tudWidth.IsValid(False)) Then SyncDimensions False
End Sub

Private Sub tudWidth_Change()
    If (Not m_unitSyncingSuspended) And (tudWidth.IsValid(False)) Then SyncDimensions True
End Sub

Private Sub PopulateDropdowns()
    
    'Prevent unit syncing until the combo boxes have been populated
    m_unitSyncingSuspended = True
    
    'Populate the width unit drop-down box
    cmbWidthUnit.Clear
    
    If PDMain.IsProgramRunning() Then
        If Not m_PercentDisabled Then cmbWidthUnit.AddItem g_Language.TranslateMessage("percent"), 0
        cmbWidthUnit.AddItem g_Language.TranslateMessage("pixels")
        cmbWidthUnit.AddItem g_Language.TranslateMessage("inches")
        cmbWidthUnit.AddItem g_Language.TranslateMessage("centimeters")
        cmbWidthUnit.AddItem g_Language.TranslateMessage("millimeters")
        cmbWidthUnit.AddItem g_Language.TranslateMessage("points")
        cmbWidthUnit.AddItem g_Language.TranslateMessage("picas")
        If (Not m_PercentDisabled) Then cmbWidthUnit.ListIndex = mu_Pixels Else cmbWidthUnit.ListIndex = mu_Pixels - 1
    End If
    
    'Rather than manually populate the height unit box, just copy whatever entries we've set for the width box
    cmbHeightUnit.Clear
    
    Dim i As Long
    For i = 0 To cmbWidthUnit.ListCount - 1
        cmbHeightUnit.AddItem cmbWidthUnit.List(i), i
    Next i
    
    cmbHeightUnit.ListIndex = cmbWidthUnit.ListIndex
    
    'Populate the resolution unit box
    cmbResolution.Clear
    
    If PDMain.IsProgramRunning() Then
        cmbResolution.AddItem g_Language.TranslateMessage("pixels / inch (PPI)"), 0
        cmbResolution.AddItem g_Language.TranslateMessage("pixels / centimeter (PPCM)"), 1
        cmbResolution.ListIndex = ru_PPI
    End If
    
    'Restore automatic unit syncing
    m_unitSyncingSuspended = False

End Sub

'When one dimension is updated, call this to synchronize the other (as necessary) and/or the aspect ratio
Private Sub SyncDimensions(ByVal useWidthAsSource As Boolean)

    'Suspend automatic text box syncing as necessary
    If useWidthAsSource Then m_allowedToUpdateWidth = False Else m_allowedToUpdateHeight = False
    
    'Cache a "preserve aspect ratio" value, which individual functions can use as necessary
    Dim preserveAspectRatio As Boolean
    preserveAspectRatio = cmdAspectRatio.Value
    
    'Because the resize dialog provides units other than pixels (e.g. "percent"), we always provide
    ' a width/height pixel equivalent, in case subsequent conversion functions needs it.
    Dim imgWidthPixels As Double, imgHeightPixels As Double
    imgWidthPixels = ConvertUnitToPixels(GetCurrentWidthUnit, tudWidth, GetResolutionAsPPI(), m_initWidth)
    imgHeightPixels = ConvertUnitToPixels(GetCurrentHeightUnit, tudHeight, GetResolutionAsPPI(), m_initHeight)
    
    'Synchronization is divided into two possible code paths: synchronizing height to match width, and width to match height.
    ' These could technically be merged down to a single path, but I find it more intuitive to handle them separately (despite
    ' the redundant code necessary to do so).
    If useWidthAsSource Then
        
        'When changing width, do not also update height unless "preserve aspect ratio" is checked
        If cmdAspectRatio.Value And m_allowedToUpdateHeight Then
            
            'The HEIGHT text value needs to be synched to the WIDTH text value.  How we do this depends on the current resize unit.
            Select Case GetCurrentWidthUnit
            
                'Percent
                Case mu_Percent
                    tudHeight = tudWidth
                
                'Anything else
                Case Else
                
                    'For all other conversions, we simply want to calculate an aspect-ratio preserved height value (in pixels),
                    ' which we can then use to populate the height up/down box.
                    tudHeight = ConvertPixelToOtherUnit(GetCurrentWidthUnit, Int((imgWidthPixels * m_hRatio) + 0.5), GetResolutionAsPPI(), m_initWidth)
            
            End Select
            
        End If
        
    Else
        
        'When changing height, do not also update width unless "preserve aspect ratio" is checked
        If cmdAspectRatio.Value And m_allowedToUpdateWidth Then
        
            'The WIDTH text value needs to be synched to the HEIGHT text value.  How we do this depends on the current resize unit.
            Select Case GetCurrentHeightUnit
        
                'Percent
                Case mu_Percent
                    tudWidth = tudHeight
                
                'Anything else
                Case Else
                    tudWidth = ConvertPixelToOtherUnit(GetCurrentHeightUnit, Int((imgHeightPixels * m_wRatio) + 0.5), GetResolutionAsPPI(), m_initWidth)
                
            End Select
            
        End If
        
    End If
    
    'Re-enable automatic text box syncing
    If useWidthAsSource Then m_allowedToUpdateWidth = True Else m_allowedToUpdateHeight = True
    
    'Display a relevant aspect ratio for the current
    UpdateAspectRatio
    
    'Update our image width/height in pixel values, so we can raise them as part of the control's Change event
    imgWidthPixels = ConvertUnitToPixels(GetCurrentWidthUnit, tudWidth, GetResolutionAsPPI(), m_initWidth)
    imgHeightPixels = ConvertUnitToPixels(GetCurrentHeightUnit, tudHeight, GetResolutionAsPPI(), m_initHeight)
    
    'Minimize events by auto-detecting when values haven't changed
    If (imgWidthPixels <> m_LastImgWidthPixels) Or (imgHeightPixels <> m_LastImgHeightPixels) Then
        m_LastImgWidthPixels = imgWidthPixels
        m_LastImgHeightPixels = imgHeightPixels
        RaiseEvent Change(imgWidthPixels, imgHeightPixels)
    End If

End Sub

'This control displays an approximate aspect ratio for the selected dimensions.  This can be helpful when
' trying to select new width/height values for a specific application with a set aspect ratio (e.g. 16:9 screens).
Private Sub UpdateAspectRatio()

    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    If tudWidth.IsValid(False) And tudHeight.IsValid(False) Then
    
        'Retrieve width and height values in pixel amounts
        Dim imgWidthPixels As Double, imgHeightPixels As Double
        imgWidthPixels = ConvertUnitToPixels(GetCurrentWidthUnit, tudWidth, GetResolutionAsPPI(), m_initWidth)
        imgHeightPixels = ConvertUnitToPixels(GetCurrentHeightUnit, tudHeight, GetResolutionAsPPI(), m_initHeight)
        
        'Convert the floating-point aspect ratio to a fraction
        Dim numerator As Long, denominator As Long
        If (imgHeightPixels > 0) Then PDMath.ConvertToFraction imgWidthPixels / imgHeightPixels, numerator, denominator
        
        'Aspect ratios are typically given in terms of base 10 if possible, so change values like 8:5 to 16:10
        If (denominator = 5) Then
            numerator = numerator * 2
            denominator = denominator * 2
        End If
        
        'In "unknown size mode", we can't display exact dimensions for PERCENT mode, so don't even try.
        If m_UnknownSizeMode And (GetCurrentWidthUnit = mu_Percent) Then
            lblAspectRatio(1).Caption = " " & g_Language.TranslateMessage("exact aspect ratio will vary by image")
            lblDimensions(1).Caption = " " & g_Language.TranslateMessage("exact size will vary by image")
        
        Else
        
            If (imgHeightPixels > 0) Then
                lblAspectRatio(1).Caption = " " & numerator & ":" & denominator & "  (" & Format$(imgWidthPixels / imgHeightPixels, "#0.00#") & ")"
            End If
            
            'While we're here, also update the dimensions caption.  Note that currency is used to cover the
            ' (rare) case of an overflow from resizing an image to larger than 2GB.
            Dim dmString As pdString
            Set dmString = New pdString
            
            'Append size, in pixels
            dmString.Append " " & Int(imgWidthPixels) & " px   X   " & Int(imgHeightPixels) & " px"
            dmString.Append "   ("
            
            'Calculate current and proposed sizes.  (If they match, we won't display both values - only the
            ' current one.)
            Dim curSize As Currency, newSize As Currency
            curSize = (CCur(m_initWidth) * CCur(m_initHeight) * 4@) / 10000@
            newSize = (CCur(imgWidthPixels) * CCur(imgHeightPixels) * 4@) / 10000@
            If (curSize = 0) Or (newSize = 0) Or (curSize = newSize) Then
                dmString.Append CStr(Files.GetFormattedFileSizeL(newSize))
            Else
                dmString.Append g_Language.TranslateMessage("%1, was %2", Files.GetFormattedFileSizeL(newSize), Files.GetFormattedFileSizeL(curSize))
            End If
            dmString.Append ")"
            lblDimensions(1).Caption = dmString.ToString()
            
        End If
    
    Else
        lblAspectRatio(1).Caption = vbNullString
        lblDimensions(1).Caption = vbNullString
    End If

End Sub

'This function is just a thin wrapper to the public convertOtherUnitToPixels function.  The only difference is that this function requests
' a reference to the actual pdSpinner control requesting conversion, and it will automatically validate that control as necessary.
Private Function ConvertUnitToPixels(ByVal curUnit As PD_MeasurementUnit, ByRef tudSource As pdSpinner, Optional ByVal srcUnitResolution As Double, Optional ByVal initPixelValue As Double) As Double

    If tudSource.IsValid(False) Then
        ConvertUnitToPixels = ConvertOtherUnitToPixels(curUnit, tudSource.Value, srcUnitResolution, initPixelValue)
    Else
        ConvertUnitToPixels = initPixelValue
    End If

End Function

'Instead of directly accessing the tudResolution box, use this function.  It will validate invalid input, and also convert PPCM to PPI
' if necessary.  (Note that all conversion functions in PD require resolution as PPI.)
Private Function GetResolutionAsPPI() As Double
    
    If tudResolution.IsValid(False) Then
    
        'cmbResolution only has two entries: inches (0), and cm (1).
        If (cmbResolution.ListIndex = ru_PPI) Then
            GetResolutionAsPPI = tudResolution.Value
        Else
            GetResolutionAsPPI = GetInchesFromCM(tudResolution)
        End If
    
    Else
        GetResolutionAsPPI = m_initDPI
    End If
    
End Function

'Percent mode may be disabled on some controls (e.g. batch processing).  To ensure proper control behavior, these wrappers
' should be used rather than accessing cmbWidth/HeightUnit.ListIndex directly.
Private Function GetCurrentWidthUnit() As Long
    If m_PercentDisabled Then GetCurrentWidthUnit = cmbWidthUnit.ListIndex + 1 Else GetCurrentWidthUnit = cmbWidthUnit.ListIndex
End Function

Private Function GetCurrentHeightUnit() As Long
    If m_PercentDisabled Then GetCurrentHeightUnit = cmbHeightUnit.ListIndex + 1 Else GetCurrentHeightUnit = cmbHeightUnit.ListIndex
End Function

'Returns all control settings in a single XML string; this is basically a serialized copy of the current
' control instance, and these settings can be restored via SetAllSettingsFromXML(), below.
Public Function GetCurrentSettingsAsXML() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    With cParams
    
        'We do store the "lock aspect ratio" setting, but note that this setting is deliberately
        ' *not used* by SetAllSettingsFromXML(), below.
        .AddParam "lockaspectratio", Me.AspectRatioLock
    
        .AddParam "sizeunit", Me.UnitOfMeasurement
        .AddParam "dpiunit", Me.UnitOfResolution
    
        .AddParam "dpi", Me.ResizeDPI
        .AddParam "width", Me.ResizeWidth
        .AddParam "height", Me.ResizeHeight
    
    End With
    
    GetCurrentSettingsAsXML = cParams.GetParamString()

End Function

Public Sub SetAllSettingsFromXML(ByVal xmlData As String)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString xmlData
    
    'Kind of funny, but we must always set the .AspectRatioLock property to FALSE in order to
    ' apply a new size to the image.  (If we don't do this, the new sizes will be clamped to
    ' the current image's aspect ratio!)
    Me.AspectRatioLock = False
    
    With cParams
        Me.UnitOfMeasurement = .GetLong("sizeunit", mu_Pixels)
        Me.UnitOfResolution = .GetLong("dpiunit", ru_PPI)
        Me.ResizeDPI = .GetDouble("dpi", 96#)
        Me.ResizeWidth = .GetDouble("width", 1920)
        Me.ResizeHeight = .GetDouble("height", 1080)
    End With
    
End Sub


'Whenever a control property changes that affects control size or layout (including internal changes, like caption adjustments),
' call this function to recalculate the control's internal layout
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'At some point, it would be *really* nice to make this control resizable (while still being center-aligned),
    ' but that's a fairly messy project, and it doesn't have immediate usefulness since PD needs some other work
    ' before resizable dialogs are feasible.
    
    'As such, just repaint the control for now
    RedrawBackBuffer
    
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDR_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDR_Background, "Background", IDE_WHITE
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    If ucSupport.ThemeUpdateRequired Then
        
        'Add the "lock aspect ratio" button
        If PDMain.IsProgramRunning() Then
            Dim buttonImageSize As Long
            buttonImageSize = Interface.FixDPI(32)
            cmdAspectRatio.AssignImage "generic_unlock", , buttonImageSize, buttonImageSize
            cmdAspectRatio.AssignImage_Pressed "generic_lock", , buttonImageSize, buttonImageSize
        End If
        
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd, False
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        
        'Manually update all sub-controls
        lblWidth.UpdateAgainstCurrentTheme
        lblHeight.UpdateAgainstCurrentTheme
        lblResolution.UpdateAgainstCurrentTheme
        
        lblDimensions(0).UpdateAgainstCurrentTheme
        lblDimensions(1).UpdateAgainstCurrentTheme
        lblAspectRatio(0).UpdateAgainstCurrentTheme
        lblAspectRatio(1).UpdateAgainstCurrentTheme
        
        tudWidth.UpdateAgainstCurrentTheme
        tudHeight.UpdateAgainstCurrentTheme
        tudResolution.UpdateAgainstCurrentTheme
        
        cmdAspectRatio.UpdateAgainstCurrentTheme
        
        cmbWidthUnit.UpdateAgainstCurrentTheme
        cmbHeightUnit.UpdateAgainstCurrentTheme
        cmbResolution.UpdateAgainstCurrentTheme
        
    End If
    
End Sub
