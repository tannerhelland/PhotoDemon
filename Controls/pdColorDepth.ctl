VERSION 5.00
Begin VB.UserControl pdColorDepth 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
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
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   ToolboxBitmap   =   "pdColorDepth.ctx":0000
   Begin PhotoDemon.pdColorSelector clsComposite 
      Height          =   735
      Left            =   3600
      TabIndex        =   7
      Top             =   1680
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      Caption         =   "compositing color"
   End
   Begin PhotoDemon.pdDropDown cboDepthGrayscale 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1296
      Caption         =   "depth"
   End
   Begin PhotoDemon.pdDropDown cboDepthColor 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1296
      Caption         =   "depth"
   End
   Begin PhotoDemon.pdDropDown cboAlphaModel 
      Height          =   735
      Left            =   3600
      TabIndex        =   4
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      Caption         =   "transparency format"
   End
   Begin PhotoDemon.pdDropDown cboColorModel 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1296
      Caption         =   "color format"
   End
   Begin PhotoDemon.pdSlider sldColorCount 
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   1680
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1296
      Caption         =   "palette size"
      Min             =   2
      Max             =   256
      Value           =   256
      NotchPosition   =   2
      NotchValueCustom=   256
   End
   Begin PhotoDemon.pdColorSelector clsAlphaColor 
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      Caption         =   "transparent color"
      curColor        =   16711935
   End
   Begin PhotoDemon.pdSlider sldAlphaCutoff 
      Height          =   735
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      Caption         =   "transparency cut-off"
      Max             =   254
      SliderTrackStyle=   1
      Value           =   64
      GradientColorRight=   1703935
      NotchPosition   =   2
      NotchValueCustom=   64
   End
End
Attribute VB_Name = "pdColorDepth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Color/Transparency depth selector User Control
'Copyright 2016-2026 by Tanner Helland
'Created: 22/April/16
'Last updated: 24/March/20
'Last update: improve UI reflow behavior when control is not yet visible (e.g. first load)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Event Change()

'For the "set alpha by color" option, the user should be allowed to click colors straight
' from the preview area.  However, if "set alpha by color" is *not* set, color selection
' should be left to the parent dialog.
Public Event ColorSelectionRequired(ByVal selectState As Boolean)

'If a format supports "use original file settings", it can specify that via this event
' (raised when the control is loaded, if not in design mode)
Public Event AreOriginalSettingsAllowed(ByRef newSetting As Boolean)

'Because this control dynamically shows/hides subcontrols, its total height can vary.  Parent controls can
' use this to reflow other controls, as necessary
Public Event SizeChanged()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'After reflowing controls, we store the final calculated "ideal" size of the control.  Our parent can ask us to
' sync to this size (although some may not care, and will ignore this).
Private m_IdealControlHeight As Long

'To ensure that we can reflow layouts successfully even with the control hidden,
' we mirror some visibility settings to local variables.
Private m_ColorCountVisible As Boolean, m_DepthColorVisible As Boolean, m_DepthGrayscaleVisible As Boolean
Private m_AlphaModelVisible As Boolean, m_AlphaCutoffVisible As Boolean, m_AlphaColorVisible As Boolean
Private m_CompositeColorVisible As Boolean

'If the calling dialog allows us to provide a "use original settings" option, this flag
' will be set by the IsOriginalSettingsAllowed event.
Private m_UseOriginalAllowed As Boolean

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_ColorDepth
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function

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
    
    'If a given text value is not valid, highlight the problem and optionally display an error message box
    If (Not sldAlphaCutoff.IsValid(showErrors)) Then
        IsValid = False
        Exit Function
    End If
    
    If (Not sldColorCount.IsValid(showErrors)) Then
        IsValid = False
        Exit Function
    End If
    
End Function

'For the "set alpha by color" option, external callers can notify us of color changes via this sub.
Public Sub NotifyNewAlphaColor(ByVal newColor As Long)
    clsAlphaColor.Color = newColor
    RaiseEvent Change
End Sub

Public Function GetIdealSize() As Long
    If (m_IdealControlHeight <= 0) Then ReflowColorPanel
    GetIdealSize = m_IdealControlHeight
End Function

Public Sub SyncToIdealSize()
    If (m_IdealControlHeight <= 0) Then ReflowColorPanel
    ucSupport.RequestNewSize ucSupport.GetBackBufferWidth, m_IdealControlHeight
End Sub

Public Sub Reset()

    cboColorModel.ListIndex = 0
    cboDepthColor.ListIndex = 1
    cboDepthGrayscale.ListIndex = 1
    cboAlphaModel.ListIndex = 0
    
    sldColorCount.Value = 256
    sldAlphaCutoff.Value = PD_DEFAULT_ALPHA_CUTOFF
    clsAlphaColor.Color = RGB(255, 0, 255)
    
End Sub

'Get/Set all control settings as a single XML packet
Public Function GetAllSettings() As String

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    'All entries use text to make it easier to pass changed/upgraded settings in the future
    Dim outputColorModel As String
    Select Case cboColorModel.ListIndex
        Case 0
            outputColorModel = "auto"
        Case 1
            outputColorModel = "color"
        Case 2
            outputColorModel = "gray"
        Case 3
            outputColorModel = "original"
    End Select
    cParams.AddParam "cd-color-model", outputColorModel
    
    'Which color depth we write is contingent on the color model, as color and gray use different button strips.
    ' (Gray supports some depths that color does not, e.g. 1-bit and 4-bit.)
    Dim colorColorDepth As String, grayColorDepth As String, outputPaletteSize As String
    
    Select Case cboDepthColor.ListIndex
        Case 0
            colorColorDepth = "color-hdr"
        Case 1
            colorColorDepth = "color-standard"
        Case 2
            colorColorDepth = "color-indexed"
    End Select
    
    cParams.AddParam "cd-color-depth", colorColorDepth
        
    Select Case cboDepthGrayscale.ListIndex
        Case 0
            grayColorDepth = "gray-hdr"
        Case 1
            grayColorDepth = "gray-standard"
        Case 2
            grayColorDepth = "gray-monochrome"
    End Select
    
    cParams.AddParam "cd-gray-depth", grayColorDepth
    
    If sldColorCount.IsValid Then outputPaletteSize = CStr(sldColorCount.Value) Else outputPaletteSize = "256"
    cParams.AddParam "cd-palette-size", outputPaletteSize
    
    'Next, we've got a bunch of possible alpha modes to deal with (uuuuuugh)
    Dim outputAlphaModel As String
    Select Case cboAlphaModel.ListIndex
        Case 0
            outputAlphaModel = "auto"
        Case 1
            outputAlphaModel = "full"
        Case 2
            outputAlphaModel = "by-cutoff"
        Case 3
            outputAlphaModel = "by-color"
        Case 4
            outputAlphaModel = "none"
    End Select
    
    cParams.AddParam "cd-alpha-model", outputAlphaModel
    If sldAlphaCutoff.IsValid Then cParams.AddParam "cd-alpha-cutoff", sldAlphaCutoff.Value Else cParams.AddParam "cd-alpha-cutoff", PD_DEFAULT_ALPHA_CUTOFF
    cParams.AddParam "cd-alpha-color", clsAlphaColor.Color
    cParams.AddParam "cd-matte-color", clsComposite.Color
    
    GetAllSettings = cParams.GetParamString()
    
End Function

'Used by the "last-used settings" manager to reset all settings to the user's previous value(s)
Public Sub SetAllSettings(ByVal newSettings As String)

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString newSettings
    
    Dim srcParam As String
    srcParam = cParams.GetString("cd-color-model", "auto")
    
    If ParamsEqual(srcParam, "color") Then
        cboColorModel.ListIndex = 1
    ElseIf ParamsEqual(srcParam, "gray") Then
        cboColorModel.ListIndex = 2
    ElseIf ParamsEqual(srcParam, "original") Then
        cboColorModel.ListIndex = 3
    Else
        cboColorModel.ListIndex = 0
    End If
    
    srcParam = cParams.GetString("cd-color-depth", "color-standard")
    
    If ParamsEqual(srcParam, "color") Then
        cboColorModel.ListIndex = 1
    ElseIf ParamsEqual(srcParam, "gray") Then
        cboColorModel.ListIndex = 2
    Else
        cboColorModel.ListIndex = 0
    End If
    
    srcParam = cParams.GetString("cd-color-depth", "color-standard")
    
    If ParamsEqual(srcParam, "color-hdr") Then
        cboDepthColor.ListIndex = 0
    ElseIf ParamsEqual(srcParam, "color-indexed") Then
        cboDepthColor.ListIndex = 2
    Else
        cboDepthColor.ListIndex = 1
    End If
    
    srcParam = cParams.GetString("cd-gray-depth", "gray-standard")
    
    If ParamsEqual(srcParam, "gray-hdr") Then
        cboDepthGrayscale.ListIndex = 0
    ElseIf ParamsEqual(srcParam, "gray-monochrome") Then
        cboDepthGrayscale.ListIndex = 2
    Else
        cboDepthGrayscale.ListIndex = 1
    End If
    
    sldColorCount.Value = cParams.GetLong("cd-palette-size", 256)
    
    srcParam = cParams.GetString("cd-alpha-model", "auto")
    
    If ParamsEqual(srcParam, "full") Then
        cboAlphaModel.ListIndex = 1
    ElseIf ParamsEqual(srcParam, "by-cutoff") Then
        cboAlphaModel.ListIndex = 2
    ElseIf ParamsEqual(srcParam, "by-color") Then
        cboAlphaModel.ListIndex = 3
    ElseIf ParamsEqual(srcParam, "none") Then
        cboAlphaModel.ListIndex = 4
    Else
        cboAlphaModel.ListIndex = 0
    End If
    
    sldAlphaCutoff.Value = cParams.GetLong("cd-alpha-cutoff")
    clsAlphaColor.Color = cParams.GetLong("cd-alpha-color")
    clsComposite.Color = cParams.GetLong("cd-matte-color", vbWhite)
    
End Sub

Private Function ParamsEqual(ByRef param1 As String, ByRef param2 As String) As Boolean
    ParamsEqual = Strings.StringsEqual(param1, param2, True)
End Function

'To support high-DPI settings properly, we expose specialized move+size functions
Public Function GetLeft() As Long
    GetLeft = ucSupport.GetControlLeft
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    ucSupport.RequestNewPosition newLeft, alsoNotifyMeViaEvent:=True
End Sub

Public Function GetTop() As Long
    GetTop = ucSupport.GetControlTop
End Function

Public Sub SetTop(ByVal newTop As Long)
    ucSupport.RequestNewPosition newTop:=newTop, alsoNotifyMeViaEvent:=True
End Sub

Public Function GetWidth() As Long
    GetWidth = ucSupport.GetControlWidth
End Function

Public Sub SetWidth(ByVal newWidth As Long)
    ucSupport.RequestNewSize newWidth, alsoNotifyMeViaEvent:=True
End Sub

Public Function GetHeight() As Long
    GetHeight = ucSupport.GetControlHeight
End Function

Public Sub SetHeight(ByVal newHeight As Long)
    ucSupport.RequestNewSize newHeight:=newHeight, alsoNotifyMeViaEvent:=True
End Sub

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

'Only some file formats support a "use original file settings" option; it's up to the caller to notify us
' if they want that option available.
Public Sub SetOriginalSettingsAvailable(ByVal newSetting As Boolean)

    m_UseOriginalAllowed = newSetting
    If m_UseOriginalAllowed And (cboColorModel.ListCount < 4) Then
        cboColorModel.AddItem "original file settings", 3
        UpdateColorDepthVisibility
    End If

End Sub

Private Sub cboAlphaModel_Click()
    UpdateTransparencyOptions
    RaiseEvent Change
End Sub

Private Sub cboColorModel_Click()
    UpdateColorDepthVisibility
    RaiseEvent Change
End Sub

Private Sub cboDepthColor_Click()
    UpdateColorDepthOptions
    RaiseEvent Change
End Sub

Private Sub cboDepthGrayscale_Click()
    UpdateColorDepthOptions
    RaiseEvent Change
End Sub

Private Sub clsAlphaColor_ColorChanged()
    RaiseEvent Change
End Sub

Private Sub clsComposite_ColorChanged()
    RaiseEvent Change
End Sub

Private Sub sldAlphaCutoff_Change()
    RaiseEvent Change
End Sub

Private Sub sldColorCount_Change()
    RaiseEvent Change
End Sub

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, False
    
    'Color model and color depth are closely related; populate all button strips, then show/hide the relevant pairings
    cboColorModel.AddItem "auto", 0
    cboColorModel.AddItem "color", 1
    cboColorModel.AddItem "grayscale", 2
    cboColorModel.ListIndex = 0
    
    cboDepthColor.AddItem "HDR", 0
    cboDepthColor.AddItem "standard", 1
    cboDepthColor.AddItem "indexed", 2
    cboDepthColor.ListIndex = 1
    
    cboDepthGrayscale.AddItem "HDR", 0
    cboDepthGrayscale.AddItem "standard", 1
    cboDepthGrayscale.AddItem "monochrome", 2
    cboDepthGrayscale.ListIndex = 1
    
    'PNGs also support a (ridiculous) amount of alpha settings
    cboAlphaModel.AddItem "auto", 0
    cboAlphaModel.AddItem "full", 1
    cboAlphaModel.AddItem "binary (by cut-off)", 2
    cboAlphaModel.AddItem "binary (by color)", 3
    cboAlphaModel.AddItem "none", 4
    cboAlphaModel.ListIndex = 0
    
    sldAlphaCutoff.NotchValueCustom = PD_DEFAULT_ALPHA_CUTOFF
    
    UpdateColorDepthVisibility
    UpdateTransparencyOptions
    ReflowColorPanel
    
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub

Private Sub UpdateColorDepthVisibility()

    Select Case cboColorModel.ListIndex
    
        'Auto
        Case 0
            m_DepthColorVisible = False
            m_DepthGrayscaleVisible = False
        
        'Color
        Case 1
            m_DepthColorVisible = True
            m_DepthGrayscaleVisible = False
        
        'Grayscale
        Case 2
            m_DepthColorVisible = False
            m_DepthGrayscaleVisible = True
            
        'Original file
        Case 3
            m_DepthColorVisible = False
            m_DepthGrayscaleVisible = False
    
    End Select
    
    cboDepthColor.Visible = m_DepthColorVisible
    cboDepthGrayscale.Visible = m_DepthGrayscaleVisible
    
    UpdateColorDepthOptions

End Sub

Private Sub UpdateColorDepthOptions()
    
    'Indexed color modes allow for variable palette sizes
    If (cboColorModel.ListIndex = 1) Then
        m_ColorCountVisible = (cboDepthColor.ListIndex = 2)
        
    'Indexed grayscale mode also allows for variable palette sizes
    ElseIf (cboColorModel.ListIndex = 2) Then
        m_ColorCountVisible = (cboDepthGrayscale.ListIndex = 1)
        
    'Other modes do not expose palette settings
    Else
        m_ColorCountVisible = False
    End If
    
    sldColorCount.Visible = m_ColorCountVisible
    
    'Alpha options are hidden when the "use original file settings" option is used
    If (cboColorModel.ListIndex = 3) Then
        m_AlphaModelVisible = False
        m_AlphaCutoffVisible = False
        m_AlphaColorVisible = False
        m_CompositeColorVisible = False
        RaiseEvent ColorSelectionRequired(False)
    Else
        m_AlphaModelVisible = True
        UpdateTransparencyOptions
    End If
    
    cboAlphaModel.Visible = m_AlphaModelVisible
    
    ReflowColorPanel
    
End Sub

Private Sub UpdateTransparencyOptions()
    
    Select Case cboAlphaModel.ListIndex
    
        'auto, full alpha
        Case 0, 1
            m_AlphaCutoffVisible = False
            m_AlphaColorVisible = False
            m_CompositeColorVisible = False
            If PDMain.IsProgramRunning() Then RaiseEvent ColorSelectionRequired(False)
            
        'alpha by cut-off
        Case 2
            m_AlphaCutoffVisible = True
            m_AlphaColorVisible = False
            m_CompositeColorVisible = True
            If PDMain.IsProgramRunning() Then RaiseEvent ColorSelectionRequired(False)
        
        'alpha by color
        Case 3
            m_AlphaCutoffVisible = False
            m_AlphaColorVisible = True
            m_CompositeColorVisible = True
            If PDMain.IsProgramRunning() Then RaiseEvent ColorSelectionRequired(True)
            
        'no alpha
        Case 4
            m_AlphaCutoffVisible = False
            m_AlphaColorVisible = False
            m_CompositeColorVisible = True
            If PDMain.IsProgramRunning() Then RaiseEvent ColorSelectionRequired(False)
    
    End Select
    
    sldAlphaCutoff.Visible = m_AlphaCutoffVisible
    clsAlphaColor.Visible = m_AlphaColorVisible
    clsComposite.Visible = m_CompositeColorVisible
    
    ReflowColorPanel
    
End Sub

'Reflow various controls contingent on their visibility.  Note that this function *does not* show or hide controls;
' instead, it relies on functions like UpdateColorDepthVisibility() to do that in advance.
Private Sub ReflowColorPanel()

    Dim curHeight As Long
    curHeight = ucSupport.GetBackBufferHeight
    
    Dim yPadding As Long, yOffset As Long
    yPadding = Interface.FixDPI(8)
    yOffset = cboColorModel.GetTop + cboColorModel.GetHeight + yPadding
    
    If m_DepthColorVisible Then
        cboDepthColor.SetTop yOffset
        yOffset = yOffset + cboDepthColor.GetHeight + yPadding
    ElseIf m_DepthGrayscaleVisible Then
        cboDepthGrayscale.SetTop yOffset
        yOffset = yOffset + cboDepthGrayscale.GetHeight + yPadding
    End If
    
    If m_ColorCountVisible Then
        sldColorCount.SetTop yOffset
        yOffset = yOffset + sldColorCount.GetHeight + yPadding
    End If
    
    Dim maxHeight As Long
    maxHeight = yOffset
    
    'Now restart at the top, and perform the same steps for the "alpha settings column" of controls
    yOffset = cboAlphaModel.GetTop + cboAlphaModel.GetHeight + yPadding
    
    If m_AlphaCutoffVisible Then
        sldAlphaCutoff.SetTop yOffset
        yOffset = yOffset + sldAlphaCutoff.GetHeight + yPadding
    ElseIf m_AlphaColorVisible Then
        clsAlphaColor.SetTop yOffset
        yOffset = yOffset + clsAlphaColor.GetHeight + yPadding
    End If
    
    If m_CompositeColorVisible Then
        If m_ColorCountVisible And (cboAlphaModel.ListIndex <> 4) Then yOffset = sldColorCount.GetTop
        clsComposite.SetTop yOffset
        yOffset = yOffset + clsComposite.GetHeight + yPadding
    End If
    
    If (yOffset > maxHeight) Then m_IdealControlHeight = yOffset Else m_IdealControlHeight = maxHeight
    RaiseEvent SizeChanged
    
End Sub

'Whenever a control property changes that affects control size or layout (including internal changes, like caption adjustments),
' call this function to recalculate the control's internal layout
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'A lot of controls on this dialog sit "side-by-side", so set their width proportionally.
    Dim halfWidth As Long, uiItemWidth As Long, pxPadding As Long
    halfWidth = bWidth \ 2
    pxPadding = Interface.FixDPI(4)
    uiItemWidth = halfWidth - pxPadding * 2
    
    'Sync all widths to match the current buffer width
    cboColorModel.SetPositionAndSize pxPadding, cboColorModel.GetTop, uiItemWidth, cboColorModel.GetHeight
    cboAlphaModel.SetPositionAndSize halfWidth + pxPadding, cboAlphaModel.GetTop, uiItemWidth, cboAlphaModel.GetHeight
    
    cboDepthColor.SetPositionAndSize pxPadding, cboDepthColor.GetTop, uiItemWidth, cboDepthColor.GetHeight
    cboDepthGrayscale.SetPositionAndSize pxPadding, cboDepthGrayscale.GetTop, uiItemWidth, cboDepthGrayscale.GetHeight
    sldColorCount.SetPositionAndSize pxPadding, sldColorCount.GetTop, uiItemWidth, sldColorCount.GetHeight
    
    sldAlphaCutoff.SetPositionAndSize halfWidth + pxPadding, sldAlphaCutoff.GetTop, uiItemWidth, sldAlphaCutoff.GetHeight
    clsAlphaColor.SetPositionAndSize halfWidth + pxPadding, clsAlphaColor.GetTop, uiItemWidth, cboDepthColor.GetHeight
    clsComposite.SetPositionAndSize halfWidth + pxPadding, clsComposite.GetTop, uiItemWidth, cboDepthColor.GetHeight
    
    'As such, just repaint the control for now
    RedrawBackBuffer
    
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    If PDMain.IsProgramRunning() Then
        Dim bufferDC As Long
        bufferDC = ucSupport.GetBackBufferDC(True, g_Themer.GetGenericUIColor(UI_Background))
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    If ucSupport.ThemeUpdateRequired Then
        
        'UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd, False
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        
        'Manually update all sub-controls
        cboColorModel.UpdateAgainstCurrentTheme
        cboDepthColor.UpdateAgainstCurrentTheme
        cboDepthGrayscale.UpdateAgainstCurrentTheme
        sldColorCount.UpdateAgainstCurrentTheme
        
        cboAlphaModel.UpdateAgainstCurrentTheme
        clsAlphaColor.UpdateAgainstCurrentTheme
        sldAlphaCutoff.UpdateAgainstCurrentTheme
        clsComposite.UpdateAgainstCurrentTheme
        
    End If
    
End Sub
