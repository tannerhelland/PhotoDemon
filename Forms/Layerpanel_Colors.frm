VERSION 5.00
Begin VB.Form layerpanel_Colors 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
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
   Icon            =   "Layerpanel_Colors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   190
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdPaletteUI palSelector 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdButton cmdSettings 
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      RenderMode      =   1
   End
   Begin PhotoDemon.pdHistory clrHistory 
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   344
   End
   Begin PhotoDemon.pdColorVariants clrVariants 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
   End
   Begin PhotoDemon.pdColorWheel clrWheel 
      Height          =   975
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      WheelWidth      =   13
   End
End
Attribute VB_Name = "layerpanel_Colors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Color Selector Tool Panel
'Copyright 2015-2026 by Tanner Helland
'Created: 15/October/15
'Last updated: 22/February/18
'Last update: when a color is selected from the image, the palette control will now select the
'             *nearest* palette index to the new color
'
'As part of the 7.0 release, PD's right-side panel gained a lot of new functionality.  To simplify the code for
' the new panel, each chunk of related settings (e.g. layer, nav, color selector) was moved to its own subpanel.
'
'This form is the subpanel for the color selector panel.  It is currently under construction.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class.
' (Normally this is declared WithEvents, but this dialog doesn't require custom settings behavior.)
Private m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

'To avoid nested resize calls, trackers are used
Private m_ResizeInProgress As Boolean

'As of Feb 2018, this panel now supports multiple rendering modes.  We can shortcut certain actions if
' certain panels are not visible, so it's important to track this.
Public Enum PD_ColorPanelMode
    cpm_Wheels = 0
    cpm_Palette = 1
End Enum

#If False Then
    Private Const cpm_Wheels = 0, cpm_Palette = 1
#End If

Private m_RenderMode As PD_ColorPanelMode

'In the "palette" rendering mode, this will be set to a non-null value
Private m_PaletteFile As String

'When various paint tools are used on the main window, they will notify us (via window message) of what
' color was used.  We will add those colors to our history list.
Private Sub clrHistory_CustomWindowMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    If (wMsg = WM_PD_PRIMARY_COLOR_APPLIED) Then clrHistory.PushNewHistoryItem CStr(wParam), , True
End Sub

Private Sub clrHistory_DrawHistoryItem(ByVal histIndex As Long, ByVal histValue As String, ByVal targetDC As Long, ByVal ptrToRectF As Long)
    
    If (LenB(histValue) <> 0) And PDMain.IsProgramRunning() And (targetDC <> 0) Then
        
        If PDMain.IsProgramRunning Then
        
            If (ptrToRectF <> 0) Then
                
                Dim tmpRectF As RectF
                CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, LenB(tmpRectF)
                
                'Note that this control *is* color-managed
                Dim cmResult As Long
                ColorManagement.ApplyDisplayColorManagement_SingleColor CLng(histValue), cmResult
            
                Dim cSurface As pd2DSurface: Dim cBrush As pd2DBrush
                Drawing2D.QuickCreateSurfaceFromDC cSurface, targetDC
                Drawing2D.QuickCreateSolidBrush cBrush, cmResult
                PD2D.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
                
                Set cSurface = Nothing: Set cBrush = Nothing
            
            End If
            
        End If
        
    End If
    
End Sub

Private Sub clrHistory_HistoryDoesntExist(ByVal histIndex As Long, histValue As String)

    Dim newColor As Long
    
    Select Case histIndex
    
        Case 0
            newColor = RGB(0, 0, 0)
        Case 1
            newColor = RGB(34, 32, 52)
        Case 2
            newColor = RGB(69, 40, 60)
        Case 3
            newColor = RGB(102, 57, 49)
        Case 4
            newColor = RGB(143, 86, 59)
        Case 5
            newColor = RGB(223, 113, 38)
        Case 6
            newColor = RGB(217, 160, 102)
        Case 7
            newColor = RGB(238, 195, 154)
        Case 8
            newColor = RGB(251, 242, 54)
        Case 9
            newColor = RGB(153, 229, 80)
        Case 10
            newColor = RGB(106, 190, 48)
        Case 11
            newColor = RGB(55, 148, 110)
        Case 12
            newColor = RGB(75, 105, 47)
        Case 13
            newColor = RGB(82, 75, 36)
        Case 14
            newColor = RGB(50, 60, 57)
        Case 15
            newColor = RGB(63, 63, 116)
        Case 16
            newColor = RGB(48, 96, 130)
        Case 17
            newColor = RGB(91, 110, 225)
        Case 18
            newColor = RGB(99, 155, 255)
        Case 19
            newColor = RGB(95, 205, 228)
        Case 20
            newColor = RGB(203, 219, 252)
        Case 21
            newColor = RGB(255, 255, 255)
        Case 22
            newColor = RGB(155, 173, 183)
        Case 23
            newColor = RGB(132, 126, 135)
        Case 24
            newColor = RGB(105, 106, 106)
        Case 25
            newColor = RGB(89, 86, 82)
        Case 26
            newColor = RGB(118, 66, 138)
        Case 27
            newColor = RGB(172, 50, 50)
        Case 28
            newColor = RGB(217, 87, 99)
        Case 29
            newColor = RGB(215, 123, 186)
        Case 30
            newColor = RGB(143, 151, 74)
        Case 31
            newColor = RGB(138, 111, 48)
        Case Else
            newColor = RGB(255, 255, 255)
            
    End Select
    
    histValue = CStr(newColor)
    
End Sub

Private Sub clrHistory_HistoryItemClicked(ByVal histIndex As Long, ByVal histValue As String)
    
    If (LenB(histValue) <> 0) Then
    
        Dim clickedColor As Long
        clickedColor = CLng(histValue)
        
        'Update the other color selectors with this color value
        clrWheel.Color = clickedColor
        clrVariants.Color = clickedColor
        
    End If
    
End Sub

'Update the color history tooltip to reflect the currently hovered color
Private Sub clrHistory_HistoryItemMouseOver(ByVal histIndex As Long, ByVal histValue As String)

    If (LenB(histValue) <> 0) Then
    
        Dim hoverColor As Long
        hoverColor = Colors.ConvertSystemColor(CLng(histValue))
        
        'Construct hex and RGB string representations of the target color
        Dim hexString As String, rgbString As String
        hexString = "#" & UCase$(Colors.GetHexStringFromRGB(hoverColor))
        rgbString = g_Language.TranslateMessage("RGB(%1, %2, %3)", Colors.ExtractRed(hoverColor), Colors.ExtractGreen(hoverColor), Colors.ExtractBlue(hoverColor))
        clrHistory.AssignTooltip hexString & vbCrLf & rgbString, , True
        
    End If

End Sub

Private Sub clrVariants_ColorChanged(ByVal newColor As Long, ByVal srcIsInternal As Boolean)
    
    'If the clrVariant control is where the color was actually changed (and it's not just syncing itself to some
    ' external color change), relay the new color to the neighboring color wheel.
    If srcIsInternal Then clrWheel.Color = newColor
    
    RelayColorChange newColor
    
End Sub

Private Sub clrWheel_ColorChanged(ByVal newColor As Long, ByVal srcIsInternal As Boolean)
    If srcIsInternal Then clrVariants.Color = newColor
End Sub

Private Sub cmdSettings_Click()
    If (Dialogs.ChooseColorPanelSettings() = vbOK) Then
        VerifyPanelUserPrefs
        ReflowInterface
    End If
End Sub

Private Sub cmdSettings_DrawButton(ByVal bufferDC As Long, ByVal buttonIsHovered As Boolean, ByVal ptrToRectF As Long)
    
    If (bufferDC = 0) Then Exit Sub
    
    If PDMain.IsProgramRunning Then
    
        Dim cSurface As pd2DSurface
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, True
        cSurface.SetSurfacePixelOffset P2_PO_Half
        
        Dim cBrush As pd2DBrush
        Drawing2D.QuickCreateSolidBrush cBrush, cmdSettings.GetCurrentCaptionColor()
        
        Dim btnRectF As RectF
        With btnRectF
            .Left = 0!
            .Top = 0!
            .Width = cmdSettings.GetWidth
            .Height = cmdSettings.GetHeight
        End With
        
        'We're now gonna create three small "dots" to render as a "..." caption
        Dim dotRectF() As RectF
        ReDim dotRectF(0 To 2) As RectF
        
        Dim i As Long
        For i = 0 To 2
            With dotRectF(i)
                .Width = Interface.FixDPIFloat(1.75)
                .Height = Interface.FixDPIFloat(1.75)
                .Top = (btnRectF.Height * 0.5) - (.Height * 0.5)
                .Left = (btnRectF.Width * 0.5) - (.Width * 0.5)
                If (i = 0) Then .Left = .Left - .Width * 2.5
                If (i = 2) Then .Left = .Left + .Width * 2.5
            End With
            PD2D.FillRectangleF_FromRectF cSurface, cBrush, dotRectF(i)
        Next i
        
        Set cBrush = Nothing: Set cSurface = Nothing
        
    End If
    
End Sub

Private Sub Form_Load()
    
    m_ResizeInProgress = True
    
    'Prep some items related to the color history UI
    clrHistory.RequestCustomSubclassing WM_PD_PRIMARY_COLOR_APPLIED, True
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
    
    'Load any relevant user settings
    VerifyPanelUserPrefs True
    
    'Update everything against the current theme.  This will also set tooltips for various controls,
    ' and reflow the interface to match.
    UpdateAgainstCurrentTheme
    
    m_ResizeInProgress = False
    
End Sub

'After the user has potentially changed settings (e.g. at first-load or after the settings panel is invoked),
' modify visibility of various panel elements to match.
Private Sub VerifyPanelUserPrefs(Optional ByVal forceRefresh As Boolean = False)
    
    Dim oldRenderMode As PD_ColorPanelMode
    oldRenderMode = m_RenderMode
    
    m_RenderMode = UserPrefs.GetPref_Long("Tools", "ColorPanelStyle", cpm_Wheels)
    m_PaletteFile = UserPrefs.GetPref_String("Tools", "ColorPanelPaletteFile")
    If (LenB(m_PaletteFile) <> 0) Then
        palSelector.PaletteFile = m_PaletteFile
        palSelector.PaletteGroup = UserPrefs.GetPref_Long("Tools", "ColorPanelPaletteGroup", 0)
    End If
    
    'If the palette file is invalid, we'll revert to the standard mode
    If (m_RenderMode = cpm_Palette) And (Not palSelector.IsPaletteValid) Then m_RenderMode = cpm_Wheels
    
    'When render mode changes, adjust visibility accordingly
    If (m_RenderMode <> oldRenderMode) Or forceRefresh Then
        
        '"Wheels" mode
        clrVariants.Visible = (m_RenderMode = cpm_Wheels)
        clrWheel.Visible = (m_RenderMode = cpm_Wheels)
        clrHistory.Visible = (m_RenderMode = cpm_Wheels)
        
        '"Palette" mode
        palSelector.Visible = (m_RenderMode = cpm_Palette)
        
    End If
    
End Sub

'When the currently active color selector experiences a color change, call this function to relay that
' change elsewhere in the program.
Private Sub RelayColorChange(ByVal newColor As Long)

    'Whenever this primary color changes, we broadcast the change throughout PD, so other color selector controls
    ' know to redraw themselves accordingly.
    UserControls.PostPDMessage WM_PD_PRIMARY_COLOR_CHANGE, newColor
    
    'We also check to see if a paint-related tool is active.  If it is, assign the new color immediately.
    Select Case g_CurrentTool
    
        Case PAINT_PENCIL
            Tools_Pencil.SetBrushColor newColor
        
        Case PAINT_SOFTBRUSH
            Tools_Paint.SetBrushSourceColor newColor
            
        Case PAINT_FILL
            Tools_Fill.SetFillBrushColor newColor
    
    End Select
    
End Sub

'Whenever this panel is resized, we must reflow all objects to fit the available space.
Private Sub ReflowInterface()
    
    Dim curFormWidth As Long, curFormHeight As Long
    If (g_WindowManager Is Nothing) Then
        curFormWidth = Me.ScaleWidth
        curFormHeight = Me.ScaleHeight
    Else
        curFormWidth = g_WindowManager.GetClientWidth(Me.hWnd)
        curFormHeight = g_WindowManager.GetClientHeight(Me.hWnd)
    End If
    
    'Failsafe to prevent IDE errors
    If (curFormWidth > 50) And (curFormHeight > 10) Then
        
        'Bottom-align the color history panel, while leaving space for the "settings" button
        Dim cmdSettingsWidth As Long
        cmdSettingsWidth = Interface.FixDPI(28)
        clrHistory.SetPositionAndSize 0, curFormHeight - clrHistory.GetHeight, curFormWidth - cmdSettingsWidth - Interface.FixDPI(4), clrHistory.GetHeight
        
        'Bottom-align the color settings button
        cmdSettings.SetPositionAndSize clrHistory.GetLeft + clrHistory.GetWidth + Interface.FixDPI(4), clrHistory.GetTop, cmdSettingsWidth, clrHistory.GetHeight
        
        'Calculate a new height available to the other controls on this panel
        curFormHeight = curFormHeight - (clrHistory.GetHeight + Interface.FixDPI(2))
        
        'Set the palette control to this height (regardless of whether or not it's visible)
        palSelector.SetPositionAndSize 0, 0, curFormWidth, cmdSettings.GetTop - FixDPI(4)
        
        'Before rendering other elements, enforce a minimum size.  During startup, form size vacillates
        ' several times as this window is "fit" against its neighbors.  This can throw GDI+ rendering
        ' error messages until a final size is arrived at.
        If (curFormHeight > 25) Then
        
            'We now have a dilemma.  If more size is available horizontally then vertically, render the color
            ' wheels side-by-side...
            If (curFormHeight < curFormWidth) Then
            
                'Right-align the color wheel
                clrWheel.SetPositionAndSize curFormWidth - (curFormHeight + Interface.FixDPI(1)), 0, curFormHeight, curFormHeight
                
                'Fit the variant selector into the remaining area.
                clrVariants.SetPositionAndSize 0, 0, clrWheel.GetLeft - Interface.FixDPI(10), curFormHeight
            
            'But if there is more space vertically, stack the controls atop each other
            Else
            
                'Bottom-align the color wheel
                clrWheel.SetPositionAndSize 0, curFormHeight - (curFormWidth + Interface.FixDPI(1)), curFormWidth, curFormWidth
                
                'Fit the variant selector into the remaining area.
                clrVariants.SetPositionAndSize 0, 0, curFormWidth, clrWheel.GetTop - Interface.FixDPI(10)
                
            End If
            
        End If
        
        'Dynamically set the width of the hue wheel (the rainbow circle surrounding the HSV picker)
        ' to be a fraction of the control's on-screen size.  This guarantees good layouts even if the user resizes
        ' this panel to be extremely large or extremely small.
        Dim newWheelWidth As Long
        newWheelWidth = (clrWheel.GetWidth * 0.15)
        If (newWheelWidth < 9) Then newWheelWidth = 9
        clrWheel.WheelWidth = newWheelWidth
        
    End If
    
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.ForceWindowRepaint Me.hWnd
    
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    ApplyThemeAndTranslations Me
    
    'Reflow the interface, to account for any language changes.
    ReflowInterface
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    If Not (m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues
        m_lastUsedSettings.SetParentForm Nothing
    End If
    
End Sub

Private Sub Form_Resize()
    If (Not m_ResizeInProgress) Then
        m_ResizeInProgress = True
        ReflowInterface
        m_ResizeInProgress = False
    End If
End Sub

Public Function GetCurrentColor() As Long
    If (m_RenderMode = cpm_Wheels) Then
        GetCurrentColor = clrVariants.Color
    ElseIf (m_RenderMode = cpm_Palette) Then
        GetCurrentColor = palSelector.GetPaletteColor()
    End If
End Function

Public Sub SetCurrentColor(ByVal newR As Long, ByVal newG As Long, ByVal newB As Long)
    
    clrVariants.Color = RGB(newR, newG, newB)
    clrWheel.Color = RGB(newR, newG, newB)
    clrHistory.PushNewHistoryItem RGB(newR, newG, newB), , True
    
    'The palette selector is a little weird here; basically, we need to find the *closest* color
    ' to the one we were passed.
    If (m_RenderMode = cpm_Palette) Then palSelector.SetPaletteColor RGB(newR, newG, newB)
    
End Sub

Private Sub palSelector_Click(ByVal palIndex As Long, ByVal palColor As Long)
    
    'Relay the new color to the other color selectors; this will automatically sync the color
    ' across the program.  (Also, we do it this so that if the user switches color selector modes,
    ' they will retain the current color correctly.)
    clrWheel.Color = palColor
    clrVariants.Color = palColor
    
End Sub
