VERSION 5.00
Begin VB.UserControl pdPaletteUI 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
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
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   260
   ToolboxBitmap   =   "pdPaletteUI.ctx":0000
End
Attribute VB_Name = "pdPaletteUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Palette-based Color Selector
'Copyright 2018-2026 by Tanner Helland
'Created: 14/February/18
'Last updated: 16/April/19
'Last update: add option for displaying alpha in palette entries; this isn't always wanted, but in the
'             Export > Palette window, it's critical for displaying user-friendly results
'
'In February 2018, PD gained extensive support for various palette file formats.  (Or as they call 'em in
' Adobe parlance: swatches.)  This UC exists to make it easier for the user to see all available palette colors,
' and rapidly select colors from a set list.
'
'In most cases, palettes should probably be treated as RGB-only entities... but there are rare cases where
' RGBA palette data is useful.  (Web-optimized PNGs are one such place.)  This control can handle both cases,
' but if you want full RGBA handling, you'll need to manually set the relevant property at design-time.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Event Click(ByVal palIndex As Long, ByVal palColor As Long)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'To simplify rendering, we pre-calculate a rectangle for the "palette" area of the control.
' (Individual items within the palette control can be resolved on-the-fly).  This rect is calculated
' by UpdateControlLayout, and it must be recalculated if the control size changes.
Private m_PaletteRect As RectF

'Once a palette has been loaded, this palette object will contain all relevant color data.
Private m_Palette As pdPalette

'Some palette formats support multiple "child" palettes (e.g. Adobe's ASE format).  To access child palettes,
' set this to a non-zero value.
Private m_ChildPaletteIndex As Long

'Individual palette items within the palette rectangle are resolved by (x, y) position.  Note that the
' number of rows is controlled by user property; columns are automatically inferred from that value.
Private m_NumPaletteRows As Long, m_NumPaletteColumns As Long
Private m_FitWidth As Single, m_FitHeight As Single
Private m_PaletteRects() As RectF

'If a palette item is currently hovered by the mouse, this will be set to some value >= 0
Private m_PaletteItemHovered As Long

'The currently selected palette entry, if any, will be stored here
Private m_PaletteItemSelected As Long

'Fast palette matching comes courtesy of a dedicated KD-tree class
Private m_ColorLookup As pdKDTree

'By default, this class only displays RGB triplets with assumed opacity of 255.  If you want full RGBA handling,
' this must be set to TRUE (currently exposed via property).
Private m_UseRGBA As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDHISTORY_COLOR_LIST
    [_First] = 0
    PDH_Background = 0
    PDH_Caption = 1
    PDH_Border = 2
    [_Last] = 2
    [_Count] = 3
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_History
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if it ever becomes relevant.
Public Property Get Caption() As String
    Caption = ucSupport.GetCaptionText()
End Property

Public Property Let Caption(ByRef newCaption As String)
    ucSupport.SetCaptionText newCaption
    PropertyChanged "Caption"
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

Public Property Get FontSize() As Single
    FontSize = ucSupport.GetCaptionFontSize()
End Property

Public Property Let FontSize(ByVal newSize As Single)
    ucSupport.SetCaptionFontSize newSize
    PropertyChanged "FontSize"
End Property

Public Property Get PaletteFile() As String
    If (Not m_Palette Is Nothing) Then PaletteFile = m_Palette.GetPaletteFilename Else PaletteFile = vbNullString
End Property

Public Property Let PaletteFile(ByRef newFile As String)
    
    Dim oldPaletteFile As String
    oldPaletteFile = Me.PaletteFile
    
    If (m_Palette Is Nothing) Then Set m_Palette = New pdPalette
    
    'NOTE: I'm not sure if ideal behavior here is to remove duplicates or not.  At present,
    ' we *do not* remove duplicate entries from imported palettes.
    If (Not m_Palette.LoadPaletteFromFile(newFile, False, True)) Then Set m_Palette = Nothing
    
    If (Not m_Palette Is Nothing) Then
    
        If Strings.StringsNotEqual(oldPaletteFile, m_Palette.GetPaletteFilename, True) Then
        
            'Whenever the palette file changes, we need to make sure various palette indices are valid
            ' against the *new* file.
            m_ChildPaletteIndex = 0
            m_PaletteItemHovered = -1
            m_PaletteItemSelected = 0
            
            'Reset our color lookup object; it will need to be re-created against the new group
            Set m_ColorLookup = Nothing
            
            UpdateControlLayout
        
        End If
        
    End If
    
End Property

'You don't have to create palettes from files; use this to do it from an existing pdPalette object
Public Function SetPDPalette(ByRef srcPalette As pdPalette) As Boolean
    
    If (Not srcPalette Is Nothing) Then
        
        Set m_Palette = srcPalette
        
        'Whenever the palette file changes, we need to make sure various palette indices are valid
        ' against the *new* file.
        m_ChildPaletteIndex = 0
        m_PaletteItemHovered = -1
        m_PaletteItemSelected = 0
        
        'Reset our color lookup object; it will need to be re-created against the new group
        Set m_ColorLookup = Nothing
        UpdateControlLayout
        
        SetPDPalette = True
            
    Else
        PDDebug.LogAction "WARNING!  pdPaletteUI.SetPDPalette was passed a null palette!"
    End If
    
End Function

Public Property Get PaletteGroup() As Long
    PaletteGroup = m_ChildPaletteIndex
End Property

Public Property Let PaletteGroup(ByVal newGroup As Long)
    If (Not m_Palette Is Nothing) Then
        If (newGroup >= 0) And (newGroup < m_Palette.GetPaletteGroupCount) And (newGroup <> m_ChildPaletteIndex) Then
            
            m_ChildPaletteIndex = newGroup
            m_PaletteItemHovered = -1
            m_PaletteItemSelected = 0
            
            'Reset our color lookup object; it will need to be re-created against the new group
            Set m_ColorLookup = Nothing
            
            RedrawBackBuffer
            
        End If
    End If
End Property

Public Property Get PaletteIndex() As Long
    PaletteIndex = m_PaletteItemSelected
End Property

Public Property Let PaletteIndex(ByVal newIndex As Long)
    If (Not m_Palette Is Nothing) Then
        If (m_ChildPaletteIndex >= 0) And (m_ChildPaletteIndex < m_Palette.GetPaletteGroupCount) Then
            If (newIndex >= 0) And (newIndex < m_Palette.GetPaletteColorCount(m_ChildPaletteIndex)) And (newIndex <> m_PaletteItemSelected) Then
                m_PaletteItemSelected = newIndex
                RedrawBackBuffer
            End If
        End If
    End If
End Property

Public Property Get UseRGBA() As Boolean
    UseRGBA = m_UseRGBA
End Property

Public Property Let UseRGBA(ByVal newState As Boolean)
    If (newState <> m_UseRGBA) Then
        m_UseRGBA = newState
        PropertyChanged "UseRGBA"
        RedrawBackBuffer
    End If
End Property

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'To support high-DPI settings properly, we expose some specialized move+size functions
Public Function GetLeft() As Long
    GetLeft = ucSupport.GetControlLeft
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    ucSupport.RequestNewPosition newLeft, , True
End Sub

Public Function GetTop() As Long
    GetTop = ucSupport.GetControlTop
End Function

Public Sub SetTop(ByVal newTop As Long)
    ucSupport.RequestNewPosition , newTop, True
End Sub

Public Function GetWidth() As Long
    GetWidth = ucSupport.GetControlWidth
End Function

Public Sub SetWidth(ByVal newWidth As Long)
    ucSupport.RequestNewSize newWidth, , True
End Sub

Public Function GetHeight() As Long
    GetHeight = ucSupport.GetControlHeight
End Function

Public Sub SetHeight(ByVal newHeight As Long)
    ucSupport.RequestNewSize , newHeight, True
End Sub

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

'If our parent control needs a redraw for some reason, it can request one here
Public Sub RequestRedraw(Optional ByVal paintImmediately As Boolean = False)
    RedrawBackBuffer paintImmediately
End Sub

Public Function GetPaletteColor() As Long
    If (Not m_Palette Is Nothing) And (m_PaletteItemSelected >= 0) Then
        GetPaletteColor = m_Palette.ChildPalette(m_ChildPaletteIndex).GetPaletteColorAsLong(m_PaletteItemSelected)
    End If
End Function

'This function doesn't work how you may think.  It accepts an arbitrary new color, and it then finds the
' *closest* palette color, and sets that as the new index.
Public Sub SetPaletteColor(ByVal newColor As Long)
    
    If (m_Palette Is Nothing) Or (m_PaletteItemSelected < 0) Then Exit Sub
    
    'If our color lookup object doesn't exist, we need to create it prior to performing lookups.
    If (m_ColorLookup Is Nothing) Then
        
        Dim tmpPalette() As RGBQuad
        If m_Palette.CopyPaletteToArray(tmpPalette, m_ChildPaletteIndex) Then
            Set m_ColorLookup = New pdKDTree
            m_ColorLookup.BuildTree tmpPalette, m_Palette.GetPaletteColorCount(m_ChildPaletteIndex)
        End If
        
    End If
    
    'Set our palette index to the *nearest* color to the target one
    Dim tmpQuad As RGBQuad
    tmpQuad.Red = Colors.ExtractRed(newColor)
    tmpQuad.Green = Colors.ExtractGreen(newColor)
    tmpQuad.Blue = Colors.ExtractBlue(newColor)
    Me.PaletteIndex = m_ColorLookup.GetNearestPaletteIndex(tmpQuad)

End Sub

Public Function IsPaletteValid() As Boolean
    IsPaletteValid = (Not m_Palette Is Nothing)
End Function

Public Sub CreateFromXML(ByRef srcXML As String)
    
    Dim cXML As pdSerialize
    Set cXML = New pdSerialize
    cXML.SetParamString srcXML
    
    With cXML
        Dim srcPalFile As String
        srcPalFile = .GetString("palette-filename", vbNullString, True)
        If (LenB(srcPalFile) <> 0) Then Me.PaletteFile = srcPalFile
        Me.PaletteGroup = .GetLong("palette-group", 0)
        Me.PaletteIndex = .GetLong("palette-index", 0)
    End With
    
End Sub

Public Function SerializeToXML() As String

    Dim cXML As pdSerialize
    Set cXML = New pdSerialize
    cXML.AddParam "palette-filename", Me.PaletteFile
    cXML.AddParam "palette-group", Me.PaletteGroup
    cXML.AddParam "palette-index", Me.PaletteIndex
    
    SerializeToXML = cXML.GetParamString()

End Function

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
    RedrawBackBuffer
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
    RedrawBackBuffer
End Sub

'Only left clicks raise Click() events
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    If Me.Enabled And ((Button And pdLeftButton) <> 0) Then
        
        'Start by seeing if the mouse is inside the history portion of the control
        Dim clickedIndex As Long
        clickedIndex = GetPaletteItemUnderMouse(x, y)
        
        If (clickedIndex >= 0) Then
            m_PaletteItemSelected = clickedIndex
            RaiseEvent Click(clickedIndex, m_Palette.ChildPalette(m_ChildPaletteIndex).GetPaletteColorAsLong(m_PaletteItemSelected))
            RedrawBackBuffer
        End If
        
    End If
    
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    ucSupport.RequestCursor IDC_HAND
    RedrawBackBuffer
End Sub

'When the mouse leaves the UC, we must repaint the button (as it's no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_PaletteItemHovered = -1
    Me.AssignTooltip vbNullString, , False
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    Dim oldHoverCheck As Long
    oldHoverCheck = m_PaletteItemHovered
    m_PaletteItemHovered = GetPaletteItemUnderMouse(x, y)
    
    If (m_PaletteItemHovered >= 0) Then
        ucSupport.RequestCursor IDC_HAND
    Else
        ucSupport.RequestCursor IDC_DEFAULT
        Me.AssignTooltip vbNullString, , False
    End If
    
    If (oldHoverCheck <> m_PaletteItemHovered) Then
        
        RedrawBackBuffer
        
        'Dynamically update our tooltip with this color's name, if any
        If (m_PaletteItemHovered >= 0) Then
            
            Dim tmpPal As PDPaletteEntry
            tmpPal = m_Palette.ChildPalette(m_ChildPaletteIndex).GetPaletteEntry(m_PaletteItemHovered)
            
            Dim targetColor As Long
            With tmpPal
                targetColor = Colors.ConvertSystemColor(RGB(.ColorValue.Red, .ColorValue.Green, .ColorValue.Blue))
            End With
            
            'Construct hex and RGB string representations of the target color
            Dim hexString As String, rgbString As String, indexString As String
            
            If m_UseRGBA Then
                hexString = "#" & UCase$(Colors.GetHexStringFromRGB(targetColor)) & UCase$(Hex$(tmpPal.ColorValue.Alpha))
                rgbString = g_Language.TranslateMessage("RGBA(%1, %2, %3, %4)", tmpPal.ColorValue.Red, tmpPal.ColorValue.Green, tmpPal.ColorValue.Blue, tmpPal.ColorValue.Alpha)
            Else
                hexString = "#" & UCase$(Colors.GetHexStringFromRGB(targetColor))
                rgbString = g_Language.TranslateMessage("RGB(%1, %2, %3)", tmpPal.ColorValue.Red, tmpPal.ColorValue.Green, tmpPal.ColorValue.Blue)
            End If
            
            indexString = g_Language.TranslateMessage("index %1", m_PaletteItemHovered)
            Me.AssignTooltip hexString & vbCrLf & rgbString & vbCrLf & indexString, tmpPal.ColorName, True
            
        End If
        
    End If
    
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Private Sub UserControl_Initialize()
    
    m_PaletteItemHovered = -1
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True
    ucSupport.RequestCaptionSupport
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDHISTORY_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDHistory", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
End Sub

'Set default properties
Private Sub UserControl_InitProperties()
    Me.Caption = vbNullString
    Me.FontSize = 12
    Me.PaletteFile = vbNullString
    Me.UseRGBA = False
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Me.Caption = .ReadProperty("Caption", vbNullString)
        Me.FontSize = .ReadProperty("FontSize", 12)
        Me.PaletteFile = .ReadProperty("PaletteFile", vbNullString)
        Me.UseRGBA = .ReadProperty("UseRGBA", False)
    End With
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", ucSupport.GetCaptionText, vbNullString
        .WriteProperty "FontSize", ucSupport.GetCaptionFontSize, 12
        .WriteProperty "PaletteFile", Me.PaletteFile, vbNullString
        .WriteProperty "UseRGBA", m_UseRGBA, False
    End With
End Sub

Private Function GetPaletteItemUnderMouse(ByVal srcX As Single, ByVal srcY As Single) As Long
    
    GetPaletteItemUnderMouse = -1
    
    'First, shortcut the function by seeing if the mouse is even inside the palette area.  (If a caption is in use,
    ' this may not be true.)
    If PDMath.IsPointInRectF(srcX, srcY, m_PaletteRect) Then
    
        If (Not m_Palette Is Nothing) Then
            
            Dim i As Long
            For i = 0 To m_Palette.ChildPalette(m_ChildPaletteIndex).GetNumOfColors - 1
                If PDMath.IsPointInRectF(srcX, srcY, m_PaletteRects(i)) Then
                    GetPaletteItemUnderMouse = i
                    Exit For
                End If
            Next i
            
        End If
    
    End If

End Function

'Call this layout calculator whenever the control size changes.  This is particularly important for this control,
' as palette color rects are pre-calculated ahead of time, to simplify rendering.
Private Sub UpdateControlLayout()

    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'The first thing we want to do is calculate an available area for the entire palette section of the control.
    ' (Spacing rules are the same as all other captioned controls.)
    If ucSupport.IsCaptionActive Then
        
        'The brush area is placed relative to the caption
        With m_PaletteRect
            .Left = FixDPI(8)
            .Top = ucSupport.GetCaptionBottom + 2
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    'If there's no caption, allow the clickable portion to fill the entire control
    Else
        
        With m_PaletteRect
            .Left = 1
            .Top = 1
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    End If
    
    'Next, we need to figure out how to place palette colors inside the current control area.  We want each palette
    ' entry to be a perfect square, so the number of items we can support is limited to how many perfectly square
    ' items we can fit in the supplied area.
    Dim colorCount As Long
    If (Not m_Palette Is Nothing) Then
        
        colorCount = m_Palette.ChildPalette(m_ChildPaletteIndex).GetNumOfColors()
        If (colorCount > 0) Then
        
            ReDim m_PaletteRects(0 To colorCount - 1) As RectF
            
            'Iterate until we find the "best fit" for this palette's color count given the
            ' available space in the control.
            Dim numColorsFit As Long, minimumSize As Long, maximumSize As Long
            minimumSize = PDMath.Min2Int(m_PaletteRect.Width, m_PaletteRect.Height)
            maximumSize = PDMath.Max2Int(m_PaletteRect.Width, m_PaletteRect.Height)
            numColorsFit = Int(CSng(maximumSize) / CSng(minimumSize))
            
            Dim fitSize As Long, numIterations As Long
            numIterations = 1
            
            Do While (numColorsFit < colorCount)
                numIterations = numIterations + 1
                fitSize = Int(minimumSize / CSng(numIterations) + 0.5)
                numColorsFit = (maximumSize \ fitSize) * (minimumSize \ fitSize)
            Loop
            
            'fitSize now represents the minimum size required to fit all palette colors in the available space
            ' for this control.  (This size may be ridiculously small, but we'll deal with that in the future.)
            m_NumPaletteColumns = m_PaletteRect.Width \ fitSize
            If (m_NumPaletteColumns < 1) Then m_NumPaletteColumns = 1
            
            'Now here's a funny thing about height - depending on the number of colors in this palette,
            ' we may not actually use all the rows we've calculated.  This depends on how the color count
            ' "fits" into the available space.  Let's figure out how many rows we'll actually be using.
            m_NumPaletteRows = Int((colorCount - 1) / m_NumPaletteColumns) + 1
            If (m_NumPaletteRows < 1) Then m_NumPaletteRows = 1
            
            'Translate this to "perfect fit" width and height values.
            m_FitWidth = m_PaletteRect.Width / m_NumPaletteColumns
            m_FitHeight = m_PaletteRect.Height / m_NumPaletteRows
            
            Dim hOffset As Single: hOffset = m_PaletteRect.Left
            Dim vOffset As Single: vOffset = m_PaletteRect.Top
            
            'With the palette area correctly identified, we can now calculate rects for each individual color.
            ' (These are pre-calculated to simplify rendering and hit-detection.)
            Dim i As Long
            For i = 0 To colorCount - 1
            
                With m_PaletteRects(i)
                
                    'When rendering palette colors, we don't want to use subpixels due to the potentially
                    ' tiny sizes involved.  Everything must be clamped to integer values.
                    .Left = Int(hOffset + 0.5)
                    .Top = Int(vOffset + 0.5)
                    .Width = Int(hOffset + m_FitWidth + 0.5) - .Left
                    .Height = Int(vOffset + m_FitHeight + 0.5) - .Top
                    
                    hOffset = hOffset + m_FitWidth
                    
                    'If a color box gets pushed past the edge of the control, move it down to the next row
                    If (hOffset > (m_PaletteRect.Left + m_PaletteRect.Width - 2)) Then
                        hOffset = m_PaletteRect.Left
                        vOffset = vOffset + m_FitHeight
                    End If
                    
                End With
            
            Next i
            
        End If
        
    End If
        
    'No other special preparation is required for this control, so proceed with recreating the back buffer
    RedrawBackBuffer
            
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDH_Background, "Background", IDE_WHITE
        .LoadThemeColor PDH_Caption, "Caption", IDE_GRAY
        .LoadThemeColor PDH_Border, "Border", IDE_GRAY
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared
' to just flipping the existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer(Optional ByVal paintImmediately As Boolean = False)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDH_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
    
    If PDMain.IsProgramRunning() Then
        
        Dim i As Long
        
        'Because this control is owner-drawn, our owner is responsible for drawing the individual history samples.
        If (Not m_Palette Is Nothing) Then
            
            Dim cSurface As pd2DSurface, cBrush As pd2DBrush
            Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, False
            Drawing2D.QuickCreateSolidBrush cBrush, m_Palette.ChildPalette(m_ChildPaletteIndex).GetPaletteColorAsLong(0)
            
            Dim colorLoopMax As Long
            colorLoopMax = m_Palette.ChildPalette(m_ChildPaletteIndex).GetNumOfColors - 1
            
            'Fill the rects already created for each palette entry
            For i = 0 To colorLoopMax
            
                cBrush.SetBrushColor m_Palette.ChildPalette(m_ChildPaletteIndex).GetPaletteColorAsLong(i)
                
                If m_UseRGBA Then
                    With m_PaletteRects(i)
                        GDI_Plus.GDIPlusFillDIBRect_Pattern Nothing, .Left, .Top, .Width, .Height, g_CheckerboardPattern, bufferDC, True
                    End With
                    cBrush.SetBrushOpacity CSng(m_Palette.ChildPalette(m_ChildPaletteIndex).GetPaletteColor(i).Alpha) / 2.55!
                End If
                
                PD2D.FillRectangleF_FromRectF cSurface, cBrush, m_PaletteRects(i)
                
            Next i
            
            'Next, draw borders around each palette entry
            Dim cPen As pd2DPen
            Drawing2D.QuickCreateSolidPen cPen, 1, m_Colors.RetrieveColor(PDH_Border)
            For i = 0 To colorLoopMax
                PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_PaletteRects(i)
            Next i
            
            'Next, highlight the hovered and/or selected color, if any
            If (m_PaletteItemSelected >= 0) Then
                Dim cOuterPen As pd2DPen
                Drawing2D.QuickCreatePairOfUIPens cOuterPen, cPen, True
                PD2D.DrawRectangleF_FromRectF cSurface, cOuterPen, m_PaletteRects(m_PaletteItemSelected)
                PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_PaletteRects(m_PaletteItemSelected)
                Set cOuterPen = Nothing
            End If
            
            If (m_PaletteItemHovered >= 0) Then
                Drawing2D.QuickCreateSolidPen cPen, 3, m_Colors.RetrieveColor(PDH_Border, True, False, True)
                PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_PaletteRects(m_PaletteItemHovered)
            End If
            
            Set cSurface = Nothing: Set cPen = Nothing
            
        End If
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint paintImmediately
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
