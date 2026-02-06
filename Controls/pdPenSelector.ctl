VERSION 5.00
Begin VB.UserControl pdPenSelector 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdPenSelector.ctx":0000
End
Attribute VB_Name = "pdPenSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Pen Selector custom control
'Copyright 2015-2026 by Tanner Helland
'Created: 04/July/15
'Last updated: 20/June/16
'Last update: migrate to pd2D for all UI rendering
'
'This thin user control is basically an empty control that when clicked, displays a pen selection window.  If a
' pen is selected (e.g. Cancel is not pressed), it updates its appearance to match, and raises a "PenChanged"
' event.
'
'Though simple, this control solves a lot of problems.  It is especially helpful for improving interaction with the
' command bar user control, as it easily supports pen reset/randomize/preset events.  It is also nice to be able
' to update a single central function for pen selection, then have the change propagate to all tool instances.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************


Option Explicit

'This control doesn't really do anything interesting, besides allow a pen to be selected.
Public Event PenChanged(ByVal isFinalChange As Boolean)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()
Public Event SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, ByRef newTargetHwnd As Long)

'The control's current pen settings
Private m_curPen As String

'The path used for the preview window
Private m_PreviewPath As pd2DPath

'When the "select pen" dialog is live, this will be set to TRUE
Private m_IsDialogLive As Boolean

'If a change event originates from the child dialog, this is set to TRUE
Private m_ChangeIsLive As Boolean

'The rectangle where the pen preview is actually rendered, and a boolean to track whether the mouse is inside that rect
Private m_PenRect As RectF, m_MouseInsidePenRect As Boolean, m_MouseDownPenRect As Boolean

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDPS_COLOR_LIST
    [_First] = 0
    PDPS_Border = 0
    [_Last] = 0
    [_Count] = 1
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_PenSelector
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if it ever becomes relevant.
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
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

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'You can retrieve the pen param string (not a pd2DPen object!) via this property
Public Property Get Pen() As String
    Pen = m_curPen
End Property

Public Property Let Pen(ByVal newPen As String)
    m_curPen = newPen
    RedrawBackBuffer
    RaiseEvent PenChanged(Not m_ChangeIsLive)
    PropertyChanged "Pen"
End Property

'Outside functions can call this to force a display of the pen selection window
Public Sub DisplayPenSelection()
    RaisePenDialog
End Sub

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition x, y
    If m_MouseInsidePenRect Then RaisePenDialog
End Sub

Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    If (wMsg = WM_PD_COLOR_MANAGEMENT_CHANGE) Then RedrawBackBuffer
End Sub

Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    markEventHandled = False
    
    If Me.Enabled And (vkCode = VK_SPACE) Then
        RaisePenDialog
        markEventHandled = True
    End If
    
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    UpdateMousePosition x, y
    If m_MouseInsidePenRect Then
        m_MouseDownPenRect = True
        RedrawBackBuffer
    End If
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition x, y
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition -100, -100
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    UpdateMousePosition x, y
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    m_MouseDownPenRect = False
    RedrawBackBuffer
End Sub

Private Sub UpdateMousePosition(ByVal mouseX As Single, ByVal mouseY As Single)
    m_MouseInsidePenRect = PDMath.IsPointInRectF(mouseX, mouseY, m_PenRect)
    If m_MouseInsidePenRect Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
End Sub

Private Sub ucSupport_GotFocusAPI()
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RedrawBackBuffer
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub RaisePenDialog()

    m_IsDialogLive = True
    
    'Backup the current pen; if the dialog is canceled, we want to restore it
    Dim newPen As String, oldPen As String
    oldPen = Me.Pen
    
    'Use the brush dialog to select a new color
    If ShowPenDialog(newPen, oldPen, Me) Then
        Me.Pen = newPen
    Else
        Me.Pen = oldPen
    End If
    
    m_IsDialogLive = False
    
End Sub

Private Sub ucSupport_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    RaiseEvent SetCustomTabTarget(shiftTabWasPressed, newTargetHwnd)
End Sub

Private Sub UserControl_Initialize()
    
    Set m_PreviewPath = New pd2DPath
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_SPACE
    ucSupport.RequestCaptionSupport
    ucSupport.SubclassCustomMessage WM_PD_COLOR_MANAGEMENT_CHANGE, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDPS_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDPenSelector", colorCount
    If Not PDMain.IsProgramRunning() Then UpdateColorList
    
End Sub

Private Sub UserControl_InitProperties()
    Caption = vbNullString
    FontSize = 12
    Pen = vbNullString
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Caption = .ReadProperty("Caption", vbNullString)
        FontSize = .ReadProperty("FontSize", 12)
        Pen = .ReadProperty("curPen", vbNullString)
    End With
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", ucSupport.GetCaptionText, vbNullString
        .WriteProperty "FontSize", ucSupport.GetCaptionFontSize, 12
        .WriteProperty "curPen", m_curPen, vbNullString
    End With
End Sub

'Whenever a control property changes that affects control size or layout (including internal changes, like caption adjustments),
' call this function to recalculate the control's internal layout
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Next, determine the positioning of the caption, if present.  (ucSupport.GetCaptionBottom tells us where the
    ' caption text ends vertically.)
    If ucSupport.IsCaptionActive Then
        
        'The clickable area is placed relative to the caption
        With m_PenRect
            .Left = FixDPI(8)
            .Top = ucSupport.GetCaptionBottom + 2
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    'If there's no caption, allow the clickable portion to fill the entire control
    Else
        
        With m_PenRect
            .Left = 1
            .Top = 1
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    End If
            
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer()
        
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True)
    If (bufferDC = 0) Then Exit Sub
    
    'NOTE: if a caption exists, it has already been drawn.  We just need to draw the clickable brush portion.
    If PDMain.IsProgramRunning() Then
        
        Dim cSurface As pd2DSurface, cPen As pd2DPen
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, True
        
        'Paint a checkerboard background first.  (Note that this brush is cached globally, so we never have to
        ' create our own version of it.)
        PD2D.FillRectangleF_FromRectF cSurface, g_CheckerboardBrush, m_PenRect
        
        'Next, create a matching GDI+ pen
        Set cPen = New pd2DPen
        cPen.SetPenPropertiesFromXML Me.Pen
        
        'If the pen is too large, it won't render correctly inside the preview area.  Cap the size at a reasonable amount.
        Dim smallestDimension As Single
        smallestDimension = m_PenRect.Width
        If (m_PenRect.Height < m_PenRect.Width) Then smallestDimension = m_PenRect.Height
        smallestDimension = smallestDimension / 3!
        If (cPen.GetPenWidth > smallestDimension) Then cPen.SetPenWidth smallestDimension
        
        'Prep the preview path.  Note that we manually pad it to make the preview look a little prettier.
        Dim hPadding As Single, vPadding As Single
        hPadding = cPen.GetPenWidth * 2#
        If (hPadding > FixDPIFloat(12)) Then hPadding = FixDPIFloat(12)
        vPadding = hPadding
        
        m_PreviewPath.ResetPath
        m_PreviewPath.CreateSamplePathForRect m_PenRect, hPadding, vPadding
        
        cSurface.SetSurfaceClip_FromRectF m_PenRect, P2_CM_Replace
        PD2D.DrawPath cSurface, cPen, m_PreviewPath
        
        'Before drawing borders around the pen results, ask our parent control to apply color-management to
        ' the pen preview.  (Note that this *will* result in the background checkerboard being color-managed.
        ' This isn't ideal, but we'll live with it for now as the alternative is messy.)
        ucSupport.RequestBufferColorManagement VarPtr(m_PenRect)
        
        'Draw borders around the brush results.
        Dim outlineColor As Long, outlineWidth As Long
        outlineColor = m_Colors.RetrieveColor(PDPS_Border, Me.Enabled, m_MouseDownPenRect, m_MouseInsidePenRect Or ucSupport.DoIHaveFocus)
        If m_MouseInsidePenRect Or ucSupport.DoIHaveFocus Then outlineWidth = 3 Else outlineWidth = 1
        
        Drawing2D.QuickCreateSolidPen cPen, outlineWidth, outlineColor
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.SetSurfaceClip_None
        PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_PenRect
        
        Set cSurface = Nothing: Set cPen = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'If a pen selection dialog is active, it will pass pen updates backward to this function, so that we can let
' our parent form display live updates *while the user is playing with pens* - very cool!
Public Sub NotifyOfLivePenChange(ByVal newPen As String)
    m_ChangeIsLive = True
    Me.Pen = newPen
    m_ChangeIsLive = False
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    m_Colors.LoadThemeColor PDPS_Border, "Border", IDE_BLACK
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    If ucSupport.ThemeUpdateRequired Then
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
    End If
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
