VERSION 5.00
Begin VB.UserControl pdPreview 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdPreview.ctx":0000
End
Attribute VB_Name = "pdPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Preview custom control
'Copyright 2013-2026 by Tanner Helland
'Created: 10/January/13
'Last updated: 09/August/19
'Last update: render a highlight+chunky border on keyboard focus
'
'For implementation details, please refer to pdFxPreviewCtl, which was the original source of most of this
' control's source code.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Preview boxes let the user switch between "full image" and "100% zoom" states; we have to let the caller know about
' these events, because a new effect preview must be generated when they change.
Public Event ViewportChanged()
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Some preview boxes will let the user click to set a new centerpoint for a filter or effect.
Public Event PointSelected(xRatio As Double, yRatio As Double)
Private m_PointSelectionAllowed As Boolean

'Some preview boxes allow the user to click and select a color from the source image
Public Event ColorSelected()
Private m_ColorSelectionAllowed As Boolean
Private m_curColor As Long, m_colorJustClicked As Long

'Because some tools believe they are always operating on a full image (e.g. perspective transform), it may be necessary
' to disable zoom toggle on those controls
Private m_disableZoomPanAbility As Boolean

'Has this control been given a copy of the original image?
Private m_HasOriginal As Boolean, m_HasFX As Boolean

'Copies of the "before" and "after" effects.  We store these internally so the user can switch between them without
' needing to invoke the underlying effect (which may be time-consuming).
Private m_OriginalImage As pdDIB, m_fxImage As pdDIB

'The control's current state: whether it is showing the original image or the fx preview, and whether the preview
' is shown at 100% zoom or "fit-to-window" zoom
Private m_ShowOriginalInstead As Boolean, m_ViewportFitMode As Boolean

'If the viewport is not set to "fit 100%", the user can click-drag around the image.  To do this successfully,
' we must track mouse position and offsets.
Private m_InitX As Long, m_InitY As Long
Private m_OffsetX As Long, m_OffsetY As Long

'Various scroll-related settings, including booleans that determine whether the user is even allowed to pan the
' preview area.
Private m_SrcImageWidth As Long, m_SrcImageHeight As Long
Private m_HScrollMax As Single, m_HScrollValue As Single
Private m_VScrollMax As Single, m_VScrollValue As Single
Private m_HScrollAllowed As Boolean, m_VScrollAllowed As Boolean

'This UniqueID is generated when the UC is first shown.  Any actions that cause the preview area to change
' (e.g. changing zoom, panning the image, etc) cause the ID to change.  This value is used by the EffectPrep module
' when generating a base preview DIB; if the UniqueID hasn't changed since the last request, the previous base preview
' DIB is copied instead of generating a new one from scratch.
Private m_UniqueID As Double

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDPREVIEW_COLOR_LIST
    [_First] = 0
    PDP_Background = 0
    PDP_PreviewBackground = 1
    PDP_PreviewBorder = 2
    [_Last] = 2
    [_Count] = 3
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'The previewed image sits slightly within the control boundaries (leaving room for a 1px border)
Private m_PreviewAreaWidth As Long, m_PreviewAreaHeight As Long
Private Const BORDER_PADDING As Long = 2&

'Some pd2D rendering items are cached to improve preview performance
Private m_Brush As pd2DBrush, m_Pen As pd2DPen

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_Preview
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'At design-time, use this property to determine whether the user is allowed to select colors directly from the
' preview window (helpful for tools like green screen, etc).
Public Property Get AllowColorSelection() As Boolean
    AllowColorSelection = m_ColorSelectionAllowed
End Property

Public Property Let AllowColorSelection(ByVal isAllowed As Boolean)
    m_ColorSelectionAllowed = isAllowed
    PropertyChanged "AllowColorSelection"
End Property

'At design-time, use this property to determine whether the user is allowed to select new center points for a filter
' or effect by clicking the preview window.
Public Property Get AllowPointSelection() As Boolean
    AllowPointSelection = m_PointSelectionAllowed
End Property

Public Property Let AllowPointSelection(ByVal isAllowed As Boolean)
    m_PointSelectionAllowed = isAllowed
    PropertyChanged "AllowPointSelection"
End Property

'At design-time, use this property to prevent the user from changing the preview area between zoom/pan and fit mode.
Public Property Get AllowZoomPan() As Boolean
    AllowZoomPan = Not m_disableZoomPanAbility
End Property

Public Property Let AllowZoomPan(ByVal isAllowed As Boolean)
    m_disableZoomPanAbility = Not isAllowed
    PropertyChanged "DisableZoomPan"
    RedrawBackBuffer
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

Public Function GetUniqueID() As Double
    GetUniqueID = m_UniqueID
End Function

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'OffsetX/Y are used when the preview is in 1:1 mode, and the user is allowed to scroll around the underlying image
Public Property Get GetOffsetX() As Long
    If m_HScrollAllowed Then GetOffsetX = ValidateXOffset(m_HScrollValue + m_OffsetX) Else GetOffsetX = 0
End Property

Public Property Get GetOffsetY() As Long
    If m_VScrollAllowed Then GetOffsetY = ValidateYOffset(m_VScrollValue + m_OffsetY) Else GetOffsetY = 0
End Property

'External functions may need to access the color selected by the preview control
Public Property Get SelectedColor() As Long
    SelectedColor = m_curColor
End Property

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

Public Property Let ShowOriginalInstead(ByVal newValue As Boolean)
    m_ShowOriginalInstead = newValue
    RedrawBackBuffer
End Property

Public Property Get ViewportFitFullImage() As Boolean
    ViewportFitFullImage = m_ViewportFitMode
End Property

Public Property Let ViewportFitFullImage(ByVal newState As Boolean)
    
    m_ViewportFitMode = newState
    
    'Note that we no longer have a valid copy of the original image data, so prepImageData must supply us with a new one
    m_HasOriginal = False
    m_HasFX = False
    
    'Change our unique ID value, so the preview engine knows to recreate the base preview DIB
    m_UniqueID = m_UniqueID - 0.1
    
    'Raise a viewport change event so the containing form can redraw itself accordingly
    RaiseEvent ViewportChanged
    
End Property

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

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub ucSupport_VisibilityChange(ByVal newVisibility As Boolean)
    m_UniqueID = Timer
    If (Not newVisibility) Then EffectPrep.ResetPreviewIDs Else RedrawBackBuffer
End Sub

'Use this to supply the preview with a copy of the original image's data.  The preview object can use this to display
' the original image when the user clicks the "show original image" link.
Public Sub SetOriginalImage(ByRef srcDIB As pdDIB)

    'Note that we have a copy of the original image, so the calling function doesn't attempt to supply it again
    m_HasOriginal = True
    
    'Make a copy of the DIB passed in
    If (m_OriginalImage Is Nothing) Then Set m_OriginalImage = New pdDIB
    m_OriginalImage.CreateFromExistingDIB srcDIB
    
    'Apply color management now; then we never have to do it again
    If (Not srcDIB Is Nothing) Then ColorManagement.ApplyDisplayColorManagement m_OriginalImage
    
    If (m_OriginalImage.GetDIBColorDepth = 32) And (Not m_OriginalImage.GetAlphaPremultiplication) Then m_OriginalImage.SetAlphaPremultiplication True
    
End Sub

'Use this to supply the object with a copy of the processed image's data.  The preview object can use this to display
' the processed image again if the user clicks the "show original image" link, then clicks it again.
Public Sub SetFXImage(ByRef srcDIB As pdDIB, Optional ByVal colorManagementAlreadyHandled As Boolean = False)

    'Note that we have a copy of the original image, so the calling function doesn't attempt to supply it again
    m_HasFX = True
    
    'Make a copy of the DIB passed in
    If (m_fxImage Is Nothing) Then Set m_fxImage = New pdDIB
    m_fxImage.CreateFromExistingDIB srcDIB
    
    'Apply color management now; then we never have to do it again
    If (Not srcDIB Is Nothing) And (Not colorManagementAlreadyHandled) Then
        If (m_fxImage.GetDIBWidth <> 0) Then ColorManagement.ApplyDisplayColorManagement m_fxImage
    End If
    
    'Redraw the on-screen image (as necessary)
    RedrawBackBuffer
    
End Sub

'Has this preview control had an original version of the image set?
Public Function HasOriginalImage() As Boolean
    HasOriginalImage = m_HasOriginal
End Function

'Return dimensions of the preview picture box
Public Function GetPreviewWidth() As Long
    GetPreviewWidth = m_PreviewAreaWidth
End Function

Public Function GetPreviewHeight() As Long
    GetPreviewHeight = m_PreviewAreaHeight
End Function

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

    'If viewport scrolling is allowed, initialize it now
    If (Not ViewportFitFullImage) Then
        If (Button = vbLeftButton) Then
            m_InitX = x
            m_InitY = y
            ucSupport.RequestCursor IDC_SIZEALL
        End If
    End If
    
    'If color selection is allowed, initialize it now
    If m_ColorSelectionAllowed Then
        
        If (Button = vbRightButton) Then
            
            'Convert the mouse coordinates to DIB coordinates.
            Dim dibX As Single, dibY As Single
            GetDIBXYFromMouseXY x, y, dibX, dibY, True
            
            Dim cRGBA As RGBQuad
            If m_OriginalImage.GetPixelRGBQuad(dibX, dibY, cRGBA) Then
                m_curColor = RGB(cRGBA.Red, cRGBA.Green, cRGBA.Blue)
                If AllowColorSelection Then m_colorJustClicked = 1
                RaiseEvent ColorSelected
            End If
            
        End If
        
    End If
    
    EvaluateMouseEvent Button, x, y

End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'If this preview control instance allows the user to select a color, display the original image upon mouse entrance
    Dim forciblyShowOriginalImage As Boolean
    forciblyShowOriginalImage = False
    
    If ViewportFitFullImage Then
        If AllowColorSelection Then
            forciblyShowOriginalImage = True
            ucSupport.RequestCursor_Resource "cursor_eyedropper", 0, 16
        Else
            ucSupport.RequestCursor IDC_DEFAULT
        End If
    Else
        ucSupport.RequestCursor IDC_HAND
    End If
    
    RedrawBackBuffer forciblyShowOriginalImage

End Sub

'If this preview control instance allows the user to select a color, restore whatever image was previously
' displayed upon mouse exit
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    ucSupport.RequestCursor IDC_DEFAULT
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'If the viewport is not set to "fit to screen", then we must determine offsets based on the mouse position
    If (Not ViewportFitFullImage) Then
        
        If (Button = vbLeftButton) Then
        
            'Make sure the move cursor remains accurate
            ucSupport.RequestCursor IDC_SIZEALL
                
            'Store new offsets for the image
            m_OffsetX = m_InitX - x
            m_OffsetY = m_InitY - y
            
            'Note that we no longer have a valid copy of the original image data, so prepImageData must supply us with a new one
            m_HasOriginal = False
            m_HasFX = False
            
            'Change our unique ID value, so the preview engine knows to recreate the base preview DIB
            m_UniqueID = m_UniqueID - 0.1
            
            'Raise an external viewport change event that tool dialogs can use to refresh their effect preview
            RaiseEvent ViewportChanged
            
        Else
            If Not m_ColorSelectionAllowed Then ucSupport.RequestCursor IDC_HAND
        End If
        
    End If
    
    If (m_colorJustClicked > 0) Then
    
        'To accomodate shaky hands, allow a few mouse movements before resetting the image
        If (m_colorJustClicked < 4) Then
            m_colorJustClicked = m_colorJustClicked + 1
        Else
            m_colorJustClicked = 0
            RedrawBackBuffer True
        End If
        
    End If
    
    EvaluateMouseEvent Button, x, y

End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)

    If (Not ViewportFitFullImage) Then
        ucSupport.RequestCursor IDC_HAND
        
        m_HScrollValue = ValidateXOffset(m_HScrollValue + m_OffsetX)
        m_OffsetX = 0
        m_VScrollValue = ValidateYOffset(m_VScrollValue + m_OffsetY)
        m_OffsetY = 0
    End If
    
    EvaluateMouseEvent Button, x, y

End Sub

Private Sub EvaluateMouseEvent(ByVal Button As PDMouseButtonConstants, ByVal x As Long, ByVal y As Long)
    
    'If point selection is allowed, evaluate this mouse position for a potential "update preview" event
    If m_PointSelectionAllowed Then
        
        ucSupport.RequestCursor IDC_HAND
        
        If (Button = vbRightButton) Or (Button = vbLeftButton) Then
        
            'Return the mouse coordinates as a ratio between 0 and 1, with 1 representing max width/height
            Dim retX As Double, retY As Double
            retX = (x - BORDER_PADDING) - ((m_PreviewAreaWidth - m_OriginalImage.GetDIBWidth) \ 2)
            retY = (y - BORDER_PADDING) - ((m_PreviewAreaHeight - m_OriginalImage.GetDIBHeight) \ 2)
            
            retX = retX / m_OriginalImage.GetDIBWidth
            retY = retY / m_OriginalImage.GetDIBHeight
            
            RaiseEvent PointSelected(retX, retY)
        
        End If
    
    End If
    
End Sub

'Given a mouse (x, y) coordinate pair, return a matching (x, y) pair for the underlying DIB.  Fit/zoom are automatically considered.
Private Sub GetDIBXYFromMouseXY(ByVal mouseX As Single, ByVal mouseY As Single, ByRef dibX As Single, ByRef dibY As Single, Optional ByVal useOriginalDIB As Boolean = True)
    
    'Because the caller may want coordinates from the original *OR* modified DIB, we use a generic reference.
    Dim srcDIB As pdDIB
    If useOriginalDIB Then Set srcDIB = m_OriginalImage Else Set srcDIB = m_fxImage
    
    'If the image is using "fit within" mode, we need to perform some extra coordinate math.
    Dim dstWidth As Double, dstHeight As Double
    dstWidth = m_PreviewAreaWidth
    dstHeight = m_PreviewAreaHeight
    
    Dim srcWidth As Double, srcHeight As Double
    srcWidth = srcDIB.GetDIBWidth
    srcHeight = srcDIB.GetDIBHeight
    
    'Calculate the aspect ratio of this DIB and the target picture box
    Dim srcAspect As Double, dstAspect As Double
    If (srcHeight > 0#) Then srcAspect = srcWidth / srcHeight Else srcAspect = 1#
    If (dstHeight > 0#) Then dstAspect = dstWidth / dstHeight Else dstAspect = 1#
        
    Dim finalWidth As Long, finalHeight As Long
    If (dstWidth <= srcWidth) Or (dstHeight <= srcHeight) Or Me.ViewportFitFullImage Then
        PDMath.ConvertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, finalWidth, finalHeight
    Else
        finalWidth = srcWidth
        finalHeight = srcHeight
    End If
        
    'Images smaller than the target area in one (or more) dimensions need to be centered in the target area
    Dim previewX As Long, previewY As Long
    If (srcAspect > dstAspect) Then
        previewY = Int((dstHeight - finalHeight) / 2) + BORDER_PADDING
        If (finalWidth = dstWidth) Then
            previewX = BORDER_PADDING
        Else
            previewX = Int((dstWidth - finalWidth) / 2) + BORDER_PADDING
        End If
    Else
        previewX = Int((dstWidth - finalWidth) / 2) + BORDER_PADDING
        If (finalHeight = dstHeight) Then
            previewY = BORDER_PADDING
        Else
            previewY = Int((dstHeight - finalHeight) / 2) + BORDER_PADDING
        End If
    End If
    
    'We now have an original DIB width/height pair, destination DIB width/height pair, preview (x, y) offset - all that's left
    ' is a source (x, y) offset.
    Dim srcX As Single, srcY As Single
    srcX = Me.GetOffsetX
    srcY = Me.GetOffsetY
    
    'Convert the destination (x, y) pair to the [0, 1] range.
    Dim dstX As Single, dstY As Single
    dstX = ((mouseX - previewX) / CDbl(finalWidth))
    dstY = ((mouseY - previewY) / CDbl(finalHeight))
    
    'Map it into the source range
    dibX = (dstX * srcWidth) + srcX
    dibY = (dstY * srcHeight) + srcY
    
End Sub

'X and Y offsets for the image preview are generated dynamically by the user's mouse movements.  As multiple functions
' need to validate those offsets to make sure they don't result in an offset outside the image, these standardized
' validation functions were created.
Private Function ValidateXOffset(ByVal currentOffset As Long) As Long
    If (currentOffset < 0) Then currentOffset = 0
    If (currentOffset > m_HScrollMax) Then currentOffset = m_HScrollMax
    ValidateXOffset = currentOffset
End Function

Private Function ValidateYOffset(ByVal currentOffset As Long) As Long
    If (currentOffset < 0) Then currentOffset = 0
    If (currentOffset > m_VScrollMax) Then currentOffset = m_VScrollMax
    ValidateYOffset = currentOffset
End Function

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True
    ucSupport.RequestCaptionSupport
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDPREVIEW_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDPreview", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
    m_ShowOriginalInstead = False
    m_curColor = 0
    m_ViewportFitMode = True
            
End Sub

Private Sub UserControl_InitProperties()
    
    m_HasOriginal = False
    
    'By default, the control *allows* the user to zoom/pan the transformation
    m_disableZoomPanAbility = False
    
    'By default, the control does *not* allow the user to select coordinate points or colors by clicking the preview area
    m_ColorSelectionAllowed = False
    m_PointSelectionAllowed = False
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        AllowColorSelection = .ReadProperty("ColorSelection", False)
        AllowPointSelection = .ReadProperty("PointSelection", False)
        AllowZoomPan = Not .ReadProperty("DisableZoomPan", False)
    End With
End Sub

'Redraw the user control after it has been resized
Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then UpdateControlLayout
End Sub

Private Sub UserControl_Show()
    
    'Generate a unique ID for this session
    m_UniqueID = Timer
    
    'Determine acceptable max/min scroll values for 100% zoom preview mode
    If PDMain.IsProgramRunning() Then
        
        ucSupport.RequestCursor IDC_DEFAULT
        
        'Some dialogs can use this control instance even without a loaded image in the main window.
        ' (e.g. the JPEG export dialog, which can be raised from the batch processor).  On such dialogs,
        ' a preview image - if any - will be loaded via the NotifyNonStandardSource function, which in turn
        ' will modify m_SrcImageWidth/Height.  If those are non-zero, assume a non-standard source, and do not
        ' auto-load dimensions from the active window.
        If PDImages.IsImageActive() And (m_SrcImageWidth = 0) And (m_SrcImageWidth = 0) Then
            If (PDImages.GetActiveImage.GetNumOfLayers <> 0) Then
                
                'An active layer is being used for this preview instance.  If a selection is active,
                ' we need to know its size so that we can calculate zoom accordingly.
                If PDImages.GetActiveImage.IsSelectionActive Then
                    Dim selBounds As RectF
                    selBounds = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect
                    m_SrcImageWidth = selBounds.Width
                    m_SrcImageHeight = selBounds.Height
                Else
                    m_SrcImageWidth = PDImages.GetActiveImage.GetActiveDIB.GetDIBWidth
                    m_SrcImageHeight = PDImages.GetActiveImage.GetActiveDIB.GetDIBHeight
                End If
                
            End If
        End If
        
        CalculateScrollMaxMin
        
    End If
    
    'Ensure the control is redrawn at least once
    UpdateControlLayout
    
End Sub

Private Sub UserControl_Terminate()
    If Not (m_OriginalImage Is Nothing) Then m_OriginalImage.EraseDIB
    If Not (m_fxImage Is Nothing) Then m_fxImage.EraseDIB
    EffectPrep.ResetPreviewIDs
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ColorSelection", AllowColorSelection, False
        .WriteProperty "DisableZoomPan", Not AllowZoomPan, False
        .WriteProperty "PointSelection", AllowPointSelection, False
    End With
End Sub

Public Sub NotifyNonStandardSource(ByVal srcWidth As Long, ByVal srcHeight As Long)
    m_SrcImageWidth = srcWidth
    m_SrcImageHeight = srcHeight
    CalculateScrollMaxMin
End Sub

Public Sub RequestImmediateRefresh()
    UpdateControlLayout True
End Sub

Private Sub CalculateScrollMaxMin()
    
    Dim maxHOffset As Long, maxVOffset As Long
    maxHOffset = m_SrcImageWidth - m_PreviewAreaWidth
    maxVOffset = m_SrcImageHeight - m_PreviewAreaHeight
    
    If (maxHOffset > 0) Then
        m_HScrollMax = maxHOffset
        m_HScrollAllowed = True
    Else
        m_HScrollMax = 1
        m_HScrollAllowed = False
    End If
    
    If (maxVOffset > 0) Then
        m_VScrollMax = maxVOffset
        m_VScrollAllowed = True
    Else
        m_VScrollMax = 1
        m_VScrollAllowed = False
    End If
        
End Sub

'After a resize or paint request, update the layout of our control
Private Sub UpdateControlLayout(Optional ByVal forceResetOfSizeParams As Boolean = False)
    
    Dim origWidth As Long, origHeight As Long
    origWidth = m_PreviewAreaWidth
    origHeight = m_PreviewAreaHeight
    
    'Cache DPI-aware control dimensions from the support class
    m_PreviewAreaWidth = ucSupport.GetControlWidth - BORDER_PADDING * 2
    m_PreviewAreaHeight = ucSupport.GetControlHeight - BORDER_PADDING * 2
    
    CalculateScrollMaxMin
    
    'If our original and new sizes don't match, request an immediate redraw.
    ' (We also need to change our unique ID, so the preview engine knows to recreate the base preview DIB;
    ' the forceResetOfSizeParams will be set to TRUE if our parent window just notified us of a window-resize-end
    ' event via WM_EXITSIZEMOVE message.)
    If (origWidth <> m_PreviewAreaWidth) Or (origHeight <> m_PreviewAreaHeight) Or forceResetOfSizeParams Then
        
        If (Not Interface.GetDialogResizeFlag()) Or forceResetOfSizeParams Then
            m_HasOriginal = False
            m_HasFX = False
            m_UniqueID = m_UniqueID - 0.1
            RaiseEvent ViewportChanged
        End If
        
    End If
    
End Sub

'This control currently handles border rendering around the preview area, so it *does* maintain a backbuffer that
' may need to be redrawn under certain circumstances.
Private Sub RedrawBackBuffer(Optional ByVal overrideWithOriginalImage As Boolean = False)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDP_Background))
    If (bufferDC = 0) Then Exit Sub
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Because the source of rendering may change, we use a temporary reference
    Dim srcDIB As pdDIB
    
    'If the user was previously examining the original image, and color selection is not allowed, be helpful and
    ' automatically restore the previewed image.
    If (m_ShowOriginalInstead Or overrideWithOriginalImage) Then
        If m_HasOriginal Then Set srcDIB = m_OriginalImage Else Set srcDIB = m_fxImage
    Else
        If m_HasFX Then Set srcDIB = m_fxImage Else Set srcDIB = m_OriginalImage
    End If
    
    'It's entirely possible for srcDIB to be nothing, particularly inside the IDE, so this check is necessary.
    If (Not srcDIB Is Nothing) Then
        
        'srcDIB points at either the original or effect image.
        ' (We don't care which; both are rendered identically.)
        
        'Start by calculating a target buffer region.
        ' (If the preview control is set to "fit" mode, we want to center it in the preview area.)
        Dim dstWidth As Double, dstHeight As Double
        dstWidth = m_PreviewAreaWidth
        dstHeight = m_PreviewAreaHeight
        
        Dim srcWidth As Double, srcHeight As Double
        srcWidth = srcDIB.GetDIBWidth
        srcHeight = srcDIB.GetDIBHeight
        
        'Calculate the aspect ratio of this DIB and the target picture box
        Dim srcAspect As Double, dstAspect As Double
        If (srcHeight > 0#) Then srcAspect = srcWidth / srcHeight Else srcAspect = 1#
        If (dstHeight > 0#) Then dstAspect = dstWidth / dstHeight Else dstAspect = 1#
            
        Dim finalWidth As Long, finalHeight As Long
        If (dstWidth <= srcWidth) Or (dstHeight <= srcHeight) Or Me.ViewportFitFullImage Then
            PDMath.ConvertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, finalWidth, finalHeight
        Else
            finalWidth = srcWidth
            finalHeight = srcHeight
        End If
        
        'Images smaller than the target area in one (or more) dimensions need to be centered in the target area
        Dim previewX As Long, previewY As Long
        If (srcAspect > dstAspect) Then
            previewY = Int((dstHeight - finalHeight) / 2 + 0.5) + BORDER_PADDING
            If (finalWidth = dstWidth) Then
                previewX = BORDER_PADDING
            Else
                previewX = Int((dstWidth - finalWidth) / 2 + 0.5) + BORDER_PADDING
            End If
        Else
            previewX = Int((dstWidth - finalWidth) / 2 + 0.5) + BORDER_PADDING
            If (finalHeight = dstHeight) Then
                previewY = BORDER_PADDING
            Else
                previewY = Int((dstHeight - finalHeight) / 2 + 0.5) + BORDER_PADDING
            End If
        End If
        
        Dim ctlBackColor As Long, ctlBorderColor As Long
        ctlBackColor = m_Colors.RetrieveColor(PDP_PreviewBackground, True)
        
        'When the control is hovered, we use an accent color and a chunky border to convey it to the user
        Dim hoverMatters As Boolean, borderWidth As Single
        borderWidth = 1!
        hoverMatters = AllowColorSelection Or AllowPointSelection Or (AllowZoomPan And Not ViewportFitFullImage)
        If hoverMatters Then
            hoverMatters = ucSupport.IsMouseInside Or ucSupport.DoIHaveFocus
            If hoverMatters Then borderWidth = 3!
        End If
        ctlBorderColor = m_Colors.RetrieveColor(PDP_PreviewBorder, True, , hoverMatters)
        
        'Use pd2D to fill the background of this window with the background color, followed by a checkerboard
        Dim dstSurface As pd2DSurface: Set dstSurface = New pd2DSurface
        dstSurface.WrapSurfaceAroundDC bufferDC
        dstSurface.SetSurfaceAntialiasing P2_AA_None
        dstSurface.SetSurfacePixelOffset P2_PO_Normal
        
        If (m_Brush Is Nothing) Then Set m_Brush = New pd2DBrush
        m_Brush.SetBrushColor ctlBackColor
        PD2D.FillRectangleI dstSurface, m_Brush, BORDER_PADDING, BORDER_PADDING, m_PreviewAreaWidth, m_PreviewAreaHeight
        PD2D.FillRectangleI dstSurface, g_CheckerboardBrush, previewX, previewY, finalWidth, finalHeight
        
        'Enable high-quality stretching, but only if the image is equal to or larger than the preview area
        Dim isZoomedIn As Boolean
        isZoomedIn = (srcWidth < dstWidth) And (srcHeight < dstHeight)
        If isZoomedIn Then
            dstSurface.SetSurfacePixelOffset P2_PO_Half
            dstSurface.SetSurfaceResizeQuality P2_RQ_Fast
        Else
            dstSurface.SetSurfacePixelOffset P2_PO_Normal
            dstSurface.SetSurfaceResizeQuality P2_RQ_Bicubic
        End If
        
        Dim srcSurface As pd2DSurface: Set srcSurface = New pd2DSurface
        srcSurface.WrapSurfaceAroundPDDIB srcDIB
        srcSurface.SetSurfaceAntialiasing P2_AA_None
        srcSurface.SetSurfacePixelOffset P2_PO_Normal
        PD2D.DrawSurfaceResizedCroppedI dstSurface, previewX, previewY, finalWidth, finalHeight, srcSurface, 0, 0, srcWidth, srcHeight
        Set srcSurface = Nothing
        srcDIB.FreeFromDC
        
        'We also draw a border around the final result
        Dim halfBorder As Long
        halfBorder = Int(BORDER_PADDING / 2 + 0.5)
        
        If (m_Pen Is Nothing) Then
            Set m_Pen = New pd2DPen
            m_Pen.SetPenLineJoin P2_LJ_Miter
        End If
        m_Pen.SetPenColor ctlBorderColor
        m_Pen.SetPenWidth borderWidth
        
        dstSurface.SetSurfacePixelOffset P2_PO_Normal
        PD2D.DrawRectangleI_AbsoluteCoords dstSurface, m_Pen, halfBorder, halfBorder, (bWidth - 1) - halfBorder, (bHeight - 1) - halfBorder
        Set dstSurface = Nothing
        
        'Paint the results to the screen!  (Note that we request an immediate redraw, rather than waiting for WM_PAINT to fire.)
        If PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
        
    End If

End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDP_Background, "Background", IDE_WHITE
        .LoadThemeColor PDP_PreviewBackground, "PreviewBackground", IDE_GRAY
        .LoadThemeColor PDP_PreviewBorder, "PreviewBorder", IDE_BLACK
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
