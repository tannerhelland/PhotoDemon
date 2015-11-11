VERSION 5.00
Begin VB.UserControl fxPreviewCtl 
   AccessKeys      =   "T"
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   ToolboxBitmap   =   "fxPreview.ctx":0000
   Begin PhotoDemon.buttonStrip btsZoom 
      Height          =   495
      Left            =   2985
      TabIndex        =   1
      Top             =   5160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      FontSize        =   8
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5100
      Left            =   0
      ScaleHeight     =   338
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   382
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   5760
      Begin VB.VScrollBar vsOffsetY 
         Height          =   1335
         Left            =   5280
         TabIndex        =   4
         Top             =   3360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar hsOffsetX 
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   4680
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin PhotoDemon.buttonStrip btsState 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      FontSize        =   8
   End
End
Attribute VB_Name = "fxPreviewCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Effect Preview custom control
'Copyright 2013-2015 by Tanner Helland
'Created: 10/January/13
'Last updated: 05/September/15
'Last update: overhaul drawing internals; the new version should be much faster, particularly in 1:1 zoom mode
'
'For the first decade of its life, PhotoDemon relied on simple picture boxes for rendering its effect previews.
' This worked well enough when there were only a handful of tools available, but as the complexity of the program
' - and its various effects and tools - has grown, it has become more and more painful to update the preview
' system, because any changes have to be mirrored across a huge number of forms.
'
'Thus, this control was born.  It is now used on every single effect form in place of a regular picture box.  This
' allows me to add preview-related features just once - to the base control - and have every tool automatically
' reap the benefits.
'
'The control is capable of storing a copy of the original image and any filter-modified versions of the image.
' The user can toggle between these by using the command link below the main picture box, or by pressing Alt+T.
' This replaces the side-by-side "before and after" of past versions.
'
'A few other extra features have been implemented, which can be enabled on a tool-by-tool basis.  Specifically:
' 1) The user can toggle between "fit image" and "100% zoom + click-drag-to-scroll" modes.  Note that 100% zoom
'    is not appropriate for some tools (i.e. perspective transformations and other algorithms that only operate
'    on the full image area).
' 2) Click-to-select color functionality.  This is helpful for tools that rely on color information within the
'    image for their operation, e.g. green screen.
' 3) Click-to-select-coordinate functionality.  This is helpful for giving the user an easy way to select a
'    location on the image as, say, a center point for a filter (e.g. vignetting works great with this).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Preview boxes can now let the user switch between "full image" and "100% zoom" states
Public Event ViewportChanged()

'Some preview boxes will let the user click to set a new centerpoint for a filter or effect.
Public Event PointSelected(xRatio As Double, yRatio As Double)
Private isPointSelectionAllowed As Boolean

'Some preview boxes allow the user to click and select a color from the source image
Public Event ColorSelected()
Private isColorSelectionAllowed As Boolean, curColor As Long
Private colorJustClicked As Long

'Flicker-free window painter for the preview area
Private WithEvents cPainter As pdWindowPainter
Attribute cPainter.VB_VarHelpID = -1

'Because some tools believe they are always operating on a full image (e.g. perspective transform), it may be necessary
' to disable zoom toggle on those controls
Private disableZoomPanAbility As Boolean

'Has this control been given a copy of the original image?
Private m_HasOriginal As Boolean, m_HasFX As Boolean

'Copies of the "before" and "after" effects.  We store these internally so the user can switch between them without
' needing to invoke the underlying effect (which may be time-consuming).
Private originalImage As pdDIB, fxImage As pdDIB

'As of PD 7.0, this control now manages its own backbuffer.  This complicates the control a bit, but it gives us more
' fine-tuned control over performance, and it will allow us to provide more preview-related features in the future.
Private m_BackBuffer As pdDIB

'The control's current state: whether it is showing the original image or the fx preview
Private m_ShowOriginalInstead As Boolean

'GetPixel is used to retrieve colors from the image
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

'Mouse events are raised with the help of the pdInputMouse class
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

'If the viewport is not set to "fit 100%", the user can click-drag around the image.  To do this successfully,
' we must track mouse position and offsets.
Private m_InitX As Long, m_InitY As Long
Private m_OffsetX As Long, m_OffsetY As Long

'Is the image large enough that the user is allowed to scroll?
Private m_HScrollAllowed As Boolean, m_VScrollAllowed As Boolean

'This UniqueID is generated when the UC is first shown.  Any actions that cause the preview area to change
' (e.g. changing zoom, panning the image, etc) cause the ID to change.  This value is used by the FastDrawing module
' when generating a base preview DIB; if the UniqueID hasn't changed since the last request, the previous base preview
' DIB is copied instead of generating a new one from scratch.
Private m_UniqueID As Double

Public Function getUniqueID() As Double
    getUniqueID = m_UniqueID
End Function

'OffsetX/Y are used when the preview is in 1:1 mode, and the user is allowed to scroll around the underlying image
Public Property Get offsetX() As Long
    If m_HScrollAllowed Then
        offsetX = validateXOffset(hsOffsetX.Value + m_OffsetX)
    Else
        offsetX = 0
    End If
End Property

Public Property Get offsetY() As Long
    If m_VScrollAllowed Then
        offsetY = validateYOffset(vsOffsetY.Value + m_OffsetY)
    Else
        offsetY = 0
    End If
End Property

Public Property Get viewportFitFullImage() As Boolean
    viewportFitFullImage = CBool(btsZoom.ListIndex = 1)
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'External functions may need to access the color selected by the preview control
Public Property Get SelectedColor() As Long
    SelectedColor = curColor
End Property

'At design-time, use this property to determine whether the user is allowed to select colors directly from the
' preview window (helpful for tools like green screen, etc).
Public Property Get AllowColorSelection() As Boolean
    AllowColorSelection = isColorSelectionAllowed
End Property

Public Property Let AllowColorSelection(ByVal isAllowed As Boolean)
    isColorSelectionAllowed = isAllowed
    PropertyChanged "AllowColorSelection"
End Property

'At design-time, use this property to determine whether the user is allowed to select new center points for a filter
' or effect by clicking the preview window.
Public Property Get AllowPointSelection() As Boolean
    AllowPointSelection = isPointSelectionAllowed
End Property

Public Property Let AllowPointSelection(ByVal isAllowed As Boolean)
    isPointSelectionAllowed = isAllowed
    PropertyChanged "AllowPointSelection"
End Property

'At design-time, use this property to prevent the user from changing the preview area between zoom/pan and fit mode.
Public Property Get AllowZoomPan() As Boolean
    AllowZoomPan = Not disableZoomPanAbility
End Property

Public Property Let AllowZoomPan(ByVal isAllowed As Boolean)
    disableZoomPanAbility = Not isAllowed
    PropertyChanged "DisableZoomPan"
    redrawControl
    UserControl.Refresh
End Property

'Use this to supply the preview with a copy of the original image's data.  The preview object can use this to display
' the original image when the user clicks the "show original image" link.
Public Sub setOriginalImage(ByRef srcDIB As pdDIB)

    'Note that we have a copy of the original image, so the calling function doesn't attempt to supply it again
    m_HasOriginal = True
    
    'Make a copy of the DIB passed in
    If (originalImage Is Nothing) Then Set originalImage = New pdDIB
    originalImage.createFromExistingDIB srcDIB
    
    If (originalImage.getDIBColorDepth = 32) And (Not originalImage.getAlphaPremultiplication) Then originalImage.setAlphaPremultiplication True
    
End Sub

'Use this to supply the object with a copy of the processed image's data.  The preview object can use this to display
' the processed image again if the user clicks the "show original image" link, then clicks it again.
Public Sub setFXImage(ByRef srcDIB As pdDIB)

    'Note that we have a copy of the original image, so the calling function doesn't attempt to supply it again
    m_HasFX = True
    
    'Make a copy of the DIB passed in
    If (fxImage Is Nothing) Then Set fxImage = New pdDIB
    fxImage.createFromExistingDIB srcDIB
    
    'Redraw the on-screen image (as necessary)
    syncPreviewImage
    
End Sub

'Render the currently active image to the preview window.  This bares some similarity to the pdDIB.renderToPictureBox function,
' but is optimized for the unique concerns of this control.
Private Sub syncPreviewImage(Optional ByVal overrideWithOriginalImage As Boolean = False)
    
    'Because the source of rendering may change, we use a temporary reference
    Dim srcDIB As pdDIB
    
    'If the user was previously examining the original image, and color selection is not allowed, be helpful and
    ' automatically restore the previewed image.
    If m_ShowOriginalInstead Or overrideWithOriginalImage Then
        If m_HasOriginal Then
            Set srcDIB = originalImage
        Else
            Set srcDIB = fxImage
        End If
    Else
        If m_HasFX Then
            Set srcDIB = fxImage
        Else
            Set srcDIB = originalImage
        End If
    End If
    
    'If we have nothing to render, exit now
    If Not (srcDIB Is Nothing) Then
        
        'srcDIB points at either the original or effect image.  We don't care which, as we render them identically.
        
        'Start by calculating a target buffer region.  If the preview control is set to "fit" mode, we want to center
        ' it in the preview area.
        Dim dstWidth As Double, dstHeight As Double
        dstWidth = m_BackBuffer.getDIBWidth
        dstHeight = m_BackBuffer.getDIBHeight
        
        Dim srcWidth As Double, srcHeight As Double
        srcWidth = srcDIB.getDIBWidth
        srcHeight = srcDIB.getDIBHeight
        
        'Calculate the aspect ratio of this DIB and the target picture box
        Dim srcAspect As Double, dstAspect As Double
        If srcHeight > 0 Then srcAspect = srcWidth / srcHeight Else srcAspect = 1
        If dstHeight > 0 Then dstAspect = dstWidth / dstHeight Else dstAspect = 1
            
        Dim finalWidth As Long, finalHeight As Long
        If (dstWidth <= srcWidth) Or (dstHeight <= srcHeight) Or Me.viewportFitFullImage Then
            convertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, finalWidth, finalHeight
        Else
            finalWidth = srcWidth
            finalHeight = srcHeight
        End If
        
        'Images smaller than the target area in one (or more) dimensions need to be centered in the target area
        Dim previewX As Long, previewY As Long
        If srcAspect > dstAspect Then
            previewY = CLng((dstHeight - finalHeight) / 2)
            
            If finalWidth = dstWidth Then
                previewX = 0
            Else
                previewX = CLng((dstWidth - finalWidth) / 2)
            End If
        Else
            previewX = CLng((dstWidth - finalWidth) / 2)
            
            If finalHeight = dstHeight Then
                previewY = 0
            Else
                previewY = CLng((dstHeight - finalHeight) / 2)
            End If
        End If
        
        'We now have a set of source and destination coordinates, allowing us to perform a StretchBlt-style copy
        GDI_Plus.GDIPlusFillDIBRect m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, picPreview.BackColor
        GDI_Plus.GDIPlusFillDIBRect_Pattern m_BackBuffer, previewX, previewY, finalWidth, finalHeight, g_CheckerboardPattern, , True
        
        'Enable high-quality stretching, but only if the image is equal to or larger than the preview area
        If (srcWidth < dstWidth) And (srcHeight < dstHeight) Then
            GDI_Plus.GDIPlus_StretchBlt m_BackBuffer, previewX, previewY, finalWidth, finalHeight, srcDIB, 0, 0, srcWidth, srcHeight, , InterpolationModeNearestNeighbor
        Else
            GDI_Plus.GDIPlus_StretchBlt m_BackBuffer, previewX, previewY, finalWidth, finalHeight, srcDIB, 0, 0, srcWidth, srcHeight, , InterpolationModeBicubic
        End If
        
        'Paint the results!  (Note that we request an immediate redraw, rather than waiting for WM_PAINT to fire.)
        If g_IsProgramRunning Then cPainter.RequestRepaint True
        
        Set srcDIB = Nothing
        
    End If

End Sub

'Has this preview control had an original version of the image set?
Public Function hasOriginalImage() As Boolean
    hasOriginalImage = m_HasOriginal
End Function

'Return a handle to our primary picture box
Public Function getPreviewPic() As PictureBox
    Set getPreviewPic = picPreview
End Function

'Return dimensions of the preview picture box
Public Function getPreviewWidth() As Long
    getPreviewWidth = picPreview.ScaleWidth
End Function

Public Function getPreviewHeight() As Long
    getPreviewHeight = picPreview.ScaleHeight
End Function

Private Sub btsState_Click(ByVal buttonIndex As Long)
    m_ShowOriginalInstead = CBool(buttonIndex = 0)
    syncPreviewImage
End Sub

Private Sub btsZoom_Click(ByVal buttonIndex As Long)
    
    'Note that we no longer have a valid copy of the original image data, so prepImageData must supply us with a new one
    m_HasOriginal = False
    m_HasFX = False
    
    'Change our unique ID value, so the preview engine knows to recreate the base preview DIB
    m_UniqueID = m_UniqueID - 0.1
    
    'Raise a viewport change event so the containing form can redraw itself accordingly
    RaiseEvent ViewportChanged
    
End Sub

Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'If viewport scrolling is allowed, initialize it now
    If Not viewportFitFullImage Then
        If Button = vbLeftButton Then
            m_InitX = x
            m_InitY = y
            cMouseEvents.setSystemCursor IDC_SIZEALL
        End If
    End If
    
    'If color selection is allowed, initialize it now
    If isColorSelectionAllowed Then
        
        If Button = vbRightButton Then
            
            'Convert the mouse coordinates to DIB coordinates.
            Dim dibX As Single, dibY As Single
            GetDIBXYFromMouseXY x, y, dibX, dibY, True
            
            Dim cRGBA As RGBQUAD
            If originalImage.GetPixelRGBQuad(dibX, dibY, cRGBA) Then
                curColor = RGB(cRGBA.Red, cRGBA.Green, cRGBA.Blue)
                If AllowColorSelection Then colorJustClicked = 1
                RaiseEvent ColorSelected
            End If
            
        End If
        
    End If
    
    'If point selection is allowed, initialize it now
    If isPointSelectionAllowed Then
    
        If (Button = vbRightButton) Or (Button = vbLeftButton) Then
        
            'Return the mouse coordinates as a ratio between 0 and 1, with 1 representing max width/height
            Dim retX As Double, retY As Double
            retX = x - ((picPreview.ScaleWidth - originalImage.getDIBWidth) \ 2)
            retY = y - ((picPreview.ScaleHeight - originalImage.getDIBHeight) \ 2)
            
            retX = retX / originalImage.getDIBWidth
            retY = retY / originalImage.getDIBHeight
            
            RaiseEvent PointSelected(retX, retY)
        
        End If
    
    End If

End Sub

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'If this preview control instance allows the user to select a color, display the original image upon mouse entrance
    If viewportFitFullImage Then
    
        If AllowColorSelection Then
            cMouseEvents.setPNGCursor "C_PIPETTE", 0, 0
            syncPreviewImage True
        Else
            cMouseEvents.setSystemCursor IDC_ARROW
        End If
        
    Else
        cMouseEvents.setSystemCursor IDC_HAND
    End If

End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'If this preview control instance allows the user to select a color, restore whatever image was previously
    ' displayed upon mouse exit
    cMouseEvents.setSystemCursor IDC_DEFAULT
    syncPreviewImage
    
End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'If the viewport is not set to "fit to screen", then we must determine offsets based on the mouse position
    If Not viewportFitFullImage Then
        
        If Button = vbLeftButton Then
        
            'Make sure the move cursor remains accurate
            cMouseEvents.setSystemCursor IDC_SIZEALL
                
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
            If Not isColorSelectionAllowed Then cMouseEvents.setSystemCursor IDC_HAND
        End If
        
    Else
        'setArrowCursor picPreview
    End If
    
    If colorJustClicked > 0 Then
    
        'To accomodate shaky hands, allow a few mouse movements before resetting the image
        If colorJustClicked < 4 Then
            colorJustClicked = colorJustClicked + 1
        Else
            colorJustClicked = 0
            syncPreviewImage True
        End If
        
    End If
    
    'If point selection is allowed, continue firing events while the mouse is moving (as a convenience to the user)
    If isPointSelectionAllowed Then
    
        cMouseEvents.setSystemCursor IDC_HAND
    
        If (Button = vbRightButton) Or (Button = vbLeftButton) Then
        
            'Return the mouse coordinates as a ratio between 0 and 1, with 1 representing max width/height
            Dim retX As Double, retY As Double
            retX = x - ((picPreview.ScaleWidth - originalImage.getDIBWidth) \ 2)
            retY = y - ((picPreview.ScaleHeight - originalImage.getDIBHeight) \ 2)
            
            retX = retX / originalImage.getDIBWidth
            retY = retY / originalImage.getDIBHeight
            
            RaiseEvent PointSelected(retX, retY)
        
        End If
    
    End If

End Sub

Private Sub cMouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)

    If Not viewportFitFullImage Then
        
        cMouseEvents.setSystemCursor IDC_HAND
        
        hsOffsetX.Value = validateXOffset(hsOffsetX.Value + m_OffsetX)
        m_OffsetX = 0
        
        vsOffsetY.Value = validateYOffset(vsOffsetY.Value + m_OffsetY)
        m_OffsetY = 0
        
    End If

End Sub

'Given a mouse (x, y) coordinate pair, return a matching (x, y) pair for the underlying DIB.  Fit/zoom are automatically considered.
Private Sub GetDIBXYFromMouseXY(ByVal mouseX As Single, ByVal mouseY As Single, ByRef dibX As Single, ByRef dibY As Single, Optional ByVal useOriginalDIB As Boolean = True)
    
    'Because the caller may want coordinates from the original *OR* modified DIB, we use a generic reference.
    Dim srcDIB As pdDIB
    If useOriginalDIB Then
        Set srcDIB = originalImage
    Else
        Set srcDIB = fxImage
    End If
    
    'If the image is using "fit within" mode, we need to perform some extra coordinate math.
    Dim dstWidth As Double, dstHeight As Double
    dstWidth = m_BackBuffer.getDIBWidth
    dstHeight = m_BackBuffer.getDIBHeight
    
    Dim srcWidth As Double, srcHeight As Double
    srcWidth = srcDIB.getDIBWidth
    srcHeight = srcDIB.getDIBHeight
    
    'Calculate the aspect ratio of this DIB and the target picture box
    Dim srcAspect As Double, dstAspect As Double
    If srcHeight > 0 Then srcAspect = srcWidth / srcHeight Else srcAspect = 1
    If dstHeight > 0 Then dstAspect = dstWidth / dstHeight Else dstAspect = 1
        
    Dim finalWidth As Long, finalHeight As Long
    If (dstWidth <= srcWidth) Or (dstHeight <= srcHeight) Or Me.viewportFitFullImage Then
        convertAspectRatio srcWidth, srcHeight, dstWidth, dstHeight, finalWidth, finalHeight
    Else
        finalWidth = srcWidth
        finalHeight = srcHeight
    End If
        
    'Images smaller than the target area in one (or more) dimensions need to be centered in the target area
    Dim previewX As Long, previewY As Long
    If srcAspect > dstAspect Then
        previewY = CLng((dstHeight - finalHeight) / 2)
        If finalWidth = dstWidth Then
            previewX = 0
        Else
            previewX = CLng((dstWidth - finalWidth) / 2)
        End If
    Else
        previewX = CLng((dstWidth - finalWidth) / 2)
        If finalHeight = dstHeight Then
            previewY = 0
        Else
            previewY = CLng((dstHeight - finalHeight) / 2)
        End If
    End If
    
    'We now have an original DIB width/height pair, destination DIB width/height pair, preview (x, y) offset - all that's left
    ' is a source (x, y) offset.
    Dim srcX As Single, srcY As Single
    srcX = Me.offsetX
    srcY = Me.offsetY
    
    'Convert the destination (x, y) pair to the [0, 1] range.
    Dim dstX As Single, dstY As Single
    dstX = ((mouseX - previewX) / CDbl(finalWidth))
    dstY = ((mouseY - previewY) / CDbl(finalHeight))
    
    'Map it into the source range
    dibX = (dstX * srcWidth) + srcX
    dibY = (dstY * srcHeight) + srcY
    
    Set srcDIB = Nothing
    
End Sub

'X and Y offsets for the image preview are generated dynamically by the user's mouse movements.  As multiple functions
' need to validate those offsets to make sure they don't result in an offset outside the image, these standardized
' validation functions were created.
Private Function validateXOffset(ByVal currentOffset As Long) As Long
    If currentOffset < 0 Then currentOffset = 0
    If currentOffset > hsOffsetX.Max Then currentOffset = hsOffsetX.Max
    validateXOffset = currentOffset
End Function

Private Function validateYOffset(ByVal currentOffset As Long) As Long
    If currentOffset < 0 Then currentOffset = 0
    If currentOffset > vsOffsetY.Max Then currentOffset = vsOffsetY.Max
    validateYOffset = currentOffset
End Function

'The pdWindowPaint class raises this event when the control needs to be redrawn.  The passed coordinates contain the
' rect returned by GetUpdateRect (but with right/bottom measurements pre-converted to width/height).
Private Sub cPainter_PaintWindow(ByVal winLeft As Long, ByVal winTop As Long, ByVal winWidth As Long, ByVal winHeight As Long)

    'Flip the relevant chunk of the buffer to the screen
    BitBlt picPreview.hDC, winLeft, winTop, winWidth, winHeight, m_BackBuffer.getDIBDC, winLeft, winTop, vbSrcCopy
    
End Sub

Private Sub picPreview_Resize()
    If (m_BackBuffer Is Nothing) Then Set m_BackBuffer = New pdDIB
    If (m_BackBuffer.getDIBWidth <> picPreview.ScaleWidth) Or (m_BackBuffer.getDIBHeight <> picPreview.ScaleHeight) Then m_BackBuffer.createBlank picPreview.ScaleWidth, picPreview.ScaleHeight, 24, picPreview.BackColor
End Sub

'When the control's access key is pressed (alt+t) , toggle the original/current image
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    btsState.ListIndex = 1 - btsState.ListIndex
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    
    'Keep the control's backcolor in sync with the parent object
    If UCase$(PropertyName) = "BACKCOLOR" Then
        BackColor = Ambient.BackColor
    End If

End Sub

Private Sub UserControl_Initialize()
    
    'A check must be made for IDE behavior so the project will compile; VB's initialization of user controls during
    ' compiling and design process causes no shortage of odd issues and errors otherwise
    If g_IsProgramRunning Then
        
        'Set up a mouse events handler
        Set cMouseEvents = New pdInputMouse
        cMouseEvents.addInputTracker picPreview.hWnd, True, True, , True
        cMouseEvents.setSystemCursor IDC_ARROW
        
        'Also start a flicker-free window painter
        Set cPainter = New pdWindowPainter
        cPainter.StartPainter picPreview.hWnd
                
    End If
    
    'Prep the various buttonstrips
    btsState.AddItem "before", 0
    btsState.AddItem "after", 1
    btsState.ListIndex = 1
    
    btsZoom.AddItem "1:1", 0
    btsZoom.AddItem "fit", 1
    btsZoom.ListIndex = 1
    
    m_ShowOriginalInstead = False
    curColor = 0
            
End Sub

'Initialize our effect preview control
Private Sub UserControl_InitProperties()
    
    'Set the background of the fxPreview to match the background of our parent object
    BackColor = Ambient.BackColor
    
    'Mark the original image as having NOT been set
    m_HasOriginal = False
    
    'By default, the control cannot be used for color selection
    isColorSelectionAllowed = False
    
    'By default, the control allows the user to zoom/pan the transformation
    disableZoomPanAbility = False
    
    'By default, the control does not allow for selecting coordinate points by clicking
    isPointSelectionAllowed = False
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        AllowColorSelection = .ReadProperty("ColorSelection", False)
        AllowPointSelection = .ReadProperty("PointSelection", False)
        disableZoomPanAbility = .ReadProperty("DisableZoomPan", False)
    End With
    
End Sub

'Redraw the user control after it has been resized
Private Sub UserControl_Resize()
    redrawControl
End Sub

Private Sub UserControl_Show()
    
    'Generate a unique ID for this session
    m_UniqueID = Timer
    
    'Ensure the control is redrawn at least once
    redrawControl
    
    'Set an initial max/min for the preview offsets if the user chooses to preview at 100% zoom
    If g_IsProgramRunning Then
    
        'Reset the mouse cursor
        cMouseEvents.setSystemCursor IDC_ARROW
    
        Dim maxHOffset As Long, maxVOffset As Long
        
        Dim srcWidth As Long, srcHeight As Long
        If pdImages(g_CurrentImage).selectionActive Then
            srcWidth = pdImages(g_CurrentImage).mainSelection.boundWidth
            srcHeight = pdImages(g_CurrentImage).mainSelection.boundHeight
        Else
            srcWidth = pdImages(g_CurrentImage).getActiveDIB.getDIBWidth
            srcHeight = pdImages(g_CurrentImage).getActiveDIB.getDIBHeight
        End If
        
        maxHOffset = srcWidth - picPreview.ScaleWidth
        maxVOffset = srcHeight - picPreview.ScaleHeight
        
        If maxHOffset > 0 Then
            hsOffsetX.Max = maxHOffset
            m_HScrollAllowed = True
        Else
            hsOffsetX.Max = 1
            m_HScrollAllowed = False
        End If
        
        If maxVOffset > 0 Then
            vsOffsetY.Max = maxVOffset
            m_VScrollAllowed = True
        Else
            vsOffsetY.Max = 1
            m_VScrollAllowed = False
        End If
        
        'Enable color management
        AssignDefaultColorProfileToObject picPreview.hWnd, picPreview.hDC
        TurnOnColorManagementForDC picPreview.hDC
    
    End If
    
End Sub

Private Sub UserControl_Terminate()

    'Release any image objects that may have been created
    If Not (originalImage Is Nothing) Then originalImage.eraseDIB
    If Not (fxImage Is Nothing) Then fxImage.eraseDIB
    
End Sub

'After a resize or paint request, update the layout of our control
Private Sub redrawControl()
    
    'The primary object in this control is the preview picture box.  Everything else is positioned relative to it.
    Dim newPicWidth As Long, newPicHeight As Long
    newPicWidth = UserControl.ScaleWidth
    newPicHeight = UserControl.ScaleHeight - (btsState.Height + FixDPI(4))
    picPreview.Move 0, 0, newPicWidth, newPicHeight
    
    'If zoom/pan is not allowed, hide that button entirely
    btsZoom.Visible = Not disableZoomPanAbility
    
    'Adjust the button strips to appear just below the preview window
    Dim newButtonTop As Long, newButtonWidth As Long
    newButtonTop = UserControl.ScaleHeight - btsState.Height
    
    'If zoom/pan is still visible, split the horizontal difference between that button strip, and the before/after strip.
    If btsZoom.Visible Then
        newButtonWidth = (newPicWidth \ 2) - FixDPI(8)
        btsZoom.Move UserControl.ScaleWidth - newButtonWidth, newButtonTop, newButtonWidth, btsState.Height
        
    'If zoom/pan is NOT visible, let the before/after button have the entire horizontal space
    Else
        newButtonWidth = newPicWidth
    End If
    
    'Move the before/after toggle into place
    btsState.Move 0, newButtonTop, newButtonWidth, btsState.Height
                
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    'Store all associated properties
    With PropBag
        .WriteProperty "ColorSelection", AllowColorSelection, False
        .WriteProperty "DisableZoomPan", disableZoomPanAbility, False
        .WriteProperty "PointSelection", AllowPointSelection, False
    End With
    
End Sub
