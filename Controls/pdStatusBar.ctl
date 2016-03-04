VERSION 5.00
Begin VB.UserControl pdStatusBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
   HasDC           =   0   'False
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   796
   ToolboxBitmap   =   "pdStatusBar.ctx":0000
   Begin PhotoDemon.pdDropDown cmbSizeUnit 
      Height          =   360
      Left            =   3630
      TabIndex        =   0
      Top             =   15
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   635
      UseCustomBackgroundColor=   -1  'True
      FontSize        =   9
   End
   Begin PhotoDemon.pdDropDown cmbZoom 
      Height          =   360
      Left            =   840
      TabIndex        =   1
      Top             =   15
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   635
      UseCustomBackgroundColor=   -1  'True
      FontSize        =   9
   End
   Begin PhotoDemon.pdLabel lblImgSize 
      Height          =   210
      Left            =   3240
      Top             =   60
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   370
      BackColor       =   -2147483626
      Caption         =   "size:"
      FontSize        =   9
      Layout          =   2
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdZoomFit 
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   609
      BackColor       =   -2147483626
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdZoomOut 
      Height          =   345
      Left            =   390
      TabIndex        =   3
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   609
      AutoToggle      =   -1  'True
      BackColor       =   -2147483626
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdZoomIn 
      Height          =   345
      Left            =   2190
      TabIndex        =   4
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   609
      AutoToggle      =   -1  'True
      BackColor       =   -2147483626
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdButtonToolbox cmdImgSize 
      Height          =   345
      Left            =   2790
      TabIndex        =   5
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   609
      AutoToggle      =   -1  'True
      BackColor       =   -2147483626
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblCoordinates 
      Height          =   210
      Left            =   5160
      Top             =   60
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   370
      BackColor       =   -2147483626
      Caption         =   "size:"
      FontSize        =   9
      Layout          =   2
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblMessages 
      Height          =   210
      Left            =   6360
      Top             =   60
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   503
      Alignment       =   1
      BackColor       =   -2147483626
      Caption         =   "(messages will appear here at run-time)"
      FontSize        =   9
      UseCustomBackColor=   -1  'True
   End
End
Attribute VB_Name = "pdStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon primary canvas status bar control
'Copyright 2002-2016 by Tanner Helland
'Created: 29/November/02
'Last updated: 03/March/16
'Last update: migrate status bar into its own dedicated control
'
'In PD, this control is never used on its own.  It is meant to be used as a component of the pdCanvas control,
' and it's split out here in an attempt to simplify the canvas's rendering code and input handling.
'
'For implementation details, please refer to pdCanvas.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Icons rendered to the scroll bar.  Rather than constantly reloading them from file, we cache them at initialization.
Private sbIconCoords As pdDIB, sbIconNetwork As pdDIB

'External functions can notify the status bar of PD's network access.  When PD is downloading various update bits, a relevant icon
' will be displayed in the status bar.  As the canvas has no knowledge of network stuff, it's imperative that the caller notify
' us of both TRUE and FALSE states.
Private m_NetworkAccessActive As Boolean

'External functions can tell us to enable or disable the status bar for various reasons (e.g. no images are loaded).  We track the
' last requested state internally, in case we need to internally refresh the status bar for some reason.
Private m_LastEnabledState As Boolean

'The status bar includes a few different separator lines.  These lines are tracked in an array, which simplifies their
' painting on refresh events.
Private m_LinePositions() As Single

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since attempted to wrap these into a single master control support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by the control, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDSTATUSBAR_COLOR_LIST
    [_First] = 0
    PDSB_Background = 0
    PDSB_Separator = 1
    [_Last] = 1
    [_Count] = 2
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
    RedrawBackBuffer
End Property

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

'To support high-DPI settings properly, we expose specialized move+size functions
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

Private Sub ucSupport_GotFocusAPI()
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_LostFocusAPI()
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout
    RedrawBackBuffer
End Sub

Private Sub ucSupport_WindowResize(ByVal newWidth As Long, ByVal newHeight As Long)
    UpdateControlLayout
End Sub

'Display the current mouse coordinates
Public Sub DisplayCanvasCoordinates(ByVal xCoord As Long, ByVal yCoord As Long, Optional ByVal clearCoords As Boolean = False)
    
    If clearCoords Then
        lblCoordinates.Caption = ""
    Else
        lblCoordinates.Caption = "(" & xCoord & "," & yCoord & ")"
    End If
    
    'Align the right-hand line control with the newly captioned label
    m_LinePositions(2) = lblCoordinates.GetLeft + lblCoordinates.GetWidth + FixDPI(10)
    
    'Make the message area shrink to match the new coordinate display size
    FitMessageArea
        
End Sub

Public Sub DisplayCanvasMessage(ByRef cMessage As String)
    lblMessages.Caption = cMessage
    lblMessages.RequestRefresh
End Sub

Public Sub DisplayImageSize(ByRef srcImage As pdImage, Optional ByVal clearSize As Boolean = False)
    
    'If the source image is irrelevant, forcibly specify a ClearSize operation
    If (srcImage Is Nothing) Then clearSize = True
    
    'The size display is cleared whenever the user has no images loaded
    If clearSize Then
        lblImgSize.Caption = ""
        FitMessageArea
        
    'When size IS displayed, we must also refresh the status bar (now that it dynamically aligns its contents)
    Else
        
        Dim iWidth As Double, iHeight As Double
        Dim sizeString As String
        
        'Convert the image size (in pixels) to whatever unit the user has currently selected from the drop-down
        Select Case cmbSizeUnit.ListIndex
            
            'Pixels
            Case 0
                sizeString = srcImage.Width & " x " & srcImage.Height
                
            'Inches
            Case 1
                iWidth = convertPixelToOtherUnit(MU_INCHES, srcImage.Width, srcImage.getDPI(), srcImage.Width)
                iHeight = convertPixelToOtherUnit(MU_INCHES, srcImage.Height, srcImage.getDPI(), srcImage.Height)
                sizeString = Format(iWidth, "0.0##") & " x " & Format(iHeight, "0.0##")
            
            'CM
            Case 2
                iWidth = convertPixelToOtherUnit(MU_CENTIMETERS, srcImage.Width, srcImage.getDPI(), srcImage.Width)
                iHeight = convertPixelToOtherUnit(MU_CENTIMETERS, srcImage.Height, srcImage.getDPI(), srcImage.Height)
                sizeString = Format(iWidth, "0.0#") & " x " & Format(iHeight, "0.0#")
            
        End Select
        
        lblImgSize.Caption = sizeString
        lblImgSize.UpdateAgainstCurrentTheme
        ReflowStatusBar True
        
    End If
        
End Sub

Public Function GetZoomDropDownReference() As pdDropDown
    Set GetZoomDropDownReference = cmbZoom
End Function

'Fill the "size units" drop-down.  We must do this relatively late in the load process, as we have to wait for the translation
' engine to initialize.
Public Function PopulateSizeUnits()
    cmbSizeUnit.Clear
    cmbSizeUnit.AddItem "px", 0
    cmbSizeUnit.AddItem "in", 1
    cmbSizeUnit.AddItem "cm", 2
    cmbSizeUnit.ListIndex = 0
End Function

'External functions can call this to set the current network state (which in turn, draws a relevant icon to the status bar)
Public Sub SetNetworkState(ByVal newNetworkState As Boolean)
    If newNetworkState <> m_NetworkAccessActive Then
        m_NetworkAccessActive = newNetworkState
        FitMessageArea
    End If
End Sub

Private Sub cmbSizeUnit_Click()
    If g_OpenImageCount > 0 Then DisplayImageSize pdImages(g_CurrentImage)
End Sub

Private Sub CmbZoom_Click()

    'Only process zoom changes if an image has been loaded
    If FormMain.mainCanvas(0).IsCanvasInteractionAllowed() Then
        
        'Before updating the current image, we need to retrieve two sets of points: the current center point
        ' of the canvas, in canvas coordinate space, and the current center point of the canvas *in image
        ' coordinate space*.  When zoom is changed, we preserve the current center of the image relative to
        ' the center of the canvas, to make the zoom operation feel more natural.
        Dim centerXCanvas As Double, centerYCanvas As Double, centerXImage As Double, centerYImage As Double
        centerXCanvas = FormMain.mainCanvas(0).GetCanvasWidth / 2
        centerYCanvas = FormMain.mainCanvas(0).GetCanvasHeight / 2
        Drawing.ConvertCanvasCoordsToImageCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), centerXCanvas, centerYCanvas, centerXImage, centerYImage, False
        
        'With those coordinates safely cached, update the currently stored zoom value in the active pdImage object
        pdImages(g_CurrentImage).currentZoomValue = cmbZoom.ListIndex
        
        'Disable the zoom in/out buttons when they reach the end of the available zoom levels
        If cmbZoom.ListIndex = 0 Then
            cmdZoomIn.Enabled = False
        Else
            If Not cmdZoomIn.Enabled Then cmdZoomIn.Enabled = True
        End If
        
        If cmbZoom.ListIndex = g_Zoom.GetZoomCount Then
            cmdZoomOut.Enabled = False
        Else
            If Not cmdZoomOut.Enabled Then cmdZoomOut.Enabled = True
        End If
        
        'Highlight the "zoom fit" button while fit mode is active
        cmdZoomFit.Value = CBool(cmbZoom.ListIndex = g_Zoom.GetZoomFitAllIndex)
        
        'Redraw the viewport (if allowed; some functions will prevent us from doing this, as they plan to request their own
        ' refresh after additional processing occurs)
        If g_AllowViewportRendering Then
            
            'If the user has selected any of the "fit xyz" zoom options, we want to re-center the viewport as part
            ' of updating zoom.  If they have *not* selected this, we want to preserve the current center point
            ' of the viewport.
            If cmbZoom.ListIndex < g_Zoom.GetZoomFitWidthIndex Then
                Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_PreservePointPosition, centerXCanvas, centerYCanvas, centerXImage, centerYImage
            Else
                Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_ResetToZero
            End If
        
            'Notify any other relevant UI elements
            FormMain.mainCanvas(0).RelayViewportChanges
            
        End If
        
    End If

End Sub

Private Sub cmdImgSize_Click()
    If FormMain.mainCanvas(0).IsCanvasInteractionAllowed() Then Process "Resize image", True
End Sub

Private Sub cmdZoomFit_Click()
    Image_Canvas_Handler.FitOnScreen
End Sub

Private Sub cmdZoomIn_Click()
    FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex = g_Zoom.GetNearestZoomInIndex(FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex)
End Sub

Private Sub cmdZoomOut_Click()
    FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex = g_Zoom.GetNearestZoomOutIndex(FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex)
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a master user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDSTATUSBAR_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDStatusBar", colorCount
    If Not g_IsProgramRunning Then UpdateColorList
    
    ReDim m_LinePositions(0 To 2) As Single
    
    'Update the control size parameters at least once
    UpdateControlLayout
                
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If Not g_IsProgramRunning Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
    If Not g_IsProgramRunning Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", Me.Enabled, True
End Sub

'When the canvas is cleared, we automatically disable portions of the status bar to match
Public Sub ClearCanvas()
    ReflowStatusBar False
End Sub

'Reposition all status bar elements.  This is time-consuming, so only call it if a large-scale state change (like unloading
' all images) requires us to do this.
Public Sub ReflowStatusBar(ByVal enabledState As Boolean)
    
    'Note the enabled state at a module level, in case we need to internally refresh the status bar for some reason
    m_LastEnabledState = enabledState
    
    'The zoom drop-down can now change width if a translation is active.  Make sure the zoom-in button is positioned accordingly.
    cmdZoomIn.SetLeft cmbZoom.GetLeft + cmbZoom.GetWidth + FixDPI(3)
    
    'Move the left-most line into position.  (This must be done dynamically, or it will be mispositioned
    ' on high-DPI displays)
    m_LinePositions(0) = (cmdZoomIn.GetLeft + cmdZoomIn.GetWidth) + FixDPI(6)
    
    'We will only draw subsequent interface elements if at least one image is currently loaded.
    If enabledState Then
        
        If (Not cmdZoomFit.Visible) Then cmdZoomFit.Visible = True
        If (Not cmdZoomIn.Visible) Then cmdZoomIn.Visible = True
        If (Not cmdZoomOut.Visible) Then cmdZoomOut.Visible = True
        If (Not cmbZoom.Visible) Then cmbZoom.Visible = True
        
        'Start with the "image size" button
        cmdImgSize.SetLeft m_LinePositions(0) + FixDPI(4)
        If (Not cmdImgSize.Visible) Then cmdImgSize.Visible = True
        
        'After the "image size" icon comes the actual image size label.  Its position is determined by the image resize button width,
        ' plus a 4px buffer on either size (contingent on DPI)
        lblImgSize.SetLeft cmdImgSize.GetLeft + cmdImgSize.GetWidth + FixDPI(4)
        
        'The image size label is autosized.  Move the "size unit" combo box next to it, and the next vertical line
        ' separator just past it.
        If (Not cmbSizeUnit.Visible) Then cmbSizeUnit.Visible = True
        cmbSizeUnit.SetLeft lblImgSize.GetLeft + lblImgSize.GetWidth + FixDPI(10)
        
        m_LinePositions(1) = cmbSizeUnit.GetLeft + cmbSizeUnit.GetWidth + FixDPI(10)
        
        'After the "image size" panel and separator comes mouse coordinates.  The basic steps from above are repeated.
        lblCoordinates.SetLeft m_LinePositions(1) + FixDPI(14) + FixDPI(16)
        
        m_LinePositions(2) = lblCoordinates.GetLeft + lblCoordinates.GetWidth + FixDPI(10)
        
    'Images are not loaded.  Hide the lines and other items.
    Else
        cmdZoomFit.Visible = False
        cmdZoomIn.Visible = False
        cmdZoomOut.Visible = False
        cmbZoom.Visible = False
        cmdImgSize.Visible = False
        cmbSizeUnit.Visible = False
    End If
    
    'We only establish positions up to the mouse coordinate label.  All items *past* that point are positioned by
    ' the dedicated message area reflow function (which is accessed much more frequently).
    FitMessageArea
    
End Sub

'Whenever this window changes size, we may need to re-align various bits of internal chrome (status bar, rulers, etc).  Call this function
' to do so.
Public Sub FitMessageArea()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Move the message label into position (right-aligned, with a slight margin)
    Dim newLeft As Long
    newLeft = m_LinePositions(2) + FixDPI(7)
    If lblMessages.GetLeft <> newLeft Then lblMessages.SetLeft newLeft
    
    'If the message label will overflow other elements of the status bar, shrink it as necessary
    Dim newMessageArea As Long
    newMessageArea = (bWidth - lblMessages.GetLeft) - FixDPI(12)
    
    If newMessageArea < 0 Then
        lblMessages.Visible = False
    Else
        If lblMessages.GetWidth <> newMessageArea Then lblMessages.SetWidth newMessageArea
        lblMessages.Visible = True
    End If
    
    RedrawBackBuffer
    
End Sub

'Because this control automatically forces all internal buttons to identical sizes, we have to recalculate a number
' of internal sizing metrics whenever the control size changes.
Private Sub UpdateControlLayout()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Center all combo boxes vertically (this is necessary for high-DPI displays)
    cmbZoom.Top = (bHeight - cmbZoom.GetHeight) \ 2
    cmbSizeUnit.Top = (bHeight - cmbSizeUnit.GetHeight) \ 2

    'If the control is resizing, the mouse cannot feasibly be over the image - so clear the coordinate box.  Note that this will
    ' also realign all chrome elements, so we don't need a manual FitMessageArea call here.
    DisplayCanvasCoordinates 0, 0, False
    
    'No other special preparation is required for this control, so proceed with recreating the back buffer
    RedrawBackBuffer
            
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer()
    
    'We can improve shutdown performance by ignoring redraw requests when the program is going down
    If g_ProgramShuttingDown Then
        If (g_Themer Is Nothing) Then Exit Sub
    End If
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long, bufferDC As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDSB_Background, Me.Enabled))
        
    If g_IsProgramRunning Then
        
        If (Not (sbIconCoords Is Nothing)) And m_LastEnabledState Then
            sbIconCoords.alphaBlendToDC bufferDC, , m_LinePositions(1) + FixDPI(8), FixDPI(4), FixDPI(sbIconCoords.getDIBWidth), FixDPI(sbIconCoords.getDIBHeight)
        End If
        
        'Render the network access icon as necessary
        If m_NetworkAccessActive Then
            If m_LastEnabledState Then
                sbIconNetwork.alphaBlendToDC bufferDC, , m_LinePositions(2) + FixDPI(8), FixDPI(4), FixDPI(sbIconNetwork.getDIBWidth), FixDPI(sbIconNetwork.getDIBHeight)
            Else
                If m_NetworkAccessActive Then sbIconNetwork.alphaBlendToDC bufferDC, , m_LinePositions(0), FixDPI(4), FixDPI(sbIconNetwork.getDIBWidth), FixDPI(sbIconNetwork.getDIBHeight)
            End If
        End If
        
        'Render all separator lines
        Dim lineTop As Single, lineBottom As Single
        lineTop = FixDPI(1)
        lineBottom = bHeight - FixDPI(2)
        
        Dim lineColor As Long
        lineColor = m_Colors.RetrieveColor(PDSB_Separator, Me.Enabled)
        
        If m_LastEnabledState Then
            Dim i As Long
            For i = 0 To UBound(m_LinePositions)
                GDI_Plus.GDIPlusDrawLineToDC bufferDC, m_LinePositions(i), lineTop, m_LinePositions(i), lineBottom, lineColor
            Next i
        End If
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDSB_Background, "Background", IDE_GRAY
        .LoadThemeColor PDSB_Separator, "Separator", IDE_BLACK
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme()

    UpdateColorList
    
    If g_IsProgramRunning Then
    
        cmdZoomFit.AssignImage "SB_ZOOM_FIT"
        cmdZoomIn.AssignImage "SB_ZOOM_IN"
        cmdZoomOut.AssignImage "SB_ZOOM_OUT"
        cmdImgSize.AssignImage "SB_IMG_SIZE"
        
        'Load various status bar icons from the resource file
        If (sbIconCoords Is Nothing) Then Set sbIconCoords = New pdDIB
        If (sbIconNetwork Is Nothing) Then Set sbIconNetwork = New pdDIB
        LoadResourceToDIB "SB_MOUSE_POS", sbIconCoords
        LoadResourceToDIB "SB_NETWORK", sbIconNetwork
        
    End If
    
    'Rebuild all drop-down boxes (so that translations can be applied)
    Dim backupZoomIndex As Long, backupSizeIndex As Long
    backupZoomIndex = cmbZoom.ListIndex
    backupSizeIndex = cmbSizeUnit.ListIndex
    
    'Repopulate zoom dropdown text
    If Not (g_Zoom Is Nothing) Then g_Zoom.InitializeViewportEngine
    If Not (g_Zoom Is Nothing) Then g_Zoom.PopulateZoomComboBox cmbZoom, backupZoomIndex
    Me.PopulateSizeUnits
    
    'Auto-size the newly populated combo boxes, according to the width of their longest entries
    cmbZoom.SetWidthAutomatically
    cmbSizeUnit.SetWidthAutomatically
    
    cmdZoomFit.AssignTooltip "Fit the image on-screen"
    cmdZoomIn.AssignTooltip "Zoom in"
    cmdZoomOut.AssignTooltip "Zoom out"
    cmdImgSize.AssignTooltip "Resize image"
    cmbZoom.AssignTooltip "Change viewport zoom"
    cmbSizeUnit.AssignTooltip "Change the image size unit displayed to the left of this box"
    
    Dim sbBackColor As Long
    sbBackColor = m_Colors.RetrieveColor(PDSB_Background, Me.Enabled)
    UserControl.BackColor = sbBackColor
    
    lblCoordinates.BackColor = sbBackColor
    lblImgSize.BackColor = sbBackColor
    lblMessages.BackColor = sbBackColor
    
    cmdZoomFit.BackColor = sbBackColor
    cmdZoomIn.BackColor = sbBackColor
    cmdZoomOut.BackColor = sbBackColor
    cmdImgSize.BackColor = sbBackColor
    cmbZoom.BackgroundColor = sbBackColor
    cmbSizeUnit.BackgroundColor = sbBackColor
    
    lblCoordinates.UpdateAgainstCurrentTheme
    lblImgSize.UpdateAgainstCurrentTheme
    lblMessages.UpdateAgainstCurrentTheme
    
    cmdZoomFit.UpdateAgainstCurrentTheme
    cmdZoomIn.UpdateAgainstCurrentTheme
    cmdZoomOut.UpdateAgainstCurrentTheme
    cmdImgSize.UpdateAgainstCurrentTheme
    
    cmbZoom.UpdateAgainstCurrentTheme
    cmbSizeUnit.UpdateAgainstCurrentTheme
    
    ucSupport.UpdateAgainstThemeAndLanguage
    
    'Fix combo box positioning (important on high-DPI displays, or if the active font has changed)
    cmbZoom.Top = (UserControl.ScaleHeight - cmbZoom.GetHeight) \ 2
    cmbSizeUnit.Top = (UserControl.ScaleHeight - cmbSizeUnit.GetHeight) \ 2
    
    'Restore zoom and size unit indices
    cmbZoom.ListIndex = backupZoomIndex
    cmbSizeUnit.ListIndex = backupSizeIndex
    
    'Note that we don't actually move the last line status bar; that is handled by DisplayImageCoordinates itself
    If g_OpenImageCount > 0 Then
        DisplayImageSize pdImages(g_CurrentImage), False
    Else
        DisplayImageSize Nothing, True
    End If
    
    DisplayCanvasCoordinates 0, 0, True
    FitMessageArea
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByVal newTooltip As String, Optional ByVal newTooltipTitle As String, Optional ByVal newTooltipIcon As TT_ICON_TYPE = TTI_NONE)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, newTooltipIcon
End Sub

