VERSION 5.00
Begin VB.UserControl pdStatusBar 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11940
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
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   796
   ToolboxBitmap   =   "pdStatusBar.ctx":0000
   Begin PhotoDemon.pdDropDown cmbSizeUnit 
      Height          =   360
      Left            =   3630
      TabIndex        =   0
      Top             =   15
      Width           =   690
      _ExtentX        =   1217
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
      Index           =   0
      Left            =   4680
      Top             =   60
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   370
      BackColor       =   -2147483626
      Caption         =   ""
      FontSize        =   9
      Layout          =   2
      UseCustomBackColor=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblCoordinates 
      Height          =   210
      Index           =   1
      Left            =   5100
      Top             =   60
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   370
      BackColor       =   -2147483626
      Caption         =   ""
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
      Caption         =   ""
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
'Copyright 2002-2026 by Tanner Helland
'Created: 29/November/02
'Last updated: 07/December/21
'Last update: new code to display selection size in status bar, even while editing a selection
'
'In PD, this control is never used on its own.  It is meant to be used as a component of the pdCanvas control,
' and it's split out here in an attempt to simplify the canvas's rendering code and input handling.
'
'For implementation details, please refer to pdCanvas.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'Icons rendered to the scroll bar.  Rather than constantly reloading them from file, we cache them at initialization.
Private sbIconCoords As pdDIB, sbIconNetwork As pdDIB, sbIconSelection As pdDIB

'The current "unit of measurement" set by the status bar dropdown.
Private m_UnitOfMeasurement As PD_MeasurementUnit

'External functions can notify the status bar of PD's network access.  When PD is downloading
' update files, a relevant icon will be displayed in the status bar.  As the canvas has no knowledge
' of network stuff, it's imperative that the caller notify us of both TRUE and FALSE states.
Private m_NetworkAccessActive As Boolean

'External functions can tell us to enable or disable the status bar for various reasons (e.g. no images are loaded).
' We track the last requested state internally, in case we need to internally refresh the status bar for some reason.
Private m_LastEnabledState As Boolean

'The status bar includes a few different separator lines.  These lines are tracked in an array, which simplifies their
' painting on refresh events.
Private m_LinePositions() As Single

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
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

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_StatusBar
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
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

'Display the current mouse coordinates
Public Sub DisplayCanvasCoordinates(ByVal xCoord As Double, ByVal yCoord As Double, Optional ByVal clearCoords As Boolean = False)
    
    If clearCoords Then
        lblCoordinates(0).Caption = vbNullString
        
    'The position displayed changes based on the current measurement unit (px, in, cm)
    Else
        If PDImages.IsImageActive() Then
            lblCoordinates(0).Caption = "(" & Units.GetValueFormattedForUnit_FromPixel(m_UnitOfMeasurement, xCoord, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Width, False, False) & "," & Units.GetValueFormattedForUnit_FromPixel(m_UnitOfMeasurement, yCoord, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Height, False, False) & ")"
        End If
    End If
    
    lblCoordinates(0).Visible = (Not clearCoords)
    
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
        lblImgSize.Caption = vbNullString
        FitMessageArea
        
    'When size IS displayed, we must also refresh the status bar (now that it dynamically aligns its contents)
    Else
        
        Const TEXT_BETWEEN_DIMENSIONS As String = " x "
        
        'Convert pixel measurements to the current unit, then convert those to a human-readable string.
        Dim sizeString As String
        sizeString = Units.GetValueFormattedForUnit_FromPixel(m_UnitOfMeasurement, srcImage.Width, srcImage.GetDPI(), srcImage.Width) & TEXT_BETWEEN_DIMENSIONS & Units.GetValueFormattedForUnit_FromPixel(m_UnitOfMeasurement, srcImage.Height, srcImage.GetDPI(), srcImage.Height)
        
        'Display the dimensions in the status bar, then reflow neighboring controls to fit
        lblImgSize.Caption = sizeString
        ReflowStatusBar True
        
    End If
        
End Sub

Public Function GetZoomDropDownIndex() As Long
    GetZoomDropDownIndex = cmbZoom.ListIndex
End Function

Public Sub SetZoomDropDownIndex(ByVal newIndex As Long)
    cmbZoom.ListIndex = newIndex
    If (cmbZoom.ListIndex = Zoom.GetZoomFitAllIndex) Then cmdZoomFit.Value = True
End Sub

Public Function IsZoomEnabled() As Boolean
    IsZoomEnabled = cmbZoom.Enabled
End Function

'Only use this function for initially populating the zoom drop-down
Public Function GetZoomDropDownReference() As pdDropDown
    Set GetZoomDropDownReference = cmbZoom
End Function

'Fill the "size units" drop-down.  We must do this relatively late in the load process, as we have to wait for the translation
' engine to initialize.
Public Sub PopulateSizeUnits()
    
    cmbSizeUnit.SetAutomaticRedraws False
    cmbSizeUnit.Clear
    
    Dim i As Long
    For i = 0 To Units.GetNumOfAvailableUnits()
        cmbSizeUnit.AddItem Units.GetNameOfUnit(i, True), i
    Next i
    
    
    cmbSizeUnit.ListIndex = 1
    cmbSizeUnit.SetAutomaticRedraws True, True
    
End Sub

'External functions can call this to set the current network state
' (which in turn, draws a relevant icon to the status bar).
Public Sub SetNetworkState(ByVal newNetworkState As Boolean)
    If (newNetworkState <> m_NetworkAccessActive) Then
        m_NetworkAccessActive = newNetworkState
        FitMessageArea
    End If
End Sub

'External functions can call this to set the current "selection" state
' (which updates the status bar with a little selection size notification).
'
'Note that it doesn't need to be used exclusively for selections - anything that represents an "area"
' can use this function to display a "selected area" region in the status bar.
Public Sub SetSelectionState(ByVal newSelectionState As Boolean)
    
    'If a message *is* displayed, we'll assemble it using a pdString instance
    Dim cString As pdString
    
    'Retrieve the contained rect into this struct
    Dim selectRect As RectF
    
    'The position displayed changes based on the current measurement unit (px, in, cm)
    If newSelectionState Then
    
        If PDImages.IsImageActive() Then
            
            'Different tools can use this area in different ways
            If Tools.IsSelectionToolActive() Then
                
                If PDImages.GetActiveImage.IsSelectionActive() Then
                    
                    'Some rectangle selections are allowed to display "in-progress" measurements.
                    newSelectionState = PDImages.GetActiveImage.MainSelection.IsLockedIn()
                    
                    'If the selection is *not* locked in, we can still display a boundary rect under
                    ' certain conditions.
                    If (Not newSelectionState) Then
                        Select Case PDImages.GetActiveImage.MainSelection.GetSelectionShape
                            
                            'Rectangle and ellipse selections are *always* okay to display, because even while
                            ' being drawn they have clear "boundary rects"
                            Case ss_Rectangle, ss_Circle
                                newSelectionState = True
                            Case ss_Polygon
                                newSelectionState = PDImages.GetActiveImage.MainSelection.GetPolygonClosedState()
                            Case ss_Lasso
                                newSelectionState = PDImages.GetActiveImage.MainSelection.GetLassoClosedState()
                            Case ss_Wand
                                newSelectionState = True
                            Case ss_Raster
                                newSelectionState = True
                        End Select
                        
                    End If
                        
                    If newSelectionState Then
                        
                        'The way we retrieve selection boundaries varies by selection type, and may also
                        ' depend on whether the selection is locked in (i.e. "finished")
                        Select Case PDImages.GetActiveImage.MainSelection.GetSelectionShape
                            
                            'Rectangle and ellipse selections are *always* okay to display, because even while
                            ' being drawn they have clear "boundary rects"
                            Case ss_Rectangle, ss_Circle
                                selectRect = PDImages.GetActiveImage.MainSelection.GetCornersLockedRect()
                            Case ss_Polygon
                                selectRect = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect()
                            Case ss_Lasso
                                selectRect = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect()
                            Case ss_Wand
                                selectRect = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect()
                            Case ss_Raster
                                selectRect = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect()
                        End Select
                        
                    End If
                    
                '/IsSelectionActive
                Else
                    newSelectionState = False
                End If
            
            '/Non-selection tools can also be handled here
            ElseIf (g_CurrentTool = ND_CROP) Then
                
                If Tools_Crop.IsValidCropActive() Then
                    newSelectionState = True
                    selectRect = Tools_Crop.GetCropRectF
                End If
                
            End If
        
        '/IsImageActive
        Else
            newSelectionState = False
        End If
        
    '/!newSelectionState
    End If
    
    'Build a display message as relevant
    If newSelectionState Then
        
        'We're also going to assemble the final display string using a pdString instance
        Set cString = New pdString
        
        cString.Append Units.GetValueFormattedForUnit_FromPixel(m_UnitOfMeasurement, selectRect.Width, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Width, False)
        Const LOWERCASE_X As String = " x "
        cString.Append LOWERCASE_X
        cString.Append Units.GetValueFormattedForUnit_FromPixel(m_UnitOfMeasurement, selectRect.Height, PDImages.GetActiveImage.GetDPI(), PDImages.GetActiveImage.Height, False)
        
        'Also append the selection's aspect ratio, in the form X : 1
        If (selectRect.Height <> 0!) Then
            cString.Append "  ("
            cString.Append Format$(selectRect.Width / selectRect.Height, "0.0#")
            cString.Append ":1)"
        End If
        
        lblCoordinates(1).Caption = cString.ToString()
    
    End If
    
    lblCoordinates(1).Visible = newSelectionState
    
    'Align both the coordinates the right-hand line control with the newly captioned label
    If newSelectionState Then
        m_LinePositions(3) = lblCoordinates(0).GetLeft + lblCoordinates(0).GetWidth + Interface.FixDPI(10)
    Else
        m_LinePositions(3) = m_LinePositions(2)
    End If
    
    'Make the message area shrink to match the new coordinate display size
    FitMessageArea
    
End Sub

Private Sub cmbSizeUnit_Click()
    m_UnitOfMeasurement = cmbSizeUnit.ListIndex
    If PDImages.IsImageActive() Then
        Me.DisplayImageSize PDImages.GetActiveImage()
        FormMain.MainCanvas(0).NotifyRulerUnitChange cmbSizeUnit.ListIndex
        If (g_CurrentTool = ND_MEASURE) Then Tools_Measure.NotifyUnitChange
        If (g_CurrentTool = NAV_MOVE) Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage, FormMain.MainCanvas(0)
    End If
End Sub

Private Sub CmbZoom_Click()

    'Only process zoom changes if an image has been loaded
    If FormMain.MainCanvas(0).IsCanvasInteractionAllowed() Then
        
        'Before updating the current image, we need to retrieve two sets of points: the current center point
        ' of the canvas, in canvas coordinate space, and the current center point of the canvas *in image
        ' coordinate space*.  When zoom is changed, we preserve the current center of the image relative to
        ' the center of the canvas, to make the zoom operation feel more natural.
        Dim centerXCanvas As Double, centerYCanvas As Double, centerXImage As Double, centerYImage As Double
        centerXCanvas = FormMain.MainCanvas(0).GetCanvasWidth * 0.5
        centerYCanvas = FormMain.MainCanvas(0).GetCanvasHeight * 0.5
        Drawing.ConvertCanvasCoordsToImageCoords FormMain.MainCanvas(0), PDImages.GetActiveImage(), centerXCanvas, centerYCanvas, centerXImage, centerYImage, False
        
        'With those coordinates safely cached, update the currently stored zoom value in the active pdImage object
        PDImages.GetActiveImage.SetZoomIndex cmbZoom.ListIndex
        
        'Redraw the viewport (if allowed; some functions will prevent us from doing this, as they plan to request their own
        ' refresh after additional processing occurs)
        If Viewport.IsRenderingEnabled Then
            
            'If the user has selected any of the "fit xyz" zoom options, we want to re-center the viewport as part
            ' of updating zoom.  If they have *not* selected this, we want to preserve the current center point
            ' of the viewport.
            If (cmbZoom.ListIndex < Zoom.GetZoomFitWidthIndex) Then
                Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0), VSR_PreservePointPosition, centerXCanvas, centerYCanvas, centerXImage, centerYImage
            Else
                Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0), VSR_ResetToZero
            End If
        
            'Notify any other relevant UI elements
            Viewport.NotifyEveryoneOfViewportChanges
            
        End If
        
    End If
    
    'Synchronization of zoom buttons *always* needs to happen, regardless of canvas interaction state,
    ' to ensure they don't fall out of sync.
    
    'Disable the zoom in/out buttons when they reach the end of the available zoom levels
    cmdZoomIn.Enabled = (cmbZoom.ListIndex <> 0)
    cmdZoomOut.Enabled = (cmbZoom.ListIndex <> Zoom.GetZoomCount)
    
    'Highlight the "zoom fit" button while fit mode is active
    cmdZoomFit.Value = (cmbZoom.ListIndex = Zoom.GetZoomFitAllIndex)
    
End Sub

Private Sub cmdImgSize_Click(ByVal Shift As ShiftConstants)
    If FormMain.MainCanvas(0).IsCanvasInteractionAllowed() Then Process "Resize image", True
End Sub

Private Sub cmdZoomFit_Click(ByVal Shift As ShiftConstants)
    If FormMain.MainCanvas(0).IsCanvasInteractionAllowed Then CanvasManager.FitOnScreen
End Sub

Private Sub cmdZoomIn_Click(ByVal Shift As ShiftConstants)
    cmbZoom.ListIndex = Zoom.GetNearestZoomInIndex(cmbZoom.ListIndex)
End Sub

Private Sub cmdZoomOut_Click(ByVal Shift As ShiftConstants)
    cmbZoom.ListIndex = Zoom.GetNearestZoomOutIndex(cmbZoom.ListIndex)
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, False
    ucSupport.RequestExtraFunctionality True
    If PDMain.IsProgramRunning() Then ucSupport.RequestCursor IDC_ARROW
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDSTATUSBAR_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDStatusBar", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
    ReDim m_LinePositions(0 To 3) As Single
    m_UnitOfMeasurement = mu_Pixels
    
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
    If Not PDMain.IsProgramRunning() Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", Me.Enabled, True
End Sub

'When the canvas is cleared, we automatically disable portions of the status bar to match
Public Sub ClearCanvas()
    ReflowStatusBar False
End Sub

'Reposition all status bar elements.  This is time-consuming, so only call it if a large-scale state change
' requires it (like unloading all images).
Public Sub ReflowStatusBar(ByVal enabledState As Boolean)
    
    'Note the enabled state at a module level, in case we need to internally refresh the status bar for some reason
    m_LastEnabledState = enabledState
    
    'The zoom drop-down can now change width if a translation is active.  Make sure the zoom-in button is positioned accordingly.
    cmdZoomIn.SetLeft cmbZoom.GetLeft + cmbZoom.GetWidth + Interface.FixDPI(3)
    
    'Move the left-most line into position.  (This must be done dynamically, or it will be mispositioned
    ' on high-DPI displays)
    m_LinePositions(0) = (cmdZoomIn.GetLeft + cmdZoomIn.GetWidth) + Interface.FixDPI(6)
    
    'We will only draw subsequent interface elements if at least one image is currently loaded.
    If enabledState Then
        
        'Ensure all relevant controls are visible.  (These controls are always shown/hidden as a group,
        ' so if one is visible, we know all are visible.)
        If (Not cmdZoomFit.Visible) Then
            cmdZoomFit.Visible = True
            cmdZoomIn.Visible = True
            cmdZoomOut.Visible = True
            cmbZoom.Visible = True
            cmdImgSize.Visible = True
            lblImgSize.Visible = True
            cmbSizeUnit.Visible = True
            lblCoordinates(0).Visible = True
            'Note that lblCoordinates(1), which describes the active selection region (if any), is not
            ' forcibly made visible here - it needs to be manually activated by a caller.
        End If
        
        'Start with the "image size" button
        cmdImgSize.SetLeft m_LinePositions(0) + Interface.FixDPI(4)
        
        'After the "image size" icon comes the actual image size label.  Its position is determined by the image resize button width,
        ' plus a 4px buffer on either size (contingent on DPI)
        lblImgSize.SetLeft cmdImgSize.GetLeft + cmdImgSize.GetWidth + Interface.FixDPI(4)
        
        'The image size label is autosized.  Move the "size unit" combo box next to it, and the next vertical line
        ' separator just past it.
        cmbSizeUnit.SetLeft lblImgSize.GetLeft + lblImgSize.GetWidth + Interface.FixDPI(10)
        m_LinePositions(1) = cmbSizeUnit.GetLeft + cmbSizeUnit.GetWidth + Interface.FixDPI(10)
        
    'Images are not loaded.  Hide the lines and other items.
    Else
        m_LinePositions(1) = 0
        m_LinePositions(2) = 0
        cmdZoomFit.Visible = False
        cmdZoomIn.Visible = False
        cmdZoomOut.Visible = False
        cmbZoom.Visible = False
        cmdImgSize.Visible = False
        lblImgSize.Visible = False
        cmbSizeUnit.Visible = False
        lblCoordinates(0).Visible = False
        lblCoordinates(1).Visible = False
        lblMessages.Caption = vbNullString
    End If
    
    'We only establish positions up to the mouse coordinate label.  All items *past* that point are positioned by
    ' the dedicated message area reflow function (which is accessed much more frequently).
    FitMessageArea
    
End Sub

'Whenever this window changes size, we may need to re-align various bits of internal chrome (status bar, rulers, etc).
Public Sub FitMessageArea()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Calculate the vertical separator line positions after the mouse coordinate area (if present)...
    Dim newLeft As Long
    
    If lblCoordinates(0).Visible Then
        newLeft = m_LinePositions(1) + Interface.FixDPI(14) + Interface.FixDPI(16)
        If (lblCoordinates(0).GetLeft <> newLeft) Then lblCoordinates(0).SetLeft newLeft
        m_LinePositions(2) = newLeft + lblCoordinates(0).GetWidth + Interface.FixDPI(10)
    Else
        m_LinePositions(2) = m_LinePositions(1)
    End If
    
    '...and the selection area (if present)
    If lblCoordinates(1).Visible Then
        newLeft = m_LinePositions(2) + Interface.FixDPI(14) + Interface.FixDPI(16)
        If (lblCoordinates(1).GetLeft <> newLeft) Then lblCoordinates(1).SetLeft newLeft
        m_LinePositions(3) = newLeft + lblCoordinates(1).GetWidth + Interface.FixDPI(10)
    Else
        m_LinePositions(3) = m_LinePositions(2)
    End If
    
    'Move the message label into position (right-aligned, with a slight margin)
    If m_LastEnabledState Then
        newLeft = m_LinePositions(3)
        If m_NetworkAccessActive Then newLeft = newLeft + Interface.FixDPI(28) Else newLeft = newLeft + Interface.FixDPI(7)
    Else
        If m_NetworkAccessActive Then newLeft = Interface.FixDPI(28) Else newLeft = 0
    End If
    If (lblMessages.GetLeft <> newLeft) Then lblMessages.SetLeft newLeft
    
    'If the message label will overflow other elements of the status bar, shrink it as necessary
    Dim newMessageArea As Long
    newMessageArea = (bWidth - lblMessages.GetLeft) - Interface.FixDPI(12)
    
    If (newMessageArea < 0) Then
        lblMessages.Visible = False
    Else
        If (lblMessages.GetWidth <> newMessageArea) Then lblMessages.SetWidth newMessageArea
        If (Not lblMessages.Visible) Then lblMessages.Visible = True
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
    cmbZoom.SetTop (bHeight - cmbZoom.GetHeight) \ 2
    cmbSizeUnit.SetTop (bHeight - cmbSizeUnit.GetHeight) \ 2

    'If the control is resizing, the mouse cannot feasibly be over the image - so clear the coordinate box.
    ' Note that this will also realign all chrome elements, so we don't need a manual FitMessageArea call here.
    DisplayCanvasCoordinates 0, 0, False
    
    'No other special preparation is required for this control, so proceed with recreating the back buffer
    RedrawBackBuffer
            
End Sub

'Use this function to completely redraw the back buffer from scratch.  Note that this is computationally expensive compared to just flipping the
' existing buffer to the screen, so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer()
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long, bufferDC As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDSB_Background, Me.Enabled))
    If (bufferDC = 0) Then Exit Sub
    
    If PDMain.IsProgramRunning() Then
        
        'Render the network access icon as necessary
        If m_NetworkAccessActive Then
            sbIconNetwork.AlphaBlendToDC bufferDC, 255&, m_LinePositions(3) + Interface.FixDPI(8), Interface.FixDPI(4)
        End If
        
        'When the control is enabled, render all separator lines and a few non-button icons
        ' (like an arrow for the mouse coordinate region)
        If m_LastEnabledState Then
            
            If (Not sbIconCoords Is Nothing) And lblCoordinates(0).Visible Then sbIconCoords.AlphaBlendToDC bufferDC, 255&, m_LinePositions(1) + FixDPI(8), FixDPI(4), sbIconCoords.GetDIBWidth, sbIconCoords.GetDIBHeight
            If (Not sbIconSelection Is Nothing) And lblCoordinates(1).Visible Then sbIconSelection.AlphaBlendToDC bufferDC, 255&, m_LinePositions(2) + FixDPI(8), FixDPI(4), sbIconSelection.GetDIBWidth, sbIconSelection.GetDIBHeight
            
            Dim lineTop As Single, lineBottom As Single
            lineTop = Interface.FixDPI(1)
            lineBottom = bHeight - Interface.FixDPI(2)
            
            Dim cSurface As pd2DSurface, cPen As pd2DPen
            Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC, False
            Drawing2D.QuickCreateSolidPen cPen, 1!, m_Colors.RetrieveColor(PDSB_Separator, Me.Enabled)
            
            Dim i As Long
            For i = 0 To UBound(m_LinePositions)
                PD2D.DrawLineF cSurface, cPen, m_LinePositions(i), lineTop, m_LinePositions(i), lineBottom
            Next i
            
            Set cSurface = Nothing: Set cPen = Nothing
            
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
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)

    If ucSupport.ThemeUpdateRequired Then
        
        UpdateColorList
        
        If PDMain.IsProgramRunning() Then
            
            Dim buttonIconSize As Long
            buttonIconSize = Interface.FixDPI(16)
            
            cmdZoomFit.AssignImage "zoom_fit", , buttonIconSize, buttonIconSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
            cmdZoomIn.AssignImage "zoom_in", , buttonIconSize, buttonIconSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
            cmdZoomOut.AssignImage "zoom_out", , buttonIconSize, buttonIconSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
            cmdImgSize.AssignImage "generic_imageportrait", , buttonIconSize, buttonIconSize, usePDResamplerInstead:=IIf(OS.IsProgramCompiled(), rf_Box, rf_Automatic)
            
            'Load various status bar icons from the resource file
            If (sbIconCoords Is Nothing) Then Set sbIconCoords = New pdDIB
            IconsAndCursors.LoadResourceToDIB "generic_cursor", sbIconCoords, buttonIconSize, buttonIconSize
            
            If (sbIconSelection Is Nothing) Then Set sbIconSelection = New pdDIB
            IconsAndCursors.LoadResourceToDIB "generic_resize", sbIconSelection, buttonIconSize, buttonIconSize, resampleAlgorithm:=GP_IM_NearestNeighbor
            sbIconSelection.SuspendDIB
            
            If (sbIconNetwork Is Nothing) Then Set sbIconNetwork = New pdDIB
            IconsAndCursors.LoadResourceToDIB "generic_network", sbIconNetwork, buttonIconSize, buttonIconSize
            sbIconNetwork.SuspendDIB
            
        End If
        
        'Rebuild all drop-down boxes (so that translations can be applied)
        Dim backupZoomIndex As Long, backupSizeIndex As Long
        backupZoomIndex = cmbZoom.ListIndex
        backupSizeIndex = cmbSizeUnit.ListIndex
        
        'Repopulate zoom dropdown text
        Zoom.InitializeZoomEngine
        Zoom.PopulateZoomDropdown cmbZoom, backupZoomIndex
        Me.PopulateSizeUnits
        
        'Auto-size the newly populated combo boxes, according to the width of their longest entries
        cmbZoom.SetWidthAutomatically
        cmbSizeUnit.SetWidthAutomatically
        
        cmdZoomFit.AssignTooltip "Fit image on screen"
        cmdZoomIn.AssignTooltip "Zoom in"
        cmdZoomOut.AssignTooltip "Zoom out"
        cmdImgSize.AssignTooltip "Resize image"
        cmbZoom.AssignTooltip "Change viewport zoom"
        cmbSizeUnit.AssignTooltip "Change the image size unit displayed to the left of this box"
        
        Dim sbBackColor As Long
        sbBackColor = m_Colors.RetrieveColor(PDSB_Background, Me.Enabled)
        UserControl.BackColor = sbBackColor
        
        lblCoordinates(0).BackColor = sbBackColor
        lblCoordinates(1).BackColor = sbBackColor
        lblImgSize.BackColor = sbBackColor
        lblMessages.BackColor = sbBackColor
        
        cmdZoomFit.BackColor = sbBackColor
        cmdZoomIn.BackColor = sbBackColor
        cmdZoomOut.BackColor = sbBackColor
        cmdImgSize.BackColor = sbBackColor
        cmbZoom.BackgroundColor = sbBackColor
        cmbSizeUnit.BackgroundColor = sbBackColor
        
        lblCoordinates(0).UpdateAgainstCurrentTheme
        lblCoordinates(1).UpdateAgainstCurrentTheme
        lblImgSize.UpdateAgainstCurrentTheme
        lblMessages.UpdateAgainstCurrentTheme
        
        cmdZoomFit.UpdateAgainstCurrentTheme
        cmdZoomIn.UpdateAgainstCurrentTheme
        cmdZoomOut.UpdateAgainstCurrentTheme
        cmdImgSize.UpdateAgainstCurrentTheme
        
        cmbZoom.UpdateAgainstCurrentTheme
        cmbSizeUnit.UpdateAgainstCurrentTheme
        
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd, False
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        
        'Fix combo box positioning (important on high-DPI displays, or if the active font has changed)
        cmbZoom.Top = (UserControl.ScaleHeight - cmbZoom.GetHeight) \ 2
        cmbSizeUnit.Top = (UserControl.ScaleHeight - cmbSizeUnit.GetHeight) \ 2
        
        'Restore zoom and size unit indices
        cmbZoom.ListIndex = backupZoomIndex
        cmbSizeUnit.ListIndex = backupSizeIndex
        
        'Note that we don't actually move the last line status bar; that is handled by DisplayImageCoordinates itself
        If PDImages.IsImageActive() Then
            DisplayImageSize PDImages.GetActiveImage(), False
        Else
            DisplayImageSize Nothing, True
        End If
        
        DisplayCanvasCoordinates 0, 0, True
        FitMessageArea
        
    End If
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
