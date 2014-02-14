VERSION 5.00
Begin VB.UserControl pdCanvas 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000003&
   ClientHeight    =   7695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   886
   Begin VB.PictureBox picProgressBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   886
      TabIndex        =   6
      Top             =   7095
      Visible         =   0   'False
      Width           =   13290
   End
   Begin VB.PictureBox picScrollV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   5520
      ScaleHeight     =   361
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picScrollH 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   361
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.PictureBox picStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   345
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   886
      TabIndex        =   0
      Top             =   7350
      Width           =   13290
      Begin PhotoDemon.jcbutton cmdZoomIn 
         Height          =   345
         Left            =   1440
         TabIndex        =   8
         Top             =   0
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483626
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "pdCanvas.ctx":0000
         PictureAlign    =   7
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "Zoom in"
         ColorScheme     =   3
      End
      Begin VB.ComboBox CmbZoom 
         CausesValidation=   0   'False
         Height          =   315
         ItemData        =   "pdCanvas.ctx":0852
         Left            =   450
         List            =   "pdCanvas.ctx":0854
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   15
         Width           =   960
      End
      Begin PhotoDemon.jcbutton cmdZoomOut 
         Height          =   345
         Left            =   30
         TabIndex        =   9
         Top             =   0
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483626
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "pdCanvas.ctx":0856
         PictureAlign    =   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ToolTip         =   "Zoom Out"
         ColorScheme     =   3
      End
      Begin VB.Line lineStatusBar 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   126
         X2              =   126
         Y1              =   1
         Y2              =   22
      End
      Begin VB.Label lblImgSize 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   2280
         TabIndex        =   3
         Top             =   60
         Width           =   345
      End
      Begin VB.Line lineStatusBar 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   236
         X2              =   236
         Y1              =   1
         Y2              =   22
      End
      Begin VB.Label lblCoordinates 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000016&
         Caption         =   "(X, Y)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   4230
         TabIndex        =   2
         Top             =   60
         Width           =   525
      End
      Begin VB.Line lineStatusBar 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   360
         X2              =   360
         Y1              =   1
         Y2              =   22
      End
      Begin VB.Label lblMessages 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(messages will appear here at run-time)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   9810
         TabIndex        =   1
         Top             =   60
         Width           =   3255
      End
   End
End
Attribute VB_Name = "pdCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Canvas User Control (previously a standalone form)
'Copyright ©2002-2014 by Tanner Helland
'Created: 11/29/02
'Last updated: 11/February/14
'Last update: move color management stuff out of this control and back into FormMain
'
'Every time the user loads an image, one of these forms is spawned. This form also interfaces with several
' specialized program components in the pdWindowManager class.
'
'As I start including more and more paint tools, this form is going to become a bit more complex. Stay tuned.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


Private Const SM_CXVSCROLL As Long = 2
Private Const SM_CYHSCROLL As Long = 3

'These are used to track use of the Ctrl, Alt, and Shift keys
Private ShiftDown As Boolean, CtrlDown As Boolean, AltDown As Boolean

'Track mouse button use on this canvas
Private lMouseDown As Boolean, rMouseDown As Boolean

'Track mouse movement on this canvas
Private hasMouseMoved As Long

'Track initial mouse button locations
Private m_initMouseX As Double, m_initMouseY As Double

'An outside class provides access to specialized mouse events (like mousewheel and forward/back keys)
Private WithEvents cMouseEvents As bluMouseEvents
Attribute cMouseEvents.VB_VarHelpID = -1

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'To improve performance, we can ask the canvas to not refresh itself until we say so.
Private m_suspendRedraws As Boolean

'API scroll bars are used in place of crappy VB ones
Private WithEvents HScroll As pdScrollAPI
Attribute HScroll.VB_VarHelpID = -1
Private WithEvents VScroll As pdScrollAPI
Attribute VScroll.VB_VarHelpID = -1

'Icons rendered to the scroll bar.  Rather than constantly reloading them from file, we cache them at initialization.
Dim sbIconSize As pdDIB, sbIconCoords As pdDIB

'Use this function to forcibly prevent the canvas from redrawing itself.  REDRAWS WILL NOT HAPPEN AGAIN UNTIL YOU RESTORE ACCESS!
Public Sub setRedrawSuspension(ByVal newRedrawValue As Boolean)
    m_suspendRedraws = newRedrawValue
End Sub

Public Property Get BackColor() As Long
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(newBackColor As Long)
    UserControl.BackColor = newBackColor
    UserControl.Refresh
End Property

Public Sub clearCanvas()
    UserControl.Picture = LoadPicture("")
    UserControl.Refresh
End Sub

'Get/Set scroll bar visibility
Public Function getScrollVisibility(ByVal barType As PD_ORIENTATION) As Boolean

    If barType = PD_HORIZONTAL Then
        getScrollVisibility = picScrollH.Visible
    Else
        getScrollVisibility = picScrollV.Visible
    End If

End Function

Public Sub setScrollVisibility(ByVal barType As PD_ORIENTATION, ByVal newVisibility As Boolean)
    
    If barType = PD_HORIZONTAL Then
        picScrollH.Visible = newVisibility
    Else
        picScrollV.Visible = newVisibility
    End If
    
End Sub

'Get/Set scroll bar value
Public Function getScrollValue(ByVal barType As PD_ORIENTATION) As Long

    If barType = PD_HORIZONTAL Then
        getScrollValue = HScroll.Value
    Else
        getScrollValue = VScroll.Value
    End If

End Function

Public Sub setScrollValue(ByVal barType As PD_ORIENTATION, ByVal newValue As Long)
    
    Select Case barType
    
        Case PD_HORIZONTAL
            HScroll.Value = newValue
            
        Case PD_VERTICAL
            VScroll.Value = newValue
        
        Case PD_BOTH
            HScroll.Value = newValue
            VScroll.Value = newValue
        
    End Select
    
End Sub

'Get/Set scroll max/min
Public Function getScrollMax(ByVal barType As PD_ORIENTATION) As Long

    If barType = PD_HORIZONTAL Then
        getScrollMax = HScroll.Max
    Else
        getScrollMax = VScroll.Max
    End If

End Function

Public Function getScrollMin(ByVal barType As PD_ORIENTATION) As Long

    If barType = PD_HORIZONTAL Then
        getScrollMin = HScroll.Min
    Else
        getScrollMin = VScroll.Min
    End If

End Function

Public Sub setScrollMax(ByVal barType As PD_ORIENTATION, ByVal newMax As Long)
    
    If barType = PD_HORIZONTAL Then
        HScroll.Max = newMax
    Else
        VScroll.Max = newMax
    End If
    
End Sub

Public Sub setScrollMin(ByVal barType As PD_ORIENTATION, ByVal newMin As Long)
    
    If barType = PD_HORIZONTAL Then
        HScroll.Min = newMin
    Else
        VScroll.Min = newMin
    End If
    
End Sub

'Get scroll bar size.  Note that scroll bar size cannot be set by external functions; it is automatically set to the system default
' upon user control initialization.
Public Function getScrollWidth(ByVal barType As PD_ORIENTATION) As Long

    If barType = PD_HORIZONTAL Then
        getScrollWidth = picScrollH.Width
    Else
        getScrollWidth = picScrollV.Width
    End If

End Function

Public Function getScrollHeight(ByVal barType As PD_ORIENTATION) As Long

    If barType = PD_HORIZONTAL Then
        getScrollHeight = picScrollH.Height
    Else
        getScrollHeight = picScrollV.Height
    End If

End Function

'Get scroll bar position (left, top).
Public Function getScrollLeft(ByVal barType As PD_ORIENTATION) As Long

    If barType = PD_HORIZONTAL Then
        getScrollLeft = picScrollH.Left
    Else
        getScrollLeft = picScrollV.Left
    End If

End Function

Public Function getScrollTop(ByVal barType As PD_ORIENTATION) As Long

    If barType = PD_HORIZONTAL Then
        getScrollTop = picScrollH.Top
    Else
        getScrollTop = picScrollV.Top
    End If

End Function

'Move a scroll bar to a new position
Public Sub moveScrollBar(ByVal barType As PD_ORIENTATION, ByVal newX As Long, ByVal newY As Long, ByVal newWidth As Long, ByVal newHeight As Long)

    If barType = PD_HORIZONTAL Then
        picScrollH.Move newX, newY, newWidth, newHeight
    Else
        picScrollV.Move newX, newY, newWidth, newHeight
    End If

End Sub

'Set scroll bar LargeChange value
Public Sub setScrollLargeChange(ByVal barType As PD_ORIENTATION, ByVal newLargeChange As Long)
        
    If barType = PD_HORIZONTAL Then
        HScroll.LargeChange = newLargeChange
    Else
        VScroll.LargeChange = newLargeChange
    End If
        
End Sub

Public Sub displayImageSize(ByVal iWidth As Long, ByVal iHeight As Long, Optional ByVal clearSize As Boolean = False)
    
    'The size display is cleared whenever the user has no images loaded
    If clearSize Then
        lblImgSize.Caption = ""
    
    'When size IS displayed, we must also refresh the status bar (now that it dynamically aligns its contents)
    Else
        lblImgSize.Caption = iWidth & " x " & iHeight
        drawStatusBarIcons True
        
    End If
    
    lblImgSize.Refresh
    
End Sub

Public Sub displayCanvasMessage(ByRef cMessage As String)
    lblMessages.Caption = cMessage
    lblMessages.Refresh
End Sub

'Display the current mouse coordinates
Public Sub displayCanvasCoordinates(ByVal xCoord As Long, ByVal yCoord As Long, Optional ByVal clearCoords As Boolean = False)
    
    If clearCoords Then
        lblCoordinates.Caption = ""
    
    'When coordinates are displayed, we must also refresh the status bar (now that it dynamically aligns its contents)
    Else
        lblCoordinates.Caption = "(" & xCoord & "," & yCoord & ")"
        
    End If
    
    'Align the right-hand line control with the newly captioned label
    lineStatusBar(2).x1 = lblCoordinates.Left + lblCoordinates.Width + fixDPI(10)
    lineStatusBar(2).x2 = lineStatusBar(2).x1
    
    lblCoordinates.Refresh
    
End Sub

Public Sub requestBufferSync()
    UserControl.Picture = UserControl.Image
    UserControl.Refresh
End Sub

Public Function getCanvasWidth() As Long
    getCanvasWidth = UserControl.ScaleWidth
End Function

Public Function getCanvasHeight() As Long
    getCanvasHeight = UserControl.ScaleHeight
End Function

Public Function getStatusBarHeight() As Long
    getStatusBarHeight = picStatusBar.ScaleHeight
End Function

Public Function getProgBarReference() As PictureBox
    Set getProgBarReference = picProgressBar
End Function

Public Property Get hWnd()
    hWnd = UserControl.hWnd
End Property

Public Property Get hDC()
    hDC = UserControl.hDC
End Property

Public Sub enableZoomIn(ByVal isEnabled As Boolean)
    cmdZoomIn.Enabled = isEnabled
End Sub

Public Sub enableZoomOut(ByVal isEnabled As Boolean)
    cmdZoomOut.Enabled = isEnabled
End Sub

Public Function getZoomDropDownReference() As ComboBox
    Set getZoomDropDownReference = CmbZoom
End Function

Private Sub CmbZoom_Click()

    'Only process zoom changes if an image has been loaded
    If g_OpenImageCount > 0 Then
    
        'Store the current zoom value in this object (so the user can switch between images without losing zoom values)
        pdImages(g_CurrentImage).currentZoomValue = CmbZoom.ListIndex
        
        'Disable the zoom in/out buttons when they reach the end of the available zoom levels
        If CmbZoom.ListIndex = 0 Then
            cmdZoomIn.Enabled = False
        Else
            If Not cmdZoomIn.Enabled Then cmdZoomIn.Enabled = True
        End If
        
        If CmbZoom.ListIndex = CmbZoom.ListCount - 1 Then
            cmdZoomOut.Enabled = False
        Else
            If Not cmdZoomOut.Enabled Then cmdZoomOut.Enabled = True
        End If
        
        'Redraw the viewport
        PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "zoom changed by primary drop-down box"
        
    End If

End Sub

Private Sub cmdZoomIn_Click()
    FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex - 1
End Sub

Private Sub cmdZoomOut_Click()
    FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex + 1
End Sub

Private Sub cMouseEvents_MouseIn()

    'Notify the window manager that toolbars need to be made translucent
    g_WindowManager.notifyMouseMoveOverCanvas
    
    'If no images have been loaded, reset the cursor
    If g_OpenImageCount = 0 Then
        setArrowCursorToHwnd Me.hWnd
    End If

End Sub

Private Sub UserControl_Initialize()

    If g_UserModeFix Then
    
        'Enable mouse subclassing for events like mousewheel, forward/back keys, enter/leave
        Set cMouseEvents = New bluMouseEvents
        cMouseEvents.Attach UserControl.hWnd
        
        'Assign tooltips manually (so theming is supported)
        Set m_ToolTip = New clsToolTip
        m_ToolTip.Create Me
        m_ToolTip.MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
        m_ToolTip.DelayTime(ttDelayShow) = 10000
        
        m_ToolTip.AddTool CmbZoom, "Click to adjust image zoom"
        
        'Allow the control to generate its own redraw requests
        m_suspendRedraws = False
        
        'Set scroll bar size to match the current system default (which changes based on DPI, theming, and other factors)
        picScrollH.Height = GetSystemMetrics(SM_CYHSCROLL)
        picScrollV.Width = GetSystemMetrics(SM_CXVSCROLL)
        
        'Initialize scroll bars
        Set HScroll = New pdScrollAPI
        Set VScroll = New pdScrollAPI
        
        HScroll.initializeScrollBarWindow picScrollH.hWnd, True, 0, 10, 0, 1, 1
        VScroll.initializeScrollBarWindow picScrollV.hWnd, False, 0, 10, 0, 1, 1
        
    End If
    
End Sub

'Mousekey back triggers the same thing as clicking Undo
Private Sub cMouseEvents_MouseBackButtonDown(ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
    If pdImages(g_CurrentImage).IsActive Then
        If pdImages(g_CurrentImage).undoManager.getUndoState Then Process "Undo", , , False
    End If
End Sub

'Mousekey forward triggers the same thing as clicking Redo
Private Sub cMouseEvents_MouseForwardButtonDown(ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
    If pdImages(g_CurrentImage).IsActive Then
        If pdImages(g_CurrentImage).undoManager.getRedoState Then Process "Redo", , , False
    End If
End Sub

Public Sub cMouseEvents_MouseHScroll(ByVal CharsScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)

    'Horizontal scrolling - only trigger if the horizontal scroll bar is visible AND a shift key has been pressed
    If picScrollH.Visible And Not (Shift And vbCtrlMask) Then
        
        If CharsScrolled < 0 Then
        
            m_suspendRedraws = True
            
            If HScroll.Value + HScroll.LargeChange > HScroll.Max Then
                HScroll.Value = HScroll.Max
            Else
                HScroll.Value = HScroll.Value + HScroll.LargeChange
            End If
            
            m_suspendRedraws = False
            
            ScrollViewport pdImages(g_CurrentImage), Me
        
        ElseIf CharsScrolled > 0 Then
        
            m_suspendRedraws = True
            
            If HScroll.Value - HScroll.LargeChange < HScroll.Min Then
                HScroll.Value = HScroll.Min
            Else
                HScroll.Value = HScroll.Value - HScroll.LargeChange
            End If
            
            m_suspendRedraws = False
            
            ScrollViewport pdImages(g_CurrentImage), Me
            
        End If
        
    End If
  
End Sub

'When the mouse leaves the window, if no buttons are down, clear the coordinate display.
' (We must check for button states because the user is allowed to do things like drag selection nodes outside the image.)
Private Sub cMouseEvents_MouseOut()
    If (Not lMouseDown) And (Not rMouseDown) Then ClearImageCoordinatesDisplay
End Sub

Public Sub cMouseEvents_MouseVScroll(ByVal LinesScrolled As Single, ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
    
    'Vertical scrolling - only trigger it if the vertical scroll bar is actually visible
    If picScrollV.Visible And Not (Shift And vbCtrlMask) Then
      
        If LinesScrolled < 0 Then
            
            m_suspendRedraws = True
            
            If VScroll.Value + VScroll.LargeChange > VScroll.Max Then
                VScroll.Value = VScroll.Max
            Else
                VScroll.Value = VScroll.Value + VScroll.LargeChange
            End If
            
            m_suspendRedraws = False
            
            ScrollViewport pdImages(g_CurrentImage), Me
        
        ElseIf LinesScrolled > 0 Then
            
            m_suspendRedraws = True
            
            If VScroll.Value - VScroll.LargeChange < VScroll.Min Then
                VScroll.Value = VScroll.Min
            Else
                VScroll.Value = VScroll.Value - VScroll.LargeChange
            End If
            
            m_suspendRedraws = False
            
            ScrollViewport pdImages(g_CurrentImage), Me
            
        End If
    
    End If
    
    'NOTE: horizontal scrolling is now handled in the separate _MouseHScroll event.  This is necessary to handle mice with
    '      a dedicated horizontal scroller.
    
    'Zooming - only trigger when Ctrl has been pressed
    If (Shift And vbCtrlMask) Then
      
        If LinesScrolled > 0 Then
            
            If CmbZoom.ListIndex > 0 Then CmbZoom.ListIndex = CmbZoom.ListIndex - 1
            'NOTE: a manual call to PrepareViewport is no longer required, as changing the combo box will automatically trigger a redraw
            
        ElseIf LinesScrolled < 0 Then
            
            If CmbZoom.ListIndex < (CmbZoom.ListCount - 1) Then CmbZoom.ListIndex = CmbZoom.ListIndex + 1
            
        End If
        
    End If
  
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    
    'If a selection is active, notify it of any changes in the shift key (which is used to request 1:1 selections)
    If pdImages(g_CurrentImage).selectionActive Then pdImages(g_CurrentImage).mainSelection.requestSquare ShiftDown
    
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    
    'If a selection is active, notify it of any changes in the shift key (which is used to request 1:1 selections)
    If pdImages(g_CurrentImage).selectionActive Then pdImages(g_CurrentImage).mainSelection.requestSquare ShiftDown
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If the main form is disabled, exit
    If Not FormMain.Enabled Then Exit Sub
    
    'If no images have been loaded, exit
    If g_OpenImageCount = 0 Then Exit Sub
    
    'If the image has not yet been loaded, exit
    If Not pdImages(g_CurrentImage).loadedSuccessfully Then Exit Sub
    
    'These variables will hold the corresponding (x,y) coordinates on the IMAGE - not the VIEWPORT.
    ' (This is important if the user has zoomed into an image, and used scrollbars to look at a different part of it.)
    Dim imgX As Double, imgY As Double
    imgX = -1
    imgY = -1
    
    'Check mouse button use
    If Button = vbLeftButton Then
        
        lMouseDown = True
            
        hasMouseMoved = 0
            
        'Remember this location
        m_initMouseX = x
        m_initMouseY = y
            
        'Display the image coordinates under the mouse pointer
        displayImageCoordinates x, y, pdImages(g_CurrentImage), Me, imgX, imgY
        
        'Any further processing depends on which tool is currently active
        Select Case g_CurrentTool
        
            'Drag-to-pan canvas
            Case NAV_DRAG
                UserControl.MousePointer = vbCustom
                setPNGCursorToHwnd Me.hWnd, "HANDCLOSED", 0, 0
                setInitialCanvasScrollValues FormMain.mainCanvas(0)
        
            'Rectangular selection
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
            
                'Check to see if a selection is already active.  If it is, see if the user is allowed to transform it.
                If pdImages(g_CurrentImage).selectionActive Then
                
                    'Check the mouse coordinates of this click.
                    Dim sCheck As Long
                    sCheck = findNearestSelectionCoordinates(x, y, pdImages(g_CurrentImage), Me)
                    
                    'If that function did not return zero, notify the selection and exit
                    If (sCheck <> 0) And pdImages(g_CurrentImage).mainSelection.isTransformable Then
                    
                        'If the selection type matches the current selection tool, start transforming the selection.
                        If (pdImages(g_CurrentImage).mainSelection.getSelectionShape = getSelectionTypeFromCurrentTool()) Then
                        
                            'Back up the current selection settings - those will be saved in a later step as part of the Undo/Redo chain
                            pdImages(g_CurrentImage).mainSelection.setBackupParamString
                            
                            'Initialize a selection transformation
                            pdImages(g_CurrentImage).mainSelection.setTransformationType sCheck
                            pdImages(g_CurrentImage).mainSelection.setInitialTransformCoordinates imgX, imgY
                            
                            Exit Sub
                            
                        'If the selection type does NOT match the current selection tool, select the proper tool, then start transforming
                        ' the selection.
                        Else
                        
                            toolbar_Tools.selectNewTool getRelevantToolFromSelectType()
                            
                            'Back up the current selection settings - those will be saved in a later step as part of the Undo/Redo chain
                            pdImages(g_CurrentImage).mainSelection.setBackupParamString
                            
                            'Initialize a selection transformation
                            pdImages(g_CurrentImage).mainSelection.setTransformationType sCheck
                            pdImages(g_CurrentImage).mainSelection.setInitialTransformCoordinates imgX, imgY
                            
                            Exit Sub
                        
                        End If
                                        
                    'If it did return zero, erase any existing selection and start a new one
                    Else
                    
                        'Back up the current selection settings - those will be saved in a later step as part of the Undo/Redo chain
                        pdImages(g_CurrentImage).mainSelection.setBackupParamString
                    
                        initSelectionByPoint imgX, imgY
                    
                    End If
                
                Else
                    
                    initSelectionByPoint imgX, imgY
                    
                End If
            
        End Select
        
    End If
    
    If Button = vbRightButton Then rMouseDown = True

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    'If the main form is disabled, exit
    If Not FormMain.Enabled Then Exit Sub
    
    'If no images have been loaded, exit
    If g_OpenImageCount = 0 Then Exit Sub
    
    If pdImages(g_CurrentImage) Is Nothing Then Exit Sub
    
    'If the image has not yet been loaded, exit
    If Not pdImages(g_CurrentImage).loadedSuccessfully Then Exit Sub
    
    hasMouseMoved = hasMouseMoved + 1
    
    'These variables will hold the corresponding (x,y) coordinates on the image - NOT the viewport
    Dim imgX As Double, imgY As Double
    imgX = -1
    imgY = -1
    
    'Check the left mouse button
    If lMouseDown Then
    
        Select Case g_CurrentTool
        
            'Drag-to-pan canvas
            Case NAV_DRAG
                UserControl.MousePointer = vbCustom
                setPNGCursorToHwnd Me.hWnd, "HANDCLOSED", 0, 0
                panImageCanvas m_initMouseX, m_initMouseY, x, y, pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
            'Selection tools
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
    
                'First, check to see if a selection is active. (In the future, we will be checking for other tools as well.)
                If pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.isTransformable Then
                                        
                    'Display the image coordinates under the mouse pointer
                    displayImageCoordinates x, y, pdImages(g_CurrentImage), Me, imgX, imgY
                    
                    'If the SHIFT key is down, notify the selection engine that a square shape is requested
                    pdImages(g_CurrentImage).mainSelection.requestSquare ShiftDown
                    
                    'Pass new points to the active selection
                    pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                    syncTextToCurrentSelection g_CurrentImage
                                        
                End If
                
                'Force a redraw of the viewport
                If hasMouseMoved > 1 Then RenderViewport pdImages(g_CurrentImage), Me
                
        End Select
    
    'This else means the LEFT mouse button is NOT down
    Else
    
        Select Case g_CurrentTool
        
            'Drag-to-navigate
            Case NAV_DRAG
            
                'If the cursor is over the image, change to an arrow cursor
                If isMouseOverImage(x, y, pdImages(g_CurrentImage)) Then
                    UserControl.MousePointer = vbCustom
                    setPNGCursorToHwnd Me.hWnd, "HANDOPEN", 0, 0
                Else
                    setArrowCursorToHwnd Me.hWnd
                End If
        
            'Standard selection tools
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
            
                'Next, check to see if a selection is active. If it is, we need to provide the user with visual cues about their
                ' ability to resize the selection.
                If Not pdImages(g_CurrentImage).mainSelection Is Nothing Then
                    If pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.isTransformable Then
                    
                        'This routine will return a best estimate for the location of the mouse.  We then pass its value
                        ' to a sub that will use it to select the most appropriate mouse cursor.
                        Dim sCheck As Long
                        sCheck = findNearestSelectionCoordinates(x, y, pdImages(g_CurrentImage), Me)
                        
                        'Based on that return value, assign a new mouse cursor to the form
                        setSelectionCursor sCheck
                        
                        'Set the active selection's transformation type to match
                        pdImages(g_CurrentImage).mainSelection.setTransformationType sCheck
                        
                    Else
                    
                        'Check the location of the mouse to see if it's over the image, and set the cursor accordingly.
                        ' (NOTE: at present this has no effect, but once paint tools are implemented, it will be more important.)
                        If isMouseOverImage(x, y, pdImages(g_CurrentImage)) Then
                            setArrowCursorToHwnd Me.hWnd
                        Else
                            setArrowCursorToHwnd Me.hWnd
                        End If
                    
                    End If
                End If
        
            Case Else
        
                'Check the location of the mouse to see if it's over the image, and set the cursor accordingly.
                ' (NOTE: at present this has no effect, but once paint tools are implemented, it will be more important.)
                If isMouseOverImage(x, y, pdImages(g_CurrentImage)) Then
                    setArrowCursorToHwnd Me.hWnd
                Else
                    setArrowCursorToHwnd Me.hWnd
                End If
            
        End Select
        
    End If
    
    'Display the image coordinates under the mouse pointer (but only if this is the currently active image)
    displayImageCoordinates x, y, pdImages(g_CurrentImage), Me
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    'If no images have been loaded, exit
    If g_OpenImageCount = 0 Then Exit Sub

    'If the image has not yet been loaded, exit
    If Not pdImages(g_CurrentImage).loadedSuccessfully Then Exit Sub
        
    'Check mouse buttons
    If Button = vbLeftButton Then
    
        lMouseDown = False
    
        Select Case g_CurrentTool
        
            'Click-to-drag navigation
            Case NAV_DRAG
                UserControl.MousePointer = vbCustom
                setPNGCursorToHwnd Me.hWnd, "HANDOPEN", 0, 0
                
            'Selection tools
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
            
                'If a selection was being drawn, lock it into place
                If pdImages(g_CurrentImage).selectionActive Then
                    
                    'Check to see if this mouse location is the same as the initial mouse press. If it is, and that particular
                    ' point falls outside the selection, clear the selection from the image.
                    If ((x = m_initMouseX) And (y = m_initMouseY) And (hasMouseMoved <= 1) And (findNearestSelectionCoordinates(x, y, pdImages(g_CurrentImage), Me) = 0)) Or ((pdImages(g_CurrentImage).mainSelection.selWidth <= 0) And (pdImages(g_CurrentImage).mainSelection.selHeight <= 0)) Then
                        Process "Remove selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, 2, g_CurrentTool
                    Else
                    
                        'Check to see if all selection coordinates are invalid.  If they are, forget about this selection.
                        If pdImages(g_CurrentImage).mainSelection.areAllCoordinatesInvalid Then
                            Process "Remove selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, 2, g_CurrentTool
                        Else
                        
                            'Depending on the type of transformation that may or may not have been applied, call the appropriate processor
                            ' function.  This has no practical purpose at present, except to give the user a pleasant name for this action.
                            Select Case pdImages(g_CurrentImage).mainSelection.getTransformationType
                            
                                'Creating a new selection
                                Case 0
                                    Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, 2, g_CurrentTool
                                    
                                'Moving an existing selection
                                Case 9
                                    Process "Move selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, 2, g_CurrentTool
                                    
                                'Anything else is assumed to be resizing an existing selection
                                Case Else
                                    Process "Resize selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, 2, g_CurrentTool
                                    
                            End Select
                            
                        End If
                        
                    End If
                    
                    'Force a redraw of the screen
                    RenderViewport pdImages(g_CurrentImage), Me
                    
                Else
                    'If the selection is not active, make sure it stays that way
                    pdImages(g_CurrentImage).mainSelection.lockRelease
                End If
                
                'Synchronize the selection text box values with the final selection
                syncTextToCurrentSelection g_CurrentImage
                
            Case Else
                    
        End Select
                        
    End If
    
    If Button = vbRightButton Then rMouseDown = False
    
    'makeFormPretty Me
    'setArrowCursorToHwnd UserControl.hWnd
        
    'Reset the mouse movement tracker
    hasMouseMoved = 0
    
End Sub

'(This code is copied from FormMain's OLEDragDrop event - please mirror any changes there)
Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub

    'Verify that the object being dragged is some sort of file or file list
    If Data.GetFormat(vbCFFiles) Then
        
        'Copy the filenames into an array
        Dim sFile() As String
        ReDim sFile(0 To Data.Files.Count) As String
        
        Dim oleFilename
        Dim tmpString As String
        
        Dim countFiles As Long
        countFiles = 0
        
        For Each oleFilename In Data.Files
            tmpString = CStr(oleFilename)
            If tmpString <> "" Then
                sFile(countFiles) = tmpString
                countFiles = countFiles + 1
            End If
        Next oleFilename
        
        'Because the OLE drop may include blank strings, verify the size of the array against countFiles
        ReDim Preserve sFile(0 To countFiles - 1) As String
        
        'Pass the list of filenames to PreLoadImage, which will load the images one-at-a-time
        PreLoadImage sFile
        
    End If
    
End Sub

'(This code is copied from FormMain's OLEDragOver event - please mirror any changes there)
Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub

    'Check to make sure the type of OLE object is files
    If Data.GetFormat(vbCFFiles) Then
        'Inform the source (Explorer, in this case) that the files will be treated as "copied"
        Effect = vbDropEffectCopy And Effect
    Else
        'If it's not files, don't allow a drop
        Effect = vbDropEffectNone
    End If
    
End Sub

Private Sub HScroll_Scroll()
    If (Not m_suspendRedraws) Then ScrollViewport pdImages(g_CurrentImage), Me
End Sub

Private Sub UserControl_Resize()
    fixChromeLayout
End Sub

Private Sub UserControl_Show()

    If g_UserModeFix Then
    
        'Convert all labels to the current interface font
        If Len(g_InterfaceFont) = 0 Then g_InterfaceFont = "Segoe UI"
        
        'XP users may not have Segoe UI available, so default to "On Error Resume NexT" for now
        On Error Resume Next
        
        lblCoordinates.FontName = g_InterfaceFont
        lblImgSize.FontName = g_InterfaceFont
        lblMessages.FontName = g_InterfaceFont
        
        'Load various status bar icons from the resource file
        Set sbIconSize = New pdDIB
        Set sbIconCoords = New pdDIB
        
        loadResourceToDIB "SB_IMG_SIZE", sbIconSize
        loadResourceToDIB "SB_MOUSE_POS", sbIconCoords
        
    End If

End Sub

Private Sub VScroll_Scroll()
    If (Not m_suspendRedraws) Then ScrollViewport pdImages(g_CurrentImage), Me
End Sub

'Selection tools utilize a variety of cursors.  To keep the main MouseMove sub clean, cursors are set separately
' by this routine.
Private Sub setSelectionCursor(ByVal transformID As Long)

    Select Case pdImages(g_CurrentImage).mainSelection.getSelectionShape()

        Case sRectangle, sCircle
        
            'For a rectangle or circle selection, the possible transform IDs are:
            ' 0 - Cursor is not near a selection point
            ' 1 - NW corner
            ' 2 - NE corner
            ' 3 - SE corner
            ' 4 - SW corner
            ' 5 - N edge
            ' 6 - E edge
            ' 7 - S edge
            ' 8 - W edge
            ' 9 - interior of selection, not near a corner or edge
            Select Case transformID
        
                Case 0
                    setArrowCursorToHwnd Me.hWnd
                Case 1
                    setSizeNWSECursor Me
                Case 2
                    setSizeNESWCursor Me
                Case 3
                    setSizeNWSECursor Me
                Case 4
                    setSizeNESWCursor Me
                Case 5
                    setSizeNSCursor Me
                Case 6
                    setSizeWECursor Me
                Case 7
                    setSizeNSCursor Me
                Case 8
                    setSizeWECursor Me
                Case 9
                    setSizeAllCursor Me
                    
            End Select
            
        'For a line selection, the possible transform IDs are:
        ' 0 - Cursor is not near an endpoint
        ' 1 - Near x1/y1
        ' 2 - Near x2/y2
        Case sLine
        
            Select Case transformID
                Case 0
                    setArrowCursorToHwnd Me.hWnd
                Case 1
                    setSizeAllCursor Me
                Case 2
                    setSizeAllCursor Me
            End Select
        
    End Select

End Sub

'Selections can be initiated several different ways.  To cut down on duplicated code, all new selection instances for this form are referred
' to this function.  Initial X/Y values are required.
Private Sub initSelectionByPoint(ByVal x As Double, ByVal y As Double)

    'I don't have a good explanation, but without DoEvents here, creating a line selection for the first
    ' time may inexplicably fail.  While I try to track down the exact cause, I'll leave this here to
    ' maintain desired behavior...
    DoEvents
    
    'Activate the attached image's primary selection
    pdImages(g_CurrentImage).selectionActive = True
    pdImages(g_CurrentImage).mainSelection.lockRelease
    
    'Populate a variety of selection attributes using a single shorthand declaration.  A breakdown of these
    ' values and what they mean can be found in the corresponding pdSelection.initFromParamString function
    pdImages(g_CurrentImage).mainSelection.initFromParamString buildParams(getSelectionTypeFromCurrentTool(), toolbar_Tools.cmbSelType(0).ListIndex, toolbar_Tools.cmbSelSmoothing(0).ListIndex, toolbar_Tools.sltSelectionFeathering.Value, toolbar_Tools.sltSelectionBorder.Value, toolbar_Tools.sltCornerRounding.Value, toolbar_Tools.sltSelectionLineWidth.Value, 0, 0, 0, 0, 0, 0, 0, 0)
    
    'Set the first two coordinates of this selection to this mouseclick's location
    pdImages(g_CurrentImage).mainSelection.setInitialCoordinates x, y
    syncTextToCurrentSelection g_CurrentImage
    pdImages(g_CurrentImage).mainSelection.requestNewMask
        
    'Make the selection tools visible
    metaToggle tSelection, True
    metaToggle tSelectionTransform, True
                        
End Sub

'Whenever this window changes size, we may need to re-align various bits of internal chrome (status bar, rulers, etc).  Call this function
' to do so.
Public Sub fixChromeLayout()
    
    'Move the message label into position (right-aligned, with a slight margin)
    Dim newLeft As Long
    newLeft = lineStatusBar(2).x1 + fixDPI(13)
    If lblMessages.Left <> newLeft Then lblMessages.Left = newLeft
    
    'If the message label will overflow other elements of the status bar, shrink it as necessary
    Dim newMessageArea As Long
    newMessageArea = (UserControl.ScaleWidth - lblMessages.Left) - fixDPI(12)
    
    If newMessageArea < 0 Then
        lblMessages.Visible = False
    Else
        lblMessages.Width = newMessageArea
        lblMessages.Visible = True
    End If

End Sub

'Dynamically render some icons onto the status bar.
Public Sub drawStatusBarIcons(ByVal enabledState As Boolean)

    'Start by clearing the status bar
    picStatusBar.Picture = LoadPicture("")
    
    'If no images are loaded, do not draw status bar icons.
    If enabledState Then
        
        lineStatusBar(0).Visible = True
        
        'Start with the "image size" icon
        sbIconSize.alphaBlendToDC picStatusBar.hDC, , lineStatusBar(0).x1 + fixDPI(8), fixDPI(4)
        
        'After the "image size" icon comes the actual image size label.  Its position is determined by 16 px for the icon,
        ' plus a 5px buffer on either size (contingent on DPI)
        lblImgSize.Left = lineStatusBar(0).x1 + fixDPI(14) + 16
        
        'The image size label is autosized.  Move the next vertical line separator just past it.
        If (Not lineStatusBar(1).Visible) Then lineStatusBar(1).Visible = True
        lineStatusBar(1).x1 = lblImgSize.Left + lblImgSize.Width + fixDPI(10)
        lineStatusBar(1).x2 = lineStatusBar(1).x1
        
        'After the "image size" panel and separator comes mouse coordinates.  The basic steps from above are repeated.
        sbIconCoords.alphaBlendToDC picStatusBar.hDC, , lineStatusBar(1).x1 + fixDPI(8), fixDPI(4)
        lblCoordinates.Left = lineStatusBar(1).x1 + fixDPI(14) + 16
        
        If (Not lineStatusBar(2).Visible) Then lineStatusBar(2).Visible = True
        
        'Note that we don't actually move the last line status bar; that is handled by DisplayImageCoordinates itself
        
    'Images are not loaded.  Hide the lines and other items.
    Else
        lineStatusBar(0).Visible = False
        lineStatusBar(1).Visible = False
        lineStatusBar(2).Visible = False
    End If
    
    'Make our painting persistent
    picStatusBar.Picture = picStatusBar.Image
    picStatusBar.Refresh
    
End Sub
