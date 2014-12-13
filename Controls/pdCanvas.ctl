VERSION 5.00
Begin VB.UserControl pdCanvas 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
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
   KeyPreview      =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   513
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   886
   ToolboxBitmap   =   "pdCanvas.ctx":0000
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
      TabIndex        =   4
      Top             =   7095
      Visible         =   0   'False
      Width           =   13290
   End
   Begin VB.PictureBox picCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000D&
      Height          =   4935
      Left            =   360
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Top             =   600
      Width           =   4575
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7350
      Width           =   13290
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
         TabIndex        =   7
         Top             =   0
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         BackColor       =   -2147483626
      End
      Begin VB.ComboBox cmbSizeUnit 
         CausesValidation=   0   'False
         Height          =   315
         ItemData        =   "pdCanvas.ctx":0312
         Left            =   3630
         List            =   "pdCanvas.ctx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   15
         Width           =   600
      End
      Begin VB.ComboBox cmbZoom 
         CausesValidation=   0   'False
         Height          =   315
         ItemData        =   "pdCanvas.ctx":0316
         Left            =   840
         List            =   "pdCanvas.ctx":0318
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   15
         Width           =   1110
      End
      Begin PhotoDemon.pdButtonToolbox cmdZoomOut 
         Height          =   345
         Left            =   390
         TabIndex        =   8
         Top             =   0
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         BackColor       =   -2147483626
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox cmdZoomIn 
         Height          =   345
         Left            =   1980
         TabIndex        =   9
         Top             =   0
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         BackColor       =   -2147483626
         AutoToggle      =   -1  'True
      End
      Begin PhotoDemon.pdButtonToolbox cmdImgSize 
         Height          =   345
         Left            =   2790
         TabIndex        =   10
         Top             =   0
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         BackColor       =   -2147483626
         AutoToggle      =   -1  'True
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
      Begin VB.Line lineStatusBar 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   184
         X2              =   184
         Y1              =   1
         Y2              =   22
      End
      Begin VB.Line lineStatusBar 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   314
         X2              =   314
         Y1              =   1
         Y2              =   22
      End
      Begin VB.Line lineStatusBar 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   402
         X2              =   402
         Y1              =   1
         Y2              =   22
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
'Last updated: 12/December/14
'Last update: performance optimizations
'
'In years past, PhotoDemon would use a separate canvas (a full VB Form) for each loaded image.  In 2013, as part of the massive
' window manager rewrite, I rewrote the program to only have a single canvas active at any time.  The canvas was rebuilt as a user
' control, and instead of living on a separate form (which required a *ton* of code to keep in sync with the main PD window), it
' was integrated directly into the main window.
'
'Technically, the primary canvas is only the first entry in an array.  This was done deliberately in case I ever added support for
' multiple canvases being usable at once.  This has some neat possibilities - for example, having side-by-side canvases at
' different locations on an image - but there's a lot of messy UI considerations with something like this, especially if the two
' viewports can support different images simultaneously.  So I have postponed this work until some later date, with the caveat
' that implementing it will be a lot of work, and likely have unexpected interactions throughout the program.
'
'This canvas relies on pdInputMouse for all mouse interactions.  See the pdInputMouse class for details on why we do our own mouse
' management instead of using VB's intrinsic mouse functions.
'
'As much as possible, I've tried to keep paint tool operation within this canvas to a minimum.  Generally speaking, the only tool
' interactions the canvas should handle is reporting mouse events to external functions that actually handle paint tool processing
' and rendering.  To that end, try to adhere to the existing tool implementation format when adding new tool support.  (Selections
' are currently the exception to this rule, because they were implemented long before other tools and thus aren't as
' well-contained.  I hope to someday remedy this.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Enum PD_MOUSEEVENT
    pMouseDown = 0
    pMouseMove = 1
    pMouseUp = 2
End Enum

#If False Then
    Private Const pMouseDown = 0, pMouseMove = 1, pMouseUp = 2
#End If

Private Const SM_CXVSCROLL As Long = 2
Private Const SM_CYHSCROLL As Long = 3

'These are used to track use of the Ctrl, Alt, and Shift keys
Private ShiftDown As Boolean, CtrlDown As Boolean, AltDown As Boolean

'Track mouse button use on this canvas
Private lMouseDown As Boolean, rMouseDown As Boolean

'Track mouse movement on this canvas
Private hasMouseMoved As Long

'If the mouse is currently over the canvas, this will be set to TRUE.
Private m_IsMouseOverCanvas As Boolean

'Track initial mouse button locations
Private m_initMouseX As Double, m_initMouseY As Double

'In the future, it may be helpful to know if the user interacted with the canvas, or if the mouse simply passed over it
' en route to something else.
Private m_UserInteractedWithCanvas As Boolean

'On the canvas's MouseDown event, mark the relevant point of interest index for this layer (if any).
' If a point of interest has not been selected, this value will be reset to -1.
Private curPointOfInterest As Long

'PD's custom input class completely replaces all mouse interfacing for this control
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1
Private WithEvents cKeyEvents As pdInputKeyboard
Attribute cKeyEvents.VB_VarHelpID = -1

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

'When no images are loaded, we instruct the user to load an image.  This generic image icon is used as a placeholder.
Dim iconLoadAnImage As pdDIB

'Some tools support the ability to auto-activate a layer beneath the mouse.  If supported, during the MouseMove event,
' this value (m_LayerAutoActivateIndex) will be updated with the index of the layer that will be auto-activated if the
' user presses the mouse button.  This can be used to modify things like cursor behavior, to make sure the user receives
' accurate feedback on what a given action will affect.
Private m_LayerAutoActivateIndex As Long

'Selection tools need to know if a selection was active before mouse events start.  If it is, creation of an invalid new selection
' will add "Remove Selection" to the Undo/Redo chain; however, if no selection was active, the working selection will simply
' be erased.
Private m_SelectionActiveBeforeMouseEvents As Boolean

'Because this control is loaded so early in the program's load process, it is initialized before the translation engine.  To make sure
' that non-English speakers have functional tooltips, this sub is called after the translation engine has loaded.
Public Sub createTooltips()
    
    'Assign tooltips to any relevant controls
    cmdZoomFit.ToolTipText = g_Language.TranslateMessage("Fit the image on-screen")
    cmdZoomIn.ToolTipText = g_Language.TranslateMessage("Zoom in")
    cmdZoomOut.ToolTipText = g_Language.TranslateMessage("Zoom out")
    cmdImgSize.ToolTipText = g_Language.TranslateMessage("Resize image")
    
End Sub

'Use this function to forcibly prevent the canvas from redrawing itself.  REDRAWS WILL NOT HAPPEN AGAIN UNTIL YOU RESTORE ACCESS!
Public Sub setRedrawSuspension(ByVal newRedrawValue As Boolean)
    m_suspendRedraws = newRedrawValue
End Sub

Public Property Get BackColor() As Long
    BackColor = picCanvas.BackColor
End Property

Public Property Let BackColor(newBackColor As Long)
    picCanvas.BackColor = newBackColor
    picCanvas.Refresh
End Property

Public Sub clearCanvas()
    
    'If no images have been loaded, draw the "load image" placeholder.
    If (g_OpenImageCount = 0) And (Not g_ProgramShuttingDown) Then
        
        fixChromeLayout
    
    'Otherwise, simply clear the user control
    Else
    
        picCanvas.Picture = LoadPicture("")
        picCanvas.Refresh
    
    End If
    
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
    
    'If the scroll bar status wasn't actually changed, we can avoid a forced screen refresh
    Dim changesMade As Boolean
    changesMade = False
    
    Select Case barType
    
        Case PD_HORIZONTAL
            If newVisibility <> picScrollH.Visible Then
                picScrollH.Visible = newVisibility
                changesMade = True
            End If
        
        Case PD_VERTICAL
            If newVisibility <> picScrollV.Visible Then
                picScrollV.Visible = newVisibility
                changesMade = True
            End If
        
        Case PD_BOTH
            If (newVisibility <> picScrollH.Visible) Or (newVisibility <> picScrollV.Visible) Then
                picScrollH.Visible = newVisibility
                picScrollV.Visible = newVisibility
                changesMade = True
            End If
    
    End Select
    
    'When scroll bar visibility is changed, we must move the main canvas picture box to match
    If changesMade Then alignCanvasPictureBox
    
End Sub

'Get/Set scroll bar value
Public Function getScrollValue(ByVal barType As PD_ORIENTATION) As Long

    If barType = PD_HORIZONTAL Then
        getScrollValue = HScroll.Value
    Else
        getScrollValue = VScroll.Value
    End If

End Function

Public Sub setScrollValue(ByVal barType As PD_ORIENTATION, ByVal NewValue As Long)
    
    Select Case barType
    
        Case PD_HORIZONTAL
            HScroll.Value = NewValue
            
        Case PD_VERTICAL
            VScroll.Value = NewValue
        
        Case PD_BOTH
            HScroll.Value = NewValue
            VScroll.Value = NewValue
        
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

Public Sub displayImageSize(ByRef srcImage As pdImage, Optional ByVal clearSize As Boolean = False)
    
    'The size display is cleared whenever the user has no images loaded
    If clearSize Then
        lblImgSize.Caption = ""
        
        'Also, clear the back of the canvas
        fixChromeLayout
        
    'When size IS displayed, we must also refresh the status bar (now that it dynamically aligns its contents)
    Else
        
        If Not srcImage Is Nothing Then
        
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
            lblImgSize.updateAgainstCurrentTheme
            drawStatusBarIcons True
            
        End If
        
    End If
        
End Sub

Public Sub displayCanvasMessage(ByRef cMessage As String)
    lblMessages.Caption = cMessage
    lblMessages.requestRefresh
End Sub

'Display the current mouse coordinates
Public Sub displayCanvasCoordinates(ByVal xCoord As Long, ByVal yCoord As Long, Optional ByVal clearCoords As Boolean = False)
    
    If clearCoords Then
        lblCoordinates.Caption = ""
    
    'When coordinates are displayed, we must also refresh the status bar (now that it dynamically aligns its contents)
    Else
        lblCoordinates.Caption = "(" & xCoord & "," & yCoord & ")"
    End If
    
    'Normally, the custom label control will not repaint until Windows requests it, but because we require its true size in order
    ' to reflow the status bar, we must request an immediate update.
    'lblCoordinates.updateAgainstCurrentTheme
    
    'Align the right-hand line control with the newly captioned label
    lineStatusBar(2).x1 = lblCoordinates.Left + lblCoordinates.PixelWidth + fixDPI(10)
    lineStatusBar(2).x2 = lineStatusBar(2).x1
    
    'Make the message area shrink to match the new coordinate display size
    fixChromeLayout
        
End Sub

Public Sub requestBufferSync()
    picCanvas.Picture = picCanvas.Image
    picCanvas.Refresh
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
    hDC = picCanvas.hDC
End Property

Public Sub enableZoomIn(ByVal isEnabled As Boolean)
    cmdZoomIn.Enabled = isEnabled
End Sub

Public Sub enableZoomOut(ByVal isEnabled As Boolean)
    cmdZoomOut.Enabled = isEnabled
End Sub

Public Sub enableZoomFit(ByVal isEnabled As Boolean)
    cmdZoomFit.Enabled = isEnabled
    cmdZoomFit.Value = (cmbZoom.ListIndex = g_Zoom.getZoomFitAllIndex)
End Sub

Public Function getZoomDropDownReference() As ComboBox
    Set getZoomDropDownReference = cmbZoom
End Function

'Key presses are handled by PhotoDemon's custom pdInputKeyboard class
Private Sub cKeyEvents_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)

    markEventHandled = False

    'Make sure canvas interactions are allowed (e.g. an image has been loaded, etc)
    If isCanvasInteractionAllowed() Then
    
        Dim hOffset As Long, vOffset As Long
        Dim canvasUpdateRequired As Boolean

        'Any further processing depends on which tool is currently active
        Select Case g_CurrentTool
        
            'Drag-to-pan canvas
            Case NAV_DRAG
                
                canvasUpdateRequired = False
                
                'Suspend automatic redraws until all arrow keys have been processed
                m_suspendRedraws = True
                
                'If scrollbars are visible, nudge the canvas in the direction of the arrows.
                If VScroll.Enabled Then
                    If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Then canvasUpdateRequired = True
                    If (vkCode = VK_UP) Then VScroll.Value = VScroll.Value - 1
                    If (vkCode = VK_DOWN) Then VScroll.Value = VScroll.Value + 1
                End If
                
                If HScroll.Enabled Then
                    If (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then canvasUpdateRequired = True
                    If (vkCode = VK_LEFT) Then HScroll.Value = HScroll.Value - 1
                    If (vkCode = VK_RIGHT) Then HScroll.Value = HScroll.Value + 1
                End If
                
                'Re-enable automatic redraws
                m_suspendRedraws = False
                
                'Redraw the viewport if necessary
                If canvasUpdateRequired Then
                    markEventHandled = True
                    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
                End If
                    
            'Move stuff around
            Case NAV_MOVE
            
                'Handle arrow keys first
                If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Or (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then
            
                    'Calculate offset modifiers for the current layer
                    If (vkCode = VK_UP) Then vOffset = vOffset - 1
                    If (vkCode = VK_DOWN) Then vOffset = vOffset + 1
                    If (vkCode = VK_LEFT) Then hOffset = hOffset - 1
                    If (vkCode = VK_RIGHT) Then hOffset = hOffset + 1
                    
                    If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Or (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then canvasUpdateRequired = True
                    
                    'Apply the offsets
                    With pdImages(g_CurrentImage).getActiveLayer
                        .setLayerOffsetX .getLayerOffsetX + hOffset
                        .setLayerOffsetY .getLayerOffsetY + vOffset
                    End With
                    
                    'Redraw the viewport if necessary
                    If canvasUpdateRequired Then
                        markEventHandled = True
                        Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
                    End If
                    
                'Handle non-arrow keys next
                Else
                
                    'Delete key: delete the active layer (if allowed)
                    If (vkCode = VK_DELETE) And pdImages(g_CurrentImage).getNumOfLayers > 1 Then
                        markEventHandled = True
                        Process "Delete layer", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_IMAGE
                    End If
                    
                    'Insert: raise Add New Layer dialog
                    If (vkCode = VK_INSERT) Then
                        markEventHandled = True
                        Process "Add new layer", True
                    End If
                
                    'Tab and Shift+Tab: move through layer stack
                    If (vkCode = VK_TAB) Then
                        
                        markEventHandled = True
                        
                        'Retrieve the active layer index
                        Dim curLayerIndex As Long
                        curLayerIndex = pdImages(g_CurrentImage).getActiveLayerIndex
                        
                        'Advance the layer index according to the Shift modifier
                        If (Shift And vbShiftMask) <> 0 Then
                            curLayerIndex = curLayerIndex + 1
                        Else
                            curLayerIndex = curLayerIndex - 1
                        End If
                        
                        If curLayerIndex < 0 Then curLayerIndex = pdImages(g_CurrentImage).getNumOfLayers - 1
                        If curLayerIndex > pdImages(g_CurrentImage).getNumOfLayers - 1 Then curLayerIndex = 0
                        
                        'Activate the new layer
                        pdImages(g_CurrentImage).setActiveLayerByIndex curLayerIndex
                        
                        'Redraw the viewport and interface to match
                        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), Me
                        syncInterfaceToCurrentImage
                        
                    End If
                
                    'Space bar: toggle active layer visibility
                    If (vkCode = VK_SPACE) Then
                        markEventHandled = True
                        pdImages(g_CurrentImage).getActiveLayer.setLayerVisibility (Not pdImages(g_CurrentImage).getActiveLayer.getLayerVisibility)
                        Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
                        syncInterfaceToCurrentImage
                    End If
                
                End If
            
            'Selections
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
            
                'Handle arrow keys first
                If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Or (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then
            
                    'If a selection is active, nudge it using the arrow keys
                    If pdImages(g_CurrentImage).selectionActive And (pdImages(g_CurrentImage).mainSelection.getSelectionShape <> sRaster) Then
                        
                        markEventHandled = True
                        
                        'Disable automatic refresh requests
                        pdImages(g_CurrentImage).mainSelection.rejectRefreshRequests = True
                        
                        'Calculate offsets
                        If (vkCode = VK_UP) Then vOffset = vOffset - 1
                        If (vkCode = VK_DOWN) Then vOffset = vOffset + 1
                        If (vkCode = VK_LEFT) Then hOffset = hOffset - 1
                        If (vkCode = VK_RIGHT) Then hOffset = hOffset + 1
                        
                        'If offsets were generated, update the selection and redraw the screen
                        If (hOffset <> 0) Or (vOffset <> 0) Then
                        
                            'Some selection types can be modified by simply updating the selection text boxes.  Others cannot.
                            
                            'Non-textbox-compatible selections are handled here
                            If (g_CurrentTool = SELECT_POLYGON) Or (g_CurrentTool = SELECT_LASSO) Or (g_CurrentTool = SELECT_WAND) Then
                                pdImages(g_CurrentImage).mainSelection.rejectRefreshRequests = False
                                pdImages(g_CurrentImage).mainSelection.nudgeSelection hOffset, vOffset
                            
                            'Textbox-compatible selections are handled here
                            Else
                            
                                'Update the selection coordinate text boxes with the new offsets
                                toolbar_Options.tudSel(0).Value = toolbar_Options.tudSel(0).Value + hOffset
                                toolbar_Options.tudSel(1).Value = toolbar_Options.tudSel(1).Value + vOffset
                                
                                If g_CurrentTool = SELECT_LINE Then
                                    toolbar_Options.tudSel(2).Value = toolbar_Options.tudSel(2).Value + hOffset
                                    toolbar_Options.tudSel(3).Value = toolbar_Options.tudSel(3).Value + vOffset
                                End If
                                
                                'Update the screen
                                pdImages(g_CurrentImage).mainSelection.rejectRefreshRequests = False
                                pdImages(g_CurrentImage).mainSelection.updateViaTextBox
                            
                            End If
                            
                            Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                            
                        End If
                    
                    End If
                
                'Handle non-arrow keys next.  (Note: most non-arrow keys are not meant to work with key-repeating, so they
                ' are handled in the KeyUp event instead.)
                Else
                
                    
                                    
                End If
            
        End Select
        
    End If

End Sub

'Key presses are handled by PhotoDemon's custom pdInputKeyboard class
Private Sub cKeyEvents_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    
    markEventHandled = False

    'Make sure canvas interactions are allowed (e.g. an image has been loaded, etc)
    If isCanvasInteractionAllowed() Then
        
        'Any further processing depends on which tool is currently active
        Select Case g_CurrentTool
        
            Case NAV_DRAG
            
            Case NAV_MOVE
            
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
                
                'Delete key: if a selection is active, erase the selected area
                If (vkCode = VK_DELETE) And pdImages(g_CurrentImage).selectionActive Then
                    markEventHandled = True
                    Process "Erase selected area", False, buildParams(pdImages(g_CurrentImage).getActiveLayerIndex), UNDO_LAYER
                End If
                
                'Escape key: if a selection is active, clear it
                If (vkCode = VK_ESCAPE) And pdImages(g_CurrentImage).selectionActive Then
                    markEventHandled = True
                    Process "Remove selection", , , UNDO_SELECTION
                End If
                
                'Backspace key: for lasso and polygon selections, retreat back one or more coordinates, giving the user a chance to
                ' correct any potential mistakes.
                If ((g_CurrentTool = SELECT_LASSO) Or (g_CurrentTool = SELECT_POLYGON)) And (vkCode = VK_BACK) And pdImages(g_CurrentImage).selectionActive And (Not pdImages(g_CurrentImage).mainSelection.isLockedIn) Then
                    
                    markEventHandled = True
                    
                    'Polygons
                    If (g_CurrentTool = SELECT_POLYGON) Then
                    
                        'Do not allow point removal if the polygon has already been successfully closed.
                        If Not pdImages(g_CurrentImage).mainSelection.getPolygonClosedState Then pdImages(g_CurrentImage).mainSelection.removeLastPolygonPoint
                    
                    'Lassos
                    Else
                    
                        'Do not allow point removal if the lasso has already been successfully closed.
                        If Not pdImages(g_CurrentImage).mainSelection.getLassoClosedState Then
                    
                            'Ask the selection object to retreat its position
                            Dim newImageX As Double, newImageY As Double
                            pdImages(g_CurrentImage).mainSelection.retreatLassoPosition newImageX, newImageY
                            
                            'The returned coordinates will be in image coordinates.  Convert them to viewport coordinates.
                            Dim newCanvasX As Double, newCanvasY As Double
                            Drawing.convertImageCoordsToCanvasCoords FormMain.mainCanvas(0), pdImages(g_CurrentImage), newImageX, newImageY, newCanvasX, newCanvasY
                            
                            'Finally, convert the canvas coordinates to screen coordinates, and move the cursor accordingly
                            setCursorToCanvasPosition newCanvasX, newCanvasY
                            
                        End If
                        
                    End If
                    
                    'Redraw the screen to reflect this new change.
                    Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                
                End If
            
        End Select
        
    End If
    
End Sub

'Set the mouse cursor to a specified position on the canvas.  This function will automatically handle the translation to screen coordinates.
Private Sub setCursorToCanvasPosition(ByVal canvasX As Double, ByVal canvasY As Double)
    cMouseEvents.moveCursorToNewPosition canvasX, canvasY
End Sub

Private Sub cmbSizeUnit_Click()
    If g_OpenImageCount > 0 Then displayImageSize pdImages(g_CurrentImage)
End Sub

Private Sub CmbZoom_Click()

    'Only process zoom changes if an image has been loaded
    If isCanvasInteractionAllowed() Then
    
        'Store the current zoom value in this object (so the user can switch between images without losing zoom values)
        pdImages(g_CurrentImage).currentZoomValue = cmbZoom.ListIndex
        
        'Disable the zoom in/out buttons when they reach the end of the available zoom levels
        If cmbZoom.ListIndex = 0 Then
            cmdZoomIn.Enabled = False
        Else
            If Not cmdZoomIn.Enabled Then cmdZoomIn.Enabled = True
        End If
        
        If cmbZoom.ListIndex = g_Zoom.getZoomCount Then
            cmdZoomOut.Enabled = False
        Else
            If Not cmdZoomOut.Enabled Then cmdZoomOut.Enabled = True
        End If
        
        'Highlight the "zoom fit" button while fit mode is active
        cmdZoomFit.Value = (cmbZoom.ListIndex = g_Zoom.getZoomFitAllIndex)
        
        'Redraw the viewport (if allowed; some functions will prevent us from doing this, as they plan to request their own
        ' refresh after additional processing occurs)
        If g_AllowViewportRendering Then Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), "zoom changed by primary drop-down box"
        
    End If

End Sub

Private Sub cmdImgSize_Click()
    If isCanvasInteractionAllowed() Then Process "Resize image", True
End Sub

Private Sub cmdZoomFit_Click()
    Image_Canvas_Handler.FitOnScreen
End Sub

Private Sub cmdZoomIn_Click()
    FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = g_Zoom.getNearestZoomInIndex(FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex)
End Sub

Private Sub cmdZoomOut_Click()
    FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = g_Zoom.getNearestZoomOutIndex(FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex)
End Sub

'At present, the only App Commands the canvas will handle are forward/back, which link to Undo/Redo
Private Sub cMouseEvents_AppCommand(ByVal cmdID As AppCommandConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If isCanvasInteractionAllowed() Then
    
        Select Case cmdID
        
            'Back button
            Case AC_BROWSER_BACKWARD, AC_UNDO
            
                If pdImages(g_CurrentImage).IsActive Then
                    If pdImages(g_CurrentImage).undoManager.getUndoState Then Process "Undo", , , UNDO_NOTHING
                End If
            
            'Forward button
            Case AC_BROWSER_FORWARD, AC_REDO
            
                If pdImages(g_CurrentImage).IsActive Then
                    If pdImages(g_CurrentImage).undoManager.getRedoState Then Process "Redo", , , UNDO_NOTHING
                End If
        
        End Select

    End If

End Sub

Private Sub cMouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
        
    'Make sure interactions with this canvas are allowed
    If Not isCanvasInteractionAllowed() Then Exit Sub
    
    'Because VB does not allow an user control to receive focus if it contains controls that can receive focus, the arrow buttons
    ' can behave unpredictably (for example, if the zoom box has focus, and the user clicks on the canvas, the canvas will not
    ' receive focus and arrow key presses will continue to interact with the zoom box instead of the viewport)
    'cmdZoomFit.SetFocus
    
    'Note that the user has attempted to interact with the canvas.
    m_UserInteractedWithCanvas = True
    
    'Note whether a selection is active when mouse interactions began
    m_SelectionActiveBeforeMouseEvents = (pdImages(g_CurrentImage).selectionActive And pdImages(g_CurrentImage).mainSelection.isLockedIn)
    
    'These variables will hold the corresponding (x,y) coordinates on the IMAGE - not the VIEWPORT.
    ' (This is important if the user has zoomed into an image, and used scrollbars to look at a different part of it.)
    Dim imgX As Double, imgY As Double
    
    'Display the image coordinates under the mouse pointer
    displayImageCoordinates x, y, pdImages(g_CurrentImage), Me, imgX, imgY
    
    'Display a relevant cursor for the current action
    setCanvasCursor pMouseUp, Button, x, y, imgX, imgY
    
    'Selection tools all use the same variable for tracking POIs
    Dim sCheck As Long
    
    'Check mouse button use
    If Button = vbLeftButton Then
        
        lMouseDown = True
        hasMouseMoved = 0
            
        'Remember this location
        m_initMouseX = x
        m_initMouseY = y
        
        'Ask the current layer if these coordinates correspond to a point of interest.  We don't always use this return value,
        ' but a number of functions could potentially ask for it, so we cache it at MouseDown time and hang onto it until
        ' the mouse is released.
        curPointOfInterest = pdImages(g_CurrentImage).getActiveLayer.checkForPointOfInterest(imgX, imgY)
        
        'Any further processing depends on which tool is currently active
        Select Case g_CurrentTool
        
            'Drag-to-pan canvas
            Case NAV_DRAG
                setInitialCanvasScrollValues FormMain.mainCanvas(0)
                
            'Move stuff around
            Case NAV_MOVE
            
                'Prior to moving or transforming a layer, we need to check the state of the "auto-activate layer beneath mouse"
                ' option; if it is set, check (and possibly modify) the active layer based on the mouse position.
                If CBool(toolbar_Options.chkAutoActivateLayer) Then
                
                    Dim layerUnderMouse As Long
                    layerUnderMouse = Layer_Handler.getLayerUnderMouse(imgX, imgY, True)
                    
                    'The "getLayerUnderMouse" function will return a layer index if the mouse is over a layer.  If the mouse is not
                    ' over a layer, it will return -1.
                    If layerUnderMouse > -1 Then
                    
                        'If the layer under the mouse is not already active, activate it now
                        If layerUnderMouse <> pdImages(g_CurrentImage).getActiveLayerIndex Then
                            Layer_Handler.setActiveLayerByIndex layerUnderMouse, False
                            Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), Me
                        End If
                    
                    End If
                
                End If
                
                'Initiate the layer transformation engine.  Note that nothing will happen until the user actually moves the mouse.
                setInitialLayerOffsets pdImages(g_CurrentImage).getActiveLayer, pdImages(g_CurrentImage).getActiveLayer.checkForPointOfInterest(imgX, imgY)
        
            'Standard selections
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO
            
                'Check to see if a selection is already active.  If it is, see if the user is allowed to transform it.
                If pdImages(g_CurrentImage).selectionActive Then
                
                    'Check the mouse coordinates of this click.
                    sCheck = findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
                    
                    'If a point of interest was clicked, initiate a transform
                    If (sCheck <> -1) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape <> sPolygon) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape <> sRaster) Then
                        
                        'Initialize a selection transformation
                        pdImages(g_CurrentImage).mainSelection.setTransformationType sCheck
                        pdImages(g_CurrentImage).mainSelection.setInitialTransformCoordinates imgX, imgY
                                        
                    'If a point of interest was *not* clicked, erase any existing selection and start a new one
                    Else
                        
                        'Polygon selections require special handling, because they don't operate on the "mouse up = complete" assumption.
                        ' They are completed when the user re-clicks the first point.  Any clicks prior to that point are treated as
                        ' an instruction to add a new points.
                        If g_CurrentTool = SELECT_POLYGON Then
                            
                            'First, see if the selection is locked in.  If it is, treat this is a regular transformation.
                            If pdImages(g_CurrentImage).mainSelection.isLockedIn Then
                                pdImages(g_CurrentImage).mainSelection.setTransformationType sCheck
                                pdImages(g_CurrentImage).mainSelection.setInitialTransformCoordinates imgX, imgY
                            
                            'Selection is not locked in, meaning the user is still constructing it.
                            Else
                            
                                'If the user clicked on the initial polygon point, attempt to close the polygon
                                If (sCheck = 0) And (pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints > 2) Then
                                    pdImages(g_CurrentImage).mainSelection.setPolygonClosedState True
                                    pdImages(g_CurrentImage).mainSelection.setTransformationType 0
                                
                                'The user did not click the initial polygon point, meaning we should add this coordinate as a new polygon point.
                                Else
                                    
                                    'Remove transformation mode (if any)
                                    pdImages(g_CurrentImage).mainSelection.setTransformationType -1
                                    pdImages(g_CurrentImage).mainSelection.overrideTransformMode False
                                    
                                    'Add the new point
                                    If pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints = 0 Then
                                        Selection_Handler.initSelectionByPoint imgX, imgY
                                    Else
                                        
                                        If (sCheck = -1) Or (sCheck = pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints) Then
                                            pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                                            pdImages(g_CurrentImage).mainSelection.setTransformationType pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints - 1
                                        Else
                                            pdImages(g_CurrentImage).mainSelection.setTransformationType sCheck
                                        End If
                                        
                                    End If
                                    
                                    'Reinstate transformation mode, using the index of the new point as the transform ID
                                    pdImages(g_CurrentImage).mainSelection.setInitialTransformCoordinates imgX, imgY
                                    pdImages(g_CurrentImage).mainSelection.overrideTransformMode True
                                    
                                    'Redraw the screen
                                    Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), Me
                                    
                                End If
                            
                            End If
                            
                        Else
                            Selection_Handler.initSelectionByPoint imgX, imgY
                        End If
                        
                    End If
                
                'If a selection is not active, start a new one
                Else
                    
                    Selection_Handler.initSelectionByPoint imgX, imgY
                    
                    'Polygon selections require special handling, as usual.  After creating the initial point, we want to immediately initiate
                    ' transform mode, because dragging the mouse will simply move the newly created point.
                    If g_CurrentTool = SELECT_POLYGON Then
                        pdImages(g_CurrentImage).mainSelection.setTransformationType pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints - 1
                        pdImages(g_CurrentImage).mainSelection.overrideTransformMode True
                    End If
                    
                End If
            
            'Magic wand selections are easy.  They never transform - they only generate anew
            Case SELECT_WAND
                Selection_Handler.initSelectionByPoint imgX, imgY
                Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), Me
                
            'In the future, other tools can be handled here
            Case Else
            
            
        End Select
    
    ElseIf Button = vbRightButton Then
    
        rMouseDown = True
        
        'TODO: right-button functionality
    
    End If
    
End Sub

'When the mouse enters the canvas, any floating toolbars must be automatically dimmed.
Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
        
    m_IsMouseOverCanvas = True
        
    'Note that the user has yet to interact with anything on the canvas.
    m_UserInteractedWithCanvas = False
    
    'Notify the window manager that toolbars need to be made translucent
    g_WindowManager.notifyMouseMoveOverCanvas
    
    'If no images have been loaded, reset the cursor
    If g_OpenImageCount = 0 Then cMouseEvents.setSystemCursor IDC_ARROW

End Sub

'When the mouse leaves the window, if no buttons are down, clear the coordinate display.
' (We must check for button states because the user is allowed to do things like drag selection nodes outside the image.)
Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    m_IsMouseOverCanvas = False
        
    If (Not lMouseDown) And (Not rMouseDown) Then ClearImageCoordinatesDisplay

End Sub

Private Sub cMouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'Make sure interactions with this canvas are allowed
    If Not isCanvasInteractionAllowed() Then Exit Sub
    
    hasMouseMoved = hasMouseMoved + 1
    
    'These variables will hold the corresponding (x,y) coordinates on the image - NOT the viewport
    Dim imgX As Double, imgY As Double
    
    'Display the image coordinates under the mouse pointer
    displayImageCoordinates x, y, pdImages(g_CurrentImage), Me, imgX, imgY
    
    'Check the left mouse button
    If lMouseDown Then
    
        Select Case g_CurrentTool
        
            'Drag-to-pan canvas
            Case NAV_DRAG
                panImageCanvas m_initMouseX, m_initMouseY, x, y, pdImages(g_CurrentImage), FormMain.mainCanvas(0)
            
            'Move stuff around
            Case NAV_MOVE
                Message "Shift key: preserve layer aspect ratio"
                transformCurrentLayer m_initMouseX, m_initMouseY, x, y, pdImages(g_CurrentImage), FormMain.mainCanvas(0), (Shift And vbShiftMask)
        
            'Basic selection tools
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON
    
                'First, check to see if a selection is both active and transformable.
                If pdImages(g_CurrentImage).selectionActive And (pdImages(g_CurrentImage).mainSelection.getSelectionShape <> sRaster) Then
                    
                    'If the SHIFT key is down, notify the selection engine that a square shape is requested
                    pdImages(g_CurrentImage).mainSelection.requestSquare (Shift And vbShiftMask)
                    
                    'Pass new points to the active selection
                    pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                    syncTextToCurrentSelection g_CurrentImage
                                        
                End If
                
                'Force a redraw of the viewport
                If hasMouseMoved > 1 Then Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), Me
            
            'Lasso selections are handled specially, because mouse move events control the drawing of the lasso
            Case SELECT_LASSO
            
                'First, check to see if a selection is active
                If pdImages(g_CurrentImage).selectionActive Then
                    
                    'Pass new points to the active selection
                    pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                                        
                End If
                
                'To spare the debug logger from receiving too many events, forcibly prevent logging of this message
                ' while in debug mode.
                #If DEBUGMODE = 1 Then
                    Message "Release the mouse button to complete the lasso selection", "DONOTLOG"
                #Else
                    Message "Release the mouse button to complete the lasso selection"
                #End If
                
                'Force a redraw of the viewport
                If hasMouseMoved > 1 Then Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), Me
            
            'Wand selections are easier than other selection types, because they don't support any special transforms
            Case SELECT_WAND
                If pdImages(g_CurrentImage).selectionActive Then
                    pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                    Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), Me
                End If
            
        End Select
    
    'This else means the LEFT mouse button is NOT down
    Else
    
        'Display a relevant cursor for the current action
        setCanvasCursor pMouseUp, Button, x, y, imgX, imgY
    
        Select Case g_CurrentTool
        
            'Drag-to-navigate
            Case NAV_DRAG
                
            'Move stuff around
            Case NAV_MOVE
            
                'If the user has the "auto-activate layer beneath mouse" option set, report the current layer name in the
                ' message bar; this is helpful for determining what layer will be affected by a given action.
                If CBool(toolbar_Options.chkAutoActivateLayer) Then
                
                    Dim layerUnderMouse As Long
                    layerUnderMouse = Layer_Handler.getLayerUnderMouse(imgX, imgY, True)
                    
                    'The "getLayerUnderMouse" function will return a layer index if the mouse is over a layer.  If the mouse is not
                    ' over a layer, it will return -1.
                    If layerUnderMouse > -1 Then
                        m_LayerAutoActivateIndex = layerUnderMouse
                        
                        'To spare the debug logger from receiving too many events, forcibly prevent logging of this message
                        ' while in debug mode.
                        #If DEBUGMODE = 1 Then
                            Message "Target layer: %1", pdImages(g_CurrentImage).getLayerByIndex(layerUnderMouse).getLayerName, "DONOTLOG"
                        #Else
                            Message "Target layer: %1", pdImages(g_CurrentImage).getLayerByIndex(layerUnderMouse).getLayerName
                        #End If
                    
                    'The mouse is not over a layer.  Default to the active layer, which allows the user to interact with the
                    ' layer even if it lies off-canvas.
                    Else
                        m_LayerAutoActivateIndex = pdImages(g_CurrentImage).getActiveLayerIndex
                    End If
                
                'Auto-activation is disabled.  Don't bother reporting the layer beneath the mouse to the user, as actions can
                ' only affect the active layer!
                Else
                    Message ""
                    m_LayerAutoActivateIndex = pdImages(g_CurrentImage).getActiveLayerIndex
                End If
                
            'Selection tools
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_POLYGON, SELECT_LASSO, SELECT_WAND
            
            Case Else
            
        End Select
        
    End If
    
End Sub

Private Sub cMouseEvents_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal ClickEventAlsoFiring As Boolean)
    
    'Make sure interactions with this canvas are allowed
    If Not isCanvasInteractionAllowed() Then Exit Sub
    
    'Display the image coordinates under the mouse pointer
    Dim imgX As Double, imgY As Double
    displayImageCoordinates x, y, pdImages(g_CurrentImage), Me, imgX, imgY
    
    'Display a relevant cursor for the current action
    setCanvasCursor pMouseUp, Button, x, y, imgX, imgY
    
    'Check mouse buttons
    If Button = vbLeftButton Then
    
        lMouseDown = False
    
        Select Case g_CurrentTool
        
            'Click-to-drag navigation
            Case NAV_DRAG
                
            'Move stuff around
            Case NAV_MOVE
            
                'Pass a final transform request to the layer handler.  This will initiate Undo/Redo creation,
                ' among other things.
                If (hasMouseMoved > 0) Then transformCurrentLayer m_initMouseX, m_initMouseY, x, y, pdImages(g_CurrentImage), FormMain.mainCanvas(0), (Shift And vbShiftMask), True
                
                'Reset the generic tool mouse tracking function
                Tool_Support.terminateGenericToolTracking
                
            'Most selection tools
            Case SELECT_RECT, SELECT_CIRC, SELECT_LINE, SELECT_LASSO
            
                'If a selection was being drawn, lock it into place
                If pdImages(g_CurrentImage).selectionActive Then
                    
                    'Check to see if this mouse location is the same as the initial mouse press. If it is, and that particular
                    ' point falls outside the selection, clear the selection from the image.
                    If ((ClickEventAlsoFiring) And (findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage)) = -1)) Or ((pdImages(g_CurrentImage).mainSelection.selWidth <= 0) And (pdImages(g_CurrentImage).mainSelection.selHeight <= 0)) Then
                        
                        If (g_CurrentTool <> SELECT_WAND) Then
                            Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                        Else
                            Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                        End If
                    
                    'The mouse is being released after a significant move event, or on a point of interest to the current selection.
                    Else
                    
                        'If the selection is not raster-type, pass these final mouse coordinates to it
                        If (pdImages(g_CurrentImage).mainSelection.getSelectionShape <> sRaster) Then

                            pdImages(g_CurrentImage).mainSelection.requestSquare (Shift And vbShiftMask)
                            pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                            syncTextToCurrentSelection g_CurrentImage
                            
                        End If
                    
                        'Check to see if all selection coordinates are invalid (e.g. off-image).  If they are, forget about this selection.
                        If pdImages(g_CurrentImage).mainSelection.areAllCoordinatesInvalid Then
                            Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                        Else
                            
                            'Depending on the type of transformation that may or may not have been applied, call the appropriate processor function.
                            ' This is required to add the current selection event to the Undo/Redo chain.
                            Select Case g_CurrentTool
                            
                                Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
                                
                                    Select Case pdImages(g_CurrentImage).mainSelection.getTransformationType
                            
                                        'Creating a new selection
                                        Case -1
                                            Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                            
                                        'Moving an existing selection
                                        Case 8
                                            Process "Move selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                            
                                        'Anything else is assumed to be resizing an existing selection
                                        Case Else
                                            Process "Resize selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                            
                                    End Select
                                    
                                Case SELECT_LASSO
                                
                                    Select Case pdImages(g_CurrentImage).mainSelection.getTransformationType
                            
                                        'Creating a new selection
                                        Case -1
                                            Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                            
                                        'Moving an existing selection
                                        Case Else
                                            Process "Move selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                            
                                    End Select
                                    
                                Case SELECT_WAND
                                    Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                            
                            End Select
                            
                        End If
                        
                    End If
                    
                    'Force a redraw of the screen
                    Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), Me
                    
                Else
                    'If the selection is not active, make sure it stays that way
                    pdImages(g_CurrentImage).mainSelection.lockRelease
                End If
                
                'Synchronize the selection text box values with the final selection
                syncTextToCurrentSelection g_CurrentImage
                
            
            'As usual, polygon selections have some special considerations.
            Case SELECT_POLYGON
            
                'If a selection was being drawn, lock it into place
                If pdImages(g_CurrentImage).selectionActive Then
                
                    'Check to see if the selection is locked in.  If it is, we need to check for an "erase selection" click.
                    If pdImages(g_CurrentImage).mainSelection.getPolygonClosedState And ClickEventAlsoFiring And (findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage)) = -1) Then
                        Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                    
                    Else
                        
                        'If the polygon is already closed, we want to lock in the newly modified polygon
                        If pdImages(g_CurrentImage).mainSelection.getPolygonClosedState Then
                        
                            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
                            
                                Case pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints
                                    Process "Move selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                    
                                Case 0
                                    If ClickEventAlsoFiring Then
                                        Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                    Else
                                        Process "Resize selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                    End If
                                    
                                Case -1
                                
                                    'If the user has clicked off the selection, we want to remove it.
                                    If ClickEventAlsoFiring Then
                                        Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                                    
                                    'If they haven't clicked, this could simply indicate that they dragged a polygon point off the polygon
                                    ' and into some new region of the image.
                                    Else
                                        
                                        pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                                        Process "Resize selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                        
                                    End If
                                
                                Case Else
                                    Process "Resize selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                                    
                            End Select
                            
                            'Check to see if all selection coordinates are invalid (e.g. off-image).  If they are, forget about this selection.
                            If pdImages(g_CurrentImage).mainSelection.isLockedIn And pdImages(g_CurrentImage).mainSelection.areAllCoordinatesInvalid Then
                                Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                            End If
                            
                        Else
                        
                            'Pass these final mouse coordinates to the selection engine
                            pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                            
                            'To spare the debug logger from receiving too many events, forcibly prevent logging of this message
                            ' while in debug mode.
                            #If DEBUGMODE = 1 Then
                                Message "Click on the first point to complete the polygon selection", "DONOTLOG"
                            #Else
                                Message "Click on the first point to complete the polygon selection"
                            #End If
                            
                        End If
                    
                    End If
                    
                    'Force a redraw of the screen
                    Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), Me
                
                Else
                    'If the selection is not active, make sure it stays that way
                    pdImages(g_CurrentImage).mainSelection.lockRelease
                End If
                
            'Magic wand selections are much easier than other selection types
            Case SELECT_WAND
                
                'If a selection was being drawn, lock it into place
                If pdImages(g_CurrentImage).selectionActive Then
                    
                    'Supply the final coords to the selection engine
                    pdImages(g_CurrentImage).mainSelection.setAdditionalCoordinates imgX, imgY
                    
                    'Check to see if all selection coordinates are invalid (e.g. off-image).  If they are, forget about this selection.
                    If pdImages(g_CurrentImage).mainSelection.areAllCoordinatesInvalid Then
                        Process "Remove selection", , , IIf(m_SelectionActiveBeforeMouseEvents, UNDO_SELECTION, UNDO_NOTHING), g_CurrentTool
                    
                    'If the selection coordinates are valid, create it now.
                    Else
                        Process "Create selection", , pdImages(g_CurrentImage).mainSelection.getSelectionParamString, UNDO_SELECTION, g_CurrentTool
                    End If
                    
                    'Force a redraw of the screen
                    Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), Me
                    
                Else
                    'If the selection is not active, make sure it stays that way
                    pdImages(g_CurrentImage).mainSelection.lockRelease
                End If
                
            Case Else
                    
        End Select
                        
    End If
    
    If Button = vbRightButton Then rMouseDown = False
    
    'Reset any tracked point of interest value for this layer
    curPointOfInterest = -1
        
    'Reset the mouse movement tracker
    hasMouseMoved = 0
    

End Sub

Public Sub cMouseEvents_MouseWheelHorizontal(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    
    'Make sure interactions with this canvas are allowed
    If Not isCanvasInteractionAllowed() Then Exit Sub
    
    'Horizontal scrolling - only trigger if the horizontal scroll bar is visible AND a shift key has been pressed BUT a ctrl
    ' button has not been pressed.
    If picScrollH.Visible Then
        
        If scrollAmount > 0 Then
        
            m_suspendRedraws = True
            
            If HScroll.Value + HScroll.LargeChange > HScroll.Max Then
                HScroll.Value = HScroll.Max
            Else
                HScroll.Value = HScroll.Value + HScroll.LargeChange
            End If
            
            m_suspendRedraws = False
            
            Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
        
        ElseIf scrollAmount < 0 Then
        
            m_suspendRedraws = True
            
            If HScroll.Value - HScroll.LargeChange < HScroll.Min Then
                HScroll.Value = HScroll.Min
            Else
                HScroll.Value = HScroll.Value - HScroll.LargeChange
            End If
            
            m_suspendRedraws = False
            
            Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
            
        End If
        
    End If

End Sub

'Vertical mousewheel scrolling.  Note that Shift+Wheel and Ctrl+Wheel modifiers do NOT raise this event; pdInputMouse automatically
' reroutes them to MouseWheelHorizontal and MouseWheelZoom, respectively.
Public Sub cMouseEvents_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    
    'Make sure interactions with this canvas are allowed
    If Not isCanvasInteractionAllowed() Then Exit Sub
    
    'PhotoDemon uses the standard photo editor convention of Ctrl+Wheel = zoom, Shift+Wheel = h_scroll, and Wheel = v_scroll.
    ' Some users (for reasons I don't understand??) expect plain mousewheel to zoom the image.  For these users, we now
    ' display a helpful message telling them to use the damn Ctrl modifier like everyone else.
    If picScrollV.Visible Then
  
        If scrollAmount < 0 Then
            
            m_suspendRedraws = True
            
            If VScroll.Value + VScroll.LargeChange > VScroll.Max Then
                VScroll.Value = VScroll.Max
            Else
                VScroll.Value = VScroll.Value + VScroll.LargeChange
            End If
            
            m_suspendRedraws = False
            
            Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
        
        ElseIf scrollAmount > 0 Then
            
            m_suspendRedraws = True
            
            If VScroll.Value - VScroll.LargeChange < VScroll.Min Then
                VScroll.Value = VScroll.Min
            Else
                VScroll.Value = VScroll.Value - VScroll.LargeChange
            End If
            
            m_suspendRedraws = False
            
            Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
            
        End If

    'The user is using the mousewheel without Ctrl/Shift modifiers, even without a visible scrollbar.
    ' Display a message about how mousewheels are supposed to work.
    Else
        Message "Mouse Wheel = VERTICAL SCROLL,  Shift + Wheel = HORIZONTAL SCROLL,  Ctrl + Wheel = ZOOM"
    End If
    
    'NOTE: horizontal scrolling via Shift+Vertical Wheel is handled in the separate _MouseWheelHorizontal event.
    'NOTE: zooming via Ctrl+Vertical Wheel is handled in the separate _MouseWheelZoom event.
    
End Sub

'The pdInputMouse class now provides a dedicated zoom event for us - how nice!
Public Sub cMouseEvents_MouseWheelZoom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal zoomAmount As Double)

    'Make sure interactions with this canvas are allowed
    If Not isCanvasInteractionAllowed() Then Exit Sub
    
    'Before doing anything else, cache the current mouse coordinates (in both Canvas and Image coordinate spaces)
    Dim imgX As Double, imgY As Double
    convertCanvasCoordsToImageCoords Me, pdImages(g_CurrentImage), x, y, imgX, imgY, True
    
    'Suspend automatic viewport redraws until we are done with our calculations
    g_AllowViewportRendering = False
    
    'Calculate a new zoom value
    If zoomAmount > 0 Then
        
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled And FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex > 0 Then
            FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = g_Zoom.getNearestZoomInIndex(FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex)
        End If
           
    ElseIf zoomAmount < 0 Then
        
        If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled And FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex <> g_Zoom.getZoomCount Then
            FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = g_Zoom.getNearestZoomOutIndex(FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex)
        End If
           
    End If
    
    'Re-enable automatic viewport redraws
    g_AllowViewportRendering = True
    
    'Request a manual redraw from Viewport_Engine.Stage1_InitializeBuffer, while supplying our x/y coordinates so that it can preserve mouse position
    ' relative to the underlying image.
    Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), "mousewheel zoom", x, y, imgX, imgY

End Sub

'(This code is copied from FormMain's OLEDragDrop event - please mirror any changes there)
Private Sub picCanvas_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub
    
    'Use the external function (in the clipboard handler, as the code is roughly identical to clipboard pasting)
    ' to load the OLE source.
    Clipboard_Handler.loadImageFromDragDrop Data, Effect, True
    
End Sub

'(This code is copied from FormMain's OLEDragOver event - please mirror any changes there)
Private Sub picCanvas_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub

    'Check to make sure the type of OLE object is files
    If Data.GetFormat(vbCFFiles) Or Data.GetFormat(vbCFText) Or Data.GetFormat(vbCFBitmap) Then
        'Inform the source that the files will be treated as "copied"
        Effect = vbDropEffectCopy And Effect
    Else
        'If it's not files or text, don't allow a drop
        Effect = vbDropEffectNone
    End If
    
End Sub

Private Sub UserControl_Initialize()

    If g_IsProgramRunning Then
        
        'Enable mouse subclassing for events like mousewheel, forward/back keys, enter/leave
        Set cMouseEvents = New pdInputMouse
        cMouseEvents.addInputTracker picCanvas.hWnd, True, True, True, True
        
        'This user control contains a lot of child controls whose key events we want to intercept (as they aren't designed to have
        ' focus on their own).  Submit these controls to the tracker, so it knows to mass any key events into the UC's master
        ' key handler function.
        'cMouseEvents.addOverrideHwnds picStatusBar.hWnd, picScrollH.hWnd, picScrollV.hWnd, picProgressBar.hWnd, cmdZoomIn.hWnd, cmdZoomOut.hWnd, cmdZoomFit.hWnd, cmdImgSize.hWnd
        
        'Enable key tracking as well
        Set cKeyEvents = New pdInputKeyboard
        cKeyEvents.createKeyboardTracker "pdCanvas", picCanvas.hWnd, VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN, VK_DELETE, VK_INSERT, VK_TAB, VK_SPACE, VK_ESCAPE, VK_BACK
                
        'Assign tooltips manually (so theming is supported)
        Set m_ToolTip = New clsToolTip
        m_ToolTip.Create Me
        m_ToolTip.MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
        m_ToolTip.DelayTime(ttDelayShow) = 10000
        
        m_ToolTip.AddTool cmbZoom, "Click to adjust image zoom"
        
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
        
        'Align the main picture box
        alignCanvasPictureBox
        
    End If
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Make sure interactions with this canvas are allowed
    If Not isCanvasInteractionAllowed() Then Exit Sub
    
    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    
    'If a selection is active, notify it of any changes in the shift key (which is used to request 1:1 selections)
    If pdImages(g_CurrentImage).selectionActive Then pdImages(g_CurrentImage).mainSelection.requestSquare ShiftDown
    
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    'Make sure interactions with this canvas are allowed
    If Not isCanvasInteractionAllowed() Then Exit Sub
    
    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    
    'If a selection is active, notify it of any changes in the shift key (which is used to request 1:1 selections)
    If pdImages(g_CurrentImage).selectionActive Then pdImages(g_CurrentImage).mainSelection.requestSquare ShiftDown
    
End Sub

Private Sub HScroll_Scroll()
    If (Not m_suspendRedraws) Then Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
End Sub

Private Sub UserControl_Resize()

    'Center all combo boxes vertically (this is necessary for high-DPI displays)
    cmbZoom.Top = (picStatusBar.ScaleHeight - cmbZoom.Height) \ 2
    cmbSizeUnit.Top = (picStatusBar.ScaleHeight - cmbSizeUnit.Height) \ 2

    'Align the canvas picture box to fill the available area
    alignCanvasPictureBox
    
    'If the control is resizing, the mouse cannot feasibly be over the image - so clear the coordinate box.  Note that this will
    ' also realign all chrome elements, so we don't need a manual fixChromeLayout call here.
    displayCanvasCoordinates 0, 0, False
    
End Sub

Public Sub alignCanvasPictureBox()
    
    Dim hOffsetRight As Long, vOffsetBottom As Long
    
    'Calculate vertical offsets
    If picScrollH.Visible Then vOffsetBottom = picScrollH.Height Else vOffsetBottom = 0
    vOffsetBottom = vOffsetBottom + picStatusBar.Height
    
    'Calculate horizontal offsets
    If picScrollV.Visible Then hOffsetRight = picScrollV.Width Else hOffsetRight = 0
    
    'Move the canvas picture box into position
    MoveWindow picCanvas.hWnd, 0, 0, UserControl.ScaleWidth - hOffsetRight, UserControl.ScaleHeight - vOffsetBottom, 1
    
End Sub

Private Sub UserControl_Show()

    If g_IsProgramRunning Then
        
        'Prep the command buttons
        cmdZoomFit.AssignImage "SB_ZOOM_FIT"
        cmdZoomIn.AssignImage "SB_ZOOM_IN"
        cmdZoomOut.AssignImage "SB_ZOOM_OUT"
        cmdImgSize.AssignImage "SB_IMG_SIZE"
        
                
        'Load various status bar icons from the resource file
        Set sbIconSize = New pdDIB
        Set sbIconCoords = New pdDIB
        
        loadResourceToDIB "SB_IMG_SIZE", sbIconSize
        loadResourceToDIB "SB_MOUSE_POS", sbIconCoords
        
        Set iconLoadAnImage = New pdDIB
        loadResourceToDIB "IMAGE_ETCH_256", iconLoadAnImage
        
        'XP users may not have Segoe UI available, which will cause the following lines to throw an error;
        ' it's not really a problem, as the labels will just keep their Tahoma font, but we must catch it anyway.
        On Error GoTo CanvasShowError
                
        'Now comes a bit of an odd case.  This control's _Show event happens very early in the load process due to it being
        ' present on FormMain.  Because of that, the global interface font value may not be loaded yet.  To avoid problems
        ' from this, we will just load Segoe UI by default, and if that fails (as it may on XP), the labels will retain
        ' their default Tahoma label.
        
        'Convert all labels to the current interface font
        If Len(g_InterfaceFont) = 0 Then g_InterfaceFont = "Segoe UI"
        
        lblCoordinates.updateAgainstCurrentTheme
        lblImgSize.updateAgainstCurrentTheme
        lblMessages.updateAgainstCurrentTheme
        
CanvasShowError:
        
        'Make all status bar lines a proper height (again, necessary on high-DPI displays)
        Dim i As Long
        For i = 0 To lineStatusBar.Count - 1
            lineStatusBar(i).y1 = fixDPI(1)
            lineStatusBar(i).y2 = picStatusBar.ScaleHeight - fixDPI(1)
        Next i
        
    End If
    
    Exit Sub

End Sub

Private Sub VScroll_Scroll()
    If (Not m_suspendRedraws) Then Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), Me
End Sub

'Whenever this window changes size, we may need to re-align various bits of internal chrome (status bar, rulers, etc).  Call this function
' to do so.
Public Sub fixChromeLayout()
    
    'Move the message label into position (right-aligned, with a slight margin)
    Dim newLeft As Long
    newLeft = lineStatusBar(2).x1 + fixDPI(7)
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
    
    'If the canvas is currently disabled (e.g. no image is loaded), let the user know that they can drag/drop files onto
    ' this space to begin editing
    If (g_OpenImageCount = 0) And g_IsProgramRunning Then
    
        'Ignore redraws if the program is being closed; this improves program termination performance
        If (Not g_ProgramShuttingDown) Then
        
            'Hide scrollbars if they aren't already
            setScrollVisibility PD_HORIZONTAL, False
            setScrollVisibility PD_VERTICAL, False
    
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createBlank picCanvas.ScaleWidth, picCanvas.ScaleHeight, 24, g_CanvasBackground
            
            Dim notifyFont As pdFont
            Set notifyFont = New pdFont
            notifyFont.setFontFace g_InterfaceFont
            
            'Set the font size dynamically.  en-US gets a larger size; other languages, whose text may be longer, use a smaller one.
            If Not (g_Language Is Nothing) Then
            
                If g_Language.translationActive Then
                    notifyFont.setFontSize 13
                Else
                    notifyFont.setFontSize 14
                End If
                
            Else
                notifyFont.setFontSize 14
            End If
            
            notifyFont.setFontBold False
            notifyFont.setFontColor RGB(41, 43, 54)
            notifyFont.setTextAlignment vbCenter
            
            'Create the font and attach it to our temporary DIB's DC
            notifyFont.createFontObject
            notifyFont.attachToDC tmpDIB.getDIBDC
            
            If Not (iconLoadAnImage Is Nothing) Then
            
                Dim modifiedHeight As Long
                modifiedHeight = tmpDIB.getDIBHeight + (iconLoadAnImage.getDIBHeight / 2) + fixDPI(24)
                
                Dim loadImageMessage As String
                loadImageMessage = g_Language.TranslateMessage("Drag an image onto this space to begin editing." & vbCrLf & vbCrLf & "You can also use the Open Image button on the left," & vbCrLf & "or the File > Open and File > Import menus.")
                notifyFont.drawCenteredText loadImageMessage, tmpDIB.getDIBWidth, modifiedHeight
                
                'Just above the text instructions, add a generic image icon
                iconLoadAnImage.alphaBlendToDC tmpDIB.getDIBDC, 192, (tmpDIB.getDIBWidth - iconLoadAnImage.getDIBWidth) / 2, (modifiedHeight / 2) - (iconLoadAnImage.getDIBHeight) - fixDPI(20)
                
            End If
            
            BitBlt picCanvas.hDC, 0, 0, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
            requestBufferSync
            Set tmpDIB = Nothing
            
        End If
        
    End If

End Sub

'Dynamically render some icons onto the status bar.
Public Sub drawStatusBarIcons(ByVal enabledState As Boolean)
    
    'Start by clearing the status bar
    picStatusBar.Picture = LoadPicture("")
    
    'If no images are loaded, do not draw status bar icons.
    If enabledState Then
        
        'Move the left-most line into position.  (This must be done dynamically, or it will be mispositioned
        ' on high-DPI displays)
        lineStatusBar(0).Visible = True
        lineStatusBar(0).x1 = (cmdZoomIn.Left + cmdZoomIn.Width) + fixDPI(6)
        lineStatusBar(0).x2 = lineStatusBar(0).x1
        
        'Start with the "image size" button
        cmdImgSize.Left = lineStatusBar(0).x1 + fixDPI(4)
        If (Not cmdImgSize.Visible) Then cmdImgSize.Visible = True
        
        'After the "image size" icon comes the actual image size label.  Its position is determined by the image resize button width,
        ' plus a 4px buffer on either size (contingent on DPI)
        lblImgSize.Left = cmdImgSize.Left + cmdImgSize.Width + fixDPI(4)
        
        'The image size label is autosized.  Move the "size unit" combo box next to it, and the next vertical line
        ' separator just past it.
        If (Not cmbSizeUnit.Visible) Then cmbSizeUnit.Visible = True
        cmbSizeUnit.Left = lblImgSize.Left + lblImgSize.PixelWidth + fixDPI(10)
        
        If (Not lineStatusBar(1).Visible) Then lineStatusBar(1).Visible = True
        lineStatusBar(1).x1 = cmbSizeUnit.Left + cmbSizeUnit.Width + fixDPI(10)
        lineStatusBar(1).x2 = lineStatusBar(1).x1
        
        'After the "image size" panel and separator comes mouse coordinates.  The basic steps from above are repeated.
        If Not sbIconCoords Is Nothing Then
            sbIconCoords.alphaBlendToDC picStatusBar.hDC, , lineStatusBar(1).x1 + fixDPI(8), fixDPI(4), fixDPI(sbIconCoords.getDIBWidth), fixDPI(sbIconCoords.getDIBHeight)
        End If
        lblCoordinates.Left = lineStatusBar(1).x1 + fixDPI(14) + fixDPI(16)
        
        If (Not lineStatusBar(2).Visible) Then lineStatusBar(2).Visible = True
        lineStatusBar(2).x1 = lblCoordinates.Left + lblCoordinates.PixelWidth + fixDPI(10)
        lineStatusBar(2).x2 = lineStatusBar(2).x1
        
        'Note that we don't actually move the last line status bar; that is handled by DisplayImageCoordinates itself
        
    'Images are not loaded.  Hide the lines and other items.
    Else
    
        cmdImgSize.Visible = False
        cmbSizeUnit.Visible = False
        lineStatusBar(0).Visible = False
        lineStatusBar(1).Visible = False
        lineStatusBar(2).Visible = False
                
    End If
    
    'Make our painting persistent
    picStatusBar.Picture = picStatusBar.Image
    picStatusBar.Refresh
    
End Sub

'Fill the "size units" drop-down.  We must do this later in the load process, as we have to wait for the translation engine to load.
Public Function populateSizeUnits()

    'Add size units to the size unit drop-down box
    cmbSizeUnit.Clear
    cmbSizeUnit.AddItem g_Language.TranslateMessage(" px"), 0
    cmbSizeUnit.AddItem g_Language.TranslateMessage(" in"), 1
    cmbSizeUnit.AddItem g_Language.TranslateMessage(" cm"), 2
    cmbSizeUnit.ListIndex = 0

End Function

'Whenever the mouse cursor needs to be reset, use this function to do so.  Also, when a new tool is created or a new tool feature
' is added, make sure to visit this sub and make any necessary cursor changes!
'
'A lot of extra values are passed to this function.  Individual tools can use those at their leisure to customize their cursor requests.
Private Sub setCanvasCursor(ByVal curMouseEvent As PD_MOUSEEVENT, ByVal Button As Integer, ByVal x As Single, ByVal y As Single, ByVal imgX As Double, ByVal imgY As Double)

    'Obviously, cursor setting is handled separately for each tool.
    Select Case g_CurrentTool
        
        Case NAV_DRAG
        
            'When click-dragging the image to scroll around it, the cursor depends on being over the image
            If isMouseOverImage(x, y, pdImages(g_CurrentImage)) Then
                
                If Button <> 0 Then
                    cMouseEvents.setPNGCursor "HANDCLOSED", 0, 0
                Else
                    cMouseEvents.setPNGCursor "HANDOPEN", 0, 0
                End If
            
            'If the cursor is not over the image, change to an arrow cursor
            Else
                cMouseEvents.setSystemCursor IDC_ARROW
            End If
        
        Case NAV_MOVE
        
            'When transforming layers, the cursor depends on several factors
            Select Case pdImages(g_CurrentImage).getActiveLayer.checkForPointOfInterest(imgX, imgY)
            
                'Mouse is not over the current layer
                Case -1
                    cMouseEvents.setSystemCursor IDC_ARROW
                    
                'Mouse is over the top-left corner
                Case 0
                    cMouseEvents.setSystemCursor IDC_SIZENWSE
                    
                'Mouse is over the top-right corner
                Case 1
                    cMouseEvents.setSystemCursor IDC_SIZENESW
                    
                'Mouse is over the bottom-right corner
                Case 2
                    cMouseEvents.setSystemCursor IDC_SIZENWSE
                    
                'Mouse is over the bottom-left corner
                Case 3
                    cMouseEvents.setSystemCursor IDC_SIZENESW
                    
                'Mouse is within the layer, but not over a specific node
                Case 4
                
                    'This case is unique because if the user has elected to ignore transparent pixels, they cannot move a layer
                    ' by dragging the mouse within a transparent region of the layer.  Thus, before changing the cursor,
                    ' check to see if the hovered layer index is the same as the current layer index; if it isn't, don't display
                    ' the Move cursor.  (Note that this works because the getLayerUnderMouse function, called during the MouseMove
                    ' event, automatically factors the transparency check into its calculation.  Thus we don't have to
                    ' re-evaluate the setting here.)
                    If m_LayerAutoActivateIndex = pdImages(g_CurrentImage).getActiveLayerIndex Then
                        cMouseEvents.setSystemCursor IDC_SIZEALL
                    Else
                        cMouseEvents.setSystemCursor IDC_ARROW
                    End If
                    
            End Select
            
        Case SELECT_RECT, SELECT_CIRC
        
            'When transforming selections, the cursor image depends on its proximity to a point of interest.
            '
            'For a rectangle or circle selection, the possible transform IDs are:
            ' -1 - Cursor is not near a selection point
            ' 0 - NW corner
            ' 1 - NE corner
            ' 2 - SE corner
            ' 3 - SW corner
            ' 4 - N edge
            ' 5 - E edge
            ' 6 - S edge
            ' 7 - W edge
            ' 8 - interior of selection, not near a corner or edge
            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
            
                Case -1
                    cMouseEvents.setSystemCursor IDC_ARROW
                Case 0
                    cMouseEvents.setSystemCursor IDC_SIZENWSE
                Case 1
                    cMouseEvents.setSystemCursor IDC_SIZENESW
                Case 2
                    cMouseEvents.setSystemCursor IDC_SIZENWSE
                Case 3
                    cMouseEvents.setSystemCursor IDC_SIZENESW
                Case 4
                    cMouseEvents.setSystemCursor IDC_SIZENS
                Case 5
                    cMouseEvents.setSystemCursor IDC_SIZEWE
                Case 6
                    cMouseEvents.setSystemCursor IDC_SIZENS
                Case 7
                    cMouseEvents.setSystemCursor IDC_SIZEWE
                Case 8
                    cMouseEvents.setSystemCursor IDC_SIZEALL
            
            End Select
        
        Case SELECT_LINE
        
            'When transforming selections, the cursor image depends on its proximity to a point of interest.
            '
            'For a line selection, the possible transform IDs are:
            ' -1 - Cursor is not near an endpoint
            ' 0 - Near x1/y1
            ' 1 - Near x2/y2
            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
            
                Case -1
                    cMouseEvents.setSystemCursor IDC_ARROW
                Case 0
                    cMouseEvents.setSystemCursor IDC_SIZEALL
                Case 1
                    cMouseEvents.setSystemCursor IDC_SIZEALL
            
            End Select
        
         Case SELECT_POLYGON
            
            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
            
                '-1: mouse is outside the lasso selection area
                Case -1
                    cMouseEvents.setSystemCursor IDC_ARROW
                
                'numOfPolygonPoints: mouse is inside the polygon, but not over a polygon node
                Case pdImages(g_CurrentImage).mainSelection.getNumOfPolygonPoints
                    If pdImages(g_CurrentImage).mainSelection.isLockedIn Then
                        cMouseEvents.setSystemCursor IDC_SIZEALL
                    Else
                        cMouseEvents.setSystemCursor IDC_ARROW
                    End If
                    
                'Everything else: mouse is over a polygon node
                Case Else
                    cMouseEvents.setSystemCursor IDC_SIZEALL
                    
            End Select
        
        Case SELECT_LASSO
            
            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
            
                '-1: mouse is outside the lasso selection area
                Case -1
                    cMouseEvents.setSystemCursor IDC_ARROW
                
                '0: mouse is inside the lasso selection area.  As a convenience to the user, we don't update the cursor
                '   if they're still in "drawing" mode - we only update it if the selection is complete.
                Case 0
                    If pdImages(g_CurrentImage).mainSelection.isLockedIn Then
                        cMouseEvents.setSystemCursor IDC_SIZEALL
                    Else
                        cMouseEvents.setSystemCursor IDC_ARROW
                    End If
                    
            End Select
            
        Case SELECT_WAND
        
            Select Case findNearestSelectionCoordinates(imgX, imgY, pdImages(g_CurrentImage))
            
                '-1: mouse is outside the lasso selection area
                Case -1
                    cMouseEvents.setSystemCursor IDC_ARROW
                
                '0: mouse is inside the lasso selection area.  As a convenience to the user, we don't update the cursor
                '   if they're still in "drawing" mode - we only update it if the selection is complete.
                Case Else
                    cMouseEvents.setSystemCursor IDC_SIZEALL
                    
            End Select
        
        
        Case Else
            cMouseEvents.setSystemCursor IDC_ARROW
                    
    End Select

End Sub

'Simple unified way to see if canvas interaction is allowed.
Private Function isCanvasInteractionAllowed() As Boolean

    'By default, canvas interaction is allowed
    isCanvasInteractionAllowed = True
    
    'Now, check a bunch of states that might indicate canvas interactions should not be allowed
    
    'If the main form is disabled, exit
    If Not FormMain.Enabled Then isCanvasInteractionAllowed = False
        
    'If user input has been forcibly disabled, exit
    If g_DisableUserInput Then isCanvasInteractionAllowed = False
    
    'If no images have been loaded, exit
    If g_OpenImageCount = 0 Then isCanvasInteractionAllowed = False
    
    'If the current image does not exist, exit
    If pdImages(g_CurrentImage) Is Nothing Then
        isCanvasInteractionAllowed = False
    Else
        
        'If an image has not yet been loaded, exit
        If Not pdImages(g_CurrentImage).IsActive Then isCanvasInteractionAllowed = False
        If Not pdImages(g_CurrentImage).loadedSuccessfully Then isCanvasInteractionAllowed = False
        If pdImages(g_CurrentImage).getNumOfLayers = 0 Then isCanvasInteractionAllowed = False
        
    End If
    
End Function
