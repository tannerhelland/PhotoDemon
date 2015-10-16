VERSION 5.00
Begin VB.UserControl gradientSelector 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   ToolboxBitmap   =   "gradientSelector.ctx":0000
End
Attribute VB_Name = "gradientSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Gradient Selector custom control
'Copyright 2014-2015 by Tanner Helland
'Created: 23/July/15
'Last updated: 23/July/15
'Last update: initial build
'
'This thin user control is basically an empty control that when clicked, displays a gradient editor window.  If a
' gradient is selected (e.g. Cancel is not pressed), it updates its appearance to match, and raises a "GradientChanged"
' event.
'
'Though simple, this control solves a lot of problems.  It is especially helpful for improving interaction with the
' command bar user control, as it easily supports gradient reset/randomize/preset events.  It is also nice to be able
' to update a single master function for gradient selection, then have the change propagate to all tool windows.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'This control doesn't really do anything interesting, besides allow a gradient to be selected.
Public Event GradientChanged()

'A specialized class handles mouse input for this control
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

'Reliable focus detection requires a specialized subclasser
Private WithEvents cFocusDetector As pdFocusDetector
Attribute cFocusDetector.VB_VarHelpID = -1
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'The control's current gradient settings
Private m_curGradient As String

'Temporary brush object, used to render the gradient preview
Private m_Brush As pdGraphicsBrush

'When the "select gradient" dialog is live, this will be set to TRUE
Private isDialogLive As Boolean

'A backing DIB is required for proper color management
Private m_BackBuffer As pdDIB

'This value will be TRUE while the mouse is inside the UC
Private m_MouseInsideUC As Boolean

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'You can retrieve the gradient param string (not a pdGradient object!) via this property
Public Property Get Gradient() As String
    Gradient = m_curGradient
End Property

Public Property Let Gradient(ByVal newGradient As String)
    
    m_curGradient = newGradient
    
    'Redraw the control to match
    drawControl
    
    PropertyChanged "Gradient"
    RaiseEvent GradientChanged
    
End Property

'Outside functions can call this to force a display of the gradient selection window
Public Sub displayGradientSelection()
    UserControl_Click
End Sub

Private Sub cMouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideUC = True
    drawControl
    cMouseEvents.setSystemCursor IDC_HAND
End Sub

Private Sub cMouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseInsideUC = False
    drawControl
    cMouseEvents.setSystemCursor IDC_DEFAULT
End Sub

'When the control receives focus, relay the event externally
Private Sub cFocusDetector_GotFocusReliable()
    RaiseEvent GotFocusAPI
End Sub

'When the control loses focus, relay the event externally
Private Sub cFocusDetector_LostFocusReliable()
    RaiseEvent LostFocusAPI
End Sub

Private Sub UserControl_Click()

    isDialogLive = True
    
    'Backup the current gradient; if the dialog is canceled, we want to restore it
    Dim newGradient As String, oldGradient As String
    oldGradient = Gradient
    
    'Display the gradient dialog, then wait for it to return
    If showGradientDialog(newGradient, oldGradient, Me) Then
        Gradient = newGradient
    Else
        Gradient = oldGradient
    End If
    
    isDialogLive = False
    
End Sub

Private Sub UserControl_Initialize()
    
    Set m_Brush = New pdGraphicsBrush
    m_Brush.setBrushProperty pgbs_BrushMode, 2
    
    'Render the initial control
    drawControl
    
    If g_IsProgramRunning Then
        
        'Initialize mouse handling
        Set cMouseEvents = New pdInputMouse
        cMouseEvents.addInputTracker UserControl.hWnd, True, , , True
        cMouseEvents.setSystemCursor IDC_HAND
        
        'Also start a focus detector
        Set cFocusDetector = New pdFocusDetector
        cFocusDetector.startFocusTracking Me.hWnd
        
    End If
    
End Sub

Private Sub UserControl_InitProperties()
    Gradient = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Gradient = PropBag.ReadProperty("curGradient", "")
End Sub

Private Sub UserControl_Resize()
    
    'Redraw the control
    drawControl
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "curGradient", m_curGradient, ""
End Sub

'Redraw the control
Private Sub drawControl()
        
    'For color management to work, we must pre-render the control onto a DIB, then copy the DIB to the screen.
    ' Using VB's internal draw commands leads to unpredictable results.
    If m_BackBuffer Is Nothing Then Set m_BackBuffer = New pdDIB
    
    If (m_BackBuffer.getDIBWidth <> UserControl.ScaleWidth) Or (m_BackBuffer.getDIBHeight <> UserControl.ScaleHeight) Then
        m_BackBuffer.createBlank UserControl.ScaleWidth, UserControl.ScaleHeight, 24, 0
    Else
        m_BackBuffer.resetDIB
    End If
    
    'Because so much of the rendering code requires GDI+, we can't do much in the IDE
    If g_IsProgramRunning Then
    
        'Render the background checkerboard first
        GDI_Plus.GDIPlusFillDIBRect_Pattern m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, g_CheckerboardPattern
        m_Brush.setBrushProperty pgbs_GradientString, m_curGradient
        
        Dim cBounds As RECTF
        With cBounds
            .Left = 0
            .Top = 0
            .Width = UserControl.ScaleWidth
            .Height = UserControl.ScaleHeight
        End With
        
        m_Brush.setBoundaryRect cBounds
        GDI_Plus.GDIPlusFillDC_Brush m_BackBuffer.getDIBDC, m_Brush.getBrushHandle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
        
        'Draw borders around the preview.
        Dim outlineColor As Long, outlineWidth As Long, outlineOffset As Long
    
        If g_IsProgramRunning And m_MouseInsideUC Then
            outlineColor = g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
            outlineWidth = 3
            outlineOffset = 1
        Else
            outlineColor = vbBlack
            outlineWidth = 1
            outlineOffset = 0
        End If
        
        GDIPlusDrawLineToDC m_BackBuffer.getDIBDC, 0, outlineOffset, UserControl.ScaleWidth - 1, outlineOffset, outlineColor, , outlineWidth, , LineCapFlat
        GDIPlusDrawLineToDC m_BackBuffer.getDIBDC, UserControl.ScaleWidth - 1 - outlineOffset, 0, UserControl.ScaleWidth - 1 - outlineOffset, UserControl.ScaleHeight - 1, outlineColor, , outlineWidth, , LineCapFlat
        GDIPlusDrawLineToDC m_BackBuffer.getDIBDC, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1 - outlineOffset, 0, UserControl.ScaleHeight - 1 - outlineOffset, outlineColor, , outlineWidth, , LineCapFlat
        GDIPlusDrawLineToDC m_BackBuffer.getDIBDC, outlineOffset, UserControl.ScaleHeight - 1, outlineOffset, 0, outlineColor, , outlineWidth, , LineCapFlat
        
        'Render the completed DIB to the control.  (This is when color management takes place.)
        TurnOnDefaultColorManagement UserControl.hDC, UserControl.hWnd
        BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy
        
    Else
        
        'TODO: failsafe backcolor correction based on the primary gradient color
        'UserControl.BackColor = m_GradientPreview.getGradientProperty()
        
    End If
        
    UserControl.Picture = UserControl.Image
    UserControl.Refresh
    
End Sub

'If a gradient selection dialog is active, it will pass gradient updates backward to this function, so that we can let
' our parent form display live updates *while the user is playing with gradients*.
Public Sub notifyOfLiveGradientChange(ByVal newGradient As String)
    Gradient = newGradient
End Sub

