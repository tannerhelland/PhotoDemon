VERSION 5.00
Begin VB.UserControl brushSelector 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
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
   MousePointer    =   99  'Custom
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ToolboxBitmap   =   "brushSelector.ctx":0000
End
Attribute VB_Name = "brushSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Brush Selector custom control
'Copyright 2013-2015 by Tanner Helland
'Created: 30/June/15
'Last updated: 30/June/15
'Last update: initial build
'
'This thin user control is basically an empty control that when clicked, displays a brush selection window.  If a
' brush is selected (e.g. Cancel is not pressed), it updates its appearance to match, and raises a "BrushChanged"
' event.
'
'Though simple, this control solves a lot of problems.  It is especially helpful for improving interaction with the
' command bar user control, as it easily supports brush reset/randomize/preset events.  It is also nice to be able
' to update a single master function for brush selection, then have the change propagate to all tool windows.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

'This control doesn't really do anything interesting, besides allow a brush to be selected.
Public Event BrushChanged()

'A specialized class handles mouse input for this control
Private WithEvents cMouseEvents As pdInputMouse
Attribute cMouseEvents.VB_VarHelpID = -1

'Reliable focus detection requires a specialized subclasser
Private WithEvents cFocusDetector As pdFocusDetector
Attribute cFocusDetector.VB_VarHelpID = -1
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'The control's current brush settings
Private m_curBrush As String

'A temporary filler object, used to render the brush preview
Private m_Filler As pdGraphicsBrush

'When the "select brush" dialog is live, this will be set to TRUE
Private isDialogLive As Boolean

'A backing DIB is required for proper color management
Private m_BackBuffer As pdDIB

'This value will be TRUE while the mouse is inside the UC
Private m_MouseInsideUC As Boolean

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'At present, all this control does is store a brush param string
Public Property Get Brush() As String
    Brush = m_curBrush
End Property

Public Property Let Brush(ByVal newBrush As String)
    
    m_curBrush = newBrush
    
    'Redraw the control to match
    drawControl
    
    PropertyChanged "Brush"
    RaiseEvent BrushChanged
    
End Property

'Outside functions can call this to force a display of the brush selection window
Public Sub displayBrushSelection()
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
    
    'Store the current brush
    Dim newBrush As String, oldBrush As String
    oldBrush = Brush
    
    'Use the brush dialog to select a new color
    If showBrushDialog(newBrush, oldBrush, Me) Then
        Brush = newBrush
    Else
        Brush = oldBrush
    End If
    
    isDialogLive = False
    
End Sub

Private Sub UserControl_Initialize()

    Set m_Filler = New pdGraphicsBrush
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
    Brush = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Brush = PropBag.ReadProperty("curBrush", "")
End Sub

Private Sub UserControl_Resize()
    drawControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "curBrush", m_curBrush, ""
End Sub

'For flexibility, we draw our own borders.  I may decide to change this behavior in the future...
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
        
        'Render the brush first.  (Gradient brushes require a target width/height, which we want to be the same size as the control.)
        Dim cBounds As RECTF
        With cBounds
            .Left = 0
            .Top = 0
            .Width = UserControl.ScaleWidth
            .Height = UserControl.ScaleHeight
        End With
        
        m_Filler.setBoundaryRect cBounds
        m_Filler.createBrushFromString Me.Brush
        
        Dim tmpBrush As Long
        tmpBrush = m_Filler.getBrushHandle
        
        GDI_Plus.GDIPlusFillDIBRect_Pattern m_BackBuffer, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight, g_CheckerboardPattern
        GDI_Plus.GDIPlusFillDC_Brush m_BackBuffer.getDIBDC, tmpBrush, 0, 0, m_BackBuffer.getDIBWidth, m_BackBuffer.getDIBHeight
        m_Filler.releaseBrushHandle tmpBrush
        
        'Draw borders around the brush results.
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
        ' (Note also that we use a g_IsProgramRunning check to prevent color management from firing at compile-time.)
        If g_IsProgramRunning Then TurnOnDefaultColorManagement UserControl.hDC, UserControl.hWnd
        BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_BackBuffer.getDIBDC, 0, 0, vbSrcCopy
    
    Else
        UserControl.BackColor = m_Filler.getBrushProperty(pgbs_PrimaryColor)
    End If
    
    UserControl.Picture = UserControl.Image
    UserControl.Refresh
    
End Sub

'If a brush selection dialog is active, it will pass brush updates backward to this function, so that we can let
' our parent form display live updates *while the user is playing with brushes* - very cool!
Public Sub notifyOfLiveBrushChange(ByVal newBrush As String)
    Brush = newBrush
End Sub
