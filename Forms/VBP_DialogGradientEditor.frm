VERSION 5.00
Begin VB.Form dialog_GradientEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Gradient editor"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   547
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   844
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Index           =   0
      Left            =   120
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   825
      TabIndex        =   4
      Top             =   3480
      Width           =   12375
      Begin PhotoDemon.pdLabel lblInstructions 
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   1296
         Alignment       =   2
         Caption         =   "yes"
         FontSize        =   9
         Layout          =   1
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   0
         Left            =   120
         Top             =   840
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   556
         Caption         =   "node settings:"
         FontSize        =   12
      End
   End
   Begin PhotoDemon.buttonStrip btsEdit 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   873
      FontSize        =   12
   End
   Begin VB.PictureBox picInteract 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   841
      TabIndex        =   2
      Top             =   2310
      Width           =   12615
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   240
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   807
      TabIndex        =   1
      Top             =   480
      Width           =   12135
   End
   Begin PhotoDemon.commandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   7455
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   1323
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14802140
      AutoloadLastPreset=   -1  'True
      dontAutoUnloadParent=   -1  'True
      dontResetAutomatically=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Index           =   1
      Left            =   240
      Top             =   120
      Width           =   9255
      _ExtentX        =   16536
      _ExtentY        =   556
      Caption         =   "preview:"
      FontSize        =   12
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Index           =   1
      Left            =   120
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   825
      TabIndex        =   5
      Top             =   3480
      Width           =   12375
   End
End
Attribute VB_Name = "dialog_GradientEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Gradient Editor Dialog
'Copyright 2014-2015 by Tanner Helland
'Created: 23/July/15 (but assembled from many bits written earlier)
'Last updated: 23/July/15
'Last update: initial build
'
'Comprehensive gradient editor.  This dialog is currently based around the properties of GDI+ gradient brushes, but it
' could easily be expanded in the future due to its modular design.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'OK/Cancel result from the dialog
Private userAnswer As VbMsgBoxResult

'The original gradient when the dialog was first loaded
Private m_OldGradient As String

'Gradient strings are generated with the help of PD's core gradient class.  This class also renders a preview of the current gradient.
Private m_GradientPreview As pdGradient

'If a user control spawned this dialog, it will pass itself as a reference.  We can then send gradient updates back
' to the control, allowing for real-time updates on the screen despite a modal dialog being raised!
Private parentGradientControl As gradientSelector

'Gradient previews are rendered using a pdGraphicsPath as the sample area
Private m_PreviewPath As pdGraphicsPath

'Recently used gradients are loaded to/saved from a custom XML file
Private m_XMLEngine As pdXML

'The file where we'll store recent gradient data when the program is closed.  (At present, this file is located in PD's
' /Data/Presets/ folder.
Private m_XMLFilename As String

'Gradient preview DIB (required for color management) and interaction DIB (where all the gradient nodes are rendered)
Private m_PreviewDIB As pdDIB, m_InteractiveDIB As pdDIB

'To prevent recursive setting changes, this value can be set to TRUE to prevent live preview updates
Private m_SuspendRedraws As Boolean

'All mouse interactions for creating/editing gradients is handled by PD's mouse manager
Private WithEvents m_MouseEvents As pdInputMouse
Attribute m_MouseEvents.VB_VarHelpID = -1

'This interface tracks its own collection of gradient points
Private m_NumOfGradientPoints As Long
Private m_GradientPoints() As pdGradientPoint

'The current gradient point (index) selected and/or hovered by the mouse.  -1 if no point is currently selected/hovered.
Private m_CurPoint As Long, m_CurHoverPoint As Long

'Size of gradient "nodes" in the interactive UI.
Private Const GRADIENT_NODE_WIDTH As Single = 12#
Private Const GRADIENT_NODE_HEIGHT As Single = 14#

'Other gradient node UI renderers
Private inactiveArrowFill As pdGraphicsBrush, activeArrowFill As pdGraphicsBrush
Private inactiveOutlinePen As pdGraphicsPen, activeOutlinePen As pdGraphicsPen

'The user's answer is returned via this property
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The newly selected gradient (if any) is returned via this property
Public Property Get newGradient() As String
    newGradient = m_GradientPreview.getGradientAsString
End Property

'The ShowDialog routine presents the user with this form.
Public Sub showDialog(ByVal initialGradient As String, Optional ByRef callingControl As gradientSelector = Nothing)
    
    'Store a reference to the calling control (if any)
    Set parentGradientControl = callingControl

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbCancel
    
    'Cache the initial gradient parameters so we can access it elsewhere
    m_OldGradient = initialGradient
    Set m_GradientPreview = New pdGradient
    m_GradientPreview.createGradientFromString initialGradient
    
    If Len(initialGradient) = 0 Then initialGradient = m_GradientPreview.getGradientAsString
    
    'Sync all controls to the initial pen parameters
    syncControlsToGradientObject
    updatePreview
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    
    'Apply extra images and tooltips to certain controls
    
    'Apply visual themes
    makeFormPretty Me
    
    'Initialize an XML engine, which we will use to read/write recent pen data to file
    Set m_XMLEngine = New pdXML
    
    'The XML file will be stored in the Preset path (/Data/Presets)
    m_XMLFilename = g_UserPreferences.getPresetPath & "Gradient_Selector.xml"
    
    'TODO: if an XML file exists, load its contents now
    'loadRecentGradientList
        
    'Display the dialog
    showPDDialog vbModal, Me, True

End Sub

Private Sub btsEdit_Click(ByVal buttonIndex As Long)
    
    Dim i As Long
    For i = picContainer.lBound To picContainer.UBound
        picContainer(i).Visible = CBool(i = buttonIndex)
    Next i
    
End Sub

'CANCEL BUTTON
Private Sub cmdBar_CancelClick()
    userAnswer = vbCancel
    Me.Hide
End Sub

'OK BUTTON
Private Sub cmdBar_OKClick()

    'Store the newGradient value (which the dialog handler will use to return the selected gradient to the caller)
    updateGradientObject
    
    'TODO: save the current list of recently used gradients
    'saveRecentGradientList
    
    userAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Reset our generic outline object
    Set m_GradientPreview = New pdGradient
    m_GradientPreview.createGradientFromString ""
    
    'Synchronize all controls to the updated settings
    syncControlsToGradientObject
    updatePreview
    
End Sub

Private Sub Form_Load()
    
    m_SuspendRedraws = True
    
    'Add the instructions label
    Dim instructionText As String
    instructionText = g_Language.TranslateMessage("Left-click to add new gradient nodes.  Right-click to remove existing nodes.")
    lblInstructions.Caption = instructionText
    
    'Populate button strips and drop-downs
    btsEdit.AddItem "manual", 0
    btsEdit.AddItem "automatic", 1
    btsEdit_Click 0
    
    If g_IsProgramRunning Then
    
        If m_GradientPreview Is Nothing Then Set m_GradientPreview = New pdGradient
        If m_PreviewDIB Is Nothing Then Set m_PreviewDIB = New pdDIB
        
        'Set up a special mouse handler for the gradient interaction window
        If m_MouseEvents Is Nothing Then Set m_MouseEvents = New pdInputMouse
        m_MouseEvents.addInputTracker picInteract.hWnd, True, True, , True
        m_MouseEvents.setSystemCursor IDC_HAND
                
        'Prep a default set of gradient points
        resetGradientPoints
        
        'Prep all gradient point tracking variables
        m_CurPoint = -1
        
        'While we're here, we'll also prep all generic drawing objects for the interactive gradient node UI bits
        Set inactiveArrowFill = New pdGraphicsBrush
        Set activeArrowFill = New pdGraphicsBrush
        
        inactiveArrowFill.setBrushProperty pgbs_BrushMode, 0
        inactiveArrowFill.setBrushProperty pgbs_PrimaryOpacity, 100
        inactiveArrowFill.setBrushProperty pgbs_PrimaryColor, g_Themer.getThemeColor(PDTC_BACKGROUND_DEFAULT)
        
        activeArrowFill.setBrushProperty pgbs_BrushMode, 0
        activeArrowFill.setBrushProperty pgbs_PrimaryOpacity, 100
        activeArrowFill.setBrushProperty pgbs_PrimaryColor, g_Themer.getThemeColor(PDTC_ACCENT_ULTRALIGHT)
        
        Set inactiveOutlinePen = New pdGraphicsPen
        Set activeOutlinePen = New pdGraphicsPen
        
        inactiveOutlinePen.setPenProperty pgps_PenMode, 0
        inactiveOutlinePen.setPenProperty pgps_PenOpacity, 100
        inactiveOutlinePen.setPenProperty pgps_PenWidth, 1#
        inactiveOutlinePen.setPenProperty pgps_PenLineJoin, LineJoinRound
        inactiveOutlinePen.setPenProperty pgps_PenColor, g_Themer.getThemeColor(PDTC_GRAY_SHADOW)
        
        activeOutlinePen.setPenProperty pgps_PenMode, 0
        activeOutlinePen.setPenProperty pgps_PenOpacity, 100
        activeOutlinePen.setPenProperty pgps_PenWidth, 1#
        activeOutlinePen.setPenProperty pgps_PenLineJoin, LineJoinRound
        activeOutlinePen.setPenProperty pgps_PenColor, g_Themer.getThemeColor(PDTC_ACCENT_DEFAULT)
        
        'The preview path is simply the rectangular area of the preview box
        If m_PreviewPath Is Nothing Then Set m_PreviewPath = New pdGraphicsPath
        syncPreviewPath
        
        'Draw the initial set of interactive gradient nodes
        drawGradientNodes
                
    End If
    
    m_SuspendRedraws = False
    
End Sub

Private Sub resetGradientPoints()
    
    m_NumOfGradientPoints = 2
    ReDim m_GradientPoints(0 To m_NumOfGradientPoints - 1) As pdGradientPoint
    
    With m_GradientPoints(0)
        .pdgp_RGB = vbBlack
        .pdgp_Opacity = 1
        .pdgp_Position = 0
    End With
    
    With m_GradientPoints(1)
        .pdgp_RGB = vbWhite
        .pdgp_Opacity = 1
        .pdgp_Position = 1
    End With
    
    m_GradientPreview.createGradientFromPointCollection m_NumOfGradientPoints, m_GradientPoints
    
End Sub

Private Sub Form_Resize()
    syncPreviewPath
    drawGradientNodes
    updatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Synchronize the preview path against the size of the current preview window.  Note that this function does not request a redraw, by design.
' The caller must do that manually.
Private Sub syncPreviewPath()
    
    If Not m_PreviewPath Is Nothing Then
        
        m_PreviewPath.resetPath
        m_PreviewPath.addRectangle_Absolute 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight
        
    End If
    
End Sub

'Update our internal gradient class against any/all changed settings.
Private Sub updateGradientObject()

    With m_GradientPreview
        .createGradientFromPointCollection m_NumOfGradientPoints, m_GradientPoints
    End With

End Sub

Private Sub updatePreview()
    
    If Not m_SuspendRedraws Then
    
        'Make sure our gradient object is up-to-date
        updateGradientObject
        
        'Retrieve a matching brush handle
        Dim gdipBrush As Long, boundsRect As RECTF
        
        With boundsRect
            .Left = 0
            .Top = 0
            .Width = picPreview.ScaleWidth
            .Height = picPreview.ScaleHeight
        End With
        
        gdipBrush = m_GradientPreview.getBrushHandle(boundsRect, True, 0)
        
        'Prep the preview DIB
        If m_PreviewDIB Is Nothing Then Set m_PreviewDIB = New pdDIB
        
        If (m_PreviewDIB.getDIBWidth <> Me.picPreview.ScaleWidth) Or (m_PreviewDIB.getDIBHeight <> Me.picPreview.ScaleHeight) Then
            m_PreviewDIB.createBlank Me.picPreview.ScaleWidth, Me.picPreview.ScaleHeight, 24, 0
        Else
            m_PreviewDIB.resetDIB
        End If
        
        'Create the preview image
        With m_PreviewDIB
            GDI_Plus.GDIPlusFillDIBRect_Pattern m_PreviewDIB, 0, 0, .getDIBWidth, .getDIBHeight, g_CheckerboardPattern
            GDI_Plus.GDIPlusFillDC_Brush .getDIBDC, gdipBrush, 0, 0, .getDIBWidth, .getDIBHeight
        End With
        
        'Copy the preview image to the screen
        m_PreviewDIB.renderToPictureBox Me.picPreview
        
        'Release our GDI+ handle
        GDI_Plus.releaseGDIPlusBrush gdipBrush
                
        'Notify our parent of the update
        If Not (parentGradientControl Is Nothing) Then parentGradientControl.notifyOfLiveGradientChange m_GradientPreview.getGradientAsString
        
    End If
    
End Sub

Private Sub syncControlsToGradientObject()
        
    m_SuspendRedraws = True
        
    With m_GradientPreview
        .getCopyOfPointCollection m_NumOfGradientPoints, m_GradientPoints
    End With
    
    drawGradientNodes
    
    m_SuspendRedraws = False
    
End Sub

Private Sub m_MouseEvents_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    Dim i As Long
    
    'Clicking the mouse either selects an existing point, or creates a new point.
    ' As such, this function will always result in a legitimate value for m_CurPoint.
    
    'See if an existing has been selected.
    Dim tmpPoint As Long
    tmpPoint = getPointAtPosition(x, y)
    
    'If this is an existing point, we will either (LMB) mark it as the active point, or (RMB) remove it
    If tmpPoint >= 0 Then
        
        If (Button And pdLeftButton) <> 0 Then
            m_CurPoint = tmpPoint
        ElseIf ((Button And pdRightButton) <> 0) And (m_NumOfGradientPoints > 2) Then
            
            m_NumOfGradientPoints = m_NumOfGradientPoints - 1
            For i = tmpPoint To m_NumOfGradientPoints
                m_GradientPoints(i) = m_GradientPoints(i + 1)
            Next i
            
        End If
        
    'If this is not an existing point, create a new one now.
    Else
    
        'Enlarge the target array as necessary
        If m_NumOfGradientPoints >= UBound(m_GradientPoints) Then ReDim Preserve m_GradientPoints(0 To m_NumOfGradientPoints * 2) As pdGradientPoint
        
        With m_GradientPoints(m_NumOfGradientPoints)
            .pdgp_Opacity = 1
            .pdgp_Position = convertPixelCoordsToNodeCoords(x)
            .pdgp_RGB = vbBlack     'TODO: calculate this color intelligently
        End With
        
        m_CurPoint = m_NumOfGradientPoints
        m_NumOfGradientPoints = m_NumOfGradientPoints + 1
        
    End If
    
    'Regardless of outcome, we need to redraw the interaction area and preview
    updatePreview
    drawGradientNodes

End Sub

Private Sub m_MouseEvents_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseEvents.setSystemCursor IDC_HAND
End Sub

Private Sub m_MouseEvents_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_CurHoverPoint = -1
    m_MouseEvents.setSystemCursor IDC_DEFAULT
    drawGradientNodes
End Sub

Private Sub m_MouseEvents_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'First, separate our handling by mouse button state
    If (Button And pdLeftButton) <> 0 Then
    
        'The left mouse button is down.  Assign the new position to the active node.
        m_GradientPoints(m_CurPoint).pdgp_Position = convertPixelCoordsToNodeCoords(x)
        
        'Redraw the gradient interaction nodes and the gradient itself
        drawGradientNodes
        updatePreview
        
    'The left mouse button is not down
    Else
    
        'See if a new point is currently being hovered.
        Dim tmpPoint As Long
        tmpPoint = getPointAtPosition(x, y)
        
        'If a new point is being hovered, highlight it and redraw the interactive area
        If tmpPoint <> m_CurHoverPoint Then
            m_CurHoverPoint = tmpPoint
            drawGradientNodes
        End If
    
    End If
    
End Sub

'Given an x-position in the interaction box, return the currently hovered point.  If multiple points are hovered, the nearest one will be returned.
Private Function getPointAtPosition(ByVal x As Long, y As Long) As Long
    
    'Start by converting the current x-position into the range [0, 1]
    Dim convPoint As Single
    convPoint = convertPixelCoordsToNodeCoords(x)
    
    'convPoint now contains the position of the mouse on the range [0, 1].  Find the nearest point.
    Dim minDistance As Single, curDistance As Single, minIndex As Long
    minDistance = 1
    minIndex = -1
    
    Dim i As Long
    For i = 0 To m_NumOfGradientPoints - 1
        curDistance = Abs(m_GradientPoints(i).pdgp_Position - convPoint)
        If curDistance < minDistance Then
            minIndex = i
            minDistance = curDistance
        End If
    Next i
    
    'The nearest point (if any) will now be in minIndex.  If it falls below the valid threshold for clicks, accept it.
    If minDistance < (GRADIENT_NODE_WIDTH / 2) / CDbl(picPreview.ScaleWidth) Then
        getPointAtPosition = minIndex
    Else
        getPointAtPosition = -1
    End If
    
End Function

'Given an (x, y) position on the gradient interaction window, convert it to the [0, 1] range used by the gradient control.
Private Function convertPixelCoordsToNodeCoords(ByVal x As Long) As Single
    
    'Start by converting the current x-position into the range [0, 1]
    Dim uiMin As Single, uiMax As Single, uiRange As Single
    uiMin = picPreview.Left + 1
    uiMax = picPreview.Left + picPreview.ScaleWidth
    uiRange = uiMax - uiMin
    
    convertPixelCoordsToNodeCoords = (CSng(x) - uiMin) / uiRange
    
    If convertPixelCoordsToNodeCoords < 0 Then
        convertPixelCoordsToNodeCoords = 0
    ElseIf convertPixelCoordsToNodeCoords > 1 Then
        convertPixelCoordsToNodeCoords = 1
    End If
    
End Function

'Draw all interactive nodes
Private Sub drawGradientNodes()

    If g_IsProgramRunning Then
        
        'Each node is basically comprised of three parts:
        ' 1) An upward arrowhead pointing at the gradient's precise position
        ' 2) a colored block representing the gradient's pure color.  (Opacity is ignored for this UI element)
        ' 3) An outline encompassing (1) and (2), which is colored based on the node's hover state
        
        'To simplify things, we assemble generic paths for (1) and (2), then simply translate and draw them for each individual node.
        Dim baseArrow As pdGraphicsPath, baseBlock As pdGraphicsPath
        Set baseArrow = New pdGraphicsPath
        Set baseBlock = New pdGraphicsPath
        
        'The base arrow is centered at 0, for convenience when translating
        Dim triangleHalfWidth As Single, triangleHeight As Single
        triangleHalfWidth = (GRADIENT_NODE_WIDTH / 2)
        triangleHeight = (picInteract.ScaleHeight - GRADIENT_NODE_HEIGHT) - 1
        baseArrow.addTriangle -1 * triangleHalfWidth, triangleHeight, 0, 0, triangleHalfWidth, triangleHeight
        
        'Next up is the colored block, also centered horizontally around position 0
        baseBlock.addRectangle_Relative -1 * GRADIENT_NODE_WIDTH \ 2, triangleHeight, GRADIENT_NODE_WIDTH, GRADIENT_NODE_HEIGHT
        
        'We also want some duplicate nodes, to remove the need to reset our base node shapes between draws
        Dim tmpArrow As pdGraphicsPath, tmpBlock As pdGraphicsPath
        Set tmpArrow = New pdGraphicsPath
        Set tmpBlock = New pdGraphicsPath
        
        'Finally, some generic scale factors to simplify the process of positioning nodes (who store their positions on the range [0, 1])
        Dim hOffset As Single, hScaleFactor As Single
        hOffset = picPreview.Left + 1
        hScaleFactor = picPreview.ScaleWidth
        
        '...and pen/fill objects for the actual rendering
        Dim blockFill As pdGraphicsBrush
        Set blockFill = New pdGraphicsBrush
        blockFill.setBrushProperty pgbs_BrushMode, 0
        blockFill.setBrushProperty pgbs_PrimaryOpacity, 100
        
        'Prep the target interaction DIB
        If m_InteractiveDIB Is Nothing Then Set m_InteractiveDIB = New pdDIB
        
        If (m_InteractiveDIB.getDIBWidth <> Me.picInteract.ScaleWidth) Or (m_InteractiveDIB.getDIBHeight <> Me.picInteract.ScaleHeight) Then
            m_InteractiveDIB.createBlank Me.picInteract.ScaleWidth, Me.picInteract.ScaleHeight, 24, 0
        Else
            m_InteractiveDIB.resetDIB
        End If
        
        'Fill the interaction DIB with white
        GDI_Plus.GDIPlusFillDIBRect m_InteractiveDIB, 0, 0, m_InteractiveDIB.getDIBWidth, m_InteractiveDIB.getDIBHeight, vbWhite, 255
        
        'Now all we do is use those to draw all the nodes in turn
        Dim i As Long
        For i = 0 To m_NumOfGradientPoints - 1
            
            'Copy the base shapes
            tmpArrow.cloneExistingPath baseArrow
            tmpBlock.cloneExistingPath baseBlock
            
            'Translate them to this node's position
            tmpArrow.translatePath hOffset + m_GradientPoints(i).pdgp_Position * hScaleFactor, 0
            tmpBlock.translatePath hOffset + m_GradientPoints(i).pdgp_Position * hScaleFactor, 0
            
            'The node's colored block is rendered the same regardless of hover
            blockFill.setBrushProperty pgbs_PrimaryColor, m_GradientPoints(i).pdgp_RGB
            tmpBlock.fillPathToDIB_BareBrush blockFill.getBrushHandle, m_InteractiveDIB
            
            'All other renders vary by hover state
            If i = m_CurPoint Then
                tmpBlock.strokePathToDIB_BarePen activeOutlinePen.getPenHandle, m_InteractiveDIB
                tmpArrow.fillPathToDIB_BareBrush activeArrowFill.getBrushHandle, m_InteractiveDIB
                tmpArrow.strokePathToDIB_BarePen activeOutlinePen.getPenHandle, m_InteractiveDIB
            ElseIf i = m_CurHoverPoint Then
                tmpBlock.strokePathToDIB_BarePen activeOutlinePen.getPenHandle, m_InteractiveDIB
                tmpArrow.fillPathToDIB_BareBrush activeArrowFill.getBrushHandle, m_InteractiveDIB
                tmpArrow.strokePathToDIB_BarePen activeOutlinePen.getPenHandle, m_InteractiveDIB
            Else
                tmpBlock.strokePathToDIB_BarePen inactiveOutlinePen.getPenHandle, m_InteractiveDIB
                tmpArrow.fillPathToDIB_BareBrush inactiveArrowFill.getBrushHandle, m_InteractiveDIB
                tmpArrow.strokePathToDIB_BarePen inactiveOutlinePen.getPenHandle, m_InteractiveDIB
            End If
            
        Next i
        
        'Finally, flip the DIB to the screen
        m_InteractiveDIB.renderToPictureBox picInteract
        
    End If

End Sub
