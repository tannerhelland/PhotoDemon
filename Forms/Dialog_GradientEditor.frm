VERSION 5.00
Begin VB.Form dialog_GradientEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gradient editor"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   375
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   636
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   844
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Index           =   0
      Left            =   0
      ScaleHeight     =   385
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   841
      TabIndex        =   3
      Top             =   3000
      Width           =   12615
      Begin PhotoDemon.smartCheckBox chkDistributeEvenly 
         Height          =   330
         Left            =   360
         TabIndex        =   12
         Top             =   5280
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   582
         Caption         =   "automatically distribute nodes evenly"
         Value           =   0
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
         TabIndex        =   11
         Top             =   855
         Width           =   12615
      End
      Begin VB.PictureBox picNodePreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   807
         TabIndex        =   10
         Top             =   360
         Width           =   12135
      End
      Begin PhotoDemon.buttonStrip btsShape 
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   4140
         Width           =   7740
         _ExtentX        =   9975
         _ExtentY        =   873
      End
      Begin PhotoDemon.sliderTextCombo sltNodeOpacity 
         Height          =   720
         Left            =   4320
         TabIndex        =   6
         Top             =   2220
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   1270
         Caption         =   "opacity"
         Max             =   100
         Value           =   100
         NotchPosition   =   2
         NotchValueCustom=   100
      End
      Begin PhotoDemon.colorSelector csNode 
         Height          =   855
         Left            =   360
         TabIndex        =   5
         Top             =   2220
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   1508
         Caption         =   "color"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   0
         Left            =   120
         Top             =   1800
         Width           =   12135
         _ExtentX        =   16536
         _ExtentY        =   556
         Caption         =   "current node settings"
         FontSize        =   12
      End
      Begin PhotoDemon.sliderTextCombo sltNodePosition 
         Height          =   720
         Left            =   8280
         TabIndex        =   7
         Top             =   2220
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   1270
         Caption         =   "position"
         Max             =   100
         SigDigits       =   2
         SliderTrackStyle=   1
         Value           =   50
         NotchPosition   =   1
         NotchValueCustom=   50
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   4
         Left            =   120
         Top             =   3360
         Width           =   12135
         _ExtentX        =   16536
         _ExtentY        =   556
         Caption         =   "full gradient settings"
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   5
         Left            =   360
         Top             =   3780
         Width           =   7740
         _ExtentX        =   16536
         _ExtentY        =   556
         Caption         =   "shape"
         FontSize        =   12
      End
      Begin PhotoDemon.sliderTextCombo sltAngle 
         Height          =   720
         Left            =   8280
         TabIndex        =   9
         Top             =   3780
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   1270
         Caption         =   "angle"
         Max             =   360
         SigDigits       =   1
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   2
         Left            =   240
         Top             =   0
         Width           =   9255
         _ExtentX        =   16536
         _ExtentY        =   556
         Caption         =   "node editor"
         FontSize        =   12
      End
      Begin PhotoDemon.pdLabel lblInstructions 
         Height          =   285
         Left            =   0
         Top             =   1410
         Width           =   12660
         _ExtentX        =   22331
         _ExtentY        =   503
         Alignment       =   2
         Caption         =   "yes"
         FontSize        =   9
         Layout          =   1
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   315
         Index           =   6
         Left            =   120
         Top             =   4800
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   556
         Caption         =   "additional tools"
         FontSize        =   12
      End
   End
   Begin PhotoDemon.buttonStrip btsEdit 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   873
      FontSize        =   12
   End
   Begin VB.PictureBox picPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   240
      ScaleHeight     =   95
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
      Top             =   8790
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   1323
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
      Caption         =   "preview"
      FontSize        =   12
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Index           =   1
      Left            =   0
      ScaleHeight     =   385
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   841
      TabIndex        =   4
      Top             =   3000
      Width           =   12615
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   315
      Index           =   3
      Left            =   240
      Top             =   2040
      Width           =   9255
      _ExtentX        =   16536
      _ExtentY        =   556
      Caption         =   "edit mode"
      FontSize        =   12
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

'Gradient strings are generated with the help of PD's core gradient class.  This class also renders two previews of the current gradient:
' 1) The main preview at the top of the page, which reflects all active settings
' 2) The node editor preview just below, which reflects only the current colors (and allows interactive editing)
Private m_GradientPreview As pdGradient, m_NodePreview As pdGradient

'If a user control spawned this dialog, it will pass itself as a reference.  We can then send gradient updates back
' to the control, allowing for real-time updates on the screen despite a modal dialog being raised!
Private parentGradientControl As gradientSelector

'Recently used gradients are loaded to/saved from a custom XML file
Private m_XMLEngine As pdXML

'The file where we'll store recent gradient data when the program is closed.  (At present, this file is located in PD's
' /Data/Presets/ folder.
Private m_XMLFilename As String

'Gradient preview DIB (required for color management) and interaction DIB (where all the gradient nodes are rendered)
Private m_MainPreviewDIB As pdDIB, m_NodePreviewDIB As pdDIB, m_InteractiveDIB As pdDIB

'To prevent recursive setting changes, this value can be set to TRUE to prevent automatic UI synchronization
Private m_SuspendUI As Boolean

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
    
    'Mirror the gradient settings across the node-editor gradient object as well
    Set m_NodePreview = New pdGradient
    m_NodePreview.createGradientFromString m_GradientPreview.getGradientAsString
    
    'TODO: force the node preview to be linear-type, angle 0
    
    'If the dialog is being initialized for the first time, there will be no "initial gradient".  In this case, the gradient class
    ' will initialize a placeholder gradient.  We make a copy of it, and use that as the basis of the editor's initial settings.
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

Private Sub btsShape_Click(ByVal buttonIndex As Long)
    
    'Show/hide the angle slider depending on the current shape
    If (buttonIndex = 0) Or (buttonIndex = 1) Then
        sltAngle.Visible = True
    Else
        sltAngle.Visible = False
    End If
    
    'Redraw the UI accordingly
    If (Not m_SuspendUI) Then redrawEverything
    
End Sub

Private Sub chkDistributeEvenly_Click()
    If (Not m_SuspendUI) Then redrawEverything
End Sub

Private Sub cmdBar_AddCustomPresetData()

    'This control (obviously) requires a lot of extra custom preset data.
    '
    'However, there's no reason to require horrible duplication code, when the gradient class is already capable of serializing
    ' all relevant data for this control!
    cmdBar.addPresetData "FullGradientDefinition", m_GradientPreview.getGradientAsString

End Sub

'CANCEL BUTTON
Private Sub cmdBar_CancelClick()
    userAnswer = vbCancel
    Me.Hide
End Sub

'OK BUTTON
Private Sub cmdBar_OKClick()

    'Store the newGradient value (which the dialog handler will use to return the selected gradient to the caller)
    updateGradientObjects
    
    'TODO: save the current list of recently used gradients
    'saveRecentGradientList
    
    userAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_ReadCustomPresetData()
    
    'This control (obviously) requires a lot of extra custom preset data.
    '
    'However, there's no reason to require horrible duplication code, when the gradient class is already capable of serializing
    ' all relevant data for this control!
    m_GradientPreview.createGradientFromString cmdBar.retrievePresetData("FullGradientDefinition")
    
    'Synchronize all controls to the updated settings
    syncControlsToGradientObject
        
End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    updatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    'Reset our generic outline object
    Set m_GradientPreview = New pdGradient
    m_GradientPreview.createGradientFromString ""
    
    'Synchronize all controls to the updated settings
    updateGradientObjects
    syncControlsToGradientObject
    updatePreview
    
End Sub

Private Sub csNode_ColorChanged()
    
    If (Not m_SuspendUI) And (m_CurPoint >= 0) Then
        m_GradientPoints(m_CurPoint).pdgp_RGB = csNode.Color
        redrawEverything
    End If
    
End Sub

Private Sub Form_Load()
    
    m_SuspendUI = True
    
    'Add the instructions label
    Dim instructionText As String
    instructionText = g_Language.TranslateMessage("Left-click to add new nodes or edit existing nodes.  Right-click a node to remove it.")
    lblInstructions.Caption = instructionText
    
    'Populate button strips and drop-downs
    btsEdit.AddItem "manual", 0
    btsEdit.AddItem "automatic", 1
    btsEdit_Click 0
    
    btsShape.AddItem "line", 0
    btsShape.AddItem "reflection", 1
    btsShape.AddItem "circle", 2
    btsShape.AddItem "rectangle", 3
    btsShape.AddItem "diamond", 4
    btsShape_Click 0
    
    If g_IsProgramRunning Then
    
        If m_GradientPreview Is Nothing Then Set m_GradientPreview = New pdGradient
        If m_MainPreviewDIB Is Nothing Then Set m_MainPreviewDIB = New pdDIB
        
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
                
        'Draw the initial set of interactive gradient nodes
        syncUIToActiveNode
        drawGradientNodes
                
    End If
    
    m_SuspendUI = False
    
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
    drawGradientNodes
    updatePreview
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Update our two internal gradient classes against any/all changed settings.
' (Note that the node-editor class only reflects the current collection of colors and positions, not things like angle or gradient type,
'  so we only sync it against the node collection.)
Private Sub updateGradientObjects()
    
    'If the "evenly distribute nodes" option is checked, assign positions automatically.
    If CBool(chkDistributeEvenly.Value) Then
        
        'Start by sorting nodes from least-to-greatest.  This has the unintended side-effect of changing the active node, unfortunately,
        ' so we must also reset the active node (if any).
        
        'Start by seeing if nodes require sorting.
        Dim i As Long
        
        Dim sortNeeded As Boolean
        sortNeeded = False
        
        For i = 1 To m_NumOfGradientPoints - 1
            If m_GradientPoints(i).pdgp_Position < m_GradientPoints(i - 1).pdgp_Position Then
                sortNeeded = True
                Exit For
            End If
        Next i
        
        'If a sort is required, perform it now
        If sortNeeded Then
            
            m_CurPoint = -1
            m_CurHoverPoint = -1
            
            syncUIToActiveNode
            
            m_GradientPreview.createGradientFromPointCollection m_NumOfGradientPoints, m_GradientPoints
            m_GradientPreview.getCopyOfPointCollection m_NumOfGradientPoints, m_GradientPoints
            
        End If
        
        'Redistribute points accordingly
        For i = 0 To m_NumOfGradientPoints - 1
            m_GradientPoints(i).pdgp_Position = i / (m_NumOfGradientPoints - 1)
        Next i
        
    End If
    
    With m_GradientPreview
        .setGradientProperty pdgs_GradientShape, btsShape.ListIndex
        .setGradientProperty pdgs_GradientAngle, sltAngle.Value
        .createGradientFromPointCollection m_NumOfGradientPoints, m_GradientPoints
    End With

    With m_NodePreview
        .setGradientProperty pdgs_GradientShape, pdgs_ShapeLinear
        .setGradientProperty pdgs_GradientAngle, 0#
        .createGradientFromPointCollection m_NumOfGradientPoints, m_GradientPoints
    End With

End Sub

Private Sub updatePreview()
    
    If Not m_SuspendUI Then
    
        'Make sure our gradient objects are up-to-date
        updateGradientObjects
        
        'Retrieve a matching brush handle for the primary preview
        Dim gdipBrushMain As Long, boundsRect As RECTF
        
        With boundsRect
            .Left = 0
            .Top = 0
            .Width = picPreview.ScaleWidth
            .Height = picPreview.ScaleHeight
        End With
        
        gdipBrushMain = m_GradientPreview.getBrushHandle(boundsRect)
        
        'Prep the preview DIB
        If m_MainPreviewDIB Is Nothing Then Set m_MainPreviewDIB = New pdDIB
        
        If (m_MainPreviewDIB.getDIBWidth <> Me.picPreview.ScaleWidth) Or (m_MainPreviewDIB.getDIBHeight <> Me.picPreview.ScaleHeight) Then
            m_MainPreviewDIB.createBlank Me.picPreview.ScaleWidth, Me.picPreview.ScaleHeight, 24, 0
        Else
            m_MainPreviewDIB.resetDIB
        End If
        
        'Create the preview image
        With m_MainPreviewDIB
            GDI_Plus.GDIPlusFillDIBRect_Pattern m_MainPreviewDIB, 0, 0, .getDIBWidth, .getDIBHeight, g_CheckerboardPattern
            GDI_Plus.GDIPlusFillDC_Brush .getDIBDC, gdipBrushMain, 0, 0, .getDIBWidth, .getDIBHeight
        End With
        
        'Copy the preview image to the screen
        m_MainPreviewDIB.renderToPictureBox Me.picPreview
        
        'Next, repeat all the above steps for the node-area preview
        Dim gdipBrushNodes As Long
        
        With boundsRect
            .Left = 0
            .Top = 0
            .Width = picNodePreview.ScaleWidth
            .Height = picNodePreview.ScaleHeight
        End With
        
        gdipBrushNodes = m_GradientPreview.getBrushHandle(boundsRect, True)
        
        If m_NodePreviewDIB Is Nothing Then Set m_NodePreviewDIB = New pdDIB
        
        If (m_NodePreviewDIB.getDIBWidth <> Me.picNodePreview.ScaleWidth) Or (m_NodePreviewDIB.getDIBHeight <> Me.picNodePreview.ScaleHeight) Then
            m_NodePreviewDIB.createBlank Me.picNodePreview.ScaleWidth, Me.picNodePreview.ScaleHeight, 24, 0
        Else
            m_NodePreviewDIB.resetDIB
        End If
        
        With m_NodePreviewDIB
            GDI_Plus.GDIPlusFillDIBRect_Pattern m_NodePreviewDIB, 0, 0, .getDIBWidth, .getDIBHeight, g_CheckerboardPattern
            GDI_Plus.GDIPlusFillDC_Brush .getDIBDC, gdipBrushNodes, 0, 0, .getDIBWidth, .getDIBHeight
        End With
        
        m_NodePreviewDIB.renderToPictureBox Me.picNodePreview
        
        'Release our GDI+ handles
        GDI_Plus.releaseGDIPlusBrush gdipBrushMain
        GDI_Plus.releaseGDIPlusBrush gdipBrushNodes
                
        'Notify our parent of the update
        If Not (parentGradientControl Is Nothing) Then parentGradientControl.notifyOfLiveGradientChange m_GradientPreview.getGradientAsString
        
    End If
    
End Sub

'Make all UI elements reflect the current gradient object.  This is typically done after the dialog loads, or after loading a
' previously created gradient.
Private Sub syncControlsToGradientObject()
        
    m_SuspendUI = True
    
    With m_GradientPreview
        btsShape.ListIndex = .getGradientProperty(pdgs_GradientShape)
        sltAngle.Value = .getGradientProperty(pdgs_GradientAngle)
        
        .getCopyOfPointCollection m_NumOfGradientPoints, m_GradientPoints
    End With
    
    drawGradientNodes
    
    m_SuspendUI = False
    
    'Also, synchronize the node-specific UI to the active node (if any)
    syncUIToActiveNode
    
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
            
            'Make sure the current point index is not invalid
            If m_CurPoint >= m_NumOfGradientPoints Then
                m_CurPoint = -1
                syncUIToActiveNode
            End If
            
        End If
        
    'If this is not an existing point, create a new one now.
    Else
    
        'Enlarge the target array as necessary
        If m_NumOfGradientPoints >= UBound(m_GradientPoints) Then ReDim Preserve m_GradientPoints(0 To m_NumOfGradientPoints * 2) As pdGradientPoint
        
        With m_GradientPoints(m_NumOfGradientPoints)
            .pdgp_Opacity = 1
            .pdgp_Position = convertPixelCoordsToNodeCoords(x)
            
            'Preset the RGB value to match whatever the gradient already is at this point
            Dim newRGBA As RGBQUAD
            m_GradientPreview.getColorAtPosition_RGBA .pdgp_Position, newRGBA
            .pdgp_RGB = RGB(newRGBA.Red, newRGBA.Green, newRGBA.Blue)
            
        End With
        
        m_CurPoint = m_NumOfGradientPoints
        m_NumOfGradientPoints = m_NumOfGradientPoints + 1
        
    End If
    
    'Regardless of outcome, we need to resync the UI to the active node, and redraw the interaction area and preview
    syncUIToActiveNode
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
        If m_CurPoint >= 0 Then m_GradientPoints(m_CurPoint).pdgp_Position = convertPixelCoordsToNodeCoords(x)
        
        'Redraw the gradient interaction nodes and the gradient itself
        syncUIToActiveNode
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

'When a new active node is selected (or its parameters somehow changed), call this sub to synchronize all UI elements to that node's properties.
Private Sub syncUIToActiveNode()
    
    If g_IsProgramRunning Then
    
        'Disable automatic UI synchronization
        m_SuspendUI = True
        
        If (m_CurPoint >= 0) And (m_CurPoint < m_NumOfGradientPoints) Then
            
            'Show all relevant controls
            If Not csNode.Visible Then
                lblTitle(0).Caption = g_Language.TranslateMessage("node settings:")
                csNode.Visible = True
                sltNodeOpacity.Visible = True
                sltNodePosition.Visible = True
            End If
            
            'Sync all UI elements to the current node's settings
            With m_GradientPoints(m_CurPoint)
                csNode.Color = .pdgp_RGB
                sltNodeOpacity.Value = .pdgp_Opacity * 100
                sltNodePosition.Value = .pdgp_Position * 100
            End With
        
        Else
        
            'Hide all relevant controls
            lblTitle(0).Caption = g_Language.TranslateMessage("please select a node")
            csNode.Visible = False
            sltNodeOpacity.Visible = False
            sltNodePosition.Visible = False
        
        End If
            
        m_SuspendUI = False
        
    End If

End Sub

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

'Some user interactions require us to redraw just about everything on the dialog.  Use this shortcut function to do so.
Private Sub redrawEverything()
    updateGradientObjects
    drawGradientNodes
    updatePreview
End Sub

Private Sub sltAngle_Change()
    If (Not m_SuspendUI) Then redrawEverything
End Sub

Private Sub sltNodeOpacity_Change()
    
    If (Not m_SuspendUI) And (m_CurPoint >= 0) Then
        m_GradientPoints(m_CurPoint).pdgp_Opacity = sltNodeOpacity.Value / 100
        redrawEverything
    End If
    
End Sub

Private Sub sltNodePosition_Change()
    
    If (Not m_SuspendUI) And (m_CurPoint >= 0) Then
        m_GradientPoints(m_CurPoint).pdgp_Position = sltNodePosition.Value / 100
        redrawEverything
    End If
    
End Sub
